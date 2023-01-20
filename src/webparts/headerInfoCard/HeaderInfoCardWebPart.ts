import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from "HeaderInfoCardWebPartStrings";
import HeaderInfoCard from "./components/HeaderInfoCard";
import { IHeaderInfoCardProps } from "./components/IHeaderInfoCardProps";

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from "@pnp/spfx-property-controls/lib/PropertyFieldColorPicker";
import { PropertyFieldIconPicker } from "@pnp/spfx-property-controls/lib/PropertyFieldIconPicker";

import { PropertyPaneAsyncDropdown } from "../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown";

import { IDropdownOption } from "office-ui-fabric-react/lib/components/Dropdown";
import { update, get } from "@microsoft/sp-lodash-subset";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { ISharePointList } from "./components/ISharePointList";
import { ISharePointLists } from "./components/ISharePointLists";
import { ISharePointListItems } from "./components/ISharePointListItems";

// import Roboto font
import "@fontsource/roboto/300.css";
import "@fontsource/roboto/400.css";
import "@fontsource/roboto/500.css";
import "@fontsource/roboto/700.css";

export interface IHeaderInfoCardWebPartProps {
  headerInfoCardTitle: string;
  headerInfoCardIconBackgroundColor: string;
  headerInfoCardIconDarkToggle: boolean;
  headerInfoCardIcon: string;
  listName: string;
  dataCount: number;
  dataFilter: string;
}

export default class HeaderInfoCardWebPart extends BaseClientSideWebPart<IHeaderInfoCardWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public render(): void {
    const element: React.ReactElement<IHeaderInfoCardProps> = React.createElement(HeaderInfoCard, {
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      headerInfoCardTitle: this.properties.headerInfoCardTitle,
      headerInfoCardIconBackgroundColor: this.properties.headerInfoCardIconBackgroundColor,
      headerInfoCardIconDarkToggle: this.properties.headerInfoCardIconDarkToggle,
      headerInfoCardIcon: this.properties.headerInfoCardIcon,
      listName: this.properties.listName,
      dataCount: this.properties.dataCount,
      dataFilter: this.properties.dataFilter,
    });

    ReactDom.render(element, this.domElement);
  }

  private getLists(): Promise<ISharePointLists> {
    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$Filter=Hidden eq false and BaseType eq 0`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private loadLists(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: unknown) => void) => {
      this.getLists().then(
        (response) => {
          const options: IDropdownOption[] = response.value.map((list: ISharePointList) => {
            return { key: list.Id, text: list.Title };
          });
          resolve(options);
        },
        (error: unknown): void => {
          reject(error);
        }
      );
    });
  }

  private async onListChange(propertyPath: string, newValue: unknown): Promise<void> {
    const oldValue: unknown = get(this.properties, propertyPath);
    // store new value in web part properties
    console.log(oldValue);
    update(this.properties, propertyPath, (): unknown => {
      return newValue;
    });

    this.properties.dataCount = 0;

    const response = await this.getDataCount(null);
    console.log(response.value.length);
    update(this.properties, "dataCount", (): unknown => {
      return response.value.length;
    });

    // refresh web part
    this.render();

    //this.properties.dataCount = response.value.length;
  }

  private getDataCount(value: string): Promise<ISharePointListItems> {
    const filterArray = value ? value.split(" ") : [];
    let filter = "?$filter=";

    if (filterArray.length >= 3) {
      filter += `${filterArray[0]} ${filterArray[1]} '${filterArray[2]}'`;
    } else {
      filter = "";
    }

    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/Lists(guid'${escape(this.properties.listName)}')/items${filter}`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private async validateDataFilter(value: string): Promise<string> {
    const response = await this.getDataCount(value);

    if (response.value.length === 0) {
      return "No data found";
    } 

    update(this.properties, "dataCount", (): unknown => {
      return response.value.length;
    });
    console.log(response);
    return "";
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = await this._getEnvironmentMessage();

    await super.onInit();

  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
        let environmentMessage: string = "";
        switch (context.app.host.name) {
          case "Office": // running in Office
            environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
            break;
          case "Outlook": // running in Outlook
            environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
            break;
          case "Teams": // running in Teams
            environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
            break;
          default:
            throw new Error("Unknown host");
        }

        return environmentMessage;
      });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty("--bodyText", semanticColors.bodyText || null);
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty("--linkHovered", semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("headerInfoCardTitle", {
                  label: strings.headerInfoCardTitleFieldLabel,
                }),
                PropertyPaneToggle("headerInfoCardIconDarkToggle", {
                  key: "headerInfoCardIconDarkToggleFieldId",
                  label: strings.headerInfoCardIconDarkToggleFieldLabel,
                  onText: strings.headerInfoCardIconDarkToggleOnText,
                  offText: strings.headerInfoCardIconDarkToggleOffText,
                }),
                PropertyFieldIconPicker("iconPicker", {
                  currentIcon: this.properties.headerInfoCardIcon,
                  key: "iconPickerId",
                  onSave: (icon: string) => {
                    this.properties.headerInfoCardIcon = icon;
                  },
                  buttonLabel: "Icon",
                  renderOption: "panel",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  label: strings.headerInfoCardIconFieldLabel,
                }),
                PropertyFieldColorPicker("headerInfoCardIconBackgroundColor", {
                  label: strings.headerInfoCardColorFieldLabel,
                  selectedColor: this.properties.headerInfoCardIconBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  showPreview: true,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: "Precipitation",
                  key: "headerInfoCardIconBackgroundColorFieldId",
                }),
              ],
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                new PropertyPaneAsyncDropdown("listName", {
                  label: strings.listFieldLabel,
                  loadOptions: this.loadLists.bind(this),
                  onPropertyChange: this.onListChange.bind(this),
                  selectedKey: this.properties.listName,
                }),
                PropertyPaneTextField("dataFilter", {
                  label: strings.dataFilterFieldLabel,
                  multiline: false,
                  placeholder: strings.dataFilterFieldPlaceholder,
                  description: strings.dataFilterFieldDescription,
                  onGetErrorMessage: this.validateDataFilter.bind(this),
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
