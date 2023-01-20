import { IDropdownOption } from "office-ui-fabric-react/lib/components/Dropdown";

export interface IPropertyPaneAsyncDropdownProps {
  label: string;
  loadOptions: () => Promise<IDropdownOption[]>;
  onPropertyChange: (propertyPath: string, newValue: unknown) => void;
  selectedKey: string | number;
  disabled?: boolean;
}
