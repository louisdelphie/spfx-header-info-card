import * as React from "react";
import styles from "./HeaderInfoCard.module.scss";
import { IHeaderInfoCardProps } from "./IHeaderInfoCardProps";
//import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from "office-ui-fabric-react/lib/Icon";

import { Card, CardContent, Typography, Avatar, colors, List, ListItem, ListItemText } from "@mui/material";
//import * as muIcon from "@mui/icons-material";

export default class HeaderInfoCard extends React.Component<IHeaderInfoCardProps, {}> {
  public render(): React.ReactElement<IHeaderInfoCardProps> {
    const { headerInfoCardTitle, headerInfoCardIconBackgroundColor, headerInfoCardIconDarkToggle, headerInfoCardIcon, dataCount } = this.props;

    // eslint-disable-next-line dot-notation
    //const SelectedIcon = muIcon["AccessAlarm"];
    const iconColor: string = headerInfoCardIconDarkToggle ? colors.grey[900] : colors.grey[50];

    return (
      <Card sx={{ maxWidth: 260, marginX: "auto", boxShadow: 6 }} raised>
        <Avatar sx={{ bgcolor: headerInfoCardIconBackgroundColor, width: 62, height: 62 }} className={styles.cardIcon}>
          <Icon iconName={headerInfoCardIcon} style={{ color: iconColor, fontSize: 42, marginTop: "2px", marginLeft: "4px" }} />
        </Avatar>
        <CardContent sx={{ marginX: 3, marginY: 0, paddingY: 0 }}>
          <List>
            <ListItem>
              <ListItemText
                primary={
                  <React.Fragment>
                    <Typography variant="h6">{headerInfoCardTitle.toUpperCase()}</Typography>
                  </React.Fragment>
                }
                secondary={
                  <React.Fragment>
                    <Typography variant="h5">{dataCount}</Typography>
                  </React.Fragment>
                }
              />
            </ListItem>
          </List>
        </CardContent>
      </Card>
    );
  }
}
