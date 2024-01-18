import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGroupByField } from "../models/IGroupByField";
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls";

export interface IListMultilevelGroupViewProps {
  listTitle: string;
  showFilter: boolean;
  lists: any;
  listColumns: any[];
  orderedListColumns: any[];
  groupByFields: IGroupByField[];
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  startDateTime: IDateTimeFieldValue;
  endDateTime: IDateTimeFieldValue;
}
