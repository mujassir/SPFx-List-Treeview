import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ListMultilevelGroupViewWebPartStrings';
import { IGroupByField } from './models/IGroupByField';
import ListMultilevelGroupView from './components/ListMultilevelGroupView';
import { IListMultilevelGroupViewProps } from './components/IListMultilevelGroupViewProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import {
  PropertyFieldColumnPicker,
  PropertyFieldColumnPickerOrderBy, IColumnReturnProperty, IPropertyFieldRenderOption
} from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import { PropertyFieldOrder } from '@pnp/spfx-property-controls/lib/PropertyFieldOrder';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/Callout';
import { PropertyFieldCheckboxWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldCheckboxWithCallout';
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls';
import { PropertyFieldDateTimePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';


export interface IListMultilevelGroupViewWebPartProps {
  listTitle: string;
  showFilter: boolean;
  lists: any;
  listColumns: any[];
  orderedListColumns: any[];
  groupByFields: IGroupByField[];
  startDateTime: IDateTimeFieldValue;
  endDateTime: IDateTimeFieldValue;
}

export default class ListMultilevelGroupViewWebPart extends BaseClientSideWebPart<IListMultilevelGroupViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  public async render(): Promise<void> {
    const element: React.ReactElement<IListMultilevelGroupViewProps> = React.createElement(
      ListMultilevelGroupView,
      {
        listTitle: this.properties.listTitle,
        showFilter: this.properties.showFilter,
        lists: this.properties.lists,
        listColumns: this.properties.listColumns,
        orderedListColumns: this.properties.orderedListColumns,
        groupByFields: this.properties.groupByFields,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        startDateTime: this.properties.startDateTime,
        endDateTime: this.properties.endDateTime,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    const message = await this._getEnvironmentMessage();
    this._environmentMessage = message;
    this.properties.orderedListColumns = this.properties.listColumns;
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldDateTimePicker('startDateTime', {
                  label: 'Start Date',
                  initialDate: this.properties.startDateTime,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  key: 'startDateTime',
                  deferredValidationTime: 0,
                  showLabels: false
                }),
                PropertyFieldDateTimePicker('endDateTime', {
                  label: 'End Date',
                  initialDate: this.properties.endDateTime,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  key: 'endDateTime',
                  deferredValidationTime: 0,
                  showLabels: false,
                }),
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  includeHidden: false,
                  selectedList: this.properties.lists,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  includeListTitleAndUrl: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldColumnPicker('listColumns', {
                  label: 'Select columns',
                  context: this.context,
                  selectedColumn: this.properties.listColumns,
                  listId: this.properties.lists ? this.properties.lists.id : null,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'multiColumnPickerFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty.Title,
                  multiSelect: true,
                  renderFieldAs: IPropertyFieldRenderOption["Multiselect Dropdown"]
                }),
                PropertyFieldCollectionData("groupByFields", {
                  key: "groupByFields",
                  label: "Group By Fields",
                  panelHeader: "Group By Field Collection",
                  manageBtnLabel: "Manage Group By Fields",
                  value: this.properties.groupByFields,
                  fields: [
                    {
                      id: "column",
                      title: "Column",
                      type: CustomCollectionFieldType.dropdown,
                      options: this.properties.listColumns ? this.properties.listColumns.map(p => { return { key: p, text: p } }) : [],
                      required: true
                    },
                    {
                      id: "sortOrder",
                      title: "Sort Order",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "ascending",
                          text: "Ascending"
                        },
                        {
                          key: "descending",
                          text: "Descending"
                        }
                      ],
                      required: true
                    },

                  ],
                  disabled: false
                }),
                PropertyFieldOrder("orderedListColumns", {
                  key: "orderedListColumns",
                  label: "Column Display Order",
                  items: this.properties.orderedListColumns,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged
                }),
                PropertyFieldCheckboxWithCallout('showFilter', {
                  calloutTrigger: CalloutTriggers.Click,
                  key: 'showFiltercheckboxWithCalloutFieldId',
                  calloutContent: React.createElement('p', {}, 'Check the checkbox to enable searching'),
                  calloutWidth: 200,
                  text: 'Enable Search',
                  checked: this.properties.showFilter
                }),
              ]
            }
          ]
        }
      ]
    };
  }

}
