import * as React from 'react';
import styles from './ListMultilevelGroupView.module.scss';
import type { IListMultilevelGroupViewProps } from './IListMultilevelGroupViewProps';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Oval } from 'react-loader-spinner' //https://www.npmjs.com/package/react-loader-spinner
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import Constants from '../common/constants';
import { repeat } from 'lodash';

export default class ListMultilevelGroupView extends React.Component<IListMultilevelGroupViewProps, {}> {
  private _sp: SPFI;

  public state = {
    isLoading: true,
    hasErrors: false,
    errors: null,
    items: [],
    viewFields: [],
    groupByFields: [],
  };

  public componentDidMount(): void {
    this._sp = spfi().using(SPFx(this.props.context));

    this.getListData();
  }

  private async getListData() {

    let listTitle = '';
    if (this.props.lists) listTitle = this.props.lists.title;
    if (!listTitle || listTitle.length == 0) return;

    const allFields = await this._sp.web.lists.getByTitle(listTitle).fields();
    const titleToInternalNameMap = new Map();

    allFields.forEach((field: { Title: any; InternalName: any; }) => {
      titleToInternalNameMap.set(field.Title, field.InternalName);
    });

    let complexFieldNamesArray: string[] = [];
    let internalNamesArray = this.props.listColumns.map(title => {
      return titleToInternalNameMap.get(title) || title; // Fallback to title if mapping not found
    });
    let viewFields: { name: any; displayName: any; isResizable: boolean; sorting: boolean; }[] = [];
    if (this.props.orderedListColumns) {

      const complexFields = allFields.filter(p => this.props.orderedListColumns.indexOf(p.Title) > -1 && p.FieldTypeKind == 20);
      complexFieldNamesArray = complexFields.map(p => p.InternalName);
      for (let index = 0; index < internalNamesArray.length; index++) {
        if (complexFieldNamesArray.indexOf(internalNamesArray[index]) > -1)
          internalNamesArray[index] = internalNamesArray[index] + "/Title";
      }

      viewFields = this.props.orderedListColumns.map(title => {
        return {
          name: titleToInternalNameMap.get(title) || title,
          displayName: title,
          isResizable: true,
          sorting: true,
          minWidth: 100,
          maxWidth: 100
        }
      });
    }
    let groupByFields: { name: string; }[] = [];
    if (this.props.groupByFields) {
      groupByFields = this.props.groupByFields.map(d => {
        return {
          name: titleToInternalNameMap.get(d.column) || d.column,
        }
      });
    }
    this.setState({ viewFields: viewFields });
    this.setState({ groupByFields: groupByFields });

    const dateColumnName = 'Date';
    if(typeof this.props.startDateTime?.value == 'string') {
      this.props.startDateTime.value = new Date(this.props.startDateTime?.value)
    }
    if(typeof this.props.endDateTime?.value == 'string') {
      this.props.endDateTime.value = new Date(this.props.endDateTime?.value)
    }
    const startDateValue = this.props.startDateTime?.value.toLocaleDateString();
    const endDateValue = this.props.endDateTime?.value.toLocaleDateString();
    const items = await this._sp.web.lists.getByTitle(listTitle).items
      .select(...internalNamesArray)
      .expand(...complexFieldNamesArray)
      .filter(`${dateColumnName} ge '${startDateValue}' and ${dateColumnName} le '${endDateValue}'`)
      .top(Constants.Defaults.MaxPageSize)();

    this.setState({ isLoading: false });

    this.setState(
      {
        items: items.map(p => {

          if (complexFieldNamesArray.length > 0) {
            complexFieldNamesArray.forEach(element => {
              p[element] = p[element].Title;
            });
          }
          return p;
        })
      });
  }

  public render(): React.ReactElement<IListMultilevelGroupViewProps> {

    let listTitle = '';
    if (this.props.lists) listTitle = this.props.lists.title;
    return listTitle && listTitle.length > 0 ? this.renderUI() : this.renderPlaceHolder();
  }

  public renderUI(): React.ReactElement<IListMultilevelGroupViewProps> {
    return (
      <div className={styles.welcome}>
        {this.state.isLoading ? this.renderLoader() : this.renderListView()}
      </div>
    );
  }

  public renderListView() {
    const items = this.state.items;
    const groupByFields = this.state.groupByFields.map((e: { name: string }) => e.name);
    if (items.length === 0) return ('Not Found')

    const groupTree = this.createTreeView(items, groupByFields)
    const tableView = this.renderTable(groupTree)
    return (tableView)
  }
  private renderTable(data: any[]): React.ReactNode {
    const viewFields = this.state.viewFields;

    return (
      <table className={styles.strippedTable}>
        <thead>
          <tr>
            {
              viewFields.map((field: any) => (
                <th key={field.name}>{field.displayName}</th>
              ))
            }
          </tr>
        </thead>
        <tbody>
          {this.renderTableRows(data, '')}
        </tbody>
      </table>
    );
  }

  private renderTableRows(data: any[], html: string, level: number = 0): React.ReactNode {
    return data.map((item: any, index: number) => {
      let levelClass
      switch (level) {
        case 0:
          levelClass = styles.listLevel0
          break;
        default:
          levelClass = styles.listLevel1
          break;

      }
      // Calculate the sum of debit and credit for each parent
      let totalAmount = 0;

      if (item?.children) {
        item.children.forEach((child: any) => {
          totalAmount += (child.Amount || 0) + (child.Credit || 0);
        });
        return (
          <>
            <tr>
              <td colSpan={100} className={`${levelClass} ${styles.parentRow}`}>
                {repeat('-- ', level)}
                {`${item.name || ''}`}
                {totalAmount > 0 ? ` (${totalAmount})` : ''}
              </td>
            </tr>
            {this.renderTableRows(item.children, html, level + 1)}
          </>
        );
      } else {
        const viewFields = this.state.viewFields;
        return (
          <tr key={index}>
            {
              viewFields.map((field: any) => (
                <td key={field.name}>
                  {
                    field.name === 'Date'
                      ? new Intl.DateTimeFormat('en-US').format(new Date(item[field.name]))
                      : item[field.name]
                  }
                </td>
              ))
            }
          </tr>
        );
      }
    });
  }

  private createTreeView(dataset: any, groupByColumns: string[]) {
    const root = { name: 'Root', children: [] };

    dataset.forEach((record: any) => {
      let currentNode: any = root;

      groupByColumns.forEach(column => {
        const key = record[column];
        let childNode: any = currentNode.children.filter((child: any) => child.name === key)[0]

        if (!childNode) {
          childNode = { name: key, children: [] };
          currentNode.children.push(childNode);
        }

        currentNode = childNode;
      });

      currentNode.children.push(record);
    });

    return root.children;
  }

  private renderLoader() {
    return (
      <Oval
        visible={true}
        height="50"
        width="50"
        secondaryColor="#4dabf5"
        color="#0078D3"
        ariaLabel="oval-loading"
        wrapperStyle={{ display: 'block' }}
      />
    );
  }

  private renderPlaceHolder(): React.ReactElement<IListMultilevelGroupViewProps> {
    return (
      <Placeholder iconName='Edit'
        iconText='Configure your web part'
        description='Please configure the web part.'
        buttonLabel='Configure'
        onConfigure={this._onConfigure} />
    );
  }

  private _onConfigure = () => {
    // Context of the web part
    this.props.context.propertyPane.open();
  }
}
