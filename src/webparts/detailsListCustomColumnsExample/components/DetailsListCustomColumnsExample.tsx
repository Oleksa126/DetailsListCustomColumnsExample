import * as React from 'react';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { DetailsList, buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { IDetailsListCustomColumnsExampleProps } from './IDetailsListCustomColumnsExampleProps';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { IExampleItem } from './IExampleItem';
import { IDetailsListCustomColumnsExampleState } from './IDetailsListCustomColumnsExampleState';
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import{UserAvatar} from "react-user-avatar";

export function createListItems() {
  var users: Array<IExampleItem> = new Array<IExampleItem>(); 
  var size = 150 + Math.round(Math.random() * 100);

  users.push({
    thumbnail: "//placehold.it/" + size + "x" + size,
    Name: "Misha Struk",
    WorkPhone:"0959544",
    Email:"test@mail.com",
    Department: "Help Desk"
},
{
  thumbnail: "//placehold.it/" + size + "x" + size,
  Name: "Misha Test",
  WorkPhone:"0959544",
  Email:"test@mail.com",
  Department: "Help Desk"
},
{
  thumbnail: "//placehold.it/" + size + "x" + size,
  Name: "new David Test",
  WorkPhone:"0959544",
  Email:"test@mail.com",
  Department: "Help Desk"
},
{
  thumbnail: "//placehold.it/" + size + "x" + size,
  Name: "new David Test",
  WorkPhone:"0959544",
  Email:"oln@sunpoint.onmicrosoft.com",
  Department: "Help Desk"
},
{
  thumbnail: "//placehold.it/" + size + "x" + size,
  Name: "Scherban Nikita",
  WorkPhone:"0959544",
  Email:"scherban@sunpoint.onmicrosoft.com",
  Department: "Help Desk"
},

);
  return users;
}

export default class DetailsListCustomColumnsExample extends React.Component<IDetailsListCustomColumnsExampleProps, IDetailsListCustomColumnsExampleState> {
  private _allItems:Array<IExampleItem>;
  constructor(props: IDetailsListCustomColumnsExampleProps, state:IDetailsListCustomColumnsExampleState) {
    super(props);

    // this._allItems = this._searchWithGraph();
    this._allItems = createListItems();
    this.state = {
      sortedItems: this._allItems,
      columns: _buildColumns(this._allItems),
    };
  }

  public render() {
    const { sortedItems, columns } = this.state;

    return (
      <Fabric>
         {/* <UserAvatar size="48" name="Will Binns-Smith" /> */}
      <TextField label="Filter by name:" onChange={this._onChangeText}/>
      <DetailsList
        items={sortedItems}
        setKey="set"
        columns={columns}
        onRenderItemColumn={_renderItemColumn}
        onColumnHeaderClick={this._onColumnClick}
        onItemInvoked={this._onItemInvoked}
        onColumnHeaderContextMenu={this._onColumnHeaderContextMenu}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
      />
       </Fabric>
    );
  }

  private _searchWithGraph() { // Log the current operation
    // Log the current operation
      console.log("Using _searchWithGraph() method");
      var users: Array<IExampleItem> = new Array<IExampleItem>();
  
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
          client
            .api("users")
            .version("v1.0")
            .select("displayName,mail,userPrincipalName,businessPhones,id")
            // .filter(`(givenName eq '${escape(this.state.searchFor)}') or (surname eq '${escape(this.state.searchFor)}') or (displayName eq '${escape(this.state.searchFor)}')`)
            .get((err, res) => {  
    
              if (err) {
                console.error(err);
                return;
              }
              res.value.map((item: any) => {
    
                this.props.context.msGraphClientFactory.getClient().then((client
                  : MSGraphClient)
                  : any => {
                  client.api("users/"+ item.id +"/photo/$value").version("v1.0").responseType('blob').get((err1, res1, raw) => {
              
                  console.log(res1);
              
                  if (err) {
                    console.error(err1);
                    return;
                  }
                  const blobUrl =res1!=null?window.URL.createObjectURL(res1):""; 
    
                  users.push( { 
                    thumbnail:blobUrl,
                    Name: item.displayName,
                    Email: item.userPrincipalName,
                    WorkPhone:item.businessPhones.length>0?item.businessPhones[0]:"",
                    Department:"",
                  });
                  console.log(users);

                  return users;


                });
              }
                );
    
              });
    
            });
        });
        this.setState(
          {
            sortedItems: users,
          }
        );
  }

  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this._searchWithGraph();
    this.setState({
      sortedItems: text ? this._allItems.filter(i => i.Name.toLowerCase().indexOf(text) > -1) : this._allItems
    });
  }

  private _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns } = this.state;
    let { sortedItems } = this.state;
    let isSortedDescending = column.isSortedDescending;

    // If we've sorted this column, flip it.
    if (column.isSorted) {
      isSortedDescending = !isSortedDescending;
    }

    // Sort the items.
    sortedItems = _copyAndSort(sortedItems, column.fieldName!, isSortedDescending);

    // Reset the items and columns to match the state.
    this.setState({
      sortedItems: sortedItems,
      columns: columns.map(col => {
        col.isSorted = col.key === column.key;

        if (col.isSorted) {
          col.isSortedDescending = isSortedDescending;
        }

        return col;
      })
    });
  }

  private _onColumnHeaderContextMenu(column: IColumn | undefined, ev: React.MouseEvent<HTMLElement> | undefined): void {
    console.log(`column ${column!.key} contextmenu opened.`);
  }

  private _onItemInvoked(item: any, index: number | undefined): void {
    alert(`Item ${item.name} at index ${index} has been invoked.`);
  }
}

function _buildColumns(items: IExampleItem[]): IColumn[] {
  const columns = buildColumns(items);

  const thumbnailColumn = columns.filter(column => column.name === 'thumbnail')[0];

  // Special case one column's definition.
  thumbnailColumn.name = '';
  thumbnailColumn.maxWidth = 50;

  return columns;
}

function _renderItemColumn(item: IExampleItem, index: number, column: IColumn) {
  const fieldContent = item[column.fieldName as keyof IExampleItem] as string;

  switch (column.key) {
    case 'thumbnail':
      return <Image src={fieldContent} width={50} height={50} imageFit={ImageFit.cover} />;

    case 'Email':
      return <Link href="#">{fieldContent}</Link>;

    case 'WorkPhone':
      return (
        <span data-selection-disabled={true} className={mergeStyles({ color: fieldContent, height: '100%', display: 'block' })}>
          {fieldContent}
        </span>
      );

    default:
      return <span>{fieldContent}</span>;
  }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}