import * as React from 'react';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { DetailsList, buildColumns, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { IDetailsListCustomColumnsExampleProps } from './IDetailsListCustomColumnsExampleProps';
import { PrimaryButton, Button } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { IExampleItem } from './IExampleItem';
import { IDetailsListCustomColumnsExampleState } from './IDetailsListCustomColumnsExampleState';
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
//import{UserAvatar} from "react-user-avatar";

import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';


export default class DetailsListCustomColumnsExample extends React.Component<IDetailsListCustomColumnsExampleProps, IDetailsListCustomColumnsExampleState> {


  constructor(props: IDetailsListCustomColumnsExampleProps, state: IDetailsListCustomColumnsExampleState) {
    super(props);
    this.setDefultUsers();
    this.searchWithGraph();
  }
  private setDefultUsers(): void {
    // console.log("Set default list items");
    var users: Array<IExampleItem> = new Array<IExampleItem>();

    this.state = {
      isLoading: true,
      loaderMessage: "",
      pageItems: users,
      sortedItems: users,
      allItems: users,
      columns: _buildColumns(users),
      pageNumber: 1,
      pageCount: 10,
      filterDepartment:"",
      filterName:"",
      isSortedDescending:false,
      columnKey:""
    };
  }
  private searchWithGraph() { // Log the current operation
    // Log the current operation
    // console.log("Using graph method");
    var users: Array<IExampleItem> = new Array<IExampleItem>();
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client.api("users")
          .version("v1.0")
          .select("displayName,mail,userPrincipalName,businessPhones,id,department")
          .top(999)
          .get((err, res) => {
              res.value.map((item: any) => {
                this.props.context.msGraphClientFactory.getClient().then((client1
                  : MSGraphClient)
                  : any => {
                    try{
                  client1.api("users/" + item.id + "/photo/$value").version("v1.0").responseType('blob').get((err1, res1, raw) => {
                    const blobUrl = res1 != null ? window.URL.createObjectURL(res1) : "";
                    users.push({
                      thumbnail: blobUrl,
                      Name: item.displayName,
                      Email: item.userPrincipalName,
                      WorkPhone: item.businessPhones.length > 0 ? item.businessPhones[0] : "",
                      Department: item.department,
                    });
                    if (res.value.length == users.length) {
                      let u1 = [...users.sort(function (a, b) { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : a.Name.toLowerCase() < b.Name.toLowerCase() ? -1 : 0 })];
                      let u2 = [...users.sort(function (a, b) { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : a.Name.toLowerCase() < b.Name.toLowerCase() ? -1 : 0 })];
                      let u3 = [...users.sort(function (a, b) { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : a.Name.toLowerCase() < b.Name.toLowerCase() ? -1 : 0 })];
  
                      this.setState({
                        isLoading: false,
                        allItems: u1,
                        sortedItems: u2,
                        pageItems: u3.splice(0, this.state.pageCount),
                        columns: _buildColumns(u3),
                      });
                    }
                  });
                  
                }
                catch(error){console.log(error)}

                }
                );
              });

          });
      });


  }
  public render() {
    // console.log("render")
    // console.log(this.state.sortedItems)
    const { sortedItems, columns} = this.state;
    const leftArrovStyle = {
      'background-image': 'url("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAAA0SURBVDhPYxgFhIGrq+t/EIZySQMwzWQZgKyZZAPQNZNsAAhQbAAIUGwACFBswCggBBgYAGTLKYMf9UF2AAAAAElFTkSuQmCC")',
      'background-repeat': 'no-repeat',
      'background-position': 'center',
      'border-style': 'solid',
      'border-width': '1px',
      'border-color': 'black',
      'min-width': '20px',
      'width': '20px',
      'height': '20px'
    };
    const rightArrovStyle = {
      'background-image': 'url("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAOCAYAAAAmL5yKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNv1OCegAAAAxSURBVDhPYxjGwNXV9T8IQ7mkA5gBZBuCbABZhqAbQLIhFGkGAYo0gwBFmkcUYGAAABqpKYOWrx73AAAAAElFTkSuQmCC")',
      'background-repeat': 'no-repeat',
      'background-position': 'center',
      'border-style': 'solid',
      'border-width': '1px',
      'border-color': 'black',
      'min-width': '20px',
      'width': '20px',
      'height': '20px'
    };
    const ms_grid_col = { 'width': '100%' };
    let displayLoader;
    let pageCountDivisor: number = this.state.pageCount;
    let pageCount: number;
    let pageButtons = [];
    let pageDivs = [];
    const nStyle = {
      display: 'none'
    };
    let _pagedButtonClick = (pageNumber: number, listData: any) => {
      let startIndex: number = (pageNumber - 1) * pageCountDivisor;
      let listItemsCollection = [...listData];
      this.setState({ pageItems: listItemsCollection.splice(startIndex, pageCountDivisor), pageNumber: pageNumber });
      // console.log("_pagedButtonClick");
      // console.log(listItemsCollection.splice(startIndex, pageCountDivisor));
      // console.log(pageNumber);

    };

    if (this.state.isLoading) {
      displayLoader = (<div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white`}>
        <div className='ms-Grid-col ms-u-lg12'>
          <Spinner size={SpinnerSize.large} label={this.state.loaderMessage} />
        </div>
      </div>);
    }
    else {
      displayLoader = (null);
    }

    if (this.state.sortedItems.length > 0) {
      pageCount = Math.ceil(this.state.sortedItems.length / pageCountDivisor);
    }
    if (pageCount > 1) {
      for (let i = 0; i < pageCount; i++) {
        pageButtons.push(<PrimaryButton onClick={() => { _pagedButtonClick(i + 1, this.state.sortedItems); }}> {i + 1} </PrimaryButton>);

      }

      pageDivs.push(<div className='ms-Grid-col ms-u-lg12'>
        {(this.state.pageNumber > 1) ? (<Button style={leftArrovStyle} onClick={() => { _pagedButtonClick((this.state.pageNumber - 1), this.state.sortedItems); }}></Button>) :
          ("")}
        <span> {(this.state.pageNumber - 1) * pageCountDivisor + 1} - {(this.state.pageNumber < pageCount) ? ((this.state.pageNumber) * pageCountDivisor) : (this.state.sortedItems.length)} </span>
        {(this.state.pageNumber < pageCount) ? (<Button style={rightArrovStyle} onClick={() => { _pagedButtonClick((this.state.pageNumber + 1), this.state.sortedItems); }}></Button>) :
          ("")}
      </div>)
    }
    return (
      <Fabric>
        <h2>Kontakte</h2>
        <TextField placeholder="Search by name..." onChange={this._onChangeText} />
        <br />
        <TextField placeholder="Search by department..." onChange={this._onChangeTextDepartment} />
        <br />
        {displayLoader}
        <DetailsList
          items={this.state.pageItems}
          setKey="set"
          columns={columns}
          onRenderItemColumn={_renderItemColumn}
          onColumnHeaderClick={this._onColumnClick}
          selectionMode={SelectionMode.none}
          //onItemInvoked={this._onItemInvoked}
          onColumnHeaderContextMenu={this._onColumnHeaderContextMenu}
        //ariaLabelForSelectionColumn="Toggle selection"
        //ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        />
        {pageDivs}
        <div className='ms-Grid-row'>
          <div style={nStyle} className='ms-Grid-col ms-u-lg12'>
            {pageButtons}
          </div>
        </div>
      </Fabric>
    );

  };

  private _onChangeText = ( ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text :string): void => {
    // console.log(text);
    let pageCountDivisor: number = this.state.pageCount;
    let sorterDepartmet: string = this.state.filterDepartment;
    // console.log(pageCountDivisor);
    let a  = [...this.state.allItems];
    let a2 = [...this.state.allItems];
    let a3=text ? (!(sorterDepartmet==null||sorterDepartmet=="")?
    a.filter(i => ((i.Name.toLowerCase().indexOf(text.toLowerCase()) > -1)&&((i.Department?i.Department:"").toLowerCase().indexOf(sorterDepartmet.toLowerCase())>-1))
    ):a.filter(i => ((i.Name.toLowerCase().indexOf(text.toLowerCase()) > -1)))) :
     (!(sorterDepartmet==null||sorterDepartmet=="")?(a.filter(i => (((i.Department?i.Department:"").toLowerCase().indexOf(sorterDepartmet.toLowerCase()) > -1)))):a);
     let a4=text ? (!(sorterDepartmet==null||sorterDepartmet=="")?
     a2.filter(i => ((i.Name.toLowerCase().indexOf(text.toLowerCase()) > -1)&&((i.Department?i.Department:"").toLowerCase().indexOf(sorterDepartmet.toLowerCase())>-1))
     ):a2.filter(i => ((i.Name.toLowerCase().indexOf(text.toLowerCase()) > -1)))) :
      (!(sorterDepartmet==null||sorterDepartmet=="")?(a2.filter(i => (((i.Department?i.Department:"").toLowerCase().indexOf(sorterDepartmet.toLowerCase()) > -1)))):a2);
    
      if(this.state.columnKey!=null && this.state.columnKey!=""){
      a3 = _copyAndSort(a3,this.state.columnKey,this.state.isSortedDescending);
      a4 = _copyAndSort(a4,this.state.columnKey,this.state.isSortedDescending);
    }

    this.setState({
      filterName:text,
      sortedItems: a3,
      pageItems: a4.splice(0, pageCountDivisor),
      pageNumber:1
    });
  }

  private _onChangeTextDepartment = ( ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text :string): void => {
    let pageCountDivisor: number = this.state.pageCount;
    let sorterName: string = this.state.filterName;

    let a = [...this.state.allItems];
    let a2 = [...this.state.allItems];
    let a3=text ? (!(sorterName==null||sorterName=="")?
    a.filter(i => (((i.Department?i.Department:"").toLowerCase().indexOf(text.toLowerCase()) > -1)&&(i.Name.toLowerCase().indexOf(sorterName.toLowerCase())>-1))
    ):a.filter(i => (((i.Department?i.Department:"").toLowerCase().indexOf(text.toLowerCase()) > -1)))) :
     (!(sorterName==null||sorterName=="")?(a.filter(i => ((i.Name.toLowerCase().indexOf(sorterName.toLowerCase()) > -1)))):a);
     let a4=text ? (!(sorterName==null||sorterName=="")?
     a2.filter(i => (((i.Department?i.Department:"").toLowerCase().indexOf(text.toLowerCase()) > -1)&&(i.Name.toLowerCase().indexOf(sorterName.toLowerCase())>-1))
     ):a2.filter(i => (((i.Department?i.Department:"").toLowerCase().indexOf(text.toLowerCase()) > -1)))) :
      (!(sorterName==null||sorterName=="")?(a2.filter(i => ((i.Name.toLowerCase().indexOf(sorterName.toLowerCase()) > -1)))):a2);
    
      if(this.state.columnKey!=null && this.state.columnKey!=""){
        a3 = _copyAndSort(a3,this.state.columnKey,this.state.isSortedDescending);
        a4 = _copyAndSort(a4,this.state.columnKey,this.state.isSortedDescending);
      }
    
      this.setState({
      filterDepartment:text,
      sortedItems: a3,
      pageItems: a4.splice(0, pageCountDivisor),
      pageNumber:1
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
    let s = [...sortedItems];
    let pageCountDivisor: number = this.state.pageCount;
    // Reset the items and columns to match the state.
    this.setState({
      columnKey:column.fieldName!,
      isSortedDescending:isSortedDescending,
      isLoading: false,
      sortedItems: sortedItems,
      columns: columns.map(col => {
        col.isSorted = col.key === column.key;

        if (col.isSorted) {
          col.isSortedDescending = isSortedDescending;
        }

        return col;
      }),
      pageNumber:1,
      pageItems: s.splice(0, pageCountDivisor)
    });
  }

  private _onColumnHeaderContextMenu(column: IColumn | undefined, ev: React.MouseEvent<HTMLElement> | undefined): void {
    // console.log(`column ${column!.key} contextmenu opened.`);
  }
}

function _buildColumns(items: IExampleItem[]): IColumn[] {

  var columns = buildColumns(items);
  if (items.length > 0) {
    var thumbnailColumn = columns.filter(column => column.name === 'thumbnail')[0];

    // Special case one column's definition.
    thumbnailColumn.name = '';
    thumbnailColumn.maxWidth = 50;
  }
  return columns;
}

function _renderItemColumn(item: IExampleItem, index: number, column: IColumn) {
  let avatar_circle = {
    "width": "50px",
    "height": "50px",
    "background-color": "#00ABA9",
    "text-align": "center",
    "border-radius": "50%",
    "-webkit-border-radius": "50%",
    "-moz-border-radius": "50%"
  };
  let initials = {
    //"position": "relative",
    "top": "25px",
    "font-size": "35px",
    "line-height": "50px",
    "color": "#fff",
    "font-family": '"Courier New", monospace',
    "font-weight": "bold"
  };
  const fieldContent = item[column.fieldName as keyof IExampleItem] as string;

  switch (column.key) {
    case 'thumbnail':
      var myStr = item["Name"] as string;
      var matches = myStr.match(/\b(\w)/g);
      var result_sub: any;
      if (fieldContent == "") {
        result_sub = <div onClick={() => { window.open("https://outlook.office.com/owa/?path=/mail/action/compose&to=" + item.Email, '_blank'); }} style={avatar_circle}><span style={initials}>{matches}</span></div>;
      }
      else {
        result_sub = <Image src={fieldContent} style={{ borderRadius: 50 }} width={50} height={50} imageFit={ImageFit.cover} onClick={() => { window.open("https://outlook.office.com/owa/?path=/mail/action/compose&to=" + item.Email, '_blank'); }} />;
      }
      return result_sub;

    case 'Email':
      return <div><Icon iconName="Mail" style={{ fontSize: 18, marginRight: '5px', float: "left" }} /><Link target="_blank" href={"https://outlook.office.com/owa/?path=/mail/action/compose&to=" + fieldContent}>{fieldContent}</Link></div>;

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
  return items.slice(0).sort((a: T, b: T) => 
    ((a[key]!=null && b[key]!=null) ?
    ((isSortedDescending ? 
      a[key].toString().toLowerCase() < b[key].toString().toLowerCase() : 
      a[key].toString().toLowerCase() > b[key].toString().toLowerCase()) ? 
      1 : -1):-1));
}
