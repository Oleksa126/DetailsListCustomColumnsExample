import * as React from 'react';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { DetailsList, buildColumns, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyles, FontSizes } from 'office-ui-fabric-react/lib/Styling';
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
import * as ReactDOM from 'react-dom';


export default class DetailsListCustomColumnsExample extends React.Component<IDetailsListCustomColumnsExampleProps, IDetailsListCustomColumnsExampleState> {


  constructor(props: IDetailsListCustomColumnsExampleProps, state: IDetailsListCustomColumnsExampleState) {
    super(props);
    this.setDefultUsers();
    this.searchWithGraph();
  }
  private setDefultUsers(): void {
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
  private searchWithGraph() { 
    console.log("searchWithGraph")
    var users: Array<IExampleItem> = new Array<IExampleItem>();
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client.api("users")
          .version("beta")
          .select("displayName,mail,userPrincipalName,businessPhones,id,department,accountEnabled")
          .top(999)
          .get((err, res) => {
              res.value.map((item: any) => {
                  console.log(item);
                    if(item.accountEnabled && item.department!=null && item.department!=""){
                      console.log(item)
                      users.push({
                        thumbnail: "",
                        Name: item.displayName,
                        Email: item.userPrincipalName,
                        WorkPhone: item.businessPhones.length > 0 ? item.businessPhones[0] : "",
                        Department: item.department,
                        thumbnailColor: _randColor()
                      });
                    }

                    // if (res.value.length == users.length) {
                      let u1 = [...users.sort(function (a, b) { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : a.Name.toLowerCase() < b.Name.toLowerCase() ? -1 : 0 })];
                      let u2 = [...users.sort(function (a, b) { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : a.Name.toLowerCase() < b.Name.toLowerCase() ? -1 : 0 })];
                      let u3 = [...users.sort(function (a, b) { return a.Name.toLowerCase() > b.Name.toLowerCase() ? 1 : a.Name.toLowerCase() < b.Name.toLowerCase() ? -1 : 0 })];

                      if(users.length>this.state.pageCount)
                      {
                        this.setState({
                          isLoading: false,
                          allItems: u1,
                          sortedItems: u2,
                          pageItems: u3.splice(0, this.state.pageCount),
                          columns: _buildColumns(u3),
                        });

                      }
                      else{
                        this.setState({
                          isLoading: false,
                          allItems: u1,
                          sortedItems: u2,
                          pageItems: u3,
                          columns: _buildColumns(u3),
                        });
                      }

                  });
                  this._loadUserPhotoWithGraph(users);

              });

          });
  }

  private _loadUserPhotoWithGraph(users: Array<IExampleItem>)
   { // Log the current operation

    users.map((item: any) => {
      this.props.context.msGraphClientFactory.getClient().then((client
        : MSGraphClient)
        : void => {
          client.api("users/" + item.Email + "/photo/$value").version("v1.0").responseType('blob').get((err, res, raw) => {
          if (res!=null) {
          var blobUrl = window.URL.createObjectURL(res);
          let a  = [...this.state.allItems];
          let s  = [...this.state.sortedItems];
          let p  = [...this.state.pageItems];

          a.filter(i=>i.Email.indexOf(item.Email)>-1)[0].thumbnail= blobUrl;

          try{
            s.filter(i=>i.Email.indexOf(item.Email)>-1)[0].thumbnail= blobUrl;
          }catch(error){}

          try{
            p.filter(i=>i.Email.indexOf(item.Email)>-1)[0].thumbnail= blobUrl;
          }catch(error){}
          this.setState({
            allItems: a,
            sortedItems: s,
            pageItems: p,
            isLoading:false
          });
          // ReactDOM.render(<Image src={blobUrl} style={{ borderRadius: 50 }} width={50} height={50} imageFit={ImageFit.cover} onClick={() => { window.open("https://outlook.office.com/owa/?path=/mail/action/compose&to=" + item.Email, '_blank'); }} />, document.getElementById(item.Email));
        }
      });
    }
      );
    });




  }

  public render() {
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
      </div>);
    }
    return (
      <Fabric>
        <h2>Kontakte</h2>
        <TextField placeholder="Suche nach Namen..." onChange={this._onChangeText} />
        <br />
        <TextField placeholder="Suche nach Filiale..." onChange={this._onChangeTextDepartment} />
        <br />
        {displayLoader}
        <DetailsList
          items={this.state.pageItems}
          setKey="set"
          columns={columns}
          onRenderItemColumn={_renderItemColumn}
          onColumnHeaderClick={this._onColumnClick}
          selectionMode={SelectionMode.none}
          onColumnHeaderContextMenu={this._onColumnHeaderContextMenu}
        />
        {pageDivs}
        <div className='ms-Grid-row'>
          <div style={nStyle} className='ms-Grid-col ms-u-lg12'>
            {pageButtons}
          </div>
        </div>
      </Fabric>
    );
  }

  private _onChangeText = ( ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text :string): void => {
    let pageCountDivisor: number = this.state.pageCount;
    let sorterDepartmet: string = this.state.filterDepartment;
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
  };

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

    var thumbnailColor = columns.filter(column => column.name === 'thumbnailColor')[0];
    var index = columns.indexOf(thumbnailColor);
    if (index > -1) {
      columns.splice(index, 1);
    }

    var emailColumn = columns.filter(column => column.name === 'Email')[0];
    emailColumn.name = "E-Mail";

    var workPhoneColumn = columns.filter(column => column.name === 'WorkPhone')[0];
    workPhoneColumn.name = "Telefon";

    var departmentColumn = columns.filter(column => column.name === 'Department')[0];
    departmentColumn.name = "Filiale";
  }
  return columns;
}

function _renderItemColumn(item: IExampleItem, index: number, column: IColumn) {


  let initials = {
    //"position": "relative",
    "top": "25px",
    "font-size": "27px",
    "line-height": "50px",
    "color": "#fff",
    "font-family": '"Courier New", monospace',
    "font-weight": "bold"
  };

  const fieldContent = item[column.fieldName as keyof IExampleItem] as string;
  switch (column.key) {
    case 'Name':
        return <p style={{ fontSize: 17, marginTop: '9px'}}>{fieldContent}</p>;
    case 'thumbnail':

      var myStr = item["Name"] as string;    
      var color = item["thumbnailColor"] as string;    

      let avatar_circle = {
        "width": "50px",
        "height": "50px",
        "background-color": color,
        "text-align": "center",
        "border-radius": "50%",
        "-webkit-border-radius": "50%",
        "-moz-border-radius": "50%"
      };
      var arr = myStr.split(' (')[0].split(' ');
      if(arr.length===3){
        arr.splice(1, 1)
        //myStr = arr[0]  + " " + arr[2];
      }
      var matches = "";
      for(var i=0; i<arr.length; i++){
        matches +=arr[i].charAt(0)
      }

      // var matches = myStr.match(/\b(\w)/g);
      var result_sub: any;
      
      if (fieldContent == "") {

        result_sub = <div id={item.Email} onClick={() => { window.open("https://outlook.office.com/owa/?path=/mail/action/compose&to=" + item.Email, '_blank'); }} style={avatar_circle}><span style={initials}>{matches}</span></div>;
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

function _randColor(): string {
  const array = ['red', 'blue', 'green'];

  const index = Math.floor(Math.random() * array.length);
  return array[index];
}