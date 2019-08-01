import { IExampleItem } from "./IExampleItem";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface IDetailsListCustomColumnsExampleState {
  pageItems: IExampleItem[];
  sortedItems: IExampleItem[];
  allItems: IExampleItem[];
  columns: IColumn[];
  pageNumber: number;
  pageCount: number;
  loaderMessage: string,
  isLoading: boolean,
  filterDepartment:string,
  filterName:string
}
