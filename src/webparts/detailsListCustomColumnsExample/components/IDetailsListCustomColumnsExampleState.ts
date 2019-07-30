import { IExampleItem } from "./IExampleItem";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface IDetailsListCustomColumnsExampleState {
    sortedItems: IExampleItem[];
    columns: IColumn[];
  }