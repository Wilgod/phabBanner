import { IList } from './IList';
import { IListColumn } from './IListColumn'
export interface IListService {
  getLists: () => Promise<IList[]>;
  getListLibrary: () => Promise<IList[]>;
  getColumns: (listName) => Promise<IListColumn[]>;
}