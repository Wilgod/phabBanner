import { IListService} from '../services/IListService';
import { IList} from '../services/IList';
import { IListColumn} from '../services/IListColumn';
import {
  SPHttpClient,
  ISPHttpClientOptions,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  IWebPartContext
} from '@microsoft/sp-webpart-base';

export class ListService implements IListService {

    constructor(private context: IWebPartContext) {
    }

    public getDocumentLibrary(): Promise<IList[]> {
      const httpClientOptions : ISPHttpClientOptions = {};
  
      httpClientOptions.headers = {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
      };

 
      return new Promise<IList[]>((resolve: (results: IList[]) => void, reject: (error: unknown) => void): void => {
        
        this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=id,title,EntityTypeName&$filter=Hidden eq false and BaseTemplate eq 101 and Title ne 'Site Assets' and Title ne 'Style Library'`,
          SPHttpClient.configurations.v1,
          httpClientOptions
          )
          .then((response: SPHttpClientResponse): Promise<{ value: IList[] }> => {
            return response.json();
          })
          .then((lists: { value: IList[] }): void => {
            resolve(lists.value);
          }, (error: unknown): void => {
            reject(error);
          });
      });
  }

    public getLists(): Promise<IList[]> {
        const httpClientOptions : ISPHttpClientOptions = {};
    
        httpClientOptions.headers = {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
        };
    
        return new Promise<IList[]>((resolve: (results: IList[]) => void, reject: (error: unknown) => void): void => {
          this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=id,title&$filter=Hidden eq false and (BaseTemplate eq 100 or BaseTemplate eq 106)`,
            SPHttpClient.configurations.v1,
            httpClientOptions
            )
            .then((response: SPHttpClientResponse): Promise<{ value: IList[] }> => {
              return response.json();
            })
            .then((lists: { value: IList[] }): void => {
              resolve(lists.value);
            }, (error: unknown): void => {
              reject(error);
            });
        });
    }

    public getDiscussBoardLists(): Promise<IList[]> {
      const httpClientOptions : ISPHttpClientOptions = {};
  
      httpClientOptions.headers = {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
      };
  
      return new Promise<IList[]>((resolve: (results: IList[]) => void, reject: (error: unknown) => void): void => {
        this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=id,title&$filter=Hidden eq false and BaseTemplate eq 108`,
          SPHttpClient.configurations.v1,
          httpClientOptions
          )
          .then((response: SPHttpClientResponse): Promise<{ value: IList[] }> => {
            return response.json();
          })
          .then((lists: { value: IList[] }): void => {
            resolve(lists.value);
          }, (error: unknown): void => {
            reject(error);
          });
      });
    }

    public getListLibrary(): Promise<IList[]> {
      const httpClientOptions : ISPHttpClientOptions = {};
  
      httpClientOptions.headers = {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
      };
  
      return new Promise<IList[]>((resolve: (results: IList[]) => void, reject: (error: unknown) => void): void => {
        this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=id,title,InternalName&$filter=Hidden eq false and (BaseTemplate eq 100 or BaseTemplate eq 106 or BaseTemplate eq 101) and Title ne 'Site Assets' and Title ne 'Style Library'`,
          SPHttpClient.configurations.v1,
          httpClientOptions
          )
          .then((response: SPHttpClientResponse): Promise<{ value: IList[] }> => {
            return response.json();
          })
          .then((lists: { value: IList[] }): void => {
            resolve(lists.value);
          }, (error: unknown): void => {
            reject(error);
          });
      });
    }
    
    public getColumns(listName: string): Promise<IListColumn[]> {
      const httpClientOptions : ISPHttpClientOptions = {};
  
      httpClientOptions.headers = {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
      };
  
      return new Promise<IListColumn[]>((resolve: (results: IListColumn[]) => void, reject: (error: unknown) => void): void => {
        this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${listName}')/fields?$filter=TypeDisplayName ne 'Attachments' and Hidden eq false and ReadOnlyField eq false`,
          SPHttpClient.configurations.v1,
          httpClientOptions
          )
          .then((response: SPHttpClientResponse): Promise<{ value: IListColumn[] }> => {
            return response.json();
          })
          .then((listColumns: { value: IListColumn[] }): void => {
            resolve(listColumns.value);
          }, (error: unknown): void => {
            reject(error);
          });
      });
    }
      
}