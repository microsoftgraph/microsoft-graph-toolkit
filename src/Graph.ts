import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Client } from '@microsoft/microsoft-graph-client/lib/es/Client';
import { IProvider } from './providers/IProvider';
import { ResponseType } from '@microsoft/microsoft-graph-client/lib/es/ResponseType';

export interface IGraph {
  me(): Promise<MicrosoftGraph.User>;
  getUser(id: string): Promise<MicrosoftGraph.User>;
  findPerson(query: string): Promise<MicrosoftGraph.Person[]>;
  myPhoto(): Promise<string>;
  getUserPhoto(id: string): Promise<string>;
  calendar(startDateTime: Date, endDateTime: Date): Promise<Array<MicrosoftGraph.Event>>;
}

export class Graph implements IGraph {
  private client: Client;

  // private token: string;
  private _provider: IProvider;
  private rootUrl: string = 'https://graph.microsoft.com/beta';

  constructor(provider: IProvider) {
    // this.token = token;
    this._provider = provider;

    this.client = Client.initWithMiddleware({
      authProvider: provider
    });
  }

  // async getJson(resource: string, scopes? : string[]) {
  //     let response = await this.get(resource, scopes);
  //     if (response) {
  //         return response.json();
  //     }

  //     return null;
  // }

  // async get(resource: string, scopes?: string[]) : Promise<Response> {
  //     if (!resource.startsWith('/')){
  //         resource = "/" + resource;
  //     }

  //     let token : string;
  //     try {
  //         token = await this._provider.getAccessToken(...scopes);
  //     } catch (error) {
  //         console.log(error);
  //         return null;
  //     }

  //     if (!token) {
  //         return null;
  //     }

  //     let response = await fetch(this.rootUrl + resource, {
  //         headers: {
  //             authorization: 'Bearer ' + token
  //         }
  //     });

  //     if (response.status >= 400) {

  //         // hit limit - need to wait and retry per:
  //         // https://docs.microsoft.com/en-us/graph/throttling
  //         if (response.status == 429) {
  //             console.log('too many requests - wait ' + response.headers.get('Retry-After') + ' seconds');
  //             return null;
  //         }

  //         let error : any = response.json();
  //         if (error.error !== undefined) {
  //             console.log(error);
  //         }
  //         console.log(response);
  //         throw 'error accessing graph';
  //     }

  //     return response;
  // }

  // private async getBase64(resource: string, scopes: string[]) : Promise<string> {
  //     try {
  //         let response = await this.get(resource, scopes);
  //         if (!response) {
  //             return null;
  //         }

  //         let blob = await response.blob();

  //         return new Promise((resolve, reject) => {
  //             const reader = new FileReader;
  //             reader.onerror = reject;
  //             reader.onload = _ => {
  //                 resolve(reader.result as string);
  //             }
  //             reader.readAsDataURL(blob);
  //         });
  //     } catch {
  //         return null;
  //     }
  // }

  private blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = reject;
      reader.onload = _ => {
        resolve(reader.result as string);
      };
      reader.readAsDataURL(blob);
    });
  }

  async me(): Promise<MicrosoftGraph.User> {
    let scopes = ['user.read'];
    return this.client
      .api('me')
      .middlewareOptions([{ scopes: scopes }])
      .get();
  }

  async getUser(userPrincipleName: string): Promise<MicrosoftGraph.User> {
    let scopes = ['user.readbasic.all'];
    return this.client
      .api(`/users/${userPrincipleName}`)
      .middlewareOptions([{ scopes: scopes }])
      .get();
  }

  async findPerson(query: string): Promise<MicrosoftGraph.Person[]> {
    let scopes = ['people.read'];
    let result = await this.client
      .api(`/me/people`)
      .search('"' + query + '"')
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return result ? result.value : null;
  }

  async myPhoto(): Promise<string> {
    let scopes = ['user.read'];
    let blob = await this.client
      .api('/me/photo/$value')
      .responseType(ResponseType.BLOB)
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return await this.blobToBase64(blob);
  }

  async getUserPhoto(id: string): Promise<string> {
    let scopes = ['user.readbasic.all'];
    let blob = await this.client
      .api(`users/${id}/photo/$value`)
      .responseType(ResponseType.BLOB)
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return await this.blobToBase64(blob);
  }

  async calendar(startDateTime: Date, endDateTime: Date): Promise<Array<MicrosoftGraph.Event>> {
    let scopes = ['calendars.read'];

    let sdt = `startdatetime=${startDateTime.toISOString()}`;
    let edt = `enddatetime=${endDateTime.toISOString()}`;
    let uri = `/me/calendarview?${sdt}&${edt}`;

    let calendarView = await this.client
      .api(uri)
      .middlewareOptions([{ scopes: scopes }])
      .get();
    return calendarView ? calendarView.value : null;
  }
}
