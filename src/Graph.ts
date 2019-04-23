import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Client } from '@microsoft/microsoft-graph-client/lib/es/Client';
import { IProvider } from './providers/IProvider';
import { ResponseType } from '@microsoft/microsoft-graph-client/lib/es/ResponseType';
import { AuthenticationHandlerOptions } from '@microsoft/microsoft-graph-client/lib/es/middleware/options/AuthenticationHandlerOptions';

export function prepScopes(...scopes: string[]) {
  const authProviderOptions = {
    scopes: scopes
  };
  return [new AuthenticationHandlerOptions(undefined, authProviderOptions)];
}

export class Graph {
  public client: Client;

  constructor(provider: IProvider) {
    if (provider) {
      this.client = Client.initWithMiddleware({
        authProvider: provider
      });
    }
  }

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
    return this.client
      .api('me')
      .middlewareOptions(prepScopes('user.read'))
      .get();
  }

  async getUser(userPrincipleName: string): Promise<MicrosoftGraph.User> {
    let scopes = 'user.readbasic.all';
    return this.client
      .api(`/users/${userPrincipleName}`)
      .middlewareOptions(prepScopes(scopes))
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
    let scopes = 'user.read';
    let blob = await this.client
      .api('/me/photo/$value')
      .responseType(ResponseType.BLOB)
      .middlewareOptions(prepScopes(scopes))
      .get();
    return await this.blobToBase64(blob);
  }

  async getUserPhoto(id: string): Promise<string> {
    let scopes = 'user.readbasic.all';
    let blob = await this.client
      .api(`users/${id}/photo/$value`)
      .responseType(ResponseType.BLOB)
      .middlewareOptions(prepScopes(scopes))
      .get();
    return await this.blobToBase64(blob);
  }

  async calendar(startDateTime: Date, endDateTime: Date): Promise<Array<MicrosoftGraph.Event>> {
    let scopes = 'calendars.read';

    let sdt = `startdatetime=${startDateTime.toISOString()}`;
    let edt = `enddatetime=${endDateTime.toISOString()}`;
    let uri = `/me/calendarview?${sdt}&${edt}`;

    let calendarView = await this.client
      .api(uri)
      .middlewareOptions([{ scopes: scopes }])
      .middlewareOptions(prepScopes(scopes))
      .get();
    return calendarView ? calendarView.value : null;
  }
}
