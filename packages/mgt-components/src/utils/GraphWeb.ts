/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
export class GraphWeb {
  private readonly worker: SharedWorker;

  constructor() {
    this.worker = new SharedWorker(
      /* webpackChunkName: "graph-web-worker" */ new URL('./graphWebWorker.js', import.meta.url)
    );
    this.worker.port.onmessage = this.onMessage;
  }

  public close() {
    this.worker.port.close();
  }

  private readonly onMessage = (e: MessageEvent) => {
    console.log('message from worker', e.data);
  };

  public getUsersForUserIds = (userIds: string[], filters: string) => {
    console.log('getUsersForUserIds', userIds);
    console.log('filters', filters);
  };
}
