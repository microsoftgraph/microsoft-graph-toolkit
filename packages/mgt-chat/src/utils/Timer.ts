/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { TimerWork } from './timerWorker';

export interface Work {
  id: string;
  callback: () => void;
}

export class Timer {
  private readonly worker: SharedWorker;
  private readonly work = new Map<string, () => void>();
  constructor() {
    this.worker = new SharedWorker(/* webpackChunkName: "timer-worker" */ new URL('./timerWorker.js', import.meta.url));
    this.worker.port.onmessage = this.onMessage;
  }

  // NOTE: this.work was increasing in size every time there was a new timeout because id was uuid()
  // but clearTimeout was only being called on teardown.
  public setTimeout(id: string, callback: () => void, delay: number): string {
    const timeoutWork: TimerWork = {
      type: 'setTimeout',
      id,
      delay
    };

    this.work.set(timeoutWork.id, callback);

    this.worker.port.postMessage(timeoutWork);

    return timeoutWork.id;
  }

  public clearTimeout(id: string): void {
    if (this.work.has(id)) {
      const timeoutWork: TimerWork = {
        type: 'clearTimeout',
        id
      };
      this.worker.port.postMessage(timeoutWork);
      this.work.delete(id);
    }
  }

  /*
  // NOTE: setInterval is not used in the library, but it is here for reference
  public setInterval(callback: () => void, delay: number): string {
    const intervalWork: TimerWork = {
      type: 'setInterval',
      id: uuid(),
      delay
    };

    this.work.set(intervalWork.id, callback);

    this.worker.port.postMessage(intervalWork);

    return intervalWork.id;
  }

  public clearInterval(id: string): void {
    if (this.work.has(id)) {
      const intervalWork: TimerWork = {
        type: 'clearInterval',
        id
      };
      this.worker.port.postMessage(intervalWork);
      this.work.delete(id);
    }
  }
  */

  private readonly onMessage = (event: MessageEvent<TimerWork>): void => {
    const intervalWork = event.data;
    if (!intervalWork) {
      return;
    }

    switch (intervalWork.type) {
      case 'runCallback': {
        if (this.work.has(intervalWork.id)) {
          const work = this.work.get(intervalWork.id);
          work?.();
        }
        break;
      }
    }
  };

  public close() {
    this.worker.port.close();
  }
}
