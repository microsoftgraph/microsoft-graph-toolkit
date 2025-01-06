/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
const ctx: SharedWorkerGlobalScope = self as unknown as SharedWorkerGlobalScope;

ctx.onconnect = (e: MessageEvent) => {
  const port = e.ports[0];
  const handleMessage = (event: MessageEvent) => {
    const data = event.data;
    console.log('worker received message', data);
  };
  port.onmessage = handleMessage;
};
