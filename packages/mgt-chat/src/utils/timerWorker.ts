export interface TimerWork {
  id: string;
  type: 'clearInterval' | 'setInterval' | 'runCallback' | 'setTimeout' | 'clearTimeout';
  delay?: number;
}

const ctx: SharedWorkerGlobalScope = self as unknown as SharedWorkerGlobalScope;

const intervals = new Map<string, ReturnType<typeof setInterval>>();

ctx.onconnect = (e: MessageEvent<unknown>) => {
  const port = e.ports[0];
  const handleMessage = (event: MessageEvent<TimerWork>) => {
    const data: TimerWork = event.data;
    const delay = data.delay;
    const jobId = data.id;
    switch (data.type) {
      case 'setInterval': {
        const interval = setInterval(() => {
          const message: TimerWork = { ...event.data, ...{ type: 'runCallback' } };
          port.postMessage(message);
        }, delay);
        intervals.set(jobId, interval);
        break;
      }
      case 'setTimeout': {
        const timeout = setTimeout(() => {
          const message: TimerWork = { ...event.data, ...{ type: 'runCallback' } };
          port.postMessage(message);
        }, delay);
        intervals.set(jobId, timeout);
        break;
      }
      case 'clearTimeout': {
        clearTimeout(intervals.get(jobId));
        intervals.delete(jobId);
        break;
      }
      case 'clearInterval':
        clearInterval(intervals.get(jobId));
        intervals.delete(jobId);
    }
  };
  port.onmessage = handleMessage;
};
