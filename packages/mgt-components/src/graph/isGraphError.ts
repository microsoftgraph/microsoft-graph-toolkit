import { GraphError } from '@microsoft/microsoft-graph-client';

export const isGraphError = (e: unknown): e is GraphError => {
  const test = e as GraphError;
  return (
    test.statusCode &&
    'code' in test &&
    'body' in test &&
    test.date &&
    'message' in test &&
    'name' in test &&
    'requestId' in test
  );
};
