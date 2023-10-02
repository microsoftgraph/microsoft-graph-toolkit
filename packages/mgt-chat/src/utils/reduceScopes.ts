/**
 * A function to calculate minimal permissions to request for a given set of operations.
 * In some circumstances this may result in requesting more scopes that are necessary
 * but should not create requests for higher permissions than the true "minimum" set.
 *
 * @param scopeSet Array<Array<string>> where each top level array represents a set of
 * permissions, ordered least to most privileged, for a given operation
 * @returns string[], the minimal set of permissions to cover all operations
 */
export const reduceScopes = (scopeSet: string[][]): string[] => {
  const valueFrequency: Record<string, number> = {};
  scopeSet.forEach(set => {
    const dupeTrack = new Set<string>();
    for (const scope of set) {
      if (dupeTrack.has(scope)) continue;
      dupeTrack.add(scope);
      if (valueFrequency[scope]) {
        valueFrequency[scope]++;
      } else {
        valueFrequency[scope] = 1;
      }
    }
  });
  const frequencyValues: Record<number, string[]> = {};
  Object.keys(valueFrequency).forEach(k => {
    const count = valueFrequency[k];
    if (frequencyValues[count]) {
      frequencyValues[count].push(k);
    } else {
      frequencyValues[count] = [k];
    }
  });
  const result: string[] = [];
  let testSet = JSON.parse(JSON.stringify(scopeSet)) as string[][];
  for (let i = scopeSet.length; i >= 0; i--) {
    if (frequencyValues[i]) {
      const scopes = frequencyValues[i];
      for (const s of scopes) {
        const preFilterCount = testSet.length;
        testSet = testSet.filter(t => !t.includes(s));
        if (preFilterCount !== testSet.length) result.push(s);
        if (testSet.length === 0) return result;
      }
    }
  }
  return result;
};
