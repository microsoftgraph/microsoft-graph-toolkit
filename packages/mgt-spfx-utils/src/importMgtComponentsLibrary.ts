export function importMgtComponentsLibrary(
  hasImportedMgtScripts: boolean,
  onSuccessCallback: () => void,
  onErrorCallback?: (e?: Error) => void
) {
  if (!hasImportedMgtScripts) {
    import('@microsoft/mgt-components')
      .then(() => {
        hasImportedMgtScripts = true;
        onSuccessCallback();
      })
      .catch(e => {
        onErrorCallback?.(e);
      });
  }
}
