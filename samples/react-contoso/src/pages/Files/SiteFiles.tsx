import * as React from 'react';
import { FileList, Picker, Providers } from '@microsoft/mgt-react';
import { makeStyles } from '@fluentui/react-components';

const useStyles = makeStyles({
  picker: {
    paddingBottom: '10px',
    display: 'block'
  }
});

export const SiteFiles: React.FunctionComponent = () => {
  const [selectedList, setSelectedList] = React.useState<any>(null);
  const [driveId, setDriveId] = React.useState<string>('');
  const [error, setError] = React.useState<string>('');
  const styles = useStyles();

  const onSelectionChanged = async (e: CustomEvent) => {
    if (e.detail.list.template === 'documentLibrary') {
      const drive = await Providers.globalProvider.graph.client.api(`/sites/root/lists/${e.detail.id}/drive`).get();
      setSelectedList(e.detail);
      setDriveId(drive.id);
      setError('');
    } else {
      setSelectedList(null);
      setDriveId('');
      setError('Please select a document library');
    }
  };

  return (
    <div>
      <Picker
        resource="/sites/root/lists"
        placeholder="Select a list"
        keyName="displayName"
        selectionChanged={onSelectionChanged}
        className={styles.picker}
      ></Picker>

      {selectedList && driveId && <FileList itemPath="/" driveId={driveId} pageSize={100}></FileList>}

      {error && <div>{error}</div>}
    </div>
  );
};
