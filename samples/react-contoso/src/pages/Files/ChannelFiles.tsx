import * as React from 'react';
import { FileList, SelectedChannel, TeamsChannelPicker } from '@microsoft/mgt-react';
import { makeStyles } from '@fluentui/react-components';

const useStyles = makeStyles({
  fileGrid: {
    paddingBottom: '10px'
  }
});

export const ChannelFiles: React.FunctionComponent = () => {
  const [selectedChannel, setSelectedChannel] = React.useState<SelectedChannel | null>(null);
  const styles = useStyles();

  return (
    <div>
      <TeamsChannelPicker
        selectionChanged={e => setSelectedChannel(e.detail)}
        className={styles.fileGrid}
      ></TeamsChannelPicker>

      {selectedChannel && (
        <FileList
          groupId={selectedChannel.team.id}
          itemPath={selectedChannel.channel.displayName}
          pageSize={100}
        ></FileList>
      )}
    </div>
  );
};
