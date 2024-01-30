import React, { FC, useCallback, useState } from 'react';
import { Chat } from '@microsoft/microsoft-graph-types';
import { IGraph, PeoplePicker, Spinner, IDynamicPerson } from '@microsoft/mgt-react';
import {
  Button,
  Field,
  FluentProvider,
  Input,
  InputOnChangeData,
  Textarea,
  TextareaOnChangeData,
  makeStyles,
  typographyStyles,
  shorthands,
  webLightTheme
} from '@fluentui/react-components';
import { createChatThread } from '../../statefulClient/graph.chat';
import { graph } from '../../utils/graph';
import { currentUserId } from '../../utils/currentUser';

interface NewChatProps {
  mode?: 'oneOnOne' | 'group' | 'auto';
  onChatCreated: (chat: Chat) => void;
  onCancelClicked: () => void;
}

const useStyles = makeStyles({
  title: {
    ...typographyStyles.subtitle2,
    ...shorthands.marginInline('32px')
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    gridRowGap: '16px',
    minWidth: '300px',
    '& textarea': {
      fontSize: 'var(--type-ramp-base-font-size)',
      '::placeholder': {
        fontSize: 'var(--type-ramp-base-font-size)'
      }
    }
  },
  formButtons: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'flex-end',
    gridColumnGap: '8px'
  }
});

const NewChat: FC<NewChatProps> = ({ mode = 'auto', onChatCreated, onCancelClicked }: NewChatProps) => {
  const styles = useStyles();
  type NewChatState = 'initial';

  const [isGroup, setIsGroup] = useState<boolean>(mode === 'group');

  const [state, setState] = useState<NewChatState | 'creating chat' | 'done'>('initial');
  // chat member data control
  const [selectedPeople, setSelectedPeople] = useState<IDynamicPerson[]>([]);
  const onSelectedPeopleChange = useCallback(
    (event: CustomEvent<IDynamicPerson[]>) => {
      if (event.detail) {
        setSelectedPeople(event.detail);
        if (mode === 'auto') {
          setIsGroup(event.detail.length > 1);
        }
      }
    },
    [mode]
  );
  // chat name data control
  const [chatName, setChatName] = useState<string>('');
  const onChatNameChanged = useCallback((_: React.FormEvent<HTMLInputElement>, data: InputOnChangeData) => {
    setChatName(data.value);
  }, []);

  // initial message data control
  const [initialMessage, setInitialMessage] = useState<string>('');
  const onInitialMessageChange = useCallback((_: React.FormEvent<HTMLTextAreaElement>, data: TextareaOnChangeData) => {
    setInitialMessage(data.value);
  }, []);
  const createChat = useCallback(() => {
    const graphClient: IGraph = graph('mgt-new-chat');
    setState('creating chat');
    const chatMembers = [currentUserId()];
    selectedPeople.reduce((acc, person) => {
      if (person.id) acc.push(person.id);
      return acc;
    }, chatMembers);
    void createChatThread(graphClient, chatMembers, isGroup, initialMessage, chatName).then(chat => {
      setState('done');
      onChatCreated(chat);
    });
  }, [onChatCreated, selectedPeople, initialMessage, chatName, isGroup]);

  return (
    <FluentProvider theme={webLightTheme}>
      {state === 'initial' ? (
        <div className={styles.form}>
          <Field label="To">
            <PeoplePicker
              disabled={(mode === 'oneOnOne' && selectedPeople?.length > 0) || selectedPeople?.length > 19}
              ariaLabel="Select people to chat with"
              selectedPeople={selectedPeople}
              selectionChanged={onSelectedPeopleChange}
            />
          </Field>

          {isGroup && (
            <Field label="Group name">
              <Input placeholder="Chat name" onChange={onChatNameChanged} value={chatName} />
            </Field>
          )}
          <Textarea
            placeholder="Type your first message"
            size="large"
            resize="vertical"
            value={initialMessage}
            onChange={onInitialMessageChange}
          />
          <div className={styles.formButtons}>
            <Button appearance="secondary" onClick={onCancelClicked}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={createChat}>
              Send
            </Button>
          </div>
        </div>
      ) : (
        <>
          {state}
          {state !== 'done' && <Spinner />}
        </>
      )}
    </FluentProvider>
  );
};

export { NewChat };
