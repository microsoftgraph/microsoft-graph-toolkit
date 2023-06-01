import React, { FC, useCallback, useState } from 'react';
import { IDynamicPerson } from '@microsoft/mgt-components';
import { Chat } from '@microsoft/microsoft-graph-types';
import { IGraph, PeoplePicker, Spinner } from '@microsoft/mgt-react';
import {
  Button,
  Divider,
  Field,
  FluentProvider,
  Input,
  InputOnChangeData,
  Text,
  Textarea,
  TextareaOnChangeData,
  teamsLightTheme,
  makeStyles,
  typographyStyles,
  shorthands,
  tokens,
  Theme
} from '@fluentui/react-components';
import { createChatThread } from '../../statefulClient/graph.chat';
import { graph } from '../../utils/graph';
import { currentUserId } from '../../utils/currentUser';

interface NewChatProps {
  hideTitle?: boolean;
  title?: string;
  mode?: 'oneOnOne' | 'group' | 'auto';
  onChatCreated: (chat: Chat) => void;
  onCancelClicked: () => void;
  theme?: Theme;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.paddingBlock('3px', '16px'),
    minWidth: '300px',
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow8
  },
  title: {
    ...typographyStyles.subtitle2,
    ...shorthands.marginInline('32px')
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    gridRowGap: '16px',
    marginBlockStart: '16px',
    ...shorthands.marginInline('32px')
  },
  formButtons: {
    display: 'flex',
    flexDirection: 'row',
    justifyContent: 'flex-end',
    gridColumnGap: '8px'
  }
});

const NewChat: FC<NewChatProps> = ({
  hideTitle,
  title,
  mode = 'auto',
  onChatCreated,
  onCancelClicked,
  theme
}: NewChatProps) => {
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
  const [chatName, setChatName] = useState<string>();
  const onChatNameChanged = useCallback(
    (_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, data: InputOnChangeData) => {
      setChatName(data.value);
    },
    []
  );
  // initial message data control
  const [initialMessage, setInitialMessage] = useState<string>();
  const onInitialMessageChange = useCallback(
    (event: React.FormEvent<HTMLTextAreaElement>, data: TextareaOnChangeData) => {
      setInitialMessage(data.value);
    },
    []
  );
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
    <FluentProvider theme={theme ? theme : teamsLightTheme}>
      {state === 'initial' ? (
        <div className={styles.container}>
          {!hideTitle && (
            <>
              <Text as="h2" className={styles.title}>
                {title ? title : 'New chat'}
              </Text>
              <Divider />
            </>
          )}
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
        </div>
      ) : (
        <>{state !== 'done' && <Spinner />}</>
      )}
    </FluentProvider>
  );
};

export { NewChat };
