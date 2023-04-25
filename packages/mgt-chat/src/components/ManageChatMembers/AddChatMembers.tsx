import React, { ChangeEvent, useCallback, useState } from 'react';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react';
import { Button } from '@fluentui/react-components';
import { Input, InputOnChangeData } from '@fluentui/react-components';
import { IDynamicPerson, PeoplePicker } from '@microsoft/mgt-react';
import { styles } from './manage-chat-members.styles';

interface AddChatMembersProps {
  closeDialog: () => void;
  addChatMembers: (userIds: string[], history?: Date) => Promise<void>;
}

const AddChatMembers = ({ addChatMembers, closeDialog }: AddChatMembersProps) => {
  const [daysOfHistory, setDaysOfHistory] = useState<number>(1);
  const handleHistoryChanged = useCallback(
    (_: ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
      if (data.value) {
        const newNumber = parseInt(data.value, 10);
        setDaysOfHistory(newNumber);
      }
    },
    [setDaysOfHistory]
  );

  const [historyOption, setHistoryOption] = useState<string | undefined>('none');
  const onHistoryOptionChange = React.useCallback(
    (_, option: IChoiceGroupOption | undefined) => {
      option && setHistoryOption(option.key);
    },
    [setHistoryOption]
  );

  const [selectedPeople, setSelectedPeople] = useState<IDynamicPerson[]>([]);
  const peopleSelectionChanged = useCallback(
    (e: CustomEvent<IDynamicPerson[]>) => {
      setSelectedPeople(e.detail);
    },
    [setSelectedPeople]
  );

  const addToChat = useCallback(async () => {
    let history: Date | undefined = undefined;
    if (historyOption === 'days') {
      history = new Date(new Date().getTime() - daysOfHistory * 24 * 60 * 60 * 1000);
    } else if (historyOption === 'all') {
      history = new Date('0001-01-01T00:00:00Z');
    }
    await addChatMembers(
      selectedPeople.map(p => p.id!),
      history
    );
    closeDialog();
  }, [selectedPeople, historyOption, daysOfHistory, addChatMembers, closeDialog]);

  const chatHistoryOptions: IChoiceGroupOption[] = [
    {
      key: 'none',
      text: "Don't include chat history"
    },
    {
      key: 'days',
      text: 'Include history from the past number of days: ',

      onRenderField: (props, render) => {
        return (
          <div className={styles.option}>
            {render!(props)}
            <Input
              type="number"
              value={daysOfHistory.toString()}
              onChange={handleHistoryChanged}
              min={1}
              className={styles.historyInput}
            />
          </div>
        );
      }
    },
    { key: 'all', text: 'Include all chat history' }
  ];
  return (
    <div className={styles.addMembers}>
      <h3>Add</h3>
      <PeoplePicker
        selectedPeople={selectedPeople}
        selectionChanged={peopleSelectionChanged}
        disabled={selectedPeople.length >= 20}
      />
      <ChoiceGroup options={chatHistoryOptions} selectedKey={historyOption} onChange={onHistoryOptionChange} />
      <div className={styles.buttonRow}>
        <Button onClick={addToChat} appearance="primary" disabled={selectedPeople.length < 1}>
          Save
        </Button>
        <Button onClick={closeDialog} appearance="secondary">
          Cancel
        </Button>
      </div>
    </div>
  );
};

export { AddChatMembers };
