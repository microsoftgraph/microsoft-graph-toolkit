import React, { ChangeEvent, useCallback, useState } from 'react';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react';
import { shorthands, makeStyles, Button, Input, InputOnChangeData } from '@fluentui/react-components';
import { IDynamicPerson, PeoplePicker } from '@microsoft/mgt-react';

interface AddChatMembersProps {
  closeDialog: () => void;
  addChatMembers: (userIds: string[], history?: Date) => Promise<void>;
}

const useStyles = makeStyles({
  addMembers: {
    display: 'flex',
    flexDirection: 'column',
    paddingBlockEnd: '18px',
    ...shorthands.paddingInline('20px')
  },
  buttonRow: {
    display: 'flex',
    columnGap: '8px',
    flexDirection: 'row',
    justifyContent: 'flex-end',
    paddingBlockStart: '18px'
  },
  dialogHeading: {
    paddingInlineStart: '4px',
    marginBlockStart: '18px',
    marginBlockEnd: '10px'
  },
  historyInput: { width: '48px' },
  option: {
    display: 'flex',
    alignItems: 'center',
    columnGap: '4px'
  },
  radio: {
    marginBlockStart: '8px',
    paddingInlineStart: '2px'
  }
});

const AddChatMembers = ({ addChatMembers, closeDialog }: AddChatMembersProps) => {
  const styles = useStyles();
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
      if (option) setHistoryOption(option.key);
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

  const addToChat = useCallback(() => {
    let history: Date | undefined;
    if (historyOption === 'days') {
      history = new Date(new Date().getTime() - daysOfHistory * 24 * 60 * 60 * 1000);
    } else if (historyOption === 'all') {
      history = new Date('0001-01-01T00:00:00Z');
    }
    void addChatMembers(
      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
      selectedPeople.map(p => p.id!),
      history
    ).then(closeDialog);
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
        return render ? (
          <div className={styles.option}>
            {render(props)}
            <Input
              type="number"
              value={daysOfHistory.toString()}
              onChange={handleHistoryChanged}
              min={1}
              className={styles.historyInput}
            />
          </div>
        ) : null;
      }
    },
    { key: 'all', text: 'Include all chat history' }
  ];
  return (
    <div className={styles.addMembers}>
      <h3 className={styles.dialogHeading}>Add</h3>
      <PeoplePicker
        selectedPeople={selectedPeople}
        selectionChanged={peopleSelectionChanged}
        disabled={selectedPeople.length >= 20}
      />
      <ChoiceGroup
        className={styles.radio}
        options={chatHistoryOptions}
        selectedKey={historyOption}
        onChange={onHistoryOptionChange}
      />
      <div className={styles.buttonRow}>
        <Button onClick={closeDialog} appearance="secondary">
          Cancel
        </Button>
        <Button onClick={addToChat} appearance="primary" disabled={selectedPeople.length < 1}>
          Add
        </Button>
      </div>
    </div>
  );
};

export { AddChatMembers };
