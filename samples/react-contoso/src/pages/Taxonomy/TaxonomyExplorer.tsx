import * as React from 'react';
import { Get, MgtTemplateProps, TaxonomyPicker } from '@microsoft/mgt-react';
import {
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  makeStyles,
  shorthands,
  Card,
  CardHeader,
  Caption1,
  tokens,
  Text
} from '@fluentui/react-components';
import { TermStore } from '@microsoft/microsoft-graph-types';
import { TagMultipleRegular, TagRegular } from '@fluentui/react-icons';

export const TaxonomyExplorer: React.FunctionComponent = () => {
  return (
    <>
      <Get resource="termStore/groups" version="beta" scopes={["TermStore.Read.All"]}>
        <GroupTemplate template="default" />
      </Get>
    </>
  );
};

const useStyles = makeStyles({
  main: {
    ...shorthands.gap('36px'),
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'wrap'
  },

  title: {
    ...shorthands.margin(0, 0, '12px')
  },

  description: {
    ...shorthands.margin(0, 0, '12px')
  },

  card: {
    width: '480px',
    maxWidth: '100%',
    height: 'fit-content',
    ...shorthands.margin('12px', 0)
  },

  caption: {
    color: tokens.colorNeutralForeground3
  },

  icon: {
    width: '24px',
    height: '24px'
  },

  text: {
    ...shorthands.margin(0)
  },

  groupPanel: {
    ...shorthands.margin('12px', '24px')
  },

  termPanel: {
    ...shorthands.margin('12px', '36px')
  }
});

const GroupTemplate = (props: MgtTemplateProps) => {
  const styles = useStyles();
  const [groups] = React.useState<any[]>(props.dataContext.value);

  return (
    <Accordion collapsible>
      {groups.map(group => (
        <AccordionItem value={group.id} key={group.id}>
          <AccordionHeader icon={<TagMultipleRegular />}>{group.displayName}</AccordionHeader>
          <AccordionPanel className={styles.groupPanel}>
            <Get resource={`termStore/groups/${group.id}/sets`} version="beta" scopes={["TermStore.Read.All"]}>
              <SetTemplate template="default" />
            </Get>
          </AccordionPanel>
        </AccordionItem>
      ))}
    </Accordion>
  );
};

const SetTemplate = (props: MgtTemplateProps) => {
  const styles = useStyles();
  const [sets] = React.useState<any[]>(props.dataContext.value);
  const [selectedTerm, setSelectedTerm] = React.useState<TermStore.Term | null>(null);

  return (
    <Accordion collapsible onToggle={() => setSelectedTerm(null)}>
      {sets.map(set => (
        <AccordionItem value={set.id} key={set.id}>
          <AccordionHeader>{set.localizedNames[0].name}</AccordionHeader>
          <AccordionPanel className={styles.termPanel}>
            <TaxonomyPicker termsetId={set.id} selectionChanged={e => setSelectedTerm(e.detail)} />
            {selectedTerm && (
              <Card className={styles.card}>
                <CardHeader
                  image={<TagRegular className={styles.icon} />}
                  header={<Text weight="semibold">{selectedTerm.labels?.[0].name}</Text>}
                  description={<Caption1 className={styles.caption}>{selectedTerm.id}</Caption1>}
                />

                {selectedTerm.descriptions?.length! > 0 && (
                  <p className={styles.text}>{selectedTerm.descriptions?.[0].description}</p>
                )}
              </Card>
            )}
          </AccordionPanel>
        </AccordionItem>
      ))}
    </Accordion>
  );
};
