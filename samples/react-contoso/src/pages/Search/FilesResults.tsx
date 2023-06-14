import { SearchResults } from '@microsoft/mgt-react/dist/es6/generated/react-preview';
import * as React from 'react';
import { IResultsProps } from './IResultsProps';
import { MgtTemplateProps, Person, PersonCardInteraction, ViewType } from '@microsoft/mgt-react';
import {
  Text,
  Button,
  Caption1,
  Card,
  CardFooter,
  CardHeader,
  CardPreview,
  makeStyles,
  shorthands,
  tokens
} from '@fluentui/react-components';
import { ArrowReplyRegular, ShareRegular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    ...shorthands.gap('16px'),
    display: 'flex',
    flexWrap: 'wrap'
  },
  card: {
    width: '280px',
    height: 'fit-content'
  },
  caption: {
    color: tokens.colorNeutralForeground3
  },
  grid: {
    ...shorthands.gap('16px'),
    display: 'flex',
    flexDirection: 'column'
  }
});

export const FilesResults: React.FunctionComponent<IResultsProps> = (props: IResultsProps) => {
  const styles = useStyles();

  return (
    <>
      {props.searchTerm && (
        <SearchResults
          entityTypes={['driveItem']}
          queryString={props.searchTerm}
          fetchThumbnail={true}
          queryTemplate="({searchTerms}) ContentTypeId:0x0101*"
          fields={['ContentTypeId']}
          version="beta"
        >
          <FileTemplate template="default"></FileTemplate>
        </SearchResults>
      )}
    </>
  );
};

const FileTemplate = (props: MgtTemplateProps) => {
  const styles = useStyles();
  const [driveItems] = React.useState<any>(props.dataContext.value);

  return (
    /*
    <Card className={styles.card}>
      <Person
        personQuery={driveItem.resource.lastModifiedBy.user.email}
        view={ViewType.twolines}
        personCardInteraction={PersonCardInteraction.hover}
        showPresence={true}
      ></Person>

      <CardPreview logo={<img src={''} alt="Microsoft Word document" />}>
        <img src={driveItem.resource.thumbnail?.url} alt={driveItem.resource.name} />
      </CardPreview>

      <CardFooter>
        <Button icon={<ArrowReplyRegular fontSize={16} />}>Reply</Button>
        <Button icon={<ShareRegular fontSize={16} />}>Share</Button>
      </CardFooter>
    </Card>
  */

    <div className={styles.container}>
      <div className={styles.grid}>
        {driveItems.map((driveItem: any) => (
          <Card className={styles.card} size="small" role="listitem">
            <CardHeader
              image={{ as: 'img', alt: 'Word app logo' }}
              header={<Text weight="semibold">{driveItem.resource.name}</Text>}
              description={<Caption1 className={styles.caption}>OneDrive &gt; Documents</Caption1>}
            />
          </Card>
        ))}
      </div>
    </div>
  );
};
