import { MgtTemplateProps, Person } from '@microsoft/mgt-react';
import { makeStyles, mergeClasses, shorthands } from '@fluentui/react-components';

const useStyles = makeStyles({
  email: {
    boxShadow: 'var(--box-shadow)',
    ...shorthands.padding('10px'),
    ...shorthands.margin('8px'),
    ':hover': {
      borderLeftWidth: '4px',
      borderLeftColor: 'var(--input-border-color--hover)',
      borderLeftStyle: 'solid',
      paddingLeft: '6px'
    },
    '& mgt-person': {
      '--font-size': '12px',
      '--person-avatar-size': '16px'
    }
  },

  link: {
    color: 'var(--color-sub1)',
    textDecorationLine: 'none'
  },

  header: {
    display: 'flex',
    justifyContent: 'space-between'
  },

  subject: {
    color: 'var(--color-sub1)',
    fontSize: '14px',
    ...shorthands.gridArea('1 / 1 / auto / 3'),
    ...shorthands.margin('0')
  },

  title: {
    display: 'flex',
    justifyContent: 'space-between',
    marginBottom: '4px',
    color: 'var(--color-sub1)'
  },

  date: {
    fontSize: '12px',
    paddingLeft: '4px',
    float: 'right'
  },

  body: {
    fontSize: '13px',
    textOverflow: 'ellipsis',
    wordWrap: 'break-word',
    ...shorthands.overflow('hidden'),
    maxHeight: '2.8em',
    lineHeight: '1.4em',
    color: 'var(--color-sub2)'
  },

  emptyBody: {
    fontStyle: 'italic'
  }
});

export function Messages(props: MgtTemplateProps) {
  const styles = useStyles();
  const email = props.dataContext;
  return (
    <div className={styles.email}>
      <a className={styles.link} href={email.webLink} target="_blank" rel="noreferrer">
        <div className={styles.header}>
          <div>
            <Person personQuery={email.sender?.emailAddress?.address} personCardInteraction="hover" view="oneline" />
          </div>
        </div>
        <div className={styles.title}>
          <h3 className={styles.subject}>{email.subject}</h3>
          <span className={styles.date}>{new Date(email.receivedDateTime).toLocaleDateString()}</span>
        </div>
        {email?.bodyPreview ? (
          <div className={styles.body}>{email.bodyPreview}</div>
        ) : (
          <div className={mergeClasses(styles.body, styles.emptyBody)}>...</div>
        )}
      </a>
    </div>
  );
}
