import { mergeStyles } from '@fluentui/react';
const styles = {
  chat: mergeStyles({
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'auto'
  }),

  chatMessages: mergeStyles({
    height: 'auto',
    overflow: 'auto'
  }),

  chatInput: mergeStyles({
    overflow: 'unset'
  })
};

export { styles };
