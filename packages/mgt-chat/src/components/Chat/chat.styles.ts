import { mergeStyleSets } from '@fluentui/react';
const chatStyles = mergeStyleSets({
  chat: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'auto'
  },

  chatMessages: {
    height: 'auto',
    overflow: 'auto'
  },

  chatInput: {
    overflow: 'unset'
  },
  fullHeight: {
    height: '100%'
  }
});

export { chatStyles };
