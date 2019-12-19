import addonAPI from '@storybook/addons';
import { STORIES_CONFIGURED, STORY_MISSING } from '@storybook/core-events';

addonAPI.register('microsoft/graph-toolkit', storybookAPI => {
  storybookAPI.on(STORIES_CONFIGURED, (kind, story) => {
    if (storybookAPI.getUrlState().path === '/story/*') {
      storybookAPI.selectStory('mgt-login', 'login');
    }
  });
  storybookAPI.on(STORY_MISSING, (kind, story) => {
    storybookAPI.selectStory('mgt-login', 'login');
  });
});
