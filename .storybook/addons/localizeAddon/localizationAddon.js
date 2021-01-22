import addons, { makeDecorator } from '@storybook/addons';
import { LocalizationHelper } from '../../../packages/mgt-element/dist/es6/utils/LocalizationHelper';
import { withKnobs, text, object, select } from '@storybook/addon-knobs';
import { STORY_CHANGED, STORY_INIT, FORCE_RE_RENDER } from '@storybook/core-events';

export const localize = makeDecorator({
  name: `localize`,
  parameterName: 'myParameter',
  decorator: [withKnobs],
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    LocalizationHelper.strings = context.parameters.strings;

    let defaultStrings = {
      _components: {
        login: {
          signInLinkSubtitle: 'Sign In',
          signOutLinkSubtitle: 'Sign Out'
        },
        'people-picker': {
          inputPlaceholderText: 'Start typing a name',
          noResultsFound: `We didn't find any matches.`,
          loadingMessage: 'Loading...'
        },
        'teams-channel-picker': {
          inputPlaceholderText: 'Select a channel',
          noResultsFound: `We didn't find any matches.`,
          loadingMessage: 'Loading...'
        },
        tasks: {
          removeTaskSubtitle: 'Delete Task',
          cancelNewTaskSubtitle: 'cancel',
          newTaskPlaceholder: 'Task...',
          addTaskButtonSubtitle: 'Add'
        },
        todo: {
          removeTaskSubtitle: 'Delete Task',
          cancelNewTaskSubtitle: 'cancel',
          newTaskPlaceholder: 'Task...',
          addTaskButtonSubtitle: 'Add'
        }
      }
    };

    for (let component in defaultStrings['_components']) {
      for (let stringKey in defaultStrings['_components'][component]) {
        let keys = stringKey;
        if (LocalizationHelper.strings['_components'][component][keys]) {
          for (let key in LocalizationHelper.strings['_components'][component]) {
            let value = text(key, LocalizationHelper.strings['_components'][component][key], component);
            if (value) {
              LocalizationHelper.strings['_components'][component][key] = value;
            }
          }
        }
      }
    }

    let channel = addons.getChannel();
    channel.on(STORY_CHANGED, () => {
      LocalizationHelper.strings = defaultStrings;
    });

    object('DefaultStrings', { ...defaultStrings }, 'DefaultStrings');
    if (context.name === 'Localization') {
      LocalizationHelper.strings = context.parameters.strings;
    }
    object('NewStrings', { ...context.parameters.strings }, 'NewStrings');

    return getStory(context);
  }
});
