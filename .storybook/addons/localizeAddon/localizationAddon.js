import addons, { makeDecorator } from '@storybook/addons';
import { LocalizationHelper } from '../../../packages/mgt/dist/es6/index.js';
import { withKnobs, text, object, select } from '@storybook/addon-knobs';
import { STORY_CHANGED } from '@storybook/core-events';

export const localize = makeDecorator({
  name: `localize`,
  parameterName: 'myParameter',
  decorator: [withKnobs],
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    const label = 'dir';
    const options = {
      LeftToRight: 'ltr',
      RightToLeft: 'rtl'
    };
    let defaultValue = 'ltr';
    const groupId = 'Document Direction';
    const value = select(label, options, defaultValue, groupId);
    document.body.setAttribute('dir', value);

    const channel = addons.getChannel();

    channel.on(STORY_CHANGED, params => {
      document.body.setAttribute('dir', 'ltr');
    });

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

    if (context.name === 'Localization') {
      LocalizationHelper.strings = {
        _components: {
          login: {
            signInLinkSubtitle: 'login',
            signOutLinkSubtitle: 'خروج'
          },
          'people-picker': {
            inputPlaceholderText: 'ابدأ في كتابة الاسم',
            noResultsFound: 'لم نجد أي قنوات',
            loadingMessage: '...جار التحميل'
          },
          'teams-channel-picker': {
            inputPlaceholderText: 'حدد قناة',
            noResultsFound: `noResultsFoundTest`,
            loadingMessage: 'Loading...'
          },
          tasks: {
            removeTaskSubtitle: 'delete',
            cancelNewTaskSubtitle: 'canceltest',
            newTaskPlaceholder: 'newTaskTest',
            addTaskButtonSubtitle: 'addme'
          },
          todo: {
            removeTaskSubtitle: 'todoremoveTEST'
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
    } else {
      LocalizationHelper.strings = defaultStrings;
    }

    let strings = LocalizationHelper.strings;

    object('DefaultStrings', { ...defaultStrings }, 'DefaultStrings');
    LocalizationHelper.strings = strings;
    return getStory(context);
  }
});
