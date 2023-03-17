import { mergeStyleSets } from '@fluentui/react/lib/Styling';

const classNames = {
  editButton: 'edit-button',
  icon: 'icon',
  iconFilled: 'icon-filled',
  iconUnfilled: 'icon-unfilled'
};

const editButtonStyle = {
  width: '2rem',
  height: '2rem'
};

const colors = {
  iconRest: '#616161',
  iconHover: '#5b5fc7'
};

const editStyles = mergeStyleSets({
  editButton: [classNames.editButton, editButtonStyle],

  icon: [
    classNames.icon,
    {
      fill: colors.iconRest,
      [`.${classNames.editButton}:hover &`]: {
        fill: colors.iconHover
      }
    }
  ],

  iconFilled: [
    classNames.iconFilled,
    {
      display: 'none',
      [`.${classNames.editButton}:hover &`]: {
        display: 'block'
      }
    }
  ],

  iconUnfilled: [
    classNames.iconUnfilled,
    {
      display: 'block',
      [`.${classNames.editButton}:hover &`]: {
        display: 'none'
      }
    }
  ]
});

export { editStyles, editButtonStyle };
