import { PersonCardInteraction } from '../components/PersonCardInteraction';

export const personCardConverter = (value: string) => {
  value = value.toLowerCase();
  if (typeof PersonCardInteraction[value] === 'undefined') {
    return PersonCardInteraction.none;
  } else {
    return PersonCardInteraction[value] as PersonCardInteraction;
  }
};
