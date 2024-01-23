import { PersonCardInteraction } from '../components/PersonCardInteraction';

export const personCardConverter = (value: string, _type: unknown) => {
  value = value.toLowerCase();
  if (typeof PersonCardInteraction[value] === 'undefined') {
    return PersonCardInteraction.none;
  } else {
    return PersonCardInteraction[value] as PersonCardInteraction;
  }
};
