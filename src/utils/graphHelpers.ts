import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

/**
 * returns a promise that resolves after specified time
 * @param time in milliseconds
 */
export function getEmailFromGraphEntity(
  entity: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact
): string {
  const person = entity as MicrosoftGraph.Person;
  const user = entity as MicrosoftGraph.User;
  const contact = entity as MicrosoftGraph.Contact;

  if (user.mail) {
    return user.mail;
  } else if (person.scoredEmailAddresses && person.scoredEmailAddresses.length) {
    return person.scoredEmailAddresses[0].address;
  } else if (contact.emailAddresses && contact.emailAddresses.length) {
    return contact.emailAddresses[0].address;
  }

  return null;
}
