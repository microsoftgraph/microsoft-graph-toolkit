import { ComponentRegistry } from '../../utils/ComponentRegistry';
import { MgtPerson } from './mgt-person';

ComponentRegistry.register('mgt-person', MgtPerson);

export * from './mgt-person';
