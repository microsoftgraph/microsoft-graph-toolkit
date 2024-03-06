import { NullableOption, Identity } from '@microsoft/microsoft-graph-types';

export const botPrefix = 'botId::';

export const buildBotId = (application: NullableOption<Identity> | undefined) =>
  application?.id ? `${botPrefix}${application.id}` : undefined;
