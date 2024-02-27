import { TeamsAppInstallation } from '@microsoft/microsoft-graph-types-beta';
import { BaseStatefulClient } from './BaseStatefulClient';
import { loadBotInChat, loadBotIcon } from './graph.chat';

export interface BotInfo {
  botInfo: Map<string, TeamsAppInstallation>;
  chatBots: Map<string, Set<TeamsAppInstallation>>;
  botIcons: Map<string, string>;
  loadBotInfo: (chatId: string, botId: string) => Promise<void>;
}

const requestKey = (chatId: string, botId: string) => `${chatId}::${botId}`;

export class BotInfoClient extends BaseStatefulClient<BotInfo> {
  private readonly infoRequestMap = new Map<string, Promise<unknown>>();

  public loadBotInfo = async (chatId: string, botId: string) => {
    const beta = this.betaGraph;
    if (!beta) return;
    const key = requestKey(chatId, botId);
    if (!this.infoRequestMap.has(key)) {
      const requestPromise = loadBotInChat(beta, chatId, botId);
      this.infoRequestMap.set(key, requestPromise);
      const botInfo = await requestPromise;
      if (botInfo?.value.length > 0) {
        // update state with the text info
        this.notifyStateChange((draft: BotInfo) => {
          botInfo.value.forEach(app => {
            if (app.teamsAppDefinition?.bot?.id) {
              draft.botInfo.set(app.teamsAppDefinition?.bot?.id, app);
              if (!draft.chatBots.has(chatId)) {
                draft.chatBots.set(chatId, new Set());
              }
              draft.chatBots.get(chatId)?.add(app);
            }
          });
        });
        // load the appIcon for each bot
        for (const app of botInfo.value) {
          const appIcon = await loadBotIcon(beta, app);
          this.notifyStateChange((draft: BotInfo) => {
            if (app.teamsAppDefinition?.bot?.id) {
              draft.botIcons.set(app.teamsAppDefinition?.bot?.id, appIcon);
            }
          });
        }
      }
    }
  };
  protected state: BotInfo = {
    botInfo: new Map<string, TeamsAppInstallation>(),
    chatBots: new Map<string, Set<TeamsAppInstallation>>(),
    botIcons: new Map<string, string>(),
    loadBotInfo: this.loadBotInfo
  };
}
