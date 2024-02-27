import { useContext, useEffect, useState } from "react";
import { BotInfoContext } from "../components/Context/BotInfoContext";

export const useBotInfo = () => {
  const botInfoClient = useContext(BotInfoContext);
  const [botInfo, setBotInfo] = useState(() => botInfoClient?.getState());
  useEffect(() => {
    if (botInfoClient) {
      botInfoClient.onStateChange(setBotInfo);
      return () => botInfoClient.offStateChange(setBotInfo);
    }
  }, [botInfoClient]);
  return botInfo;
};
