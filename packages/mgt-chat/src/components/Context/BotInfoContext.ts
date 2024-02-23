import { createContext } from 'react';
import { BotInfoClient } from '../../statefulClient/BotInfoClient';

export const BotInfoContext = createContext<BotInfoClient | undefined>(undefined);
