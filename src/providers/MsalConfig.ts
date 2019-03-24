import { LoginType } from "./IProvider";
export interface MsalConfig {
    clientId: string;
    scopes?: string[];
    authority?: string;
    loginType?: LoginType;
    options?: any;
}