import { LoginType } from "./IAuthProvider";
export interface MSALConfig {
    clientId: string;
    scopes?: string[];
    authority?: string;
    loginType?: LoginType;
    options?: any;
}