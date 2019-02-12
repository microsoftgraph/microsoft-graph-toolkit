jest.mock('msal');

import { MSALProvider } from './MSALProvider';
import { MSALConfig } from './MSALConfig';

describe('MSALProvider', () => {

    beforeEach(() => {
        jest.resetAllMocks();
    });

    it('undefined clientId should throw exception', () => {
        const config: MSALConfig = {
            clientId: undefined,
        };
        expect(() => {
            new MSALProvider(config);
        }).toThrowError("ClientID must be a valid string");
    });
    
});