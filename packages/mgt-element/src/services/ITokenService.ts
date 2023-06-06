export interface ITokenService {
  /**
   * Sets the value for a specific token
   * @param token the token name
   * @param value the value to set
   */
  setTokenValue(token: string, value: string): void;

  /**
   * Gets the value of a specific token
   * @param token the token name to retrieve
   */
  getTokenValue(token: string): string;

  /**
   * Resolves tokens for the specified input string.
   * @param string
   */
  resolveTokens(string: string): string;
}
