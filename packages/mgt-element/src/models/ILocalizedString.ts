export interface ILocalizedString {
  /**
   * The default label to use
   */
  default: string;

  /**
   * Any other locales for the string (ex: 'fr-fr')
   */
  [locale: string]: string;
}
