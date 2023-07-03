export class DateHelper {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private dayJsLibrary: any;

  private culture: string;

  public constructor(culture?: string) {
    this.culture = culture ? culture.toLocaleLowerCase() : 'en-us';
  }

  public async dayJs(): Promise<unknown> {
    if (this.dayJsLibrary) {
      return Promise.resolve(this.dayJsLibrary);
    } else {
      const dayjs = await import('dayjs/esm');

      const localizedFormat = await import('dayjs/esm/plugin/localizedFormat');

      this.dayJsLibrary = dayjs.default;
      this.dayJsLibrary.extend(localizedFormat.default);

      const twoLetterLanguageName = [
        'af',
        'az',
        'be',
        'bg',
        'bm',
        'bo',
        'br',
        'bs',
        'ca',
        'cs',
        'cv',
        'cy',
        'da',
        'de-de',
        'dv',
        'el',
        'eo',
        'es-es',
        'et',
        'eu',
        'fa',
        'fi',
        'fil',
        'fo',
        'fy',
        'fr-fr',
        'ga',
        'gd',
        'gl',
        'gu',
        'he',
        'hi',
        'hr',
        'hu',
        'id',
        'is',
        'it-it',
        'ja',
        'jv',
        'ka',
        'kk',
        'km',
        'kn',
        'ko',
        'ku',
        'ky',
        'lb',
        'lo',
        'lt',
        'lv',
        'me',
        'mi',
        'mk',
        'ml',
        'mn',
        'mr',
        'mt',
        'my',
        'nb',
        'ne',
        'nn',
        'nl-nl',
        'pl',
        'pt-pt',
        'ro',
        'ru',
        'sd',
        'se',
        'si',
        'sk',
        'sl',
        'sq',
        'ss',
        'sv',
        'sw',
        'ta',
        'te',
        'tet',
        'tg',
        'th',
        'tk',
        'tlh',
        'tr',
        'tzl',
        'uk',
        'ur',
        'vi',
        'yo'
      ];

      // DayJs is by default "en-us"
      if (!this.culture.startsWith('en-us')) {
        for (let i = 0; i < twoLetterLanguageName.length; i++)
          if (this.culture.startsWith(twoLetterLanguageName[i])) {
            this.culture = this.culture.split('-')[0];
            break;
          }

        await import(`dayjs/esm/locale/${this.culture}.js`);
      }

      // Set default locale
      this.dayJsLibrary.locale(this.culture);
      return this.dayJsLibrary;
    }
  }

  public isDST() {
    const today = new Date();
    const jan = new Date(today.getFullYear(), 0, 1);
    const jul = new Date(today.getFullYear(), 6, 1);
    const stdTimeZoneOffset = Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
    return today.getTimezoneOffset() < stdTimeZoneOffset;
  }

  public addMinutes(isDst: boolean, date: Date, minutes: number, dst: number) {
    if (isDst) {
      minutes += dst;
    }
    return new Date(date.getTime() + minutes * 60000);
  }
}
