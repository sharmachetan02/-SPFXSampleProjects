import moment from 'moment';
import { isPossiblePhoneNumber } from 'react-phone-number-input';
export class UtilHelper {

  public static MIN_DATE = new Date('0001-01-01');
  public static MAX_DATE = new Date('9999-12-31');

  public static isToday(date: Date) {
    const today = new Date();
    return date.getDate() === today.getDate() &&
      date.getMonth() === today.getMonth() &&
      date.getFullYear() === today.getFullYear();
  }

  public static getTomorrowDate() {
    const today = new Date();
    const tomorrow = new Date();
    tomorrow.setDate(today.getDate() + 1);
    tomorrow.setHours(0, 0, 0, 0);
    return tomorrow;
  }

  public static capitalizeFirstLetter = (string: string): string => {
    if (string)
      return string.charAt(0).toUpperCase() + string.toLocaleLowerCase().slice(1);
  };

  public static capitalizeAllFirstLetter = (string: string): string => {
    if (string) {
      const strings = string.toLocaleLowerCase().split(' ');
      for (let i = 0; i < strings.length; i++) {
        strings[i] = UtilHelper.capitalizeFirstLetter(strings[i]);
      }
      return strings.join(' ');
    }
  };
  public static toLocaleDateString = (date: Date, locale: string, dateFormat?: Intl.DateTimeFormatOptions): string => {
    dateFormat = dateFormat ?? { year: 'numeric', month: 'long', day: '2-digit' };
    return date.toLocaleDateString(locale, dateFormat);
  };
  public static formatDate(date: Date, format: string, local: string, utc = false) {
    if (!date) return '';
    moment.updateLocale('en', {
      monthsShort: [
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
      ]
    });
    moment.updateLocale('fr', {
      monthsShort: [
        'Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin',
        'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'
      ]
    });
    const momentLocal = (local === 'en-us') ? 'en' : 'fr';
    moment.locale(momentLocal);
    return !utc ? moment(date).format(format) : moment(date).utc().format(format);
  }
  public static toUtcDate(date: Date): Date {
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(),
      date.getUTCDate(), date.getUTCHours(),
      date.getUTCMinutes(), date.getUTCSeconds()));
  }
  public static getMaxDate(dates: Date[]): Date {
    const moments = dates.map((d) => moment(d));
    return moment.max(moments).toDate();
  }
  public static sliceIntoChunks(arr, chunkSize) {
    const res = [];
    for (let i = 0; i < arr.length; i += chunkSize) {
      const chunk = arr.slice(i, i + chunkSize);
      res.push(chunk);
    }
    return res;
  }
  public static compareObjectArrays(arr1, arr2) {
    return (
      arr1.length === arr2.length &&
      arr1.every((element1) =>
        arr2.some((element2) =>
          Object.keys(element1).every((key) => element1[key] === element2[key])
        )
      )
    );
  }
  public static toNormalForm(str: string) {
    return str?.normalize('NFD')?.replace(/[\u0300-\u036f]/g, '')?.trim();
  }
  // public static groupBy(array, property)  {
  //   return array.reduce((grouped, element) => ({
  //     ...grouped,
  //     [element[property]]: [...(grouped[element[property]] || []), element]
  //   }), {});
  // }
  public static sanitizeSharePointTitle(input: string): string {
    if (!input) return '';

    return input
      // Replace forbidden SP characters with underscore
      .replace(/[~"'#%&*:<>?/\\{|}]/g, '')
      // Replace non-ASCII characters with underscore
      .replace(/[^\u0020-\u007E]/g, '')  // printable ASCII only
      // Collapse multiple spaces and trim
      .replace(/\s+/g, ' ')
      .trim();
  }
  public static sanitizeSharePointFolderName(input: string): string {
    if (!input) return '';

    return input
    .replace(/[~"'#%&*:<>?/\\{|}]/g, '')
    .replace(/[^\u0020-\u007E]/g, '')  // printable ASCII only
    .replace(/^[.\s]+|[.\s]+$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  }
  public static validateURL(str: string): boolean {
    if (!str) return true;
    const urlRegex = new RegExp(
      '^https?:\\/\\/' +                                   // protocol
      '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' + // domain
      '((\\d{1,3}\\.){3}\\d{1,3})' +                       // OR IPv4
      ')' +                                                // <-- closing domain/ip group
      '(\\:\\d+)?' +                                       // optional port
      '(\\/[-a-z\\d%_.~+]*)*' +                            // path
      '(\\?[;&a-z\\d%_.~+=-]*)?' +                         // query
      '(\\#[-a-z\\d_]*)?$',                                // fragment
      'i'
    );
    return !!urlRegex.test(str) && str.length < 256;
  }
  public static validateEmail(str: string): boolean {
    const expression = new RegExp(
      '(?!.*\\.\\.)^([a-z\\d!#$%&\'*+\\-\\/=?^_`{|}~\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]+' +
      '(\\.[a-z\\d!#$%&\'*+\\-\\/=?^_`{|}~\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]+)*|' +
      '"((([\\t]*\\r\\n)?[\\t]+)?([\\x01-\\x08\\x0b\\x0c\\x0e-\\x1f\\x7f\\x21\\x23-\\x5b\\x5d-\\x7e\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]|' +
      '\\\\[\\x01-\\x09\\x0b\\x0c\\x0d-\\x7f\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]))*' +
      '(([\\t]*\\r\\n)?[\\t]+)?")@(([a-z\\d\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]|' +
      '[a-z\\d\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF][a-z\\d\\-._~\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]*' +
      '[a-z\\d\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF])\\.)+' +
      '([a-z\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]|' +
      '[a-z\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF][a-z\\d\\-._~\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF]*' +
      '[a-z\\u00A0-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFEF])\\.?$',
      'i'
    );


    return expression.test(String(str).toLowerCase());
  }
  public static validatePhoneNumber(str: string): boolean {
    return isPossiblePhoneNumber(str);
  }
}
