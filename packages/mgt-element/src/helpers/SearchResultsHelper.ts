import { ThemeCSSVariables } from '../Constants';

/**
 * Fields to be added to results
 */
enum SearchResponseEnhancedFields {
  FileTypeFamily = 'filefamily',
  OpenInBrowserUrl = 'previewurl'
}

/**
 * Well known SharePoint managed properties (lower cased)
 */
enum WellKnownSearchProperties {
  FileType = 'filetype',
  Title = 'title',
  Summary = 'summary'
}

export class SearchResultsHelper {
  public static readonly FileTypeAssociations = {
    word: ['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm'],
    excel: ['xls', 'xlsx', 'csv', 'xlsm', 'xlsb', 'xlx', 'xml', 'csv', 'xltm', 'xlt', 'xltx'],
    powerpoint: ['ppt', 'pptx', 'pptm', 'pps', 'ppsm', 'ppsx', 'potx', 'potm', 'pot'],
    onenote: ['one'],
    text: ['txt', 'rtf'],
    visio: ['vsd', 'vsdx', 'vsdm'],
    webpage: ['aspx', 'html'],
    pdf: ['pdf'],
    archive: ['zip', '7z', 'rar']
  };

  /**
   * Get the file famnily for its extension
   * @param fileExtension the file extension
   * @returns the file family corresponding to this extensions
   */
  private static getFileIconType(fileExtension: string): string {
    let fileType = 'generic';
    if (fileExtension) {
      fileType = Object.keys(SearchResultsHelper.FileTypeAssociations).filter(key => {
        return SearchResultsHelper.FileTypeAssociations[key].indexOf(fileExtension) !== -1;
      })[0];
    }

    return fileType;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public static enhanceResults(items: any[]): any[] {
    return items.map(item => {
      // Title
      if (item[WellKnownSearchProperties.Title])
        item[WellKnownSearchProperties.Title] = item[WellKnownSearchProperties.Title].replace(/\r|\t|\n/g, '');

      // File family
      if (item[WellKnownSearchProperties.FileType])
        item[SearchResponseEnhancedFields.FileTypeFamily] = this.getFileIconType(
          item[WellKnownSearchProperties.FileType]
        );

      // Summary
      if (item[WellKnownSearchProperties.Summary])
        item[WellKnownSearchProperties.Summary] = this.getItemSummary(item[WellKnownSearchProperties.Summary]);

      return item;
    });
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public static getItemTitle(item: any) {
    if (item.resource?.fields?.name) {
      return item.resource?.fields?.name;
    } else if (item.resource?.fields?.title) {
      return item.resource?.fields?.title;
    } else if (item.resource?.name) {
      return item.resource?.name;
    }

    return '';
  }

  public static getItemSummary(summary: string) {
    // Special case with HitHighlightedSummary field
    // eslint-disable-next-line no-useless-escape
    return summary
      .replace(/<c0\>/g, `<span style='color:var(${ThemeCSSVariables.colorPrimary});font-weight:900'>`)
      .replace(/<\/c0\>/g, '</span>')
      .replace(/<ddd\/>/g, '&#8230;');
  }
}
