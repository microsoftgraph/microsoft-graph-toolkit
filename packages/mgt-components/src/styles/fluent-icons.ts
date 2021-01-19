const fileTypeMap = {
  PowerPoint: 'pptx',
  Word: 'docx',
  Excel: 'xlsx',
  Pdf: 'pdf',
  OneNote: 'onetoc',
  OneNotePage: 'one',
  InfoPath: 'xsn',
  Visio: 'vstx',
  Publisher: 'pub',
  Project: 'mpp',
  Access: 'accdb',
  Mail: 'email',
  Csv: 'csv',
  Archive: 'archive',
  Xps: 'vector',
  Audio: 'audio',
  Video: 'video',
  Image: 'photo',
  Web: 'html',
  Text: 'txt',
  Xml: 'xml',
  Story: 'genericfile',
  ExternalContent: 'genericfile',
  Folder: 'folder',
  Spsite: 'spo',
  Other: 'genericfile'
};

// tslint:disable-next-line:completed-docs
const baseUri = 'https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201008.001/assets/item-types';

export type IconSize = 16 | 20 | 24 | 32 | 40 | 48 | 64 | 96;

export function getFileTypeIconUri(type: string, size: IconSize, extension: 'png' | 'svg') {
  const fileType = fileTypeMap[type] || 'genericfile';
  return `${baseUri}/${size.toString()}/${fileType}.${extension}`;
}
