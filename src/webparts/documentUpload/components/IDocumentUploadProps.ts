import { IDocumentLibrary, IField } from '../DocumentUploadWebPart';

export interface IDocumentUploadProps {
  // Hold all of the list fields used to generate form controls
  fields: Array<IField>;

  // Contains a list of document libraries defined by the web part
  documentLibraries: Array<IDocumentLibrary>;

  // Contains a list of libraries the user has access to
  librariesWithPermissions: Array<string>;

  // Web URL that contains the lists the webpart is working with  
  webUrl: string;
}
