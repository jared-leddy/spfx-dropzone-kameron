import * as React from 'react';
import styles from './DocumentUpload.module.scss';
import { IDocumentUploadProps } from './IDocumentUploadProps';
import Dropzone from 'react-dropzone';
import { IDocumentLibrary, IField, IFileInfo } from '../DocumentUploadWebPart';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { ActivityItem } from 'office-ui-fabric-react/lib/ActivityItem';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Web } from '@pnp/sp/webs';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/webs';

export interface IDocumentUploadState {
  // Contains a list of files to be uploaded, populated by dropzone after a drop or manual upload
  filesToUpload: Array<IFileInfo>;

  // Contains the selected the library from the UI choice group
  selectedLibrary: IDocumentLibrary;

  // Contains a list of fields that the user must fill out (defined in property pane)
  fields: Array<IField>;

  // Loading indicator helper to show progress indicator when documents are being uploaded
  uploading: boolean;

  // A list of activity/log messages to show the user as documents are uploaded and updated
  messages: Array<ActivityItem>;
}

export default class DocumentUpload extends React.Component<IDocumentUploadProps, IDocumentUploadState> {

  constructor(props: IDocumentUploadProps) {
    super(props);

    this.state = {
      filesToUpload: [],
      selectedLibrary: null,
      fields: props.fields,
      uploading: false,
      messages: []
    };

    this.setFiles = this.setFiles.bind(this);
    this.uploadDocuments = this.uploadDocuments.bind(this);
    this.setFieldValue = this.setFieldValue.bind(this);
    this.setDocumentLibrary = this.setDocumentLibrary.bind(this);
  }

  /**
   * 
   * @param internalName Internal field name
   * @param value Value that the user selected via the UI
   * 
   * This method sets the metadata values as the user selects then 
   * via the UI form controls.
   */
  public setFieldValue(internalName: string, value: any): void {
    this.setState((prevState) => {
      return {
        fields: prevState.fields.map(fld => {
          if (fld.InternalName === internalName) {
            fld.Value = value;
          }

          return fld;
        })
      };
    });
  }

  /**
  * Returns image url for the given filename.
  * The urls points to https://spoprod-a.akamaihd.net.....
  */
  public GetImgUrl(fileName: string): string {
    const fileNameItems: any[] = !fileName ? ["file", "bad"] : fileName.split(".");
    const fileExtenstion: string = fileNameItems[fileNameItems.length - 1];

    return this.GetImgUrlByFileExtension(fileExtenstion);
  }

  /**
 * Returns image url for the given extension.
 * The urls points to https://spoprod-a.akamaihd.net.....
 */
  public GetImgUrlByFileExtension(extension: string): string {
    // cuurently in SPFx with React I didn't find different way of getting the image
    // feel free to improve this
    let imgRoot: string =
      "https://spoprod-a.akamaihd.net/files/odsp-next-prod_ship-2017-04-21-sts_20170503.001/odsp-media/images/filetypes/16/";
    let imgType: string = "genericfile.png";
    imgType = extension + ".png";

    switch (extension) {
      case "aspx":
      case "htm":
      case "html":
        imgRoot = 'https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20210115.001/assets/item-types/32/spo.svg';
        imgType = '';
        break;
      case "jpg":
      case "jpeg":
      case "jfif":
      case "gif":
      case "png":
        imgType = "photo.png";
        break;
      case "folder":
        imgType = "folder.svg";
        break;
    }
    return imgRoot + imgType;
  }


  /**
   * 
   * @param name Document Library display name
   * 
   * Sets the selected document library name that is used
   * when uploading the documents
   */
  public setDocumentLibrary(name: string): void {

    const { documentLibraries } = this.props;

    this.setState({
      selectedLibrary: documentLibraries.filter(d => d.Title === name)[0]
    });
  }

  /**
   * 
   * @param acceptedFiles File array frmo drop zone control
   * 
   * Sets the current working files based on the Drop/Upload of 
   * file usig the DropZone control
   */
  public setFiles(acceptedFiles): void {

    this.setState({
      filesToUpload: acceptedFiles
    });
  }

  /**
   * Clears all data for the form including files, controls
   * and any activity messages
   */
  public clearForm(): void {
    this.setState((prevState) => {

      return {
        filesToUpload: [],
        messages: [],
        fields: prevState.fields.map(fld => {
          fld.Value = null;

          return fld;
        })
      };
    });
  }

  /**
   * 
   * @param spField SharePoit field data. 
   * @returns Office UI fabric control corresponding to SharePoint field data type
   * 
   * This method will read the field type tha is defined for the form and render the appropriate control
   */
  public renderField(spField: any): JSX.Element {

    switch (spField.TypeAsString) {
      case "Text":
      case "Note":
      case "Number":
      case "Currency":
        return (
          <div>
            <TextField
              label={spField.Title}
              multiline={spField.TypeAsString === 'Note'}
              onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setFieldValue(spField.InternalName, newValue)}
            />;
            {
              spField.Description && <small>{spField.Description}</small>
            }
          </div>
        );
      case "DateTime":
        return (
          <div>
            <DatePicker
              label={spField.Title}
              allowTextInput
              value={spField.Value}
              ariaLabel="Select a date"
              onSelectDate={(date: Date) => this.setFieldValue(spField.InternalName, date)}
            />
            {
              spField.Description && <small>{spField.Description}</small>
            }
          </div>
        );
      case "Choice":
      case "MultiChoice":
        return (
          <div>
            <Dropdown
              label={spField.Title}
              multiSelect={spField.TypeAsString === 'MultiChoice'}
              onChange={(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => this.setFieldValue(spField.InternalName, option.key.toString())}
              options={!(spField as any).Choices ? [] : (spField as any).Choices.map(c => {
                const option: IDropdownOption = {
                  key: c,
                  text: c
                };

                return option;
              })}
            />
            {
              (spField as any).FillInChoice &&
              <TextField
                label={`${spField.Title} - Custom Value`}
                onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setFieldValue(spField.InternalName, newValue)}
              />
            }
            {
              spField.Description && <small>{spField.Description}</small>
            }
          </div>
        );
      case "Boolean":
        return (
          <div>
            <Toggle label={spField.Title} onText='Yes' offText='No' onChange={(event: React.MouseEvent<HTMLElement, MouseEvent>, checked?: boolean) => this.setFieldValue(spField.InternalName, checked)} />;
            {
              spField.Description && <small>{spField.Description}</small>
            }
          </div>
        );
      case "URL":
        return (
          <div>
            <TextField
              label="Url"
              onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setFieldValue(spField.InternalName, newValue)}
            />
            <TextField
              label='Description'
              onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.setFieldValue(spField.InternalName, newValue)}
            />
            {
              spField.Description && <small>{spField.Description}</small>
            }
          </div>
        );
      default:
        return <div>Field type {spField.TypeAsString} is not supported.</div>;
    }

  }

  /**
   * 
   * @returns True/False based on validation of form
   * 
   * The form must have 1 or moer files, all metadata/form fields completed and a library selected.
   */
  public validForm(): boolean {
    const { fields, filesToUpload, selectedLibrary } = this.state;

    return (filesToUpload.length > 0 && fields.filter(fld => !!fld.Value).length === fields.length && !!selectedLibrary);
  }

  /**
   * 
   * @param library Target document library
   * @param files List of files to upload
   * 
   * This method will first upload each of the files to the target document library.
   * Next, it will fetch the new file to obain the Item metdata (ID).
   * Using the ID it will update the metadata of the new item using the form fields
   * defined by the user.
   * 
   * NOTE: Overwriting files is disabled and items that have the same name will fail
   */
  public uploadDocuments(library: IDocumentLibrary, files: Array<IFileInfo>): void {

    const { webUrl } = this.props;
    const { fields } = this.state;
    const web = Web(webUrl);

    this.setState({
      uploading: true
    });

    files.forEach(f => {

      let messages = [];
      web.getFolderByServerRelativeUrl(library.ServerRelativeUrl).files.add(f.name, f, false)
        .then(() => {
          messages.push({
            activityDescription: [
              <div>Successfully <strong style={{ color: 'green', fontWeight: 700 }}>uploaded</strong> {f.name}.</div>
            ],
            activityIcon: <Icon iconName={'Up'} />,
            isCompact: true,
          });


          web.getFolderByServerRelativeUrl(library.ServerRelativeUrl).files.getByName(f.name).getItem()
            .then((item: any) => {

              let update = {};
              fields.forEach(fld => {
                update[fld.InternalName] = fld.Value;
              });

              web.lists.getByTitle(library.Title).items.getById(item.Id).update(update)
                .then(() => {
                  messages.push({
                    activityDescription: [
                      <div>Successfully <strong style={{ color: 'green', fontWeight: 700 }}>updated</strong> metadata for file {f.name}.</div>
                    ],
                    activityIcon: <Icon iconName={'Save'} />,
                    isCompact: true,
                  });

                  this.setState((prevState) => {
                    return {
                      filesToUpload: prevState.filesToUpload.map(file => {
                        if (file.name === f.name) {
                          file.success = true;
                        }

                        return file;
                      }),
                      uploading: prevState.filesToUpload.filter(file => file.success === true || file.success === false).length !== prevState.filesToUpload.length,
                      messages: [...prevState.messages, ...messages]
                    };
                  });
                });
            });

        })
        .catch((error: any) => {

          if (JSON.stringify(error.message).toLowerCase().indexOf('already exists') >= 0) {
            messages.push({
              activityDescription: [
                <div>Error <strong style={{ color: 'crimson', fontWeight: 700 }}>uploading</strong> {f.name}. The file already exists.</div>
              ],
              activityIcon: <Icon iconName={'ErrorBadge'} />,
              isCompact: true,
            });
          }

          this.setState((prevState) => {
            return {
              filesToUpload: prevState.filesToUpload.map(file => {
                if (file.name === f.name) {
                  file.success = false;
                }

                return file;
              }),
              uploading: prevState.filesToUpload.filter(file => file.success === true || file.success === false).length !== prevState.filesToUpload.length,
              messages: [...prevState.messages, ...messages]
            };
          });
        });
    });
  }

  /**
   * 
   * @param prevProps prev component props
   * 
   * Check to see if the fields defined by the web part have changed and update our 
   * state with the new field definitions
   */
  public componentDidUpdate(prevProps: Readonly<IDocumentUploadProps>): void {

    const { fields } = this.props;

    if (JSON.stringify(prevProps.fields) !== JSON.stringify(fields)) {
      this.setState({
        fields
      });
    }
  }


  /**
   * 
   * @returns The document upload user interface
   */
  public render(): React.ReactElement<IDocumentUploadProps> {

    const {
      filesToUpload,
      fields,
      messages,
      selectedLibrary,
      uploading,
    } = this.state;

    const {
      documentLibraries,
      librariesWithPermissions
    } = this.props;

    return (
      <div className={styles.documentUpload}>
        <div className={styles.dropZone}>
          <Dropzone onDrop={acceptedFiles => this.setFiles(acceptedFiles)} multiple noDragEventsBubbling>
            {({ getRootProps, getInputProps }) => (
              <section>
                <div {...getRootProps()}>
                  <input {...getInputProps()} />
                  <div className={styles.zones}>
                    <div className={styles.dnd}>
                      {
                        filesToUpload && filesToUpload.length === 0 &&
                        <p>Drag 'n' drop files here, or click to select files</p>
                      }
                      {
                        filesToUpload && filesToUpload.length > 0 &&
                        <div>
                          <aside>
                            <div className={styles.files}>
                              {
                                filesToUpload.map(f => {
                                  return (
                                    <div className={styles.pill} style={{ border: f.success === true ? '1px solid green' : f.success === false ? '1px solid crimson' : null }}>
                                      <div>
                                        <img width={16} src={this.GetImgUrl(f.name)} /> {f.name}
                                        {
                                          (f.success === true || f.success === false) &&
                                          <Icon iconName={f.success === true ? 'CheckMark' : f.success === false ? 'ErrorBadge' : null}
                                            style={{ color: f.success === true ? 'green' : f.success === false ? 'crimson' : null }}
                                          />
                                        }
                                      </div>
                                    </div>
                                  );
                                })
                              }
                            </div>
                          </aside>
                        </div>
                      }
                    </div>

                  </div>
                </div>
              </section>
            )}
          </Dropzone>
        </div>
        {
          filesToUpload.length > 0 &&
          <div>
            <div className={styles.fields}>
              {
                fields.map(fld => {
                  return this.renderField(fld);
                })
              }
            </div>
            <div className={styles.libraries}>
              <div>
                <ChoiceGroup onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => this.setDocumentLibrary(option.key.toString())} label="Pick a library" options={!documentLibraries ? [] : documentLibraries.filter(lib => librariesWithPermissions.indexOf(lib.Title) >= 0).map(lib => {
                  const option = {
                    key: lib.Title,
                    text: lib.Title,
                  };

                  return option;
                })} />
              </div>
              <div className={styles.messages}>

                {
                  messages && messages.length > 0 &&
                  <div>
                    <Label>Activity Log</Label>
                    {
                      messages.map((item, index) => <ActivityItem {...item} key={index} />)
                    }
                  </div>
                }
              </div>
            </div>
            <div className={styles.submit}>
              {
                uploading &&
                <ProgressIndicator label="Uploading documents" description={`Uploading ${filesToUpload.length} documents (${Math.round((filesToUpload.map(f => f.size).reduce((partialSum, a) => partialSum + a, 0)) / 1024)}KB)`} />
              }
              <div className={styles.buttons}>
                <DefaultButton text='Reset' onClick={() => this.clearForm()} iconProps={{ iconName: "ClearFormattingEraser" }} />
                <PrimaryButton disabled={!this.validForm()} text={`Submit (${filesToUpload.filter(file => !file.success).length})`} iconProps={{ iconName: "Up" }} onClick={() => this.uploadDocuments(selectedLibrary, filesToUpload)} />
              </div>
            </div>
          </div>
        }
      </div>
    );
  }
}
