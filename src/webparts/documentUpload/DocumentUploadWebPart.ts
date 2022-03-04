import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PermissionKind } from '@pnp/sp/security';
import { IWeb, Web } from '@pnp/sp/webs';
import DocumentUpload from './components/DocumentUpload';
import { IDocumentUploadProps } from './components/IDocumentUploadProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { sp } from "@pnp/sp";
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/webs';

export interface IField {
  InternalName: string;
  SortOrder: number;
  Value: any;
}

export interface IDocumentLibrary {
  Title: string;
  ServerRelativeUrl: string;
}

export interface IFileInfo {
  path: string;
  lastModified: number;
  lastModifiedDate: Date;
  name: string;
  size: number;
  type: string;
  success: boolean;
}

export interface IDocumentUploadWebPartProps {
  fields: Array<IField>;
  webUrl: string;
  listId: string;
  documentLibraries: Array<IDocumentLibrary>;
}

export default class DocumentUploadWebPart extends BaseClientSideWebPart<IDocumentUploadWebPartProps> {

  // Hold all of the lists for the property pane drop down
  private lists: Array<any> = [];

  // Holds all the fields for the property pane drop down
  private fields: Array<any> = [];

  // Basic string array that contains a list of libraries that the user has access to
  private librariesWithPermissions: Array<string> = [];

  protected async onInit(): Promise<void> {
    await super.onInit().then(() => {

      /**
       * Setup SPFX context for PNP
       */
      sp.setup({
        spfxContext: this.context
      });

      /**
       * If no web url is present then set the web URL to the current site
       */
      if (!this.properties.webUrl) {
        this.properties.webUrl = this.context.pageContext.web.absoluteUrl;
      }


      /**
       * Check to see if the web part has a web URL configured
       */
      if (this.properties.webUrl) {

        /**
         * Get a list of SharePoint lists that area availale via the above 
         * web URL. This list is used to generate the desired metadata fields.
         */
        this.getLists();

        /**
         * Check to see if a list has been selected
         */
        if (this.properties.listId) {

          /**
           * Fetch all of the list fields for the selected list.
           */
          this.getFields();
        }


        /**
         * Check to see if the web part has been configured with one or more
         * document library destinations
         */
        if (this.properties.documentLibraries) {

          /**
           * Iterate each document library as the current user and check to 
           * see if the user has AddListItems permission. If they have permissions
           * we add the library to an array that allows the user to select it
           * when uploading docs via the UI.
           * 
           * Re-render the web part when we finish.
           */
          this.properties.documentLibraries.forEach(lib => {
            const web: IWeb = Web(this.properties.webUrl);
            web.lists.getByTitle(lib.Title).currentUserHasPermissions(PermissionKind.AddListItems)
              .then((hasPermission: boolean) => {
                // if (hasPermission) {
                //   this.librariesWithPermissions.push(lib.Title);
                // }
                  this.librariesWithPermissions.push(lib.Title);
                this.render();
              });
          });
        }
      }



      Promise.resolve();
    });
  }


  /**
   * Get SharePoint lists via the rest API
   */
  private async getLists(): Promise<void> {

    const web: IWeb = Web(this.properties.webUrl);
    this.lists = await web.lists.get();
  }

  /**
   * Get SharePoint list fields via the rest API
   */
  private async getFields(): Promise<void> {

    const web: IWeb = Web(this.properties.webUrl);
    this.fields = await web.lists.getById(this.properties.listId).fields.filter(this.properties.fields.map(fld => `(InternalName eq '${fld.InternalName}')`).join(' or ')).get();

    this.render();
  }

  public render(): void {
    const element: React.ReactElement<IDocumentUploadProps> = React.createElement(
      DocumentUpload,
      {
        fields: this.fields,
        documentLibraries: this.properties.documentLibraries,
        librariesWithPermissions: this.librariesWithPermissions,
        webUrl: this.properties.webUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * 
   * @param propertyPath property name
   * @param oldValue old vlaue of property
   * @param newValue new value of property
   * 
   * Ovrrride the onPropertyPaneFieldChanged method to fetch the lists if the web URL property poane 
   * field changes.
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'webUrl' && newValue && (oldValue !== newValue)) {
      this.getLists();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Document Library Connection",
              groupFields: [
                PropertyPaneTextField('webUrl', {
                  label: 'Web URL'
                }),
                PropertyPaneDropdown('listId', {
                  label: 'Select a List',
                  options: !this.lists ? [] : this.lists.map(l => {
                    const option = {
                      key: l.Id,
                      text: l.Title
                    };

                    return option;
                  })
                }),
                PropertyFieldCollectionData("fields", {
                  key: "fieldsData",
                  label: "Document Library Fields",
                  panelHeader: "Document Library Fields",
                  manageBtnLabel: "Document Library Fields",
                  value: this.properties.fields,
                  fields: [
                    {
                      id: "InternalName",
                      title: "Internal Field Name",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "SortOrder",
                      title: "Sort Order",
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false
                }),
                PropertyFieldCollectionData("documentLibraries", {
                  key: "documentLibrariesData",
                  label: "Configure Available Libraries",
                  panelHeader: "Configure Available Libraries",
                  manageBtnLabel: "Configure Available Libraries",
                  value: this.properties.documentLibraries,
                  fields: [
                    {
                      id: "Title",
                      title: "Document Library Title",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "ServerRelativeUrl",
                      title: "Server Relative URL",
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
