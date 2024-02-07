import { Guid, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './DocxFillerWebPart.module.scss';
import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption, IPropertyPaneField, PropertyPaneCheckbox, PropertyPaneDropdown, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
//import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls';
import { SPService } from '../../shared/SPService';
import { IHtmlComponent } from './components/IHtmlComponent';
import { DocxFillerTiles } from './components/DocxFillerTiles';
import { DocxFillerModalWnd } from './components/DocxFillerModalWnd';
import * as jq from "jquery";

export interface IDocxFillerWebPartProps {
  
  libraryFields: IPropertyPaneDropdownOption[],
  documentLibraries: IPropertyPaneDropdownOption[],
  lists: IPropertyPaneDropdownOption[],
  targetListFields: IPropertyPaneDropdownOption[],

  displayStyle: string,

  templateLibrary: string,
  targetList: string,
  filterField: string,
  displayFields: string[],
  useDisplayFields: boolean,
  tokenStyle: string

  exportType: string,
  exportFormat: string,

  exportFilename: string,
  exportPath: string,
}

export default class DocxFillerWebPart extends BaseClientSideWebPart<IDocxFillerWebPartProps> {

  private service: SPService;
  private guid: Guid;

  public async render(): Promise<void> {
    this.guid = Guid.newGuid();

    await this.fillDocumentLibraries();
    await this.fillLists();
    await this.fillTargetListFields();

    const configured = this.properties.displayStyle && this.properties.templateLibrary && this.properties.tokenStyle;
    if (!configured){
      this.domElement.innerHTML = `<div>This web part was not configured properly. Please update configuration of this web part.</div>`;
    }
    else {
      let components: IHtmlComponent[] = [];
      
      const modal = new DocxFillerModalWnd(this.guid, this.service, this.properties.targetList, this.properties.templateLibrary, this.properties);
      components.push(modal);      

      const items = (await this.service.getAllDocuments(this.properties.templateLibrary, this.properties.displayFields))
        .filter((x) => x["FileLeafRef"].toLocaleLowerCase().endsWith(".docx")); //docx files filtration

      switch (this.properties.displayStyle) {
        default:
          components.push(new DocxFillerTiles(items));
          break;
      }
      
      this.domElement.innerHTML = `<link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css" />
      <div class='${styles.docxFiller}' wpId='${this.guid}'>
        ${components.map((x) => {return x.render()}).join("")}
      </div>`;
      this.registerButtons(modal);
    }
  }
  
  private registerButtons(modal: DocxFillerModalWnd) {
    jq(`div[wpId='${this.guid}'] .${styles.templateClickable}`).on("click", (event) => { modal.open(parseInt(jq(event.target).attr("templateId") ?? ""));});
    modal.registerClose();
  }    

  private async fillDocumentLibraries(): Promise<void> {    
    this.properties.documentLibraries = (await this.service.getSiteLists())
      .filter((list) => {return list.BaseTemplate == 101 && !list.Hidden})
      .map((list) => {return {key: list.Id, text: list.Title}})
  }

  private async fillLists(): Promise<void> {    
    this.properties.lists = (await this.service.getSiteLists())
      .filter((list) => {return !list.Hidden})
      .map((list) => {return {key: list.Id, text: list.Title}})
  }

  private async fillTargetListFields(): Promise<void> {
    if (this.properties.targetList) {
      this.properties.targetListFields = (await this.service.getListFields(this.properties.targetList))
        .filter((f) => { return !f.Hidden; })
        .map((f) => { return {key: f.InternalName, text: f.Title};});
    }
  }

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      this.service = new SPService(this.context);
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {


    let exportFields: IPropertyPaneField<any>[] = [PropertyPaneDropdown("exportType", {
      label: "Export type",
      options: [
        {
          key: "download",
          text: "Download"
        },
        {
          key: "attachment",
          text: "Attachment"
        },
        {
          key: "sp_file",
          text: "Sharepoint File"
        },
        {
          key: "*",
          text: "Any"
        },
      ],
      selectedKey: this.properties.exportType
    }),
    PropertyPaneDropdown("exportFormat", {
      label: "Export format",
      options: [
        {
          key: ".docx",
          text: ".docx"
        },
      ],
      selectedKey: this.properties.exportFormat
    })];

    switch (this.properties.exportType) {

      case "*":
        exportFields.push(PropertyPaneTextField("exportFilename", {
          label: "Filename",
          value: this.properties.exportFilename
        }));
        exportFields.push(PropertyPaneTextField("exportPath", {
          label: "Path",
          value: this.properties.exportPath
        }));
        break;

      case "sp_file":
        exportFields.push(PropertyPaneTextField("exportPath", {
          label: "Path",
          value: this.properties.exportPath
        }));
        break;

      default:
      case "attachment":
        exportFields.push(PropertyPaneTextField("exportFilename", {
          label: "Filename",
          value: this.properties.exportFilename
        }));
        break;
    }
    

    return {
      pages: [
        {                    
          groups: [
            {
              groupName: "Style",
              groupFields: [
                PropertyPaneDropdown("displayStyle", {
                  label: "Display style",
                  options: [
                    {
                      key: "tiles",
                      text: "Tiles"
                    },
                  ],
                  selectedKey: this.properties.displayStyle
                }),
              ]
            },
            {
              groupName: "Library settings",
              groupFields: [
                PropertyPaneDropdown("templateLibrary", {
                  label: "Template library",
                  options: this.properties.documentLibraries,
                  selectedKey: this.properties.templateLibrary                
                }),
                PropertyPaneDropdown("targetList", {
                  label: "Target list",
                  options: this.properties.lists,
                  selectedKey: this.properties.targetList                
                }),
                PropertyPaneDropdown("filterField", {
                  label: "List field",
                  options: this.properties.targetListFields,
                  selectedKey: this.properties.filterField                
                }),
                PropertyPaneDropdown("tokenStyle", {
                  label: "Token style",
                  options: [
                    {
                      key: "{token}",
                      text: "{Token}"
                    },
                    {
                      key: "\\[token\\]",
                      text: "[Token]"
                    },
                    {
                      key: "<token>",
                      text: "<Token>"
                    },
                    {
                      key: "_token_",
                      text: "_Token_"
                    },
                    {
                      key: "*",
                      text: "Any"
                    },
                  ],
                  selectedKey: this.properties.tokenStyle
                }),
                PropertyPaneLabel("useDisplayFieldsLabel", {
                  text: "Use field display names",
                }),
                PropertyPaneCheckbox("useDisplayFields", {
                  checked: this.properties.useDisplayFields,
                }),      
              ]
            },
            {
              groupName: "Export settings",
              groupFields: exportFields
            },
          ]
        }
      ]
    };
  }  
}
