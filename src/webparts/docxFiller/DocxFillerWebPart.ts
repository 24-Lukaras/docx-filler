import { Guid, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './DocxFillerWebPart.module.scss';
import { IPropertyPaneConfiguration, IPropertyPaneDropdownOption, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
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

  displayStyle: string,

  templateLibrary: string,
  targetList: string,
  displayFields: string[],
  tokenStyle: string

  exportType: string
}

export default class DocxFillerWebPart extends BaseClientSideWebPart<IDocxFillerWebPartProps> {

  private service: SPService;
  private guid: Guid;

  public async render(): Promise<void> {
    this.guid = Guid.newGuid();

    await this.fillDocumentLibraries();
    await this.fillLists();

    const configured = this.properties.displayStyle && this.properties.templateLibrary && this.properties.tokenStyle;
    if (!configured){
      this.domElement.innerHTML = `<div>This web part was not configured properly. Please update configuration of this web part.</div>`;
    }
    else {
      let components: IHtmlComponent[] = [];
      
      const modal = new DocxFillerModalWnd(this.guid, this.service, this.properties.targetList, this.properties.templateLibrary);
      components.push(modal);      

      const items = await this.service.getAllDocuments(this.properties.templateLibrary, this.properties.displayFields);

      switch (this.properties.displayStyle) {

        default:
          components.push(new DocxFillerTiles(items));
          break;

      }
      
      this.domElement.innerHTML = `<div class='${styles.docxFiller}' wpId='${this.guid}'>${components.map((x) => {return x.render()}).join("")}</div>`;
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


  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      this.service = new SPService(this.context);
    });
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
                PropertyPaneDropdown("tokenStyle", {
                  label: "Token style",
                  options: [
                    {
                      key: "{token}",
                      text: "{Token}"
                    },
                    {
                      key: "[token]",
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
              ]
            },
            {
              groupName: "Export settings",
              groupFields: [
                PropertyPaneDropdown("exportType", {
                  label: "Export type",
                  options: [
                    {
                      key: "download",
                      text: "Download"
                    },
                  ],
                  selectedKey: this.properties.exportType
                }),
              ]
            },
          ]
        }
      ]
    };
  }  
}