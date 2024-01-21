import { IHtmlComponent } from "./IHtmlComponent";
import styles from "./DocxFillerTiles.module.scss";
import root from "../DocxFillerWebPart.module.scss";

export class DocxFillerTiles implements IHtmlComponent {

    private templateItems: any[];

    constructor (templateItems: any[]) {
        this.templateItems = templateItems;
    }

    render(): string
    {        
        return `<div class='${styles.docxTiles}'>${this.templateItems.map((x) => {return this.renderTile(x);}).join("")}</div>`;
    }

    private renderTile(item: any): string {
        return `<div templateId='${item["ID"]}' class='${styles.docxTile} ${root.templateClickable}'>
            ${item["FileLeafRef"]}
            <hr>            
        </div>`;
    }

}