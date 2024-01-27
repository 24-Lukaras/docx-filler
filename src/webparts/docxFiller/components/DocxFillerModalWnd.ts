import { Guid } from "@microsoft/sp-core-library";
import { IHtmlComponent } from "./IHtmlComponent";
import * as jq from "jquery";
import styles from "./DocxFillerModalWnd.module.scss";
import { SPService } from "../../../shared/SPService";
import { DocxTokenFiller } from "../../../shared/DocxTokenFiller";
import { IDocxFillerWebPartProps } from "../DocxFillerWebPart";
import { saveAs } from "file-saver";

export class DocxFillerModalWnd implements IHtmlComponent{

    private parentGuid: Guid;
    private guid: Guid;
    private listGuid: string;
    private templateLibrary: string;
    private opened: boolean;
    private service: SPService;
    private templateId: number;
    private props: IDocxFillerWebPartProps;

    constructor (guid: Guid, service: SPService, listGuid: string, templateLibrary: string, props: IDocxFillerWebPartProps){
        this.guid = Guid.newGuid();
        this.parentGuid = guid;
        this.service = service;
        this.listGuid = listGuid;
        this.templateLibrary = templateLibrary;
        this.props = props;
    }
    
    public open(templateId: number) {
        if (!this.opened){
            jq("." + styles.modalBg).css("display", "block");
            jq("input#search").trigger("change");
            this.templateId = templateId;
            this.opened = true;
        }
    }

    public close() {
        if (this.opened) {
            jq("." + styles.modalBg).css("display", "none");
            this.opened = false;
        }
    }

    public registerClose() {
        jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] .${styles.modalClose}`).on("click", () => {this.close()});
        jq("input#search").on("change", () => this.populateItems());
    }

    private async populateItems(): Promise<void> {
        this.service.getAllListItems(this.listGuid).then((items) => {
            const itemsHtml = items.map((item) => {
                return `<tr><td>${item["Title"]}</td><td><a itemId='${item["ID"]}'>Download</a></td></tr>`;
            }).join("");
            jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] table tbody`).append(itemsHtml);
            jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] table tbody a`).on("click", (event) => { this.tryDownload(parseInt(jq(event.target).attr("itemId") ?? "")); });
        });
    }

    private async tryDownload(itemId: number) {
        const fields = await this.service.getListFields(this.listGuid);
        const lookupFields = fields.filter((x) => x.TypeAsString == "Lookup" && !x.Hidden && !x.FromBaseType).map((x) => x.InternalName);
        const item = await this.service.getListItem(this.listGuid, itemId);
        const lookups = await this.service.getListItemLookups(this.listGuid, itemId, lookupFields);
        lookupFields.forEach((x) => {
            item[x] = lookups[x]["Title"];
        });
        const doc = await this.service.getDocument(this.templateLibrary, this.templateId);
        const docxFiller = new DocxTokenFiller(this.props);
        await docxFiller.loadDocument(doc);
        await docxFiller.replace(item, fields, this.props.useDisplayFields);

        const result = await docxFiller.export();

        switch (this.props.exportType) {

            case "attachment":
                await this.service.uploadAttachment(this.listGuid, itemId, this.props.exportFilename, result);
                break;

            default:
                saveAs(result, this.props.exportFilename);
                break;
        }
    }

    render(): string
    {
        return `<div wpId='${this.parentGuid}' wndId='${this.guid}' class='${styles.modalBg}'>
            <div class="${styles.modalContent}">
                <div class="${styles.content}">
                    <div class="${styles.row}">
                        <span class="${styles.modalClose}">&times;</span>
                    </div>
                    <div class="${styles.rowCenter}">
                        <label>Search:</label>&nbsp;<input id="search" />
                    </div>
                    <div class="${styles.row}">
                        <table>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>`;
    }

}