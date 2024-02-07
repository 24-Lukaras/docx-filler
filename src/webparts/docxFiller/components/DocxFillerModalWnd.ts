import { Guid } from "@microsoft/sp-core-library";
import { IHtmlComponent } from "./IHtmlComponent";
import * as jq from "jquery";
import styles from "./DocxFillerModalWnd.module.scss";
import { SPService } from "../../../shared/SPService";
import { DocxTokenFiller } from "../../../shared/DocxTokenFiller";
import { IDocxFillerWebPartProps } from "../DocxFillerWebPart";
import { saveAs } from "file-saver";

export class DocxFillerModalWnd implements IHtmlComponent
{

    private parentGuid: Guid;
    private guid: Guid;
    private listGuid: string;
    private templateLibrary: string;
    private opened: boolean;
    private service: SPService;
    private templateId: number;
    private props: IDocxFillerWebPartProps;
    private processing: boolean;

    constructor(guid: Guid, service: SPService, listGuid: string, templateLibrary: string, props: IDocxFillerWebPartProps)
    {
        this.guid = Guid.newGuid();
        this.parentGuid = guid;
        this.service = service;
        this.listGuid = listGuid;
        this.templateLibrary = templateLibrary;
        this.props = props;
    }

    public open(templateId: number)
    {
        if (!this.opened) {
            jq("." + styles.modalBg).css("display", "block");
            jq("input#search").trigger("change");
            this.templateId = templateId;
            this.opened = true;
        }
    }

    public close()
    {
        if (this.opened) {
            jq("." + styles.modalBg).css("display", "none");
            this.opened = false;
        }
    }

    public registerClose()
    {
        jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] .${styles.modalClose}`).on("click", () => { this.close() });
        jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] input#search`).on("change", (event) => this.populateItems(event));
    }

    private async populateItems(event: JQuery.ChangeEvent): Promise<void>
    {
        const tBody = jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] table tbody`);
        tBody.empty();
        tBody.append("<div>Searching...</div>");
        this.service.getListItemsContaining(this.listGuid, this.props.filterField, event.target.value).then((items) =>
        {
            const btn = this.getItemButtons();
            const itemsHtml = items.map((item) =>
            {
                return `<tr itemId='${item["ID"]}'>
                <td>${item["Title"]}</td>
                ${btn}
                </tr>`;
            }).join("");
            tBody.empty();
            tBody.append(itemsHtml);
            jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] table tbody a.docx-download`).on("click", (event) => { this.tryDownload(parseInt(jq(event.target).closest("tr").attr("itemId") ?? "")).catch((error) => { this.setProcessingError(error); }); });
            jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] table tbody a.docx-attachment`).on("click", (event) => { this.tryAttachment(parseInt(jq(event.target).closest("tr").attr("itemId") ?? "")).catch((error) => { this.setProcessingError(error); }); });
            jq(`div[wpId="${this.parentGuid}"][wndId="${this.guid}"] table tbody a.docx-file`).on("click", (event) => { this.trySpFile(parseInt(jq(event.target).closest("tr").attr("itemId") ?? "")).catch((error) => { this.setProcessingError(error); }); });
        });
    }

    private getItemButtons(): string
    {
        switch (this.props.exportType) {

            case "download":
                return `<td><a class="docx-download ${styles.docxAction}"><i class="ms-Icon ms-Icon--Download" aria-hidden="true"></i></a></td>`;

            case "attachment":
                return `<td><a class="docx-attachment ${styles.docxAction}"><i class="ms-Icon ms-Icon--Link12" aria-hidden="true"></i></a></td>`;

            case "sp_file":
                return `<td><a class="docx-file ${styles.docxAction}"><i class="ms-Icon ms-Icon--OpenFile" aria-hidden="true"></i></a></td>`;

            default:
                return `<td><a class="docx-download ${styles.docxAction}"><i class="ms-Icon ms-Icon--Download" aria-hidden="true"></i></a></td>
                <td><a class="docx-attachment ${styles.docxAction}"><i class="ms-Icon ms-Icon--Link12" aria-hidden="true"></i></a></td>
                <td><a class="docx-file ${styles.docxAction}"><i class="ms-Icon ms-Icon--OpenFile" aria-hidden="true"></i></a></td>`;
        }
    }

    private async tryDownload(itemId: number) : Promise<void>
    {

        if (this.processing) return;

        this.setProcessing(true);
        const result = await this.getFilledDocument(itemId);
        saveAs(result.file, result.fileName);
        this.setProcessing(false);
    }

    private async tryAttachment(itemId: number) : Promise<void>
    {
        if (this.processing) return;

        this.setProcessing(true);
        const result = await this.getFilledDocument(itemId);
        await this.service.uploadAttachment(this.listGuid, itemId, result.fileName, result.file);
        this.setProcessing(false);
    }

    private async trySpFile(itemId: number) : Promise<void>
    {

        if (this.processing) return;

        this.setProcessing(true);
        const result = await this.getFilledDocument(itemId);
        await this.service.uploadFile(result.filePath, result.file);
        this.setProcessing(false);
    }

    private setProcessing(value: boolean)
    {
        this.processing = value;
        const table = jq(`div[wpId='${this.parentGuid}'][wndId='${this.guid}'] .${styles.itemsTable}`);
        const row = jq(`div[wpId='${this.parentGuid}'][wndId='${this.guid}'] .${styles.processingRow} .${styles.rowCenter}`);
        if (value) {
            table.addClass(styles.itemsTableProcessing);
            row.removeClass(styles.processingInit);
            row.removeClass(styles.processingFinished);
            row.removeClass(styles.processingError);
            row.addClass(styles.processingProgress);
        }
        else {
            table.removeClass(styles.itemsTableProcessing);
            row.removeClass(styles.processingInit);
            row.removeClass(styles.processingProgress);
            row.removeClass(styles.processingError);
            row.addClass(styles.processingFinished);
        }
    }

    private setProcessingError(exception: any)
    {
        const table = jq(`div[wpId='${this.parentGuid}'][wndId='${this.guid}'] .${styles.itemsTable}`);
        const row = jq(`div[wpId='${this.parentGuid}'][wndId='${this.guid}'] .${styles.processingRow} .${styles.rowCenter}`);
        table.removeClass(styles.itemsTableProcessing);
        row.removeClass(styles.processingInit);
        row.removeClass(styles.processingProgress);
        row.removeClass(styles.processingFinished);
        row.addClass(styles.processingError);
        this.processing = false;
    }

    private async getFilledDocument(itemId: number): Promise<ExportProperties>
    {
        const fields = await this.service.getListFields(this.listGuid);
        const lookupFields = fields.filter((x) => x.TypeAsString == "Lookup" && !x.Hidden && !x.FromBaseType).map((x) => x.InternalName);
        const item = await this.service.getListItem(this.listGuid, itemId);
        const lookups = await this.service.getListItemLookups(this.listGuid, itemId, lookupFields);
        lookupFields.forEach((x) =>
        {
            item[x] = lookups[x]["Title"];
        });
        const doc = await this.service.getDocument(this.templateLibrary, this.templateId);
        const docxFiller = new DocxTokenFiller(this.props);
        await docxFiller.loadDocument(doc);
        await docxFiller.replace(item, fields, this.props.useDisplayFields);
        const file = await docxFiller.export();

        return {
            fileName: this.replaceTokens(this.props.exportFilename, "item", item),
            filePath: this.replaceTokens(this.props.exportPath, "item", item),
            file: file,
        };
    }

    private replaceTokens(original: string, prefix: string, item: any): string
    {
        let result = original;

        if (!original) {
            return result;
        }

        const keys = Object.keys(item);

        for (let i = 0; i < keys.length; i++) {
            result = result.replace(new RegExp("{" + prefix + ":" + keys[i] + "}"), item[keys[i]]);
        }

        return result;
    }

    render(): string
    {
        return `        
        <div wpId='${this.parentGuid}' wndId='${this.guid}' class='${styles.modalBg}'>
            <div class="${styles.modalContent}">
                <div class="${styles.content}">
                    <div class="${styles.row}">
                        <span class="${styles.modalClose}">&times;</span>
                    </div>
                    <div class="${styles.rowCenter}">
                        <label>Search:</label>&nbsp;<input id="search" />
                    </div>
                    <div class="${styles.row}" style="min-height: 300px">
                        <table class="${styles.itemsTable}">
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="${styles.row}">
                        <div class="${styles.processingRow}">
                            <div class="${styles.rowCenter} ${styles.processingInit}">
                                <span class="${styles.processingInit}">Choose an action...</span>
                                <span class="${styles.processingProgress}">Processing...</span>
                                <span class="${styles.processingFinished}">Finished</span>
                                <span class="${styles.processingError}">An error occured...</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>`;
    }

}

class ExportProperties
{
    public fileName: string;
    public filePath: string;
    public file: Blob;
}