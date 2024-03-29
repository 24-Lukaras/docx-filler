import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFieldInfo, IListInfo, ISiteUser, spfi, SPFI, SPFx } from '@pnp/sp/presets/all';

export class SPService {        

    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    public async getCurrentUser(): Promise<ISiteUser> {
        const user = await this.sp.web.currentUser;
        return user;
    }

    public async getSiteLists(): Promise<IListInfo[]> {
        const lists = await this.sp.web.lists();
        return lists;
    }

    public async getListFields(listId: string): Promise<IFieldInfo[]> {
        const fields = await this.sp.web.lists.getById(listId).fields();
        return fields;
    }

    public async getAllListItems(listId: string): Promise<any[]> {
        const items = await this.sp.web.lists.getById(listId).items();
        return items;
    }

    public async getListItemsContaining(listId: string, field: string, filter: string): Promise<any[]> {
        if (!filter || filter.length == 0) {
            return await this.sp.web.lists.getById(listId).items.top(10)();
        }
        const items = await this.sp.web.lists.getById(listId).getItemsByCAMLQuery({
            ViewXml: `<View><Query><Where><Contains><FieldRef Name="${field}" /><Value Type="Text">${filter}</Value></Contains></Where></Query></View>`
        });
        return items;
    }

    public async getListItem(listId: string, itemId: number): Promise<any> {
        const item = await this.sp.web.lists.getById(listId).items.getById(itemId)();
        return item;
    }

    public async getListItemLookups(listId: string, itemId: number, expands: string[]): Promise<any> {
        const item = await this.sp.web.lists.getById(listId).items.getById(itemId).select(expands.map((x) => x + "/Title").join(",")).expand(expands.join(","))();
        return item;
    }

    public async getAllDocuments(listId: string, fields: string[]): Promise<any[]> {
        const items = await this.sp.web.lists.getById(listId).items.select("ID", "FileLeafRef")();
        return items;
    }

    public async getDocument(listId: string, id: number): Promise<Blob> {
        const file = await this.sp.web.lists.getById(listId).items.getById(id).file.getBlob();
        return file;
    }

    public async uploadFile(filePath: string, file: Blob) : Promise<void> {
        const pathArr = filePath.split("/");
        const filename = pathArr.pop() ?? "404";
        const folderpath = pathArr.join("/");
        await this.sp.web.getFolderByServerRelativePath(folderpath).files.addUsingPath(filename, file, {Overwrite: true});
    }

    public async uploadAttachment(listId: string, itemId: number, filename: string, file: Blob) : Promise<void> {
        await this.sp.web.lists.getById(listId).items.getById(itemId).attachmentFiles.add(filename, file);
    }
}