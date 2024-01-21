import * as zip from 'jszip';
import { saveAs } from 'file-saver';
import { IFieldInfo } from '@pnp/sp/fields';

export class DocxTokenFiller {
    
    private zipFile: zip;
    private sourceDocumentString: string;

    constructor () {        
    }

    public async loadDocument(document: Blob): Promise<void> {
        const loader = new zip();
        const zipFile = await loader.loadAsync(document);
        this.zipFile = zipFile;
        this.sourceDocumentString = await this.zipFile.files['word/document.xml'].async('text');
    }

    public replace(item: any, fields: IFieldInfo[], useDisplayNames: boolean) {
        let xmlDocument = new DOMParser().parseFromString(this.sourceDocumentString, "application/xml");

        let coll = Array.from(xmlDocument.getElementsByTagName("w:r")).reduce((previous, currentItem) => {
            const group = <Element>currentItem.parentNode;
            let obj = previous.filter((x) => x["parent"] == group)[0];
            if (!obj) {
                obj = {
                    parent: group,
                    items: [] as Element[]
                };
                previous.push(obj);
            }
            obj.items.push(currentItem);
            return previous;
        }, [] as any[]);    

        for (let i = 0; i < coll.length; i++) {
            const wRunCollItem = coll[i];
            const wRunColl = new WRunCollection(wRunCollItem["parent"] as Element, wRunCollItem["items"] as Element[]);

            const text = wRunColl.getCollectionText();
            fields.forEach((field) => {

                const key = useDisplayNames ? field.Title : field.InternalName;
                const regex = new RegExp(`{${key}}`);
                const match = regex.exec(text);

                if (match){
                    const value = this.getFieldString(item, field);

                    wRunColl.split([match.index, match.index + match[0].length]);
                    wRunColl.mergeReplace(match.index,  match.index + match[0].length, value);
                }

            });
        }

        const wordFolder = this.zipFile.folder("word");
        wordFolder?.file("document.xml", xmlDocument.documentElement.innerHTML);

        this.zipFile.generateAsync({type: "blob"}).then((content) => {
            saveAs(content, "test.docx");
        });
    }

    private getFieldString(item: any, field: IFieldInfo) : string {
        const val = item[field.InternalName];

        if (val == null || val == undefined){
            return "";
        }

        switch (field.TypeAsString) {      
            
            case "Note":
                return val.toString().replace(/<\/?[^>]+(>|$)/g, "");

            case "DateTime":
                return new Date(val).toLocaleString();

            default:
                return val.toString();

        }
    }

}

class WRunCollection {

    public parent: Element;
    public nodes: WRun[];

    constructor (node: Element, runNodes: Element[] | null) {

        this.parent = node;

        const childNodes = Array.from(node.childNodes).map((x) => <Element>x);
        if (!runNodes) {
            runNodes = childNodes.filter((x) => x.nodeName == "w:r");
        }

        let collection: WRun[] = [];

        let index = 0;
        for (let i = 0; i < runNodes.length; i++) {
            const node = runNodes[i];
            const wRun = new WRun(node);
            wRun.collectionIndex = index;
            index += wRun.text.length;
            collection.push(wRun);
        }
        this.nodes = collection;
    }

    public split(indexes: number[]) {

        let newRuns: NumDictionary<WRun[]> = {};

        for (let i = 0; i < this.nodes.length; i++) {

            const item = this.nodes[i];

            const currentIndexes = indexes.filter((x) => x > item.collectionIndex && x < item.collectionIndex + item.text.length).map((x) => x - item.collectionIndex);

            if (currentIndexes.length > 0) {
                const newRunColl = item.split(currentIndexes);
                newRuns[i + 1] = newRunColl;
            }            
        }


        Object.keys(newRuns).forEach((index) => {
            const intIndex = parseInt(index);
            newRuns[intIndex].forEach((run) => {
                this.nodes.splice(intIndex, 0, run);
            });
        });

    }

    public mergeReplace(start: number, end: number, value: any) {
        const mergeRuns = this.nodes.filter((x) => x.collectionIndex >= start && x.collectionIndex + x.text.length <= end).sort((x) => x.collectionIndex);

        const text: string = value.toString();

        const firstRun = mergeRuns.splice(0, 1)[0];
        const runsToDelete = mergeRuns;

        const textNode = Array.from(firstRun.node.childNodes).filter((x) => x.nodeName == "w:t")[0];
        textNode.textContent = text;
        firstRun.text = text;

        runsToDelete.forEach((x) => {const index = this.nodes.indexOf(x); this.nodes = this.nodes.splice(index, 1);});
    }

    public getCollectionText() {
        return this.nodes.map((x) => x.text).join("");
    }
}

interface NumDictionary<T> {
    [Key: number]: T;
}

class WRun {

    public text: string;
    public properties: WRunProperty[];
    public node: Element;
    public collectionIndex: number;

    constructor (node: Element) {

        const childNodes = Array.from(node.childNodes);
        const textNodes = childNodes.filter((x) => x.nodeName == "w:t");
        const propertyParents = childNodes.filter((x) => x.nodeName == "w:rPr");

        this.node = node;
        this.text = Array.from(textNodes).map((x) => x.textContent).join("");

        this.properties = propertyParents.map((parent) => {
            return Array.from(parent.childNodes).map((x) => new WRunProperty(<Element>x));
        }).reduce((prev, next) => prev.concat(next), []);
    }

    public split(indexes: number[]): WRun[] {

        const orderedIndexes = indexes.sort().reverse();

        let result: WRun[] = [];

        let current = this.text;
        for (let i = 0; i < orderedIndexes.length; i++){
            const index = orderedIndexes[i];
            const left = current.slice(0, index);
            const right = current.slice(index);
            const newNode = this.node.cloneNode(true);
            const textNode = <Element>Array.from(newNode.childNodes).filter((x) => x.nodeName == "w:t")[0];
            textNode.textContent = right;
            if (/\s/.test(right.charAt(0)) || /\s/.test(right.charAt(right.length - 1))) {
                textNode.setAttribute("xml:space", "preserve");
            }
            this.node.after(newNode);
            const wRun = new WRun(<Element>newNode);
            wRun.collectionIndex = this.collectionIndex + index;
            result.unshift(wRun);
            current = left;
        }

        const currentTextNode = <Element>Array.from(this.node.childNodes).filter((x) => x.nodeName == "w:t")[0];
        currentTextNode.textContent = current;
        if (/\s/.test(current.charAt(0)) || /\s/.test(current.charAt(current.length - 1))) {
            currentTextNode.setAttribute("xml:space", "preserve");
        }
        this.text = current;

        return result;
    }
    
}

class WRunProperty {

    public name: string;
    public attributes: string[];

    constructor (node: Element) {
        this.name = node.nodeName;
        this.attributes = Array.from(node.attributes).map((x) => x.value);
    }
}