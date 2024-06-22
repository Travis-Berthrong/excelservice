export interface IExcelSheet {
    "@odata.context": string;
    "@odata.id": string;
    id: string;
    name: string;
    position: number;
    visibility: string;
    tables?: Array<{ id: string, name: string}>
}


