import { IIep504DocumentUiWebPartProps } from "../../webparts/iep504DocumentUi/Iep504DocumentUiWebPart";
export declare class SPData {
    webPartProperties: IIep504DocumentUiWebPartProps;
    siteUrl: String;
    constructor(webPartProperties: IIep504DocumentUiWebPartProps, siteUrl: String);
    GetUserById(id: number): Promise<any>;
    GetFirstAccessedDate(spFileName: string, spFileModifiedDate: string, spUserId: number): Promise<any>;
}
