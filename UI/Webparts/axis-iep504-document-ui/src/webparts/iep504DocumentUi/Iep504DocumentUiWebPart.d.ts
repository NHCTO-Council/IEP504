import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from "@microsoft/sp-webpart-base";
export interface IIep504DocumentUiWebPartProps {
    docLibraryName: string;
    auditListName: string;
    uiMode: string;
    treatPreviewAsRead: boolean;
    allowExport: boolean;
    debugMode: boolean;
    authorLink: string;
}
export default class Iep504DocumentUiWebPart extends BaseClientSideWebPart<IIep504DocumentUiWebPartProps> {
    render(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    protected readonly disableReactivePropertyChanges: boolean;
}
