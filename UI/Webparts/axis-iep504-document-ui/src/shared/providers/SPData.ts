import * as $ from "jquery";
import { IIep504DocumentUiWebPartProps } from "../../webparts/iep504DocumentUi/Iep504DocumentUiWebPart";
import { sp, Items } from "@pnp/sp";
export class SPData {
  public webPartProperties: IIep504DocumentUiWebPartProps;
  public siteUrl: String;
  constructor(
    webPartProperties: IIep504DocumentUiWebPartProps,
    siteUrl: String
  ) {
    this.webPartProperties = webPartProperties;
    this.siteUrl = siteUrl;
  }
  public async GetFirstAccessedDate(spFileName: string, spFileModifiedDate: string, spUserId: number) {
    try {
      if (!spUserId) { console.warn(`Skipping Access Audit query of file '${spFileName}', as there was no supplied user.`); return null; }
      let filter = `(AuditSourceFileName eq '${spFileName}')`;
      filter += `and ((AuditDate ge datetime'${new Date(spFileModifiedDate).toISOString()}')`;
      filter += `and ((AuditOperation eq 'FileAccessed')`;
      filter += (this.webPartProperties.treatPreviewAsRead) ? "or (AuditOperation eq 'FilePreviewed'))" : ")";
      filter += `and (AuditUserIdId eq ${spUserId}))`;

      let items = await sp.web.lists.getByTitle(this.webPartProperties.auditListName).items
        .select("AuditDate")
        .filter(filter)
        .orderBy("AuditDate", true)
        .getPaged();
      return (items.results.length > 0) ? items.results[0].AuditDate : null;
    }
    catch (e) { console.log(e); }
  }
}
