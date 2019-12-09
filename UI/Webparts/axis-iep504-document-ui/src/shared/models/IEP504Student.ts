import { IEP504CaseManager } from "../../shared/models";
export class IEP504Student {
  public id: string;
  public firstName: string;
  public lastName: string;
  public caseManager: IEP504CaseManager; // TODO: Research-Should Case Manager be relative to the student, instead of document?
}
