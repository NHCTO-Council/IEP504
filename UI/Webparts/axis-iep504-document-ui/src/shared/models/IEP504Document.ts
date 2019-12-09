import { IEP504Student, IEP504Teacher } from "../../shared/models";

export class IEP504Document {
  public fileName: string;
  public student: IEP504Student;
  public teachers: Array<IEP504Teacher>;
  public modified: string;
  public headerRow: string;
  public SetHeaderRow() {
    try {
      this.headerRow =
        this.student.lastName +
        ", " +
        this.student.firstName +
        " (ID: " +
        this.student.id +
        ")";
    } catch (err) {
      return "Error: " + err;
    }
  }
}
