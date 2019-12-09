import { IEP504Student, IEP504Teacher } from "../../shared/models";
export declare class IEP504Document {
    fileName: string;
    student: IEP504Student;
    teachers: Array<IEP504Teacher>;
    modified: string;
    headerRow: string;
    SetHeaderRow(): string;
}
