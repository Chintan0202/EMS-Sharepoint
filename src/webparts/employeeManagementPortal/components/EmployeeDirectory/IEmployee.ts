export interface IEmployee {
  Id?: number;
  Title: string;        // Name
  EmployeeID: string;
  Email: string;
  Designation: string;
  PhoneNumber: string;
  IsActive: boolean;
  DepartmentId?: number;   // Lookup reference
  Department?: { Id: number; Title: string };
  PhotoUrl?: { Url: string; Description: string };
}
