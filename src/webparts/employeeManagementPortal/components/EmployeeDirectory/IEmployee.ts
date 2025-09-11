export interface IEmployee {
Title?: string; // SharePoint's Title (we can map Name -> Title)
Name: string;
Email: string;
Designation: string;
Department: string;
Status: 'Active' | 'Inactive' | string;
PhotoURL?: string;
}
