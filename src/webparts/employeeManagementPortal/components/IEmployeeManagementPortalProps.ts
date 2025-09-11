import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { EmployeeHttpService } from "../services/EmployeeHttpService";

export interface IEmployeeManagementPortalProps {
  title: string;
  context: WebPartContext;
  sp: SPFI; 
  employeeHttpService: EmployeeHttpService
}
