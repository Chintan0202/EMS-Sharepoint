import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface ILeaveRequest {
  Employee: string;
  LeaveType: string;
  StartDate: string;
  EndDate: string;
  Status: string;
}

export class LeaveTrackerService {
  private sp: SPFI;
  private listName: string;

  constructor(sp: SPFI, listName: string) {
    this.sp = sp;
    this.listName = listName;
  }
  public async getLeaveRequests(): Promise<ILeaveRequest[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.listName).items.select(
        'Title',
        'LeaveType',
        'StartDate',
        'EndDate',
        'Status'
      )();

      return items.map((item: any) => ({
        Employee: item.Title,
        LeaveType: item.LeaveType,
        StartDate: new Date(item.StartDate).toLocaleDateString(),
        EndDate: new Date(item.EndDate).toLocaleDateString(),
        Status: item.Status
      }));
    } catch (error) {
      console.error('Error fetching leave requests:', error);
      return [];
    }
  }
}
