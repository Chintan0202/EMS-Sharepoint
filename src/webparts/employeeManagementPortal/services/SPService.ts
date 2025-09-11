import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { IEmployee } from "../components/EmployeeDirectory/IEmployee";

export class EmployeeService {
  private sp: SPFI;
  private listName: string;

  constructor(sp: SPFI, listName: string) {
    this.sp = sp;
    this.listName = listName;
  }

  public async getEmployees(): Promise<IEmployee[]> {
    try {
    const items = await this.sp.web.lists.getByTitle(this.listName).items
      .select("Id", "Title", "Department", "Email")();

      return items;
    } catch (error) {
      console.error("Error fetching employees:", error);
      return [];
    }
  }

  public async addEmployee(employee: Partial<IEmployee>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.listName).items.add(employee);
    } catch (error) {
      console.error("Error adding employee:", error);
    }
  }

  public async updateEmployee(id: number, employee: Partial<IEmployee>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.listName).items.getById(id).update(employee);
    } catch (error) {
      console.error("Error updating employee:", error);
    }
  }

  public async deleteEmployee(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.listName).items.getById(id).delete();
    } catch (error) {
      console.error("Error deleting employee:", error);
    }
  }
}
