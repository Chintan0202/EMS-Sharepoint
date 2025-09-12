import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export class EmployeeHttpService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getEmployees(listName: string, searchtext?: string): Promise<any[]> {
    let filterQuery = "";
    if (searchtext && searchtext.trim().length > 0) {
      filterQuery = `&$filter=substringof('${searchtext}',Title) or substringof('${searchtext}',Email) or substringof('${searchtext}',Designation) or substringof('${searchtext}',EmployeeID)`;
    }
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=EmployeeID,Title,Designation,Email,DepartmentId,PhotoUrl${filterQuery}`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Error fetching employees: ${response.statusText}`);
    }

    const data = await response.json();
    return data.value;
  }

  public async addEmployee(listName: string, employee: any): Promise<any> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    console.log(this.context.pageContext.web.absoluteUrl);
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": ""
        },
        body: JSON.stringify(employee),
      }
    );

    if (!response.ok) {
      throw new Error(`Error adding employee: ${response.statusText}`);
    }

    return response.json();
  }

  public async updateEmployee(listName: string, id: number, employee: any): Promise<any> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify(employee),
      }
    );

    if (!response.ok) {
      throw new Error(`Error updating employee: ${response.statusText}`);
    }

    return response.json();
  }

  public async getCurrentUserDetails(): Promise<any> {

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Error fetching employees: ${response.statusText}`);
    }

    const data = await response.json();
    return data;
  }
  public async getDepartments(listName: string): Promise<any[]> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Error fetching departments: ${response.statusText}`);
    }

    const data = await response.json();
    return data.value.map((d: any) => ({
      key: d.Id,
      text: d.Title,
    }));
  }

}
