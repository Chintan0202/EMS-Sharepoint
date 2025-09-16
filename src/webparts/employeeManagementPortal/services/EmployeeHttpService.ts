/* eslint-disable @typescript-eslint/no-explicit-any */
import { ISPHttpClientOptions, MSGraphClientV3, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
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

  const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items` +
    `?$select=Id,Title,EmployeeID,DepartmentId,Email,Designation,PhoneNumber,IsActive,Department/Id,Department/Title,PhotoUrl` +
    `&$expand=Department${filterQuery}`;

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

  public async uploadEmployeePhoto(file: File): Promise<string> {
    const makeSafeFileName = (name: string): string => {
      const dot = name.lastIndexOf('.');
      const base = (dot > 0 ? name.slice(0, dot) : name).replace(/[^a-zA-Z0-9_.-]/g, '_');
      const ext = dot > 0 ? name.slice(dot) : '';
      return `${base}_${Date.now()}${ext}`;
    }

    const folderUrl = `${this.context.pageContext.web.serverRelativeUrl}/EmployeePhotos/Images`;
    const fileName = makeSafeFileName(file.name);
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${folderUrl}')/files/add(overwrite=false,url='${fileName}')`;


    const fileBuffer = await file.arrayBuffer();

    const options: ISPHttpClientOptions = {
      body: fileBuffer,
      headers: {
        "Content-Type": "application/octet-stream"
      }
    };

    const response = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      options
    );
    console.log(response);
    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Upload failed (${response.status}): ${err}`);
    }

    const json = await response.json();
    return json.ServerRelativeUrl || json.d?.ServerRelativeUrl;
  }


  public async updateEmployee(listName: string, id: number, employee: any): Promise<any> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          // "Accept": "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify(employee),
      }
    );

    return "success";
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

  public async getCurrentUserDetail(): Promise<any> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient("3");

    // Step 1: Get basic user profile info
    const user = await client
      .api('/me')
      .select('id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone')
      .get();

    let photoUrl = '';
    try {
      const photoResponse = await client.api('/me/photo/$value').get();
      const blob = await photoResponse.blob();
      photoUrl = URL.createObjectURL(blob);
    } catch (error) {
      console.warn(error);
    }

    return { ...user, photoUrl };
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
  public async deleteEmployee(listName: string, id: number): Promise<void> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          // "Accept": "application/json;odata=nometadata",
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Error deleting employee: ${response.statusText}`);
    }
  }

}
