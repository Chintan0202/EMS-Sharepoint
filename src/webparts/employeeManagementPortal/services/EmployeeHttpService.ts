import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export class EmployeeHttpService {
  private spHttpClient: SPHttpClient;
  private siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
  }

public async getEmployees(listName: string, searchtext?: string): Promise<any[]> {
  let filterQuery = "";
  if (searchtext && searchtext.trim().length > 0) {
    filterQuery = `&$filter=substringof('${searchtext}',LinkTitle) or substringof('${searchtext}',Email)`;
  }

  const endpoint = `${this.siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,Designation,Email,PhoneNumber${filterQuery}`;

  const response: SPHttpClientResponse = await this.spHttpClient.get(
    endpoint,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    throw new Error(`Error fetching employees: ${response.statusText}`);
  }

  const data = await response.json();
  return data.value;
}

}
