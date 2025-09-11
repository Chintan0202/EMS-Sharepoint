import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

import * as strings from "EmployeeManagementPortalWebPartStrings";
import EmployeeManagementPortal from "./components/EmployeeManagementPortal";
import { IEmployeeManagementPortalProps } from "./components/IEmployeeManagementPortalProps";
import { EmployeeHttpService } from "./services/EmployeeHttpService";

export interface IEmployeeManagementPortalWebPartProps {
  title: string;
}

export default class EmployeeManagementPortalWebPart extends BaseClientSideWebPart<IEmployeeManagementPortalWebPartProps> {
  private _sp: SPFI;
  private employeeHttpService: EmployeeHttpService

  public onInit(): Promise<void> {
  this._sp = spfi("https://tatvasoft0.sharepoint.com/sites/ems").using(SPFx(this.context));
  this.employeeHttpService = new EmployeeHttpService(this.context.spHttpClient, "https://tatvasoft0.sharepoint.com/sites/CustomPOCSite");
  return Promise.resolve();
}


  public render(): void {
    const element: React.ReactElement<IEmployeeManagementPortalProps> =
      React.createElement(EmployeeManagementPortal, {
        title: this.properties.title,
        sp: this._sp,
        context: this.context,
        employeeHttpService: this.employeeHttpService
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleFieldLabel,
                  value: "Employee Management System",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
