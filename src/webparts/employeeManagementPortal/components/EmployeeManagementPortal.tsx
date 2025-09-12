import * as React from "react";
import { useState } from "react";
import type { IEmployeeManagementPortalProps } from "./IEmployeeManagementPortalProps";
import styles from "./EmployeeManagementPortal.module.scss";

import EmployeeDirectory from "./EmployeeDirectory/EmployeeDirectory";
import { Navbar } from "./Navbar/NavbarComponent";

const EmployeeManagementPortal: React.FC<IEmployeeManagementPortalProps> = (
  props
) => {
  const [activePage, setActivePage] = useState<string>("directory");

  const renderPage = (): JSX.Element => {
    switch (activePage) {
      case "directory":
        return (
          <EmployeeDirectory
            listName="Employees"
            sp={props.sp}
            employeeHttpService={props.employeeHttpService}
          />
        );
      case "leave":
        return <p>Leave Tracker</p>;
      case "announcements":
        return <p>Announcements</p>;
      default:
        return <h3>Page not found</h3>;
    }
  };

  return (
    <div>
      <h2 className={styles.webpartTitle}>Employee Management System</h2>

      <div className={styles.employeeManagementPortal}>
        <div style={{ flex: "0 0 230px" }}>
          <Navbar onSelectPage={setActivePage} />
        </div>

        <div className={styles.contentArea}>{renderPage()}</div>
      </div>
    </div>
  );
};

export default EmployeeManagementPortal;
