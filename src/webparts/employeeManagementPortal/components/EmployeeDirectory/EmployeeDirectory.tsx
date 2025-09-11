import * as React from "react";
import {
  SearchBox,
  DetailsList,
  IColumn,
  SelectionMode,
  DetailsListLayoutMode,
  Spinner,
} from "@fluentui/react";
import styles from "./EmployeeDirectory.module.scss";
import { EmployeeHttpService } from "../../services/EmployeeHttpService";
import { IEmployee } from "./IEmployee";
import { useEffect, useState } from "react";
import { SPFI } from "@pnp/sp";

export interface IEmployeeDirectoryProps {
  listName: string;
    sp: SPFI;
  employeeHttpService: EmployeeHttpService;
}

const EmployeeDirectory: React.FC<IEmployeeDirectoryProps> = ({
  listName,
  sp,
  employeeHttpService,
}) => {
  const [employees, setEmployees] = useState<IEmployee[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [searchValue, setSearchValue] = useState<string>("");

  const columns: IColumn[] = [
    { key: "id", name: "ID", fieldName: "Id", minWidth: 50, maxWidth: 70, isResizable: true },
    { key: "name", name: "Name", fieldName: "Title", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "designation", name: "Designation", fieldName: "Designation", minWidth: 100, maxWidth: 200, isResizable: true },
    { key: "email", name: "Email", fieldName: "Email", minWidth: 150, maxWidth: 250, isResizable: true },
    { key: "phone", name: "Phone", fieldName: "PhoneNumber", minWidth: 100, maxWidth: 200, isResizable: true },
  ];

  // fetch function (calls your Http service)
  const fetchEmployees = async (filterText?: string) => {
    setLoading(true);
    try {
      const items = await employeeHttpService.getEmployees(listName, filterText ?? "");
      setEmployees(items || []);
    } catch (err) {
      console.error("Error fetching employees:", err);
      setEmployees([]);
    } finally {
      setLoading(false);
    }
  };

  // initial load (once)
  useEffect(() => {
    void fetchEmployees("");
  }, [listName]);

  useEffect(() => {
    const handler = setTimeout(() => {
      void fetchEmployees(searchValue.trim());
    }, 400); 

    return () => clearTimeout(handler);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [searchValue]);

  return (
    <div className={styles.employeeDirectory}>
      <h2 className={styles.title}>Employee Directory</h2>

      {/* Controlled SearchBox (always rendered) */}
      <SearchBox
        placeholder="Search by name or email"
        value={searchValue}
        onChange={(_, newValue) => setSearchValue(newValue ?? "")}
        styles={{ root: { marginBottom: 10, maxWidth: 360 } }}
      />

      {/* inline spinner (doesn't unmount the input) */}
      {loading && <Spinner label="Loading..." />}

      <DetailsList
        items={employees}
        columns={columns}
        setKey="Id"
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.justified}
        compact={true}
      />
    </div>
  );
};

export default EmployeeDirectory;
