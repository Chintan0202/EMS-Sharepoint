import * as React from "react";
import {
  SearchBox,
  DetailsList,
  IColumn,
  SelectionMode,
  DetailsListLayoutMode,
  Spinner,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  Toggle,
  Checkbox,
} from "@fluentui/react";

// import styles from "./EmployeeDirectory.module.scss";
import { EmployeeHttpService } from "../../services/EmployeeHttpService";
import { IEmployee } from "./IEmployee";
import { useEffect, useState } from "react";
import { SPFI } from "@pnp/sp";
import UserProfileMenu from "../UserProfile/UserProfle";

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
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [isCardView, setIsCardView] = useState<boolean>(false);
  const [selectedEmployee, setSelectedEmployee] = useState<IEmployee | null>(
    null
  );

  const [formData, setFormData] = useState<any>({
    Title: "",
    EmployeeID: "",
    Email: "",
    Designation: "",
    PhoneNumber: "",
    IsActive: true,
    DepartmentId: "",
  });

  const [errors, setErrors] = useState<{ [key: string]: string }>({});

  const columns: IColumn[] = [
    {
      key: "EmployeeID",
      name: "Employee ID",
      fieldName: "EmployeeID",
      minWidth: 70,
      maxWidth: 90,
      isResizable: true,
    },
    {
      key: "name",
      name: "Name",
      fieldName: "Title",
      minWidth: 120,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "email",
      name: "Email",
      fieldName: "Email",
      minWidth: 150,
      maxWidth: 250,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "Designation",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "department",
      name: "Department",
      fieldName: "DepartmentId",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
  ];

  const [departments, setDepartments] = useState<IDropdownOption[]>([]);
  const [selectedDepartment, setSelectedDepartment] = useState<string>("All");
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize] = useState<number>(5);
  const [totalPages, setTotalPages] = useState<number>(1);

  const fetchEmployees = async (filterText?: string, department?: string) => {
    setLoading(true);
    try {
      let items = await employeeHttpService.getEmployees(
        listName,
        filterText ?? ""
      );
      if (department && department !== "All") {
        items = items.filter((e) => e.DepartmentId === department);
      }

      const totalPages = Math.ceil(items.length / pageSize);
      setTotalPages(totalPages);
      // ensure we don't go out of range
      const safePage = Math.min(currentPage, totalPages || 1);
      setCurrentPage(safePage);

      const paginatedEmployees = items.slice(
        (safePage - 1) * pageSize,
        safePage * pageSize
      );
      setEmployees(paginatedEmployees || []);
    } catch (err) {
      console.error("Error fetching employees:", err);
      setEmployees([]);
    } finally {
      setLoading(false);
    }
  };

  const handleChange = (field: string, value?: string | number | boolean) => {
    setFormData({ ...formData, [field]: value ?? "" });
  };

  const validateForm = () => {
    const newErrors: { [key: string]: string } = {};
    if (!formData.Title.trim()) newErrors.Title = "Name is required";
    if (!formData.Email.trim()) newErrors.Email = "Email is required";
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSave = async () => {
    if (!validateForm()) return;
    try {
      await employeeHttpService.addEmployee(listName, formData);
      setIsDialogOpen(false);
      setFormData({ Title: "", Designation: "", Email: "", DepartmentId: "" });
      await fetchEmployees("");
    } catch (err) {
      console.error("Error saving employee:", err);
    }
  };

  useEffect(() => {
    const fetchDepartments = async () => {
      try {
        const deptItems = await employeeHttpService.getDepartments(
          "Departments"
        );
        setDepartments([{ key: "All", text: "All Departments" }, ...deptItems]);
      } catch (err) {
        console.error("Error fetching departments:", err);
      }
    };

    void fetchDepartments();
    void fetchEmployees("");
  }, [listName]);

  useEffect(() => {
    const handler = setTimeout(() => {
      void fetchEmployees(searchValue.trim(), selectedDepartment);
    }, 400);
    return () => clearTimeout(handler);
  }, [searchValue, selectedDepartment, currentPage]);

  return (
    <div style={{ display: "flex", height: "100%" }}>
      <div
        style={{ flex: 2, padding: "0 1rem", borderRight: "1px solid #ddd" }}
      >
        <UserProfileMenu employeeHttpService={employeeHttpService} />
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: "10px",
            margin: 10,
          }}
        >
          <SearchBox
            placeholder="Search name, designation, id, email..."
            value={searchValue}
            styles={{
              root: { flex: 1, maxWidth: "400px" },
            }}
            onChange={(_, newValue) => {
              setSearchValue(newValue ?? "");
              setCurrentPage(1);
            }}
          />

          <Dropdown
            selectedKey={selectedDepartment}
            options={departments}
            styles={{
              root: { width: 200 },
            }}
            onChange={(_, option) => {
              setSelectedDepartment((option?.key as string) || "All");
              setCurrentPage(1);
            }}
          />

          <div style={{ marginLeft: "auto" }}>
            <Toggle
              checked={isCardView}
              onChange={(_, checked) => setIsCardView(!!checked)}
              onText="Card View"
              offText="Table View"
            />
          </div>
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end" }}>
          <PrimaryButton
            text="Add Employee"
            onClick={() => setIsDialogOpen(true)}
            styles={{ root: { marginBottom: 10 } }}
          />
        </div>
        <Dialog
          hidden={!isDialogOpen}
          onDismiss={() => setIsDialogOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Add Employee",
          }}
          styles={{ main: { minWidth: "650px" } }}
        >
          {" "}
          <TextField
            label="Name"
            value={formData.Title}
            onChange={(_, newValue) =>
              setFormData({ ...formData, Title: newValue || "" })
            }
            required
            errorMessage={errors.Title}
          />{" "}
          <TextField
            label="Email"
            value={formData.Email}
            onChange={(_, newValue) =>
              setFormData({ ...formData, Email: newValue || "" })
            }
            required
            errorMessage={errors.Email}
          />{" "}
          <TextField
            label="Designation"
            value={formData.Designation}
            onChange={(_, val) => handleChange("Designation", val)}
          />{" "}
          <Dropdown
            label="Department"
            selectedKey={formData.DepartmentId || null}
            options={departments.filter((d) => d.key !== "All")}
            onChange={(_, option) =>
              handleChange("DepartmentId", option?.key as number)
            }
            required
          />{" "}
          <Checkbox
            label="Status"
            checked={formData.IsActive}
            onChange={(event, data) => handleChange("IsActive", data)}
            styles={{ root: { marginTop: 10 } }}
          />{" "}
          <DialogFooter>
            {" "}
            <PrimaryButton onClick={handleSave} text="Save" />{" "}
            <DefaultButton
              onClick={() => setIsDialogOpen(false)}
              text="Cancel"
            />{" "}
          </DialogFooter>{" "}
        </Dialog>
        {loading && <Spinner label="Loading..." />}

        {isCardView ? (
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fill, minmax(250px, 1fr))",
              gap: "16px",
            }}
          >
            {employees.map((emp) => (
              <div
                key={emp.Id}
                onClick={() => setSelectedEmployee(emp)}
                style={{
                  border: "1px solid #ddd",
                  borderRadius: 8,
                  padding: 16,
                  boxShadow: "0 1px 4px rgba(0,0,0,0.1)",
                  backgroundColor:
                    selectedEmployee?.Id === emp.Id ? "#f3f2f1" : "#fff",
                  cursor: "pointer",
                }}
              >
                <h3>{emp.Title}</h3>
                <p>
                  <strong>Email:</strong> {emp.Email}
                </p>
                <p>
                  <strong>Department:</strong> {emp.DepartmentId}
                </p>
                <p>
                  <strong>Designation:</strong> {emp.Designation}
                </p>
              </div>
            ))}
          </div>
        ) : (
          <DetailsList
            items={employees}
            columns={columns}
            setKey="Id"
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            compact
            onItemInvoked={(item) => setSelectedEmployee(item)}
          />
        )}

        {totalPages > 1 && (
          <div
            style={{
              display: "flex",
              justifyContent: "start",
              marginTop: 16,
              gap: 8,
            }}
          >
            <DefaultButton
              text="Previous"
              disabled={currentPage === 1}
              onClick={() => setCurrentPage((prev) => prev - 1)}
            />
            <span style={{ alignSelf: "center" }}>
              Page {currentPage} of {totalPages}
            </span>
            <DefaultButton
              text="Next"
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage((prev) => prev + 1)}
            />
          </div>
        )}
      </div>

      {/* RIGHT PANEL (SELECTED EMPLOYEE DETAILS) */}
      <div style={{ flex: 1, padding: "1rem", backgroundColor: "#fafafa" }}>
        {selectedEmployee ? (
          <div style={{ textAlign: "center" }}>
            <div
              style={{
                width: 100,
                height: 100,
                borderRadius: "50%",
                background: "#a4262c",
                color: "#fff",
                fontSize: 36,
                fontWeight: "bold",
                margin: "0 auto 10px",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              {selectedEmployee.Title.charAt(0)}
            </div>
            <h2>{selectedEmployee.Title}</h2>
            <p>{selectedEmployee.Email}</p>
            <p>
              <strong>Designation:</strong> {selectedEmployee.Designation}
            </p>
            <p>
              <strong>Department:</strong> {selectedEmployee.DepartmentId}
            </p>
            <p>
              <strong>Employee ID:</strong> {selectedEmployee.EmployeeID}
            </p>
          </div>
        ) : (
          <div style={{ textAlign: "center", marginTop: "30%" }}>
            <p>Select an employee to view details</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default EmployeeDirectory;
