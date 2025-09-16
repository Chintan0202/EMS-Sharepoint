/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
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
  IconButton,
} from "@fluentui/react";

import { EmployeeHttpService } from "../../services/EmployeeHttpService";
import { IEmployee } from "./IEmployee";
import { useEffect, useState } from "react";
import { SPFI } from "@pnp/sp";
// import UserProfileMenu from "../UserProfile/UserProfle";

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
  const [isDeleteDialogOpen, setIsDeleteDialogOpen] = useState(false);
  const [isCardView, setIsCardView] = useState<boolean>(false);

  const [editingEmployee, setEditingEmployee] = useState<IEmployee>(
    {} as IEmployee
  );
  const [formData, setFormData] = useState<any>({
    Title: "",
    EmployeeID: "",
    Email: "",
    Designation: "",
    PhoneNumber: "",
    PhotoUrl: "",
    IsActive: true,
    DepartmentId: "",
  });

  const [errors, setErrors] = useState<{ [key: string]: string }>({});

  const [departments, setDepartments] = useState<IDropdownOption[]>([]);
  const [selectedDepartment, setSelectedDepartment] = useState<string>("All");
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize] = useState<number>(5);
  const [totalPages, setTotalPages] = useState<number>(1);
  const [currentUser, setCurrentUser] = useState<any>(null);

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

  const validateForm = (): boolean => {
    const newErrors: { [key: string]: string } = {};
    if (!formData.Title.trim()) newErrors.Title = "Name is required";
    if (!formData.Email.trim()) newErrors.Email = "Email is required";
    if (!formData.EmployeeID.trim())
      newErrors.EmployeeID = "Employee ID is required";
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const openAddDialog = (): void => {
    setFormData({
      Title: "",
      EmployeeID: "",
      Email: "",
      Designation: "",
      PhoneNumber: "",
      PhotoUrl: "",
      IsActive: true,
      DepartmentId: "",
    });
    setIsDialogOpen(true);
  };

  const handleSave = async (): Promise<void> => {
    if (!validateForm()) return;
    try {
      if (editingEmployee && editingEmployee.Id) {
        await employeeHttpService.updateEmployee(
          listName,
          editingEmployee.Id,
          formData
        );
      } else {
        await employeeHttpService.addEmployee(listName, formData);
      }
      setIsDialogOpen(false);
      await fetchEmployees("");
    } catch (err) {
      console.error("Error saving employee:", err);
    }
  };

  const handleEdit = (employee: IEmployee) => {
    setEditingEmployee(employee);
    setFormData(employee);
    setIsDialogOpen(true);
  };

  const handleDelete = async () => {
    if (!editingEmployee) return;
    try {
      if (editingEmployee && editingEmployee.Id) {
        await employeeHttpService.deleteEmployee(listName, editingEmployee.Id);
        setIsDeleteDialogOpen(false);
      }
      await fetchEmployees("");
    } catch (err) {
      console.error("Error deleting employee:", err);
    }
  };
  useEffect(() => {
    const fetchCurrentUser = async () => {
      try {
        const user = await employeeHttpService.getCurrentUserDetails();
        console.log(user);
        setCurrentUser(user);
      } catch (error) {
        console.error("Error fetching current user:", error);
      }
    };
    void fetchCurrentUser();
  }, []);

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

  const columns: IColumn[] = [
    {
      key: "photo",
      name: "Photo",
      minWidth: 50,
      maxWidth: 60,
      onRender: (item: IEmployee) =>
        item.PhotoUrl?.Url ? (
          <img
            src={item.PhotoUrl.Url}
            alt="Employee"
            style={{ width: 32, height: 32, borderRadius: "50%" }}
          />
        ) : (
          <div
            style={{
              width: 32,
              height: 32,
              borderRadius: "50%",
              background: "#ccc",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 14,
              fontWeight: "bold",
            }}
          >
            {item.Title?.charAt(0)}
          </div>
        ),
    },
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
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IEmployee) => (
        <span>{item.Department?.Title ?? "—"}</span>
      ),
    },
    {
      key: "actions",
      name: "Actions",
      minWidth: 100,
      isResizable: true,
      onRender: (item: IEmployee) => (
        <>
          <IconButton
            iconProps={{ iconName: "Edit" }}
            title="Edit"
            ariaLabel="Edit"
            onClick={() => handleEdit(item)}
          />
          <IconButton
            iconProps={{ iconName: "Delete" }}
            title="Delete"
            ariaLabel="Delete"
            onClick={() => {
              setEditingEmployee(item);
              setIsDeleteDialogOpen(true);
            }}
            styles={{ root: { color: "red" } }}
          />
        </>
      ),
    },
  ];

  return (
    <div style={{ display: "flex", height: "100%" }}>
      {/* LEFT SECTION - 70% */}
      <div
        style={{
          flexBasis: "70%",
          display: "flex",
          flexDirection: "column",
          borderRight: "1px solid #ddd",
          overflow: "hidden",
        }}
      >
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
            styles={{ root: { flex: 1, maxWidth: "400px" } }}
            onChange={(_, newValue) => {
              setSearchValue(newValue ?? "");
              setCurrentPage(1);
            }}
          />
          <Dropdown
            selectedKey={selectedDepartment}
            options={departments}
            styles={{ root: { width: 200 } }}
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

        {/* ADD EMPLOYEE BUTTON */}
        <div style={{ display: "flex", justifyContent: "flex-end" }}>
          <PrimaryButton
            text="Add Employee"
            onClick={openAddDialog}
            styles={{ root: { margin: 10 } }}
          />
        </div>
        <Dialog
          hidden={!isDialogOpen}
          onDismiss={() => setIsDialogOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: editingEmployee ? "Edit Employee" : "Add Employee",
          }}
          styles={{ main: { minWidth: "650px" } }}
        >
          <TextField
            label="Employee ID"
            value={formData.EmployeeID}
            onChange={(_, v) => handleChange("EmployeeID", v)}
            required
            errorMessage={errors.EmployeeID}
          />
          <TextField
            label="Name"
            value={formData.Title}
            onChange={(_, v) => handleChange("Title", v)}
            required
            errorMessage={errors.Title}
          />
          <TextField
            label="Email"
            value={formData.Email}
            onChange={(_, v) => handleChange("Email", v)}
            required
            errorMessage={errors.Email}
          />
          <TextField
            label="Designation"
            value={formData.Designation}
            onChange={(_, v) => handleChange("Designation", v)}
          />
          <TextField
            label="Phone Number"
            value={formData.PhoneNumber}
            onChange={(_, v) => handleChange("PhoneNumber", v)}
          />
          <Dropdown
            label="Department"
            selectedKey={formData.DepartmentId || null}
            options={departments.filter((d) => d.key !== "All")}
            onChange={(_, option) =>
              handleChange("DepartmentId", option?.key as number)
            }
            required
          />
          <Checkbox
            label="Status"
            checked={formData.IsActive}
            onChange={(_, data) => handleChange("IsActive", data)}
            styles={{ root: { marginTop: 10 } }}
          />
          <input
            type="file"
            accept="image/*"
            onChange={async (e) => {
              if (e.target.files && e.target.files.length > 0) {
                const file = e.target.files[0];
                try {
                  const photoUrl =
                    await employeeHttpService.uploadEmployeePhoto(file);
                  setFormData({ ...formData, PhotoUrl: { Url: photoUrl } });
                } catch (err) {
                  console.error("Photo upload failed", err);
                }
              }
            }}
          />

          {formData.PhotoUrl?.Url && (
            <img
              src={formData.PhotoUrl.Url}
              alt="Employee"
              style={{
                width: 80,
                height: 80,
                borderRadius: "50%",
                marginTop: 8,
              }}
            />
          )}

          <DialogFooter>
            <PrimaryButton
              onClick={handleSave}
              text={editingEmployee.Id ? "Update" : "Save"}
            />
            <DefaultButton
              onClick={() => setIsDialogOpen(false)}
              text="Cancel"
            />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={!isDeleteDialogOpen}
          onDismiss={() => setIsDeleteDialogOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Delete Employee",
            subText: "Are you sure you want to delete this employee?",
          }}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={handleDelete}
              text="Yes, Delete"
              styles={{ root: { background: "red" } }}
            />
            <DefaultButton
              onClick={() => setIsDeleteDialogOpen(false)}
              text="Cancel"
            />
          </DialogFooter>
        </Dialog>
        {/* TABLE / CARD VIEW - WITH HORIZONTAL SCROLL */}
        <div
          style={{
            flex: 1,
            minHeight: 0,
            overflowX: "auto",
            overflowY: "hidden",
          }}
        >
          {loading && <Spinner label="Loading..." />}

          {!loading && (
            <div
              style={{ minWidth: "900px" /* Force scroll if table is wide */ }}
            >
              {isCardView ? (
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(3, 1fr)",
                    gap: "16px",
                  }}
                >
                  {employees.map((emp) => (
                    <div
                      key={emp.Id}
                      style={{
                        border: "1px solid #ddd",
                        borderRadius: 8,
                        padding: "16px",
                        boxShadow: "0 1px 4px rgba(0,0,0,0.08)",
                        background: "#fff",
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "flex-start",
                        textAlign: "left",
                        transition: "box-shadow 0.2s ease, transform 0.2s ease",
                        cursor: "pointer",
                      }}
                      onMouseEnter={(e) => {
                        (e.currentTarget as HTMLDivElement).style.boxShadow =
                          "0 2px 8px rgba(0,0,0,0.15)";
                        (e.currentTarget as HTMLDivElement).style.transform =
                          "translateY(-2px)";
                      }}
                      onMouseLeave={(e) => {
                        (e.currentTarget as HTMLDivElement).style.boxShadow =
                          "0 1px 4px rgba(0,0,0,0.08)";
                        (e.currentTarget as HTMLDivElement).style.transform =
                          "none";
                      }}
                    >
                      {emp.PhotoUrl?.Url ? (
                        <img
                          src={emp.PhotoUrl.Url}
                          alt="Employee"
                          style={{
                            width: 70,
                            height: 70,
                            borderRadius: "50%",
                            objectFit: "cover",
                            marginBottom: 8,
                            border: "2px solid #f0f0f0",
                          }}
                        />
                      ) : (
                        <div
                          style={{
                            width: 70,
                            height: 70,
                            borderRadius: "50%",
                            background: "#a4262c",
                            color: "#fff",
                            fontSize: 24,
                            fontWeight: "bold",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            marginBottom: 8,
                          }}
                        >
                          {emp.Title?.charAt(0)}
                        </div>
                      )}

                      {/* NAME */}
                      <h3
                        style={{
                          margin: "4px 0",
                          fontSize: "1rem",
                          fontWeight: 600,
                          color: "#323130",
                        }}
                      >
                        {emp.Title}
                      </h3>

                      {/* STATUS */}
                      <span
                        style={{
                          fontSize: "0.75rem",
                          color: emp.IsActive ? "#107c10" : "#a4262c",
                          fontWeight: 500,
                          marginBottom: 8,
                        }}
                      >
                        {emp.IsActive ? "Active" : "Inactive"}
                      </span>

                      {/* DETAILS */}
                      <div
                        style={{
                          fontSize: "0.85rem",
                          color: "#444",
                          lineHeight: 1.5,
                          width: "100%",
                        }}
                      >
                        <p style={{ margin: "2px 0" }}>
                          <strong>ID:</strong> {emp.EmployeeID}
                        </p>
                        <p style={{ margin: "2px 0" }}>
                          <strong>Email:</strong> {emp.Email}
                        </p>
                        <p style={{ margin: "2px 0" }}>
                          <strong>Department:</strong>{" "}
                          {emp.Department?.Title || "—"}
                        </p>
                        <p style={{ margin: "2px 0" }}>
                          <strong>Designation:</strong> {emp.Designation}
                        </p>
                        {emp.PhoneNumber && (
                          <p style={{ margin: "2px 0" }}>
                            <strong>Phone:</strong> {emp.PhoneNumber}
                          </p>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <DetailsList
                  items={employees}
                  columns={columns}
                  setKey="Id"
                  selectionMode={SelectionMode.single}
                  layoutMode={DetailsListLayoutMode.justified}
                  compact
                />
              )}
            </div>
          )}

          {totalPages > 1 && (
            <div
              style={{
                borderTop: "1px solid #ddd",
                padding: "8px 0",
                display: "flex",
                justifyContent: "start",
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
      </div>

      {/* RIGHT SECTION - 30% */}
      {currentUser && (
        <div
          style={{
            flexBasis: "30%",
            minWidth: "150px",
            padding: "1rem",
            backgroundColor: "#fafafa",
            overflowY: "auto",
          }}
        >
          {currentUser.PictureUrl ? (
            <img
              src={currentUser.PictureUrl}
              alt="Employee"
              style={{
                width: 70,
                height: 70,
                borderRadius: "50%",
                objectFit: "cover",
                marginBottom: 8,
                border: "2px solid rgb(214, 199, 199)",
              }}
            />
          ) : (
            <div
              style={{
                width: 70,
                height: 70,
                borderRadius: "50%",
                background: "#a4262c",
                color: "#fff",
                fontSize: 24,
                fontWeight: "bold",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                marginBottom: 8,
              }}
            >
              {currentUser.DisplayName.charAt(0)}
            </div>
          )}
          <h3
            style={{
              margin: "4px 0",
              fontSize: "1rem",
              fontWeight: 600,
              color: "#323130",
            }}
          >
            {currentUser.DisplayName}
          </h3>
          <div
            style={{
              fontSize: "0.85rem",
              color: "#444",
              lineHeight: 1.5,
              width: "100%",
            }}
          >
            <p style={{ margin: "2px 0",textWrap: "wrap" }}>{currentUser.Email}</p>
          </div>
        </div>
      )}
    </div>
  );
};

export default EmployeeDirectory;
