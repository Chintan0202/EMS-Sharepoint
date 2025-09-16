import * as React from "react";
import { useEffect, useState } from "react";
import styles from "./LeaveTrackerComponent.module.scss";
import { ILeaveRequest, LeaveTrackerService } from "./LeaveTrackerService";

import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
} from "@fluentui/react/lib/DetailsList";

import { mergeStyles } from "@fluentui/react/lib/Styling";
import { SPFI } from "@pnp/sp";

export interface ILeaveTrackerProps {
  listName: string;
  sp: SPFI;
}

export const LeaveTrackerComponent: React.FC<ILeaveTrackerProps> = ({
  listName,
  sp,
}) => {
  const [leaves, setLeaves] = useState<ILeaveRequest[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  const leaveTrackerService = new LeaveTrackerService(sp, listName);

  useEffect(() => {
    const fetchData = async (): Promise<void> => {
      try {
        setLoading(true);
        const data = await leaveTrackerService.getLeaveRequests();
        setLeaves(data);
      } catch (error) {
        console.error("Error fetching leave data:", error);
      } finally {
        setLoading(false);
      }
    };

    void fetchData();
  }, [listName, sp]); 

  const statusBadgeClass = (status: string): string => {
    switch (status?.toLowerCase()) {
      case "approved":
        return mergeStyles(styles.statusBadge, styles.approved);
      case "pending":
        return mergeStyles(styles.statusBadge, styles.pending);
      case "rejected":
        return mergeStyles(styles.statusBadge, styles.rejected);
      default:
        return styles.statusBadge;
    }
  };

  const columns: IColumn[] = [
    { key: "employee", name: "Employee", fieldName: "Employee", minWidth: 120, isResizable: true },
    { key: "leaveType", name: "Leave Type", fieldName: "LeaveType", minWidth: 100, isResizable: true },
    { key: "startDate", name: "Start Date", fieldName: "StartDate", minWidth: 100, isResizable: true },
    { key: "endDate", name: "End Date", fieldName: "EndDate", minWidth: 100, isResizable: true },
    {
      key: "status",
      name: "Status",
      minWidth: 100,
      onRender: (item: ILeaveRequest) => (
        <span className={statusBadgeClass(item.Status)}>{item.Status}</span>
      ),
      isResizable: true,
    },
  ];

  return (
    <div className={styles.leaveTracker}>
      <h2 className={styles.title}>Leave Tracker</h2>

      {loading ? (
        <div>Loading...</div>
      ) : (
        <>
          <DetailsList
            items={leaves}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            styles={{ root: { overflowX: "auto" } }}
          />
          {leaves.length === 0 && (
            <div className={styles.noData}>No leave records found.</div>
          )}
        </>
      )}
    </div>
  );
};
