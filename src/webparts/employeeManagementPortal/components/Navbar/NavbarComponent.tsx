import * as React from "react";
import { Nav, INavLinkGroup } from "@fluentui/react";
import { useState } from "react";

export interface INavbarProps {
  onSelectPage: (page: string) => void;
}

export const Navbar: React.FC<INavbarProps> = ({ onSelectPage }) => {
  const [activeMenu, setActiveMenu] = useState<string>("directory");
  const navLinkGroups: INavLinkGroup[] = [
    {
      links: [
        {
          key: "directory",
          name: "Employee Directory",
          url: "#",
          icon: "Contact",
        },
        {
          key: "leave",
          name: "Leave Tracker",
          url: "#",
          icon: "Calendar",
        },
        {
          key: "announcements",
          name: "Announcements",
          url: "#",
          icon: "Megaphone",
        },
      ],
    },
  ];

  return (
    <Nav
      groups={navLinkGroups}
      selectedKey={activeMenu}
      styles={{
        root: {
          width: "180px",
          height: "100%",
          boxSizing: "border-box",
          borderRight: "1px solid #ddd",
          overflowY: "auto",
        },
      }}
      onLinkClick={(ev, item) => {
        if (item) {
          setActiveMenu(item.key as string);
          onSelectPage(item.key as string);
        }
      }}
    />
  );
};
