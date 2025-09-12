import * as React from "react";
import { useState, useEffect } from "react";
import { IconButton, Callout, Persona, PersonaSize, Stack, Text } from "@fluentui/react";
import { EmployeeHttpService } from "../../services/EmployeeHttpService";

interface IUserProfile {
  DisplayName: string;
  Email: string;
  JobTitle: string;
  UserUrl: string;
}

interface IUserProfileProps {
  employeeHttpService: EmployeeHttpService;
}

const UserProfileMenu: React.FC<IUserProfileProps> = ({ employeeHttpService }) => {
  const [user, setUser] = useState<IUserProfile | null>(null);
  const [showCallout, setShowCallout] = useState<boolean>(false);
  const [target, setTarget] = useState<HTMLElement | null>(null);

  // Fetch user profile data
  useEffect(() => {
    const fetchUser = async () => {
      try {
        const data = await employeeHttpService.getCurrentUserDetails();
        const UserUrl = data?.UserUrl || "/_layouts/15/images/person.gif";

        setUser({
          DisplayName: data.DisplayName,
          Email: data.Email,
          JobTitle: data.Title || "Software Engineer",
          UserUrl: UserUrl,
        });
      } catch (error) {
        console.error("Error fetching user profile:", error);
      }
    };

    void fetchUser();
  }, []);

  return (
    <div style={{ position: "absolute", top: 40, right: 20 }}>
      <IconButton
        id="userProfileBtn"
        styles={{ root: { borderRadius: "50%", overflow: "hidden", backgroundColor: "white"} }}
        onClick={(e) => {
          setTarget(e.currentTarget as HTMLElement);
          setShowCallout(!showCallout);
        }}
      >
        {user?.UserUrl && (
          <img
            src={user.UserUrl}
            alt="User"
            style={{ width: "100%", height: "100%", borderRadius: "50%" }}
            className="ms-Image-image ms-Image-image--cover ms-Image-image--portrait is-notLoaded is-fadeIn is-error image-230"
          />
        )}
      </IconButton>

      {/* Callout with user details */}
      {showCallout && target && (
        <Callout
          target={target}
          onDismiss={() => setShowCallout(false)}
          setInitialFocus
          directionalHint={4}
          styles={{
            root: { padding: 12, borderRadius: 8, boxShadow: "0 2px 8px rgba(0,0,0,0.2)" },
          }}
        >
          {user && (
            <Stack tokens={{ childrenGap: 8 }}>
              <Persona
                text={user.DisplayName}
                secondaryText={user.JobTitle}
                tertiaryText={user.Email}
                size={PersonaSize.size56}
                imageUrl={user.UserUrl}
              />
              <Text variant="small" styles={{ root: { fontStyle: "italic", marginTop: 4 } }}>
                {user.Email}
              </Text>
            </Stack>
          )}
        </Callout>
      )}
    </div>
  );
};

export default UserProfileMenu;
