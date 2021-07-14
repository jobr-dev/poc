import {
  Checkbox,
  DocumentCardActivity,
  Panel,
  PrimaryButton,
  Stack,
} from "@fluentui/react";
import { Web } from "@pnp/sp/webs";
import { DefaultButton, Label } from "office-ui-fabric-react";
import * as React from "react";
import { OpportunityFile } from "../../models/OpportunityFile";
import * as msTeams from "@microsoft/teams-js";

export interface NotifyMembersPanelProps {
  visible: boolean;
  oppFile: OpportunityFile;
  dismissPanel: () => void;
}

const NotifyMembersPanel: React.FunctionComponent<NotifyMembersPanelProps> = (
  props
) => {
  const [users, setUsers] = React.useState([]);

  React.useEffect(() => {
    if (props.oppFile) {
      const getUsers = async () => {
        const web = Web(props.oppFile.opportunityUrl);
        const siteUsers = await web.siteUsers.get();
        const currentUser = await web.currentUser.get();
        const chatMembers = siteUsers.filter(
          (x) =>
            x.Email != "" &&
            x.PrincipalType == 1 &&
            x.Email != currentUser.Email
        );
        setUsers(
          chatMembers.map((x) => {
            return {
              selected: true,
              title: x.Title,
              email: x.UserPrincipalName,
            };
          })
        );
      };

      getUsers();
    }
  }, [props.visible]);

  const anySelected = users.filter((x) => x.selected == true).length > 0;
  const allSelected =
    users.filter((x) => x.selected == true).length == users.length;

  return (
    <Panel
      headerText="Notify Members"
      isOpen={props.visible}
      onDismiss={() => props.dismissPanel()}
      closeButtonAriaLabel="Close"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <div>
          <DefaultButton
            styles={{ root: { marginRight: 8 } }}
            onClick={(_) => {
              setUsers(
                users.map((x) => {
                  return { ...x, selected: !anySelected };
                })
              );
            }}
          >
            {anySelected ? "Clear all" : "Select all"}
          </DefaultButton>
          <PrimaryButton
            styles={{ root: { marginRight: 8 } }}
            disabled={!anySelected}
            onClick={() =>
              msTeams.executeDeepLink(
                `https://teams.microsoft.com/l/chat/0/0?users=${users
                  .filter((x) => x.selected == true)
                  .map((x) => x.email)
                  .join(",")}${
                  allSelected
                    ? `&topicName=${props.oppFile.opportunityName}`
                    : ""
                }&message=${encodeURI(props.oppFile.url.replace(" ", "%20"))}`
              )
            }
          >
            Chat
          </PrimaryButton>
        </div>
      )}
    >
      {props.oppFile ? (
        <Label disabled style={{ marginTop: "20px", marginBottom: "20px" }}>
          Select users and start a conversation about {props.oppFile.name}!
        </Label>
      ) : null}
      {users.map((x) => {
        return (
          <Stack horizontal tokens={{ childrenGap: 5 }}>
            <Stack.Item grow={1}>
              <Checkbox
                checked={x.selected}
                onChange={() => {
                  const newList = users.map((user) => {
                    if (x.email === user.email) {
                      return {
                        ...x,
                        selected: !x.selected,
                      };
                    }
                    return user;
                  });
                  setUsers(newList);
                }}
                label=""
                title=""
                styles={{ root: { margin: "10px" } }}
              />
            </Stack.Item>

            <Stack.Item grow={6}>
              <DocumentCardActivity
                styles={{ root: {} }}
                activity={""}
                people={[
                  {
                    name: x.title,
                    profileImageSrc: `/_layouts/15/userphoto.aspx?size=S&username=${x.email}`,
                  },
                ]}
              />
            </Stack.Item>
          </Stack>
        );
      })}
    </Panel>
  );
};

export default NotifyMembersPanel;
