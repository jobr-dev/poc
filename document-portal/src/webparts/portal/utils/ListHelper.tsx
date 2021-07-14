import * as React from "react";
import { Opportunity } from "../models/Opportunity";
import {
  ContextualMenuItemType,
  DirectionalHint,
  IconButton,
  SelectionMode,
} from "@fluentui/react";
import { OpportunityFile } from "../models/OpportunityFile";
import { IViewField, ListView } from "@pnp/spfx-controls-react/lib/ListView";
import {
  FileTypeIcon,
  IconType,
} from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { BrandIcons } from "./BrandIcons";

export default class ListHelper {
  public static clientsToList(
    clients: string[],
    redirectFunc: (opp: Opportunity) => void
  ): React.ReactElement {
    const viewFields: IViewField[] = [
      {
        name: "name",
        displayName: "Name",
        linkPropertyName: "name",
        render: (item: any) => {
          return (
            <a
              onClick={() => {
                const opp = new Opportunity();
                opp.url = item["name"];
                opp.userHasAccess = true;
                redirectFunc(opp);
              }}
              style={{ cursor: "pointer", fontSize: "14px" }}
            >
              {item["name"]}
            </a>
          );
        },
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        sorting: true,
      },
      {
        name: " ",
        sorting: false,
        render: (_) => {
          return (
            <IconButton
              id="ContextualMenu"
              text=""
              split={false}
              style={{ visibility: "hidden" }}
              menuIconProps={{ iconName: "MoreVertical" }}
              menuProps={{
                shouldFocusOnMount: true,
                directionalHint: DirectionalHint.bottomRightEdge,
                items: [],
              }}
            />
          );
        },
      },
    ];
    return (
      <ListView
        items={clients.map((x) => {
          return { name: x };
        })}
        viewFields={viewFields}
        compact={true}
        selectionMode={SelectionMode.none}
        dragDropFiles={false}
        stickyHeader={true}
      />
    );
  }

  public static opportunitiesToList(
    opps: Opportunity[],
    refreshFunc: (opp: Opportunity) => void,
    redirectFunc: (opp: Opportunity) => void
  ): React.ReactElement {
    const viewFields: IViewField[] = [
      {
        name: "name",
        displayName: "Name",
        linkPropertyName: "title",
        render: (opp: Opportunity) => {
          return (
            <a
              onClick={() => redirectFunc(opp)}
              style={{ cursor: "pointer", fontSize: "14px" }}
            >
              {opp.title}
            </a>
          );
        },
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        sorting: true,
      },
      {
        name: " ",
        sorting: false,
        render: (opp: Opportunity) => {
          return (
            <IconButton
              id="ContextualMenu"
              text=""
              split={false}
              menuIconProps={{ iconName: "MoreVertical" }}
              menuProps={{
                shouldFocusOnMount: true,
                directionalHint: DirectionalHint.bottomRightEdge,
                items: [
                  {
                    key: "view",
                    name: "Go to opportunity",
                    iconProps: { iconName: "View" },
                    onClick: () => redirectFunc(opp),
                  },
                  {
                    key: "pin",
                    name: opp.pinned ? "Unpin" : "Pin",
                    iconProps: { iconName: opp.pinned ? "Unpin" : "Pin" },
                    onClick: () => {
                      refreshFunc(opp);
                    },
                  },
                ],
              }}
            />
          );
        },
      },
    ];
    return (
      <ListView
        items={opps}
        viewFields={viewFields}
        compact={true}
        selectionMode={SelectionMode.none}
        //showFilter={true}
        filterPlaceHolder="Search..."
        dragDropFiles={false}
        stickyHeader={true}
      />
    );
  }

  public static filesToList(
    files: OpportunityFile[],
    refreshFunc: (oppFile: OpportunityFile) => void,
    redirectFunc: (opp: Opportunity) => void,
    notifyMembers: (opp: OpportunityFile) => void
  ): React.ReactElement {
    const viewFields: IViewField[] = [
      {
        name: "",
        displayName: "",
        linkPropertyName: "url",
        render: (item: OpportunityFile) => {
          const img = BrandIcons[item.type];
          return <img src={img} style={{ maxWidth: "18px" }} />;
        },
        minWidth: 20,
        maxWidth: 20,
        isResizable: false,
        sorting: false,
      },
      {
        name: "name",
        displayName: "Name",
        linkPropertyName: "name",
        render: (item: OpportunityFile) => {
          return (
            <a
              target="_blank"
              href={item.editUrl}
              style={{ cursor: "pointer", fontSize: "14px" }}
            >
              {item.name}
            </a>
          );
        },
        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        sorting: true,
      },
      {
        name: " ",
        sorting: false,
        maxWidth: 40,
        render: (file: OpportunityFile) => {
          return (
            <IconButton
              id="ContextualMenu"
              text=""
              split={false}
              menuIconProps={{ iconName: "MoreVertical" }}
              menuProps={{
                shouldFocusOnMount: true,
                directionalHint: DirectionalHint.bottomRightEdge,
                items: [
                  {
                    key: "view",
                    name: "Go to opportunity",
                    iconProps: { iconName: "View" },
                    onClick: () => {
                      const opp = new Opportunity();
                      opp.id = file.opportunityUrl;
                      opp.url = file.opportunityUrl;
                      opp.userHasAccess = true;
                      redirectFunc(opp);
                    },
                  },
                  {
                    key: "edit",
                    name: "Edit",
                    iconProps: { iconName: "Edit" },
                    onClick: () => window.open(file.editUrl, "_blank"),
                  },
                  {
                    key: "pin",
                    name: file.pinned ? "Unpin" : "Pin",
                    iconProps: { iconName: file.pinned ? "Unpin" : "Pin" },
                    onClick: () => {
                      refreshFunc(file);
                    },
                  },
                  {
                    key: "divider_1",
                    itemType: ContextualMenuItemType.Divider,
                  },
                  {
                    key: "header",
                    name: "Members",
                    itemType: ContextualMenuItemType.Header,
                  },
                  {
                    key: "header",
                    name: "Notify members",
                    iconProps: { iconName: "Chat" },
                    onClick: () => {
                      notifyMembers(file);
                    },
                  },
                ],
              }}
            />
          );
        },
      },
      {
        name: "lastModifiedUserDisplayName",
        displayName: "Modified By",
        linkPropertyName: "lastModifiedUserDisplayName",
        render: (item: OpportunityFile) => {
          return (
            <span style={{ fontSize: "14px" }}>
              {item.lastModifiedUserDisplayName}
            </span>
          );
        },
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        sorting: true,
      },
      {
        name: "lastModified",
        displayName: "Modified",
        linkPropertyName: "lastModified",
        render: (item: OpportunityFile) => {
          return (
            <span style={{ fontSize: "14px" }}>
              {item.lastModified.replace("Z", "").replace("T", " ")}
            </span>
          );
        },
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        sorting: true,
      },
    ];
    return (
      <ListView
        items={files}
        viewFields={viewFields}
        compact={true}
        selectionMode={SelectionMode.none}
        //showFilter={true}
        filterPlaceHolder="Search..."
        dragDropFiles={false}
        stickyHeader={true}
      />
    );
  }
}
