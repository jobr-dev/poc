import * as React from "react";
import { Opportunity } from "../models/Opportunity";
import {
  ContextualMenuItemType,
  DirectionalHint,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardDetails,
  DocumentCardImage,
  DocumentCardPreview,
  IconButton,
  ImageFit,
  Label,
  Stack,
} from "@fluentui/react";
import { OpportunityFile } from "../models/OpportunityFile";
import { BrandIcons } from "./BrandIcons";

export default class CardHelper {
  public static clientToCard(
    client: string,
    redirectFunc: (opp: Opportunity) => void
  ): React.ReactElement {
    return (
      <DocumentCard
        styles={{
          root: {
            display: "inline-block",
            marginRight: 20,
            marginBottom: 20,
            width: 170,
            cursor: "pointer",
          },
        }}
        onClick={() => {
          const opp = new Opportunity();
          opp.url = client;
          opp.userHasAccess = true;
          redirectFunc(opp);
        }}
      >
        <Stack horizontal>
          <Stack.Item grow={11}>
            <DocumentCardActivity
              activity={""}
              people={[
                {
                  name: client,
                  profileImageSrc: ``,
                },
              ]}
            />
          </Stack.Item>
        </Stack>
        <DocumentCardImage
          height={150}
          imageFit={ImageFit.cover}
          iconProps={{
            iconName: "Teamwork",
            styles: {
              root: {
                color: "inherit",
                fontSize: "60px",
                width: "60px",
                height: "60px",
              },
            },
          }}
        />
        <DocumentCardDetails>
          <Label styles={{ root: { marginLeft: "10px" } }} disabled>
            {client}
          </Label>
        </DocumentCardDetails>
      </DocumentCard>
    );
  }

  public static opportunityToCard(
    opp: Opportunity,
    refreshFunc: (opp: Opportunity) => void,
    redirectFunc: (opp: Opportunity) => void
  ): React.ReactElement {
    return (
      <DocumentCard
        styles={{
          root: {
            display: "inline-block",
            marginRight: 20,
            marginBottom: 20,
            width: 220,
            height: 240,
            cursor: "pointer",
          },
        }}
        onClick={() => {
          redirectFunc(opp);
        }}
      >
        <Stack horizontal>
          <Stack.Item grow={11}>
            <DocumentCardActivity
              activity={""}
              people={[
                {
                  name: opp.title,
                  profileImageSrc: ``,
                },
              ]}
            />
          </Stack.Item>
          <Stack.Item>
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
          </Stack.Item>
        </Stack>
        <DocumentCardImage
          height={120}
          imageFit={ImageFit.cover}
          iconProps={{
            iconName: "SharepointLogo",
            imageProps: {
              src: "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/sharepoint_48x1.svg",
            },
            styles: {
              root: {
                color: "inherit",
                fontSize: "100px",
                width: "100px",
                height: "100px",
              },
            },
          }}
        />
        {opp.activities["value"].length > 0 ? (
          <>
            <DocumentCardActivity
              styles={{ root: {} }}
              activity={CardHelper._relativeDate(
                opp.activities["value"][0]["times"]["recordedDateTime"]
              )}
              people={opp.activities["value"].map((x) => {
                return {
                  name: x.actor.user.displayName,
                  profileImageSrc: `/_layouts/15/userphoto.aspx?size=S&username=${x.actor.user["userPrincipalName"]}`,
                };
              })}
            />
          </>
        ) : (
          <DocumentCardActivity
            styles={{ root: {} }}
            activity={""}
            people={[
              {
                name: "",
                profileImageSrc: "",
              },
            ]}
          />
        )}
        <DocumentCardDetails>
          <Label styles={{ root: { marginLeft: "10px" } }} disabled>
            {opp.client}
          </Label>
        </DocumentCardDetails>
      </DocumentCard>
    );
  }

  public static fileToCard(
    file: OpportunityFile,
    refreshFunc: (oppFile: OpportunityFile) => void,
    redirectFunc: (opp: Opportunity) => void,
    notifyMembers: (opp: OpportunityFile) => void
  ): React.ReactElement {
    return (
      <DocumentCard
        styles={{
          root: {
            display: "inline-block",
            marginRight: 20,
            marginBottom: 20,
            width: 220,
            height: 240,
            cursor: "pointer",
            overflowWrap: "unset",
          },
        }}
        onClick={() => {
          window.open(file.editUrl, "_blank");
        }}
      >
        <Stack horizontal>
          <Stack.Item grow={11}>
            <DocumentCardActivity
              activity={CardHelper._relativeDate(file.lastAccessed)}
              people={[
                {
                  name: file.lastModifiedUserDisplayName,
                  profileImageSrc: `/_layouts/15/userphoto.aspx?size=S&username=${file.lastModifiedUserLoginName}`,
                },
              ]}
            />
          </Stack.Item>
          <Stack.Item>
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
          </Stack.Item>
        </Stack>
        <DocumentCardPreview
          previewImages={[
            {
              previewImageSrc: `${file.opportunityUrl}/_layouts/15/getpreview.ashx?path=${file.url}`,
              imageFit: ImageFit.cover,
              iconSrc: BrandIcons[file.type],
              height: 130,
              name: file.name,
            },
          ]}
        />
        <img
          src={BrandIcons[file.type]}
          style={{
            display: "inlineBlock",
            maxWidth: "15%",
            float: "right",
            margin: "5px",
          }}
        />
        <DocumentCardDetails>
          <Label
            styles={{
              root: {
                marginLeft: "10px",
                fontSize: "1vw",
                overflow: "hidden",
                whiteSpace: "nowrap",
              },
            }}
          >
            {file.name}
          </Label>
          <Label styles={{ root: { marginLeft: "10px" } }} disabled>
            {file.opportunityName}
          </Label>
        </DocumentCardDetails>
      </DocumentCard>
    );
  }

  private static getTimeStrings(): any {
    return {
      DateJustNow: "just now",
      DateMinute: "1 minute ago",
      DateMinutesAgo: "minutes ago",
      DateHour: "1 hour ago",
      DateHoursAgo: "hours ago",
      DateDay: "yesterday",
      DateDaysAgo: "days ago",
      DateWeeksAgo: "weeks ago",
    };
  }
  private static _relativeDate(crntDate: string): string {
    const timeStrings = CardHelper.getTimeStrings();
    const date = new Date(
      (crntDate || "").replace(/-/g, "/").replace(/[TZ]/g, " ")
    );

    const compareDate = new Date();

    compareDate.setHours(
      compareDate.getHours() + compareDate.getTimezoneOffset() / 60
    );

    const diff =
      (compareDate.getTime() - new Date(date.toLocaleString()).getTime()) /
      1000;
    const day_diff = Math.floor(diff / 86400);

    if (isNaN(day_diff) || day_diff < 0) {
      return;
    }

    return (
      (day_diff === 0 &&
        ((diff < 60 && timeStrings.DateJustNow) ||
          (diff < 120 && timeStrings.DateMinute) ||
          (diff < 3600 &&
            `${Math.floor(diff / 60)} ${timeStrings.DateMinutesAgo}`) ||
          (diff < 7200 && timeStrings.DateHour) ||
          (diff < 86400 &&
            `${Math.floor(diff / 3600)} ${timeStrings.DateHoursAgo}`))) ||
      (day_diff == 1 && timeStrings.DateDay) ||
      (day_diff <= 30 && `${day_diff} ${timeStrings.DateDaysAgo}`) ||
      (day_diff > 30 &&
        `${Math.ceil(day_diff / 7)} ${timeStrings.DateWeeksAgo}`)
    );
  }
}
