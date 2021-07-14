import * as React from "react";
import {
  ITreeItem,
  TreeItemActionsDisplayMode,
  TreeView,
  TreeViewSelectionMode,
} from "@pnp/spfx-controls-react/lib/TreeView";
import { Opportunity } from "../../../models/Opportunity";
import styles from "./BrowseTreeView.module.scss";

export interface BrowseTreeViewProps {
  opportunities: Opportunity[];
  allOpps: boolean;
  selectedOpportunityChanged: (opportunity: Opportunity) => void;
  selectedOpportunity: string;
  changePinForOpportunity: (opp: Opportunity) => void;
}

const BrowseTreeView: React.FunctionComponent<BrowseTreeViewProps> = (
  props
) => {
  const onTreeItemSelected = (items: ITreeItem[]) => {
    if (items[0] == undefined) return;
    props.selectedOpportunityChanged(items[0].data);
    var b = document.getElementById("BrowseContainer");
    b.scrollTop = 0;
  };

  const pinAction = (opportunity: Opportunity) => {
    return {
      title: opportunity.pinned == true ? "Unpin" : "Pin",
      iconProps: {
        iconName: opportunity.pinned ? "Unpin" : "Pin",
      },
      hidden: !opportunity.userHasAccess,
      id: "pinItem",
      actionCallback: (treeItem: ITreeItem) => {
        props.changePinForOpportunity(treeItem.data);
      },
    };
  };
  const goToSite = (opportunity: Opportunity) => {
    return {
      title: "Go to document library",
      iconProps: {
        iconName: "DocLibrary",
      },
      id: "gotosite",
      hidden: !opportunity.userHasAccess,
      actionCallback: (treeItem: ITreeItem) => {
        window.open(`${treeItem.key}/Shared%20Documents/`, "_blank");
      },
    };
  };

  const renderTreeViewItems = (opportunities: Opportunity[]): ITreeItem[] => {
    var treeViewItems: ITreeItem[] = [];

    opportunities.forEach((opportunity) => {
      var existingLetters = treeViewItems.filter(
        (x) => x.label == opportunity.client[0]
      );
      if (existingLetters.length == 0) {
        treeViewItems.push({
          key: opportunity.client[0],
          label: opportunity.client[0],
          selectable: false,
          children: [
            {
              key: opportunity.client,
              label: opportunity.client,
              selectable: false,
              children: [
                {
                  key: opportunity.url,
                  label: opportunity.title,
                  data: opportunity,
                  actions: [pinAction(opportunity), goToSite(opportunity)],
                },
              ],
            },
          ],
        });
      } else {
        var existingClient = existingLetters[0].children.filter(
          (x) => x.label == opportunity.client
        );

        if (existingClient.length == 0) {
          existingLetters[0].children.push({
            key: opportunity.client,
            label: opportunity.client,
            selectable: false,
            children: [
              {
                key: opportunity.url,
                label: opportunity.title,
                data: opportunity,
                actions: [pinAction(opportunity), goToSite(opportunity)],
              },
            ],
          });
        } else {
          existingClient[0].children.push({
            key: opportunity.url,
            label: opportunity.title,
            data: opportunity,
            actions: [pinAction(opportunity), goToSite(opportunity)],
          });
        }
      }
    });

    return treeViewItems;
  };

  return (
    <div className={styles.base}>
      <TreeView
        items={renderTreeViewItems(props.opportunities)}
        selectionMode={TreeViewSelectionMode.Single}
        defaultSelectedKeys={
          props.selectedOpportunity !== "" ? [props.selectedOpportunity] : null
        }
        selectChildrenIfParentSelected={false}
        showCheckboxes={false}
        treeItemActionsDisplayMode={TreeItemActionsDisplayMode.ContextualMenu}
        expandToSelected={true}
        onSelect={onTreeItemSelected}
        defaultExpandedChildren={false}
      />
    </div>
  );
};

export default BrowseTreeView;
