import * as React from "react";
import {
  ContextualMenuItemType,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IconButton,
  PrimaryButton,
} from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import FileService from "../../../services/fileService";
import { OpportunityFile } from "../../../models/OpportunityFile";

export interface BrowseFileViewContextMenuProps {
  item: OpportunityFile;
  selectedOpportunity: string;
  refreshFiles: () => void;
  showError: () => void;
  changePinForOpportunityFile: (oppFile: OpportunityFile) => void;
  currentFolder: string;
  fileService: FileService;
  notifyMembers: (oppFile: OpportunityFile) => void;
}

const BrowseFileViewContextMenu: React.FunctionComponent<BrowseFileViewContextMenuProps> =
  (props) => {
    const [dialogOpen, { setTrue: showDialog, setFalse: closeDialog }] =
      useBoolean(false);
    const [currentItem, setCurrentItem] = React.useState<OpportunityFile>();

    const handleDelete = (item: OpportunityFile) => {
      showDialog();
      setCurrentItem(item);
    };

    const handleConfirm = async () => {
      var result = await props.fileService.deleteFile(
        props.selectedOpportunity,
        currentItem.name,
        props.currentFolder
      );

      if (!result) {
        props.showError();
      } else if (props.item.pinned) {
        props.changePinForOpportunityFile(props.item);
      }

      closeDialog();
      props.refreshFiles();
    };

    return (
      <>
        <IconButton
          id="ContextualMenu"
          text=""
          width="30"
          split={false}
          iconProps={{ iconName: "MoreVertical" }}
          menuIconProps={{ iconName: "" }}
          menuProps={{
            shouldFocusOnMount: true,
            items: [
              {
                key: "pin",
                name: props.item.pinned === true ? "Unpin" : "Pin",
                iconProps: {
                  iconName: props.item.pinned ? "Unpin" : "Pin",
                },
                disabled: props.item.type == "Folder",
                onClick: () => {
                  props.changePinForOpportunityFile(props.item);
                  props.refreshFiles();
                },
              },
              {
                key: "delete",
                name: "Delete",
                iconProps: { iconName: "Delete" },
                onClick: () => handleDelete(props.item),
              },
              {
                key: "divider_1",
                itemType: ContextualMenuItemType.Divider,
                disabled: props.item.type == "Folder",
              },
              {
                key: "header",
                name: "Members",
                itemType: ContextualMenuItemType.Header,
                disabled: props.item.type == "Folder",
              },
              {
                key: "header",
                name: "Notify members",
                iconProps: { iconName: "Chat" },
                disabled: props.item.type == "Folder",
                onClick: () => {
                  props.notifyMembers(props.item);
                },
              },
            ],
          }}
        />
        <Dialog
          hidden={!dialogOpen}
          onDismiss={closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Delete",
            closeButtonAriaLabel: "Close",
            subText: "Are you sure you want to delete?",
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={handleConfirm} text="Delete" />
            <DefaultButton onClick={closeDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>
      </>
    );
  };

export default BrowseFileViewContextMenu;
