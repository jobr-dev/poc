import * as React from "react";
import {
  CommandBarButton,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IStackStyles,
  PrimaryButton,
  ProgressIndicator,
  Stack,
  TextField,
} from "@fluentui/react";
import { LocalFile } from "../../../models/LocalFile";
import FileService from "../../../services/fileService";
import Mapper from "../../../utils/Mapper";
import { useBoolean } from "@fluentui/react-hooks";

export interface BrowseFileViewCommands {
  selectedOpportunity: string;
  refreshFileList: () => void;
  onNewFilesToUpload: (files: LocalFile[]) => void;
  currentFolder: string;
  fileService: FileService;
}

const stackStyles: Partial<IStackStyles> = { root: { height: 44 } };

const BrowseFileViewCommands: React.FunctionComponent<BrowseFileViewCommands> =
  (props) => {
    const [isLoading, { setTrue: loading, setFalse: notLoading }] =
      useBoolean(false);
    const [dialogOpen, { setTrue: showDialog, setFalse: closeDialog }] =
      useBoolean(false);
    const [folderName, setFolderName] = React.useState("");
    var inputFileUploadRef = React.useRef<HTMLInputElement>();

    const newFile = async (fileExtension: string) => {
      loading();
      await props.fileService.addNewFileToOpportunity(
        props.selectedOpportunity,
        fileExtension,
        props.currentFolder
      );

      props.refreshFileList();
      notLoading();
    };

    const newFolderDialog = async () => {
      showDialog();
    };

    const newFolder = async () => {
      if (folderName == "") return;

      loading();
      await props.fileService.addNewFolderToOpportunity(
        props.selectedOpportunity,
        folderName,
        props.currentFolder
      );

      setFolderName("");
      closeDialog();
      props.refreshFileList();
      notLoading();
    };

    const uploadFileChange = async (e: any) => {
      var fileList =
        e.target !== undefined ? (e.target.files as FileList) : (e as FileList);

      var result: LocalFile[] = Mapper.inputFilesToLocalFiles(fileList);

      if (result.length > 0) {
        props.onNewFilesToUpload(result);
      }
    };

    const menuProps = {
      items: [
        {
          key: "word",
          text: "Word document",
          onClick: () => {
            newFile(".docx");
          },
          iconProps: {
            imageProps: {
              src: "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001//assets/item-types/16/docx.svg",
            },
          },
        },
        {
          key: "excel",
          text: "Excel workbook",
          onClick: () => {
            newFile(".xlsx");
          },
          iconProps: {
            imageProps: {
              src: "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001//assets/item-types/16/xlsx.svg",
            },
          },
        },
        {
          key: "presentation",
          text: "PowerPoint presentation",
          onClick: () => {
            newFile(".pptx");
          },
          iconProps: {
            imageProps: {
              src: "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001//assets/item-types/16/pptx.svg",
            },
          },
        },
        {
          key: "folder",
          text: "Folder",
          onClick: () => {
            newFolderDialog();
          },
          iconProps: {
            imageProps: {
              src: "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001//assets/item-types/16/folder.svg",
            },
          },
        },
      ],
    };
    return (
      <>
        <Stack horizontal styles={stackStyles}>
          <CommandBarButton
            iconProps={{ iconName: "Add" }}
            text="New"
            menuProps={menuProps}
          />
          <CommandBarButton
            iconProps={{ iconName: "Upload" }}
            text="Upload"
            type="file"
            onClick={() => {
              if (null !== inputFileUploadRef.current) {
                inputFileUploadRef.current.click();
              }
            }}
          />
          <input
            style={{ display: "none" }}
            type="file"
            ref={inputFileUploadRef}
            onChange={(e) => uploadFileChange(e)}
            multiple={true}
          />
          <CommandBarButton
            iconProps={{ iconName: "Refresh" }}
            text="Refresh"
            onClick={async () => props.refreshFileList()}
          />
          <Dialog
            hidden={!dialogOpen}
            onDismiss={closeDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: "Create Folder",
              closeButtonAriaLabel: "Close",
            }}
          >
            <TextField
              label="Folder name"
              onChange={(_, val) => setFolderName(val)}
            />
            <DialogFooter>
              <PrimaryButton
                disabled={folderName == ""}
                onClick={newFolder}
                text="Create"
              />
              <DefaultButton onClick={closeDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </Stack>
        {isLoading ? (
          <ProgressIndicator />
        ) : (
          <ProgressIndicator
            styles={{
              root: {
                visibility: "hidden",
              },
            }}
          />
        )}
      </>
    );
  };

export default BrowseFileViewCommands;
