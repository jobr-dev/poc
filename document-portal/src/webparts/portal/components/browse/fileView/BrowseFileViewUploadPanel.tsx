import * as React from "react";
import {
  Checkbox,
  Panel,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import { LocalFile } from "../../../models/LocalFile";
import FileService from "../../../services/fileService";

export interface BrowseFileViewUploadPanelProps {
  selectedOpportunity: string;
  filesToBeUploaded: LocalFile[];
  dismissPanel: () => void;
  panelVisible: boolean;
  fileService: FileService;
  folder: string;
}

const BrowseFileViewUploadPanel: React.FunctionComponent<BrowseFileViewUploadPanelProps> =
  (props) => {
    const [
      uploadInProgress,
      { setTrue: uploadStarted, setFalse: uploadCompleted },
    ] = useBoolean(false);
    const [uploadingFiles, setUploadingFiles] = React.useState<LocalFile[]>([]);

    React.useEffect(() => {
      setUploadingFiles(props.filesToBeUploaded);
    }, [props.filesToBeUploaded]);

    const checkBoxOnChange = (file: LocalFile) => {
      const newList = uploadingFiles.map((item) => {
        if (item.file.name === file.file.name) {
          const updatedItem = {
            ...item,
            selected: !item.selected,
          };

          return updatedItem;
        }

        return item;
      });

      setUploadingFiles(newList);
    };

    const uploadLocalFiles = async () => {
      uploadStarted();
      var filesToUpload = uploadingFiles.filter((x) => x.selected);
      var result = await props.fileService.uploadFiles(
        props.selectedOpportunity,
        filesToUpload,
        props.folder
      );

      if (result) {
        props.dismissPanel();
      }
      uploadCompleted();
    };

    return (
      <Panel
        headerText="Upload"
        isOpen={props.panelVisible}
        onDismiss={props.dismissPanel}
        closeButtonAriaLabel="Close"
      >
        <div>
          {uploadingFiles.length > 0 ? (
            <div>
              {uploadingFiles.map((x) => {
                return (
                  <div>
                    <Checkbox
                      title={
                        x.file["fullPath"] && x.file["fullPath"].length > 0
                          ? x.file["fullPath"]
                          : x.file.name
                      }
                      label={
                        x.file["fullPath"] && x.file["fullPath"].length > 0
                          ? x.file["fullPath"]
                          : x.file.name
                      }
                      checked={x.selected}
                      onChange={() => checkBoxOnChange(x)}
                    />
                    <br />
                  </div>
                );
              })}
              {uploadInProgress ? (
                <Spinner label="Uploading files..." size={SpinnerSize.large} />
              ) : (
                <PrimaryButton onClick={uploadLocalFiles}>Upload</PrimaryButton>
              )}
            </div>
          ) : null}
        </div>
      </Panel>
    );
  };

export default BrowseFileViewUploadPanel;
