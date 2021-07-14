import * as React from "react";
import {
  Breadcrumb,
  IBreadcrumbItem,
  MessageBar,
  MessageBarType,
  SelectionMode,
} from "@fluentui/react";
import { IViewField, ListView } from "@pnp/spfx-controls-react/lib/ListView";
import { useBoolean } from "@fluentui/react-hooks";
import FileService from "../../../services/fileService";
import BrowseFileViewCommands from "./BrowseFileViewCommands";
import BrowseFileViewUploadPanel from "./BrowseFileViewUploadPanel";
import BrowseFileViewContextMenu, {
  BrowseFileViewContextMenuProps,
} from "./BrowseFileViewContextMenu";
import { LocalFile } from "../../../models/LocalFile";
import styles from "./BrowseFileView.module.scss";
import { OpportunityFile } from "../../../models/OpportunityFile";
import Mapper from "../../../utils/Mapper";
import {
  FileTypeIcon,
  IconType,
} from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { BrandIcons } from "../../../utils/BrandIcons";

export interface BrowseFileViewProps {
  selectedOpportunity: string;
  hasAccess: boolean;
  changePinForOpportunityFile: (oppFile: OpportunityFile) => void;
  fileService: FileService;
  notifyMembers: (oppFile: OpportunityFile) => void;
}

const BrowseFileView: React.FunctionComponent<BrowseFileViewProps> = (
  props
) => {
  const [opportunityFiles, setOpportunityFiles] = React.useState<
    OpportunityFile[]
  >([]);
  const [currentFolder, setcurrentFolder] = React.useState("/Shared Documents");

  const [filesToBeUploaded, setUploadingFiles] = React.useState<LocalFile[]>(
    []
  );
  const [breadcrumbItems, setBreadcrumbItems] = React.useState<
    IBreadcrumbItem[]
  >([
    {
      text: "Documents",
      key: "/Shared Documents",
      onClick: (e, item) => breadcrumbClicked(e, item),
    },
  ]);
  const [uploadPanelOpen, { setTrue: openPanel, setFalse: closePanel }] =
    useBoolean(false);

  const [errorIsShowing, { setTrue: showError, setFalse: dismissError }] =
    useBoolean(false);

  const refreshFiles = async () => {
    try {
      var files = await props.fileService.getFilesFromOpportunity(
        props.selectedOpportunity
      );
      setOpportunityFiles(files);
    } catch {
      setOpportunityFiles([]);
    }
  };

  const onNewFilesToUpload = (files: LocalFile[]) => {
    setUploadingFiles(files);
    openPanel();
  };
  const dismissPanel = async () => {
    closePanel();
    setUploadingFiles([]);
    await refreshFiles();
  };

  const showErrorMessage = () => {
    showError();
    setTimeout(dismissError, 3500);
  };

  const breadcrumbClicked = (_: any, item: IBreadcrumbItem) => {
    const currentBreadcrumb = breadcrumbItems;

    setcurrentFolder(item.key);
    setBreadcrumbItems(
      currentBreadcrumb.filter((x) => x.key.length <= item.key.length)
    );
  };

  React.useEffect(() => {
    setcurrentFolder("/Shared Documents");
    refreshFiles();
  }, [props.selectedOpportunity]);

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
        if (item.type == "Folder") {
          return (
            <a
              style={{ cursor: "pointer", fontSize: "14px" }}
              onClick={() => {
                setcurrentFolder(`${currentFolder}/${item.name}`);
                const current = breadcrumbItems;
                current.push({
                  text: item.name,
                  key: `${currentFolder}/${item.name}`,
                  onClick: (e, item) => breadcrumbClicked(e, item),
                });
                setBreadcrumbItems(current);
              }}
            >
              {item.name}
            </a>
          );
        } else {
          return (
            <a target="_blank" href={item.editUrl} style={{ fontSize: "14px" }}>
              {item.name}
            </a>
          );
        }
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
      render: (item: OpportunityFile) => {
        const element: React.ReactElement<BrowseFileViewContextMenuProps> =
          React.createElement(BrowseFileViewContextMenu, {
            item: item,
            selectedOpportunity: props.selectedOpportunity,
            refreshFiles: () => refreshFiles(),
            showError: () => showErrorMessage(),
            changePinForOpportunityFile: props.changePinForOpportunityFile,
            currentFolder: currentFolder.slice(1),
            fileService: props.fileService,
            notifyMembers: props.notifyMembers,
          });
        return element;
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
    <>
      {props.hasAccess ? (
        <>
          {errorIsShowing ? (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={dismissError}
              dismissButtonAriaLabel="Close"
            >
              An error occured. Try deleting the file later...
            </MessageBar>
          ) : null}
          <BrowseFileViewCommands
            selectedOpportunity={props.selectedOpportunity}
            refreshFileList={refreshFiles}
            onNewFilesToUpload={(files: LocalFile[]) => {
              onNewFilesToUpload(files);
            }}
            currentFolder={currentFolder.slice(1)}
            fileService={props.fileService}
          />
          <Breadcrumb
            items={breadcrumbItems}
            maxDisplayedItems={3}
            ariaLabel="Folder structure"
            overflowAriaLabel="More folders"
          />
          <div className={styles.base}>
            <ListView
              items={opportunityFiles.filter((x) => {
                return x.url.split(currentFolder).pop() == `/${x.name}`;
              })}
              viewFields={viewFields}
              compact={true}
              selectionMode={SelectionMode.none}
              //showFilter={true}
              filterPlaceHolder="Search..."
              dragDropFiles={true}
              onDrop={(e: any) => {
                onNewFilesToUpload(Mapper.inputFilesToLocalFiles(e));
              }}
              stickyHeader={true}
            />
          </div>
          <BrowseFileViewUploadPanel
            selectedOpportunity={props.selectedOpportunity}
            filesToBeUploaded={filesToBeUploaded}
            dismissPanel={dismissPanel}
            panelVisible={uploadPanelOpen}
            fileService={props.fileService}
            folder={currentFolder.substring(1)}
          />
        </>
      ) : (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          You do not have access to this opportunity.
        </MessageBar>
      )}
    </>
  );
};

export default BrowseFileView;
