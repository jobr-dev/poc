import * as React from "react";
import {
  createTheme,
  Label,
  Shimmer,
  ShimmerElementType,
  Spinner,
  SpinnerSize,
  Stack,
  Toggle,
} from "@fluentui/react";
import { Opportunity } from "../../models/Opportunity";
import BrowseTreeView from "./treeView/BrowseTreeView";
import BrowseFileView from "./fileView/BrowseFileView";
import { OpportunityFile } from "../../models/OpportunityFile";
import FileService from "../../services/fileService";

export interface BrowseProps {
  opportunities: Opportunity[];
  loading: boolean;
  selectedOpportunity: Opportunity;
  setShowAllOpps: (showAll: boolean) => void;
  showAllOpps: boolean;
  selectedOpportunityChanged: (opportunity: Opportunity) => void;
  changePinForOpportunity: (opp: Opportunity) => void;
  changePinForOpportunityFile: (oppFile: OpportunityFile) => void;
  fileService: FileService;
  notifyMembers: (oppFile: OpportunityFile) => void;
}

const Browse: React.FunctionComponent<BrowseProps> = (props) => {
  const [selectedOpportunity, setSelectedOpportunity] = React.useState(
    props.selectedOpportunity
  );

  const onTreeItemSelected = async (opportunity: Opportunity) => {
    setSelectedOpportunity(opportunity ? opportunity : null);
    props.selectedOpportunityChanged(opportunity);
  };
  const shimmer = [...Array(26).keys()].map((i) => i + 1);

  return (
    <div>
      <Stack
        horizontal
        horizontalAlign="space-between"
        styles={{ root: { overflowY: "auto", height: "80vh" } }}
        id="BrowseContainer"
      >
        <Stack.Item
          /* style={{ minWidth: "48vw", maxWidth: "48vw" }} */ grow={3}
        >
          <Stack horizontal styles={{ inner: { overflow: "hidden" } }}>
            <Stack.Item>
              <Label>Clients &gt; Opportunities</Label>
            </Stack.Item>
            <Stack.Item>
              {props.loading ? (
                <Spinner label="" size={SpinnerSize.medium} />
              ) : (
                <Toggle
                  label=""
                  disabled={props.loading}
                  checked={props.showAllOpps}
                  inlineLabel
                  offText="Your opportunities"
                  onText="All opportunities"
                  onChange={(_, checked) => {
                    props.setShowAllOpps(checked);
                  }}
                />
              )}
            </Stack.Item>
          </Stack>
          {props.loading ? (
            <>
              <div style={{ marginTop: "22px" }}>
                {shimmer.map((_) => {
                  return (
                    <Shimmer
                      styles={{
                        root: {
                          marginLeft: "20px",
                          maxWidth: "20%",
                        },
                      }}
                      theme={createTheme({
                        palette: {
                          neutralLight: (window as any).__themeState__.theme
                            .neutralLight,
                          neutralLighter: (window as any).__themeState__.theme
                            .neutralLighter,
                          white: (window as any).__themeState__.theme.white,
                        },
                      })}
                      shimmerElements={[
                        {
                          type: ShimmerElementType.line,
                          width: "100%",
                          height: 15,
                        },
                      ]}
                      style={{ marginTop: "10px" }}
                    />
                  );
                })}
              </div>
            </>
          ) : (
            <BrowseTreeView
              opportunities={props.opportunities}
              allOpps={props.showAllOpps}
              selectedOpportunityChanged={onTreeItemSelected}
              selectedOpportunity={props.selectedOpportunity.url}
              changePinForOpportunity={props.changePinForOpportunity}
            />
          )}
        </Stack.Item>
        <Stack.Item grow={3}>
          <div>
            {selectedOpportunity.url !== "" ? (
              <>
                <Label>Files</Label>
                <div style={{ padding: "15px" }}>
                  <BrowseFileView
                    selectedOpportunity={selectedOpportunity.url}
                    hasAccess={selectedOpportunity.userHasAccess}
                    changePinForOpportunityFile={
                      props.changePinForOpportunityFile
                    }
                    fileService={props.fileService}
                    notifyMembers={props.notifyMembers}
                  />
                </div>
              </>
            ) : null}
          </div>
        </Stack.Item>
      </Stack>
    </div>
  );
};

export default Browse;
