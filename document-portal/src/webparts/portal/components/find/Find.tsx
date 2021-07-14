import * as React from "react";
import {
  Checkbox,
  Label,
  SearchBox,
  Slider,
  Spinner,
  SpinnerSize,
  Stack,
} from "@fluentui/react";
import CardHelper from "../../utils/CardHelper";
import { Opportunity } from "../../models/Opportunity";
import { OpportunityFile } from "../../models/OpportunityFile";
import styles from "./Find.module.scss";
import SearchService from "../../services/searchService";
import { SearchSettings } from "../../models/SearchSettings";
import { ChoiceGroup } from "office-ui-fabric-react";
import LayoutSelector from "../shared/LayoutSelector";
import ListHelper from "../../utils/ListHelper";
const img: string = require("../../assets/Illustration.png");

export interface FindProps {
  opportunities: Opportunity[];
  clients: string[];
  pinnedFiles: OpportunityFile[];
  gridViewLayout: boolean;
  changeLayout: (isGridView: boolean) => void;
  redirectFunc: (opp: Opportunity) => void;
  changePinForOpportunity: (opp: Opportunity) => void;
  changePinForOpportunityFile: (oppFile: OpportunityFile) => void;
  searchService: SearchService;
  searchSettings: SearchSettings;
  updateSearchSettings: (currentSettings: SearchSettings) => void;
  notifyMembers: (oppFile: OpportunityFile) => void;
}

const Find: React.FunctionComponent<FindProps> = (props) => {
  const [searchFiles, setSearchFiles] = React.useState<OpportunityFile[]>([]);
  const [isSearching, setSearching] = React.useState<boolean>(false);
  const [lastEditedBy, setlastEditedBy] = React.useState([]);
  const [filterOpps, setfilterOpps] = React.useState([]);
  const [fileTypes, setFileTypeFilter] = React.useState([]);
  const [selectedLastEditedBy, setSelectedLastEditedBy] = React.useState("All");
  const [selectedFilterOpp, setSelectedFilterOpp] = React.useState("All");
  const [selectedFileType, setSelectedFileType] = React.useState("All");

  const searchByText = async (query: string) => {
    if (
      props.searchSettings.searchFiles &&
      props.searchSettings.searchQuery.length > 0
    ) {
      setSearching(true);
      var foundFiles = await props.searchService.search(
        query,
        props.searchSettings.numberOfResults
      );
      foundFiles.forEach((x) => {
        x.pinned = props.pinnedFiles.some((pinned) => x.url == pinned.url);
      });

      setlastEditedBy([
        ...new Set(foundFiles.map((x) => x.lastModifiedUserDisplayName)),
        "All",
      ]);
      setfilterOpps([
        ...new Set(foundFiles.map((x) => x.opportunityName)),
        "All",
      ]);
      setFileTypeFilter([...new Set(foundFiles.map((x) => x.type)), "All"]);

      setSearchFiles(foundFiles);
      setSearching(false);
    } else if (props.searchSettings.searchQuery.length == 0) {
      setlastEditedBy([]);
      setfilterOpps([]);
      setFileTypeFilter([]);
      setSelectedLastEditedBy("All");
      setSelectedFilterOpp("All");
      setSelectedFileType("All");
      setSearchFiles([]);
    }
  };

  const setPinForSearchFile = (oppFile: OpportunityFile) => {
    const newList = searchFiles.map((item) => {
      if (item.url === oppFile.url) {
        return {
          ...item,
          pinned: !oppFile.pinned,
        };
      }
      return item;
    });

    setSearchFiles(newList);

    props.changePinForOpportunityFile(oppFile);
  };

  const updateSearchSettings = (property: string, value: any) => {
    var settings = props.searchSettings;
    settings[property] = value;

    props.updateSearchSettings(settings);
  };

  React.useEffect(() => {
    searchByText(props.searchSettings.searchQuery);
  }, [props.searchSettings.searchQuery, props.searchSettings.numberOfResults]);

  const clientsFound = props.clients.filter(
    (x) =>
      x.toLowerCase().indexOf(props.searchSettings.searchQuery.toLowerCase()) >
      -1
  );
  const oppsFound = props.opportunities.filter(
    (x) =>
      x.title
        .toLowerCase()
        .indexOf(props.searchSettings.searchQuery.toLowerCase()) > -1
  );

  return (
    <div>
      <Stack
        horizontal
        horizontalAlign="space-between"
        className={styles.baseStack}
      >
        <Stack.Item grow={3}>
          <Stack>
            <Stack.Item className={styles.searchArea} grow={1}>
              <SearchBox
                placeholder="Search"
                defaultValue={props.searchSettings.searchQuery}
                className={styles.searchBox}
                onChange={(_, searchText) => {
                  updateSearchSettings("searchQuery", searchText);
                  searchByText(searchText);
                }}
                onSearch={(searchText) => {
                  updateSearchSettings("searchQuery", searchText);
                  searchByText(searchText);
                }}
              />
            </Stack.Item>
            {props.searchSettings.searchQuery.length == 0 ? (
              <Stack horizontal>
                <Stack.Item
                  style={{
                    minWidth: "100%",
                    textAlign: "center",
                    marginTop: "15px",
                  }}
                >
                  <img src={img} />
                  <Label>You haven't searched for anything yet</Label>
                </Stack.Item>
              </Stack>
            ) : (
              <div style={{ margin: "2%", height: "70vh", overflowY: "auto" }}>
                <>
                  <LayoutSelector
                    isGridViewLayout={props.gridViewLayout}
                    changeLayout={props.changeLayout}
                  />
                  <Stack
                    horizontal={true}
                    horizontalAlign="space-between"
                    tokens={{
                      childrenGap: 20,
                    }}
                  >
                    {props.searchSettings.searchClients &&
                    selectedFilterOpp == "All" &&
                    selectedLastEditedBy == "All" &&
                    selectedFileType == "All" ? (
                      <Stack.Item grow={1}>
                        <div>
                          {props.searchSettings.searchQuery.length > 0 &&
                          clientsFound.length > 0 ? (
                            <>
                              <Label>Clients</Label>
                              {props.gridViewLayout
                                ? clientsFound
                                    .slice(
                                      0,
                                      props.searchSettings.numberOfResults
                                    )
                                    .map((x) => {
                                      return CardHelper.clientToCard(
                                        x,
                                        props.redirectFunc
                                      );
                                    })
                                : ListHelper.clientsToList(
                                    clientsFound.slice(
                                      0,
                                      props.searchSettings.numberOfResults
                                    ),
                                    props.redirectFunc
                                  )}
                            </>
                          ) : null}
                        </div>
                      </Stack.Item>
                    ) : null}
                    {props.searchSettings.searchOpportunities &&
                    selectedFilterOpp == "All" &&
                    selectedLastEditedBy == "All" &&
                    selectedFileType == "All" ? (
                      <Stack.Item grow={2}>
                        <div>
                          {props.searchSettings.searchQuery.length > 0 &&
                          oppsFound.length > 0 ? (
                            <>
                              <Label>Opportunities</Label>
                              {props.gridViewLayout
                                ? oppsFound
                                    .slice(
                                      0,
                                      props.searchSettings.numberOfResults
                                    )
                                    .map((x) => {
                                      return CardHelper.opportunityToCard(
                                        x,
                                        props.changePinForOpportunity,
                                        props.redirectFunc
                                      );
                                    })
                                : ListHelper.opportunitiesToList(
                                    oppsFound.slice(
                                      0,
                                      props.searchSettings.numberOfResults
                                    ),
                                    props.changePinForOpportunity,
                                    props.redirectFunc
                                  )}
                            </>
                          ) : null}
                        </div>
                      </Stack.Item>
                    ) : null}
                  </Stack>
                </>
                <Stack.Item
                  style={
                    !props.gridViewLayout &&
                    props.searchSettings.searchOpportunities &&
                    props.searchSettings.searchClients &&
                    selectedFilterOpp == "All" &&
                    selectedFileType == "All" &&
                    selectedLastEditedBy == "All" &&
                    oppsFound.length > 0 &&
                    clientsFound.length > 0
                      ? {
                          marginTop: `${
                            -350 + props.searchSettings.numberOfResults * 45
                          }px`,
                        }
                      : {}
                  }
                >
                  <div style={{ maxWidth: "70vw" }}>
                    {props.searchSettings.searchQuery.length > 0 &&
                    props.searchSettings.searchFiles &&
                    searchFiles.length > 0 ? (
                      <>
                        <Label>Files</Label>
                        {isSearching ? (
                          <Spinner
                            label="Searching..."
                            size={SpinnerSize.large}
                          />
                        ) : (
                          <>
                            {props.gridViewLayout
                              ? searchFiles
                                  .filter(
                                    (x) =>
                                      selectedFilterOpp == "All" ||
                                      x.opportunityName == selectedFilterOpp
                                  )
                                  .filter(
                                    (x) =>
                                      x.lastModifiedUserDisplayName ==
                                        selectedLastEditedBy ||
                                      selectedLastEditedBy == "All"
                                  )
                                  .filter(
                                    (x) =>
                                      x.type == selectedFileType ||
                                      selectedFileType == "All"
                                  )
                                  .slice(
                                    0,
                                    props.searchSettings.numberOfResults
                                  )
                                  .map((x) => {
                                    return CardHelper.fileToCard(
                                      x,
                                      setPinForSearchFile,
                                      props.redirectFunc,
                                      props.notifyMembers
                                    );
                                  })
                              : ListHelper.filesToList(
                                  searchFiles
                                    .filter(
                                      (x) =>
                                        selectedFilterOpp == "All" ||
                                        x.opportunityName == selectedFilterOpp
                                    )
                                    .filter(
                                      (x) =>
                                        x.lastModifiedUserDisplayName ==
                                          selectedLastEditedBy ||
                                        selectedLastEditedBy == "All"
                                    )
                                    .filter(
                                      (x) =>
                                        x.type == selectedFileType ||
                                        selectedFileType == "All"
                                    )
                                    .slice(
                                      0,
                                      props.searchSettings.numberOfResults
                                    ),
                                  setPinForSearchFile,
                                  props.redirectFunc,
                                  props.notifyMembers
                                )}
                          </>
                        )}
                      </>
                    ) : null}
                  </div>
                </Stack.Item>
              </div>
            )}
          </Stack>
        </Stack.Item>
        <Stack.Item grow={1} className={styles.filtersBox}>
          <div
            style={{
              padding: "5%",
              overflowY: "auto",
              height: "80vh",
            }}
          >
            <Label style={{ fontSize: "20px" }}>Filters</Label>
            <Label>What are you searching for?</Label>
            <Checkbox
              checked={props.searchSettings.searchClients}
              onChange={() =>
                updateSearchSettings(
                  "searchClients",
                  !props.searchSettings.searchClients
                )
              }
              label="Client"
              title="client"
            />
            <Checkbox
              checked={props.searchSettings.searchOpportunities}
              onChange={() =>
                updateSearchSettings(
                  "searchOpportunities",
                  !props.searchSettings.searchOpportunities
                )
              }
              label="Opportunity"
              title="Opportunity"
            />
            <Checkbox
              checked={props.searchSettings.searchFiles}
              onChange={() =>
                updateSearchSettings(
                  "searchFiles",
                  !props.searchSettings.searchFiles
                )
              }
              label="File"
              title="file"
            />
            <Slider
              label="Number of search results"
              defaultValue={props.searchSettings.numberOfResults}
              onChange={(x) => updateSearchSettings("numberOfResults", x)}
              min={1}
              max={10}
              styles={{ root: { marginTop: "5px", marginBottom: "20px" } }}
            />
            {props.searchSettings.searchFiles ? (
              <>
                <Label
                  style={{
                    marginTop: "5px",
                    marginBottom: "10px",
                    fontSize: "20px",
                    display:
                      props.searchSettings.searchQuery.length == 0
                        ? "none"
                        : "block",
                  }}
                >
                  File refiners
                </Label>
                {lastEditedBy.length > 0 ? (
                  <ChoiceGroup
                    label="Last Edited By"
                    defaultSelectedKey="All"
                    onChange={(_, x) => setSelectedLastEditedBy(x.key)}
                    options={lastEditedBy.map((x) => {
                      return {
                        key: x,
                        text: x,
                      };
                    })}
                    styles={{
                      root: { marginTop: "5px", marginBottom: "10px" },
                    }}
                  />
                ) : null}
                {filterOpps.length > 0 ? (
                  <ChoiceGroup
                    label="Opportunity"
                    defaultSelectedKey="All"
                    onChange={(_, x) => setSelectedFilterOpp(x.key)}
                    options={filterOpps.map((x) => {
                      return {
                        key: x,
                        text: x,
                      };
                    })}
                    styles={{
                      root: { marginTop: "5px", marginBottom: "10px" },
                    }}
                  />
                ) : null}
                {fileTypes.length > 0 ? (
                  <ChoiceGroup
                    label="Type"
                    defaultSelectedKey="All"
                    onChange={(_, x) => setSelectedFileType(x.key)}
                    options={fileTypes.map((x) => {
                      return {
                        key: x,
                        text: x,
                      };
                    })}
                    styles={{
                      root: { marginTop: "5px", marginBottom: "10px" },
                    }}
                  />
                ) : null}
              </>
            ) : null}
          </div>
        </Stack.Item>
      </Stack>
    </div>
  );
};

export default Find;
