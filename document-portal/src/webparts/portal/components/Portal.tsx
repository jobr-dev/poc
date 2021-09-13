import * as React from "react";
import styles from "./Portal.module.scss";
import SearchService from "../services/searchService";
import { Pivot, PivotItem } from "@fluentui/react";
import Browse from "./browse/Browse";
import Pinned from "./pinned/Pinned";
import { Opportunity } from "../models/Opportunity";
import MenuSeperator from "./shared/MenuSeperator";
import { OpportunityFile } from "../models/OpportunityFile";
import PinService from "../services/pinService";
import { SocialActorType } from "@pnp/sp/social";
import { MSGraphClient } from "@microsoft/sp-http";
import ActivitiesHelper from "../utils/ActivitiesHelper";
import Find from "./find/Find";
import { SearchSettings } from "../models/SearchSettings";
import FileService from "../services/fileService";
import RecentService from "../services/recentService";
import Recent from "./recent/Recent";
import NotifyMembersPanel from "./shared/NotifyMembersPanel";

export interface IPortalProps {
  graphClient: MSGraphClient;
}
export interface IPortalState {
  opportunities: Opportunity[];
  showAllOpps: boolean;
  clients: string[];
  pinnedFiles: OpportunityFile[];
  loading: boolean;
  selectedKey: string;
  gridViewLayout: boolean;
  key: string;
  selectedOpportunity: Opportunity;
  searchSettings: SearchSettings;
  recentFiles: OpportunityFile[];
  recentOpportunities: Opportunity[];
  notifyMembersPanelOpen: boolean;
  notifyMembersOppFile: OpportunityFile;
}

export default class Portal extends React.Component<
  IPortalProps,
  IPortalState
> {
  private activitiesInterval: number;
  // private recentInterval: number;
  private searchService: SearchService;
  private pinService: PinService;
  private activityHelper: ActivitiesHelper;
  private fileService: FileService;
  private recentService: RecentService;

  constructor(props: any) {
    super(props);
    this.state = {
      opportunities: [],
      showAllOpps: true,
      key: "",
      clients: [],
      pinnedFiles: [],
      loading: true,
      selectedKey: "0",
      gridViewLayout: true,
      selectedOpportunity: {
        id: "",
        url: "",
        pinned: false,
        title: "",
        client: "",
        userHasAccess: false,
        activities: [],
        listId: "",
        siteId: "",
        webId: "",
      },
      searchSettings: {
        searchFiles: true,
        searchClients: true,
        searchOpportunities: true,
        searchQuery: "",
        numberOfResults: 5,
      },
      recentFiles: [],
      recentOpportunities: [],
      notifyMembersPanelOpen: false,
      notifyMembersOppFile: null,
    };
    this.searchService = new SearchService();
    this.pinService = new PinService();
    this.activityHelper = new ActivitiesHelper(this.props.graphClient);
    this.fileService = new FileService();
    this.recentService = new RecentService();
  }

  public async componentDidMount() {
    await this.getAllOpportunities();
    await this.getPinnedFiles();
    // await this.getRecentContent();
    this.setState({ loading: false });

    this.activitiesInterval = setInterval(() => this.updateActivities(), 60000);
    // this.recentInterval = setInterval(() => this.getRecentContent(), 90000);
  }

  public componentWillUnmount() {
    clearInterval(this.activitiesInterval);
    // clearInterval(this.recentInterval);
  }

  private async getAllOpportunities() {
    var opportunities =
      await this.searchService.getAllOpportunititesFromStructure();
    var clients = [
      ...new Set(
        opportunities.filter((x) => x.userHasAccess).map((x) => x.client)
      ),
    ];

    for (
      let index = 0;
      index < opportunities.filter((x) => x.pinned).length;
      index++
    ) {
      var element = opportunities[index];
      if (element.pinned) {
        element = await this.activityHelper.getOpportunityActivity(element);
      }
    }

    this.setState({
      opportunities,
      clients,
    });
  }

  private async getPinnedFiles() {
    var pinnedFiles = await this.pinService.getPinnedFiles();
    this.setState({ pinnedFiles });
  }

  private async getRecentContent() {
    var recentFiles = await this.recentService.getRecentFiles();
    var opportunityFilter = [
      ...new Set(recentFiles.map((x) => x.opportunityUrl)),
    ];

    var recentOpportunities = this.state.opportunities.filter(
      (x) => opportunityFilter.indexOf(x.url) > -1
    );

    recentFiles.forEach((x) => {
      x.pinned = this.state.pinnedFiles.some((pinned) => x.url == pinned.url);
    });

    this.setState({
      recentFiles: recentFiles,
      recentOpportunities: recentOpportunities,
    });
  }

  private updateRecentContent(
    url: string,
    pinStatus: boolean,
    isDocument: boolean
  ) {
    if (isDocument) {
      const recentFiles = this.state.recentFiles.map((item) => {
        if (item.url === url) {
          return {
            ...item,
            pinned: pinStatus,
          };
        }
        return item;
      });

      this.setState({ recentFiles });
    } else {
      const recentOpportunities = this.state.recentOpportunities.map((item) => {
        if (item.url === url) {
          return {
            ...item,
            pinned: pinStatus,
          };
        }
        return item;
      });
      this.setState({ recentOpportunities });
    }
  }

  private async updateActivities(): Promise<void> {
    const opps = this.state.opportunities;
    for (let index = 0; index < opps.length; index++) {
      var currentOpp = opps[index];
      if (currentOpp.pinned) {
        currentOpp = await this.activityHelper.getOpportunityActivity(
          currentOpp
        );
      }
    }

    const pinnedFiles = this.state.pinnedFiles;
    for (let index = 0; index < pinnedFiles.length; index++) {
      const currentFile = pinnedFiles[index];
      await this.activityHelper.getOpportunityFileActivity(currentFile);
    }

    this.setState({ ...this.state, opportunities: opps, pinnedFiles });
  }

  private redirect(opp: Opportunity): void {
    this.setState({ selectedKey: "0", selectedOpportunity: opp });
  }

  private changePinForOpportunity(opp: Opportunity): void {
    const newList = this.state.opportunities.map((item) => {
      if (item.url === opp.url) {
        return {
          ...item,
          pinned: !opp.pinned,
        };
      }
      return item;
    });

    this.setState({
      ...this.state,
      opportunities: newList,
    });

    if (!opp.pinned) {
      this.pinService.pin(opp.url, SocialActorType.Site);
    } else {
      this.pinService.unpin(opp.url, SocialActorType.Site);
    }

    // this.updateRecentContent(opp.url, !opp.pinned, false);
  }

  private changePinForOpportunityFile(oppFile: OpportunityFile): void {
    if (!oppFile.pinned) {
      this.setState({
        pinnedFiles: this.state.pinnedFiles.concat({
          ...oppFile,
          pinned: !oppFile.pinned,
        }),
      });
      this.pinService.pin(oppFile.url, SocialActorType.Document);
    } else {
      this.setState({
        pinnedFiles: this.state.pinnedFiles.filter(
          (item) => item.editUrl !== oppFile.editUrl
        ),
      });
      this.pinService.unpin(oppFile.url, SocialActorType.Document);
    }
    // this.updateRecentContent(oppFile.url, !oppFile.pinned, true);
  }

  private setShowAllOpps(showAll: boolean): void {
    this.setState({
      ...this.state,
      showAllOpps: showAll,
      key: Math.random().toString(36).substring(7),
    });
  }

  private setNotifyMembersOppFile(opp: OpportunityFile): void {
    this.setState({ notifyMembersOppFile: opp, notifyMembersPanelOpen: true });
  }

  public render(): React.ReactElement<IPortalProps> {
    return (
      <>
        <div className={styles.container}>
          <div className={styles.base}>
            <Pivot
              aria-label="Portal Pivot"
              selectedKey={this.state.selectedKey}
              onLinkClick={(item: PivotItem) => {
                if (this.state.loading) return;
                this.setState({
                  ...this.state,
                  selectedKey: item.props["itemKey"],
                });
              }}
            >
              <PivotItem headerText="Browse" itemKey="0">
                <MenuSeperator title="Browse" />
                <Browse
                  key={this.state.key}
                  loading={this.state.loading}
                  opportunities={
                    this.state.showAllOpps
                      ? this.state.opportunities
                      : this.state.opportunities.filter((x) => x.userHasAccess)
                  }
                  selectedOpportunity={this.state.selectedOpportunity}
                  setShowAllOpps={(allOpps) => this.setShowAllOpps(allOpps)}
                  showAllOpps={this.state.showAllOpps}
                  selectedOpportunityChanged={(opp) =>
                    this.setState({
                      selectedOpportunity: opp,
                    })
                  }
                  changePinForOpportunity={(opp) =>
                    this.changePinForOpportunity(opp)
                  }
                  changePinForOpportunityFile={(oppFile) =>
                    this.changePinForOpportunityFile(oppFile)
                  }
                  fileService={this.fileService}
                  notifyMembers={(oppFile: OpportunityFile) =>
                    this.setNotifyMembersOppFile(oppFile)
                  }
                />
              </PivotItem>
              <PivotItem headerText="Pinned" itemKey="1">
                <MenuSeperator title="Pinned" />
                <Pinned
                  opportunities={this.state.opportunities.filter(
                    (x) => x.pinned
                  )}
                  files={this.state.pinnedFiles}
                  changeLayout={(isGridView: boolean) =>
                    this.setState({ gridViewLayout: isGridView })
                  }
                  gridViewLayout={this.state.gridViewLayout}
                  redirectFunc={(opp) => this.redirect(opp)}
                  changePinForOpportunity={(opp) =>
                    this.changePinForOpportunity(opp)
                  }
                  changePinForOpportunityFile={(oppFile) =>
                    this.changePinForOpportunityFile(oppFile)
                  }
                  notifyMembers={(oppFile: OpportunityFile) =>
                    this.setNotifyMembersOppFile(oppFile)
                  }
                />
              </PivotItem>
              {/*               <PivotItem headerText="Recent" itemKey="2">
                <MenuSeperator title="Recent" />
                <Recent
                  files={this.state.recentFiles}
                  opportunities={this.state.recentOpportunities}
                  clients={[
                    ...new Set(
                      this.state.recentOpportunities.map((x) => x.client)
                    ),
                  ]}
                  changeLayout={(isGridView: boolean) =>
                    this.setState({ gridViewLayout: isGridView })
                  }
                  gridViewLayout={this.state.gridViewLayout}
                  redirectFunc={(opp) => this.redirect(opp)}
                  changePinForOpportunity={(opp) =>
                    this.changePinForOpportunity(opp)
                  }
                  changePinForOpportunityFile={(oppFile) =>
                    this.changePinForOpportunityFile(oppFile)
                  }
                  notifyMembers={(oppFile: OpportunityFile) =>
                    this.setNotifyMembersOppFile(oppFile)
                  }
                />
              </PivotItem> */}
              <PivotItem headerText="Find" itemKey="3">
                <MenuSeperator title="Find" />
                <Find
                  opportunities={this.state.opportunities.filter(
                    (x) => x.userHasAccess
                  )}
                  clients={this.state.clients}
                  pinnedFiles={this.state.pinnedFiles}
                  gridViewLayout={this.state.gridViewLayout}
                  changeLayout={(isGridView: boolean) =>
                    this.setState({ gridViewLayout: isGridView })
                  }
                  redirectFunc={(opp) => this.redirect(opp)}
                  changePinForOpportunity={(opp) =>
                    this.changePinForOpportunity(opp)
                  }
                  changePinForOpportunityFile={(oppFile) =>
                    this.changePinForOpportunityFile(oppFile)
                  }
                  searchService={this.searchService}
                  searchSettings={this.state.searchSettings}
                  updateSearchSettings={(x) =>
                    this.setState({ ...this.state, searchSettings: x })
                  }
                  notifyMembers={(oppFile: OpportunityFile) =>
                    this.setNotifyMembersOppFile(oppFile)
                  }
                />
              </PivotItem>
            </Pivot>
          </div>
          <NotifyMembersPanel
            visible={this.state.notifyMembersPanelOpen}
            oppFile={this.state.notifyMembersOppFile}
            dismissPanel={() =>
              this.setState({ notifyMembersPanelOpen: false })
            }
          />
        </div>
      </>
    );
  }
}