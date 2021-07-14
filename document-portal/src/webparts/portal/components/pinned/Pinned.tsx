import * as React from "react";
import { Label, Stack } from "@fluentui/react";
import { Opportunity } from "../../models/Opportunity";
import { OpportunityFile } from "../../models/OpportunityFile";
import CardHelper from "../../utils/CardHelper";
import LayoutSelector from "../shared/LayoutSelector";
import ListHelper from "../../utils/ListHelper";

export interface PinnedProps {
  opportunities: Opportunity[];
  files: OpportunityFile[];
  gridViewLayout: boolean;
  changeLayout: (isGridView: boolean) => void;
  redirectFunc: (opp: Opportunity) => void;
  changePinForOpportunity: (opp: Opportunity) => void;
  changePinForOpportunityFile: (oppFile: OpportunityFile) => void;
  notifyMembers: (oppFile: OpportunityFile) => void;
}

const Pinned: React.FunctionComponent<PinnedProps> = (props) => {
  return (
    <div
      style={{
        overflowY: "auto",
        height: "80vh",
        marginRight: "-1%",
      }}
    >
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
        <Stack.Item grow={2}>
          <div>
            <Label>Opportunities</Label>
            {props.opportunities.length > 0
              ? props.gridViewLayout
                ? props.opportunities.map((x) => {
                    return CardHelper.opportunityToCard(
                      x,
                      props.changePinForOpportunity,
                      props.redirectFunc
                    );
                  })
                : ListHelper.opportunitiesToList(
                    props.opportunities,
                    props.changePinForOpportunity,
                    props.redirectFunc
                  )
              : null}
          </div>
        </Stack.Item>
        <Stack.Item grow={4}>
          <div>
            <Label>Files</Label>
            {props.files
              ? props.gridViewLayout
                ? props.files.map((x) => {
                    return CardHelper.fileToCard(
                      x,
                      props.changePinForOpportunityFile,
                      props.redirectFunc,
                      props.notifyMembers
                    );
                  })
                : ListHelper.filesToList(
                    props.files,
                    props.changePinForOpportunityFile,
                    props.redirectFunc,
                    props.notifyMembers
                  )
              : null}
          </div>
        </Stack.Item>
      </Stack>
    </div>
  );
};

export default Pinned;
