import { FontIcon, Stack } from "@fluentui/react";
import * as React from "react";
import styles from "./MenuSeperator.module.scss";

export interface LayoutSelector {
  isGridViewLayout: boolean;
  changeLayout: (isGridView: boolean) => void;
}

const LayoutSelector: React.FunctionComponent<LayoutSelector> = (props) => {
  return (
    <Stack horizontal style={{ float: "right", position: "sticky", zIndex: 1 }}>
      <Stack.Item>
        <FontIcon
          iconName="GroupedList"
          className={!props.isGridViewLayout ? styles.isActive : styles.base}
          onClick={() => props.changeLayout(false)}
        />
      </Stack.Item>
      <Stack.Item>
        <FontIcon
          iconName="GridViewMedium"
          className={props.isGridViewLayout ? styles.isActive : styles.base}
          onClick={() => props.changeLayout(true)}
        />
      </Stack.Item>
    </Stack>
  );
};

export default LayoutSelector;
