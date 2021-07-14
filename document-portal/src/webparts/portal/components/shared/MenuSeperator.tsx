import { Label, Separator } from "@fluentui/react";
import * as React from "react";
import styles from "./MenuSeperator.module.scss";

export interface MenuSeperatorProps {
  title: string;
}

const MenuSeperator: React.FunctionComponent<MenuSeperatorProps> = (props) => {
  return (
    <div>
      <Label
        style={{
          fontSize: "24px",
          marginTop: "10px",
          padding: "0 8px",
        }}
      >
        {props.title}
      </Label>
      <Separator className={styles.shadow} alignContent="start"></Separator>
    </div>
  );
};

export default MenuSeperator;
