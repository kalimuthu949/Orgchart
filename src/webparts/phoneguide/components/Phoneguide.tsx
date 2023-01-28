import * as React from "react";
import styles from "./Phoneguide.module.scss";
import { IPhoneguideProps } from "./IPhoneguideProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";
import { Dashboard } from "./Dashboard";
import MaterialDtabs from "./Materialtabs";
import DepartmentPivot from "./DepartmentPivot";
import { OrgChart } from "./OrgChart";
import App from "./App";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
export default class Phoneguide extends React.Component<IPhoneguideProps, {}> {
  constructor(prop: IPhoneguideProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<IPhoneguideProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div style={{ padding: 26 }}>
        
        <App context={this.props.context} />
      </div>
    );
  }
}
