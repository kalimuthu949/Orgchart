import * as React from 'react';
import styles from './Phoneguide.module.scss';
import { IPhoneguideProps } from './IPhoneguideProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";
import { Dashboard } from './Dashboard';
import MaterialDtabs from './Materialtabs';
import DepartmentPivot from './DepartmentPivot';
import { OrgChart } from './OrgChart';

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
      userDisplayName
    } = this.props;

    return (<div><OrgChart context={this.props.context}/></div>);
  }
}
