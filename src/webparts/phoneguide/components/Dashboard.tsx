import * as React from 'react';
import { IStyleSet, Label, ILabelStyles, Pivot, PivotItem } from '@fluentui/react';
import DepartmentList from "./DepartmentList";
import MaterialDtabs from './Materialtabs';
const labelStyles: Partial<IStyleSet<ILabelStyles>> = 
{
  root: { marginTop: 10 },
};

export const Dashboard: React.FunctionComponent = () => {
  return (
    <Pivot aria-label="Basic Pivot Example">
      <PivotItem
        headerText="My Files"
        headerButtonProps={{
          'data-order': 1,
          'data-title': 'My Files Title',
        }}
      >
        
      </PivotItem>
      <PivotItem headerText="Recent">

      </PivotItem>
      <PivotItem headerText="Shared with me">

      </PivotItem>
    </Pivot>
  );
};
