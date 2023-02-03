import * as React from "react";
import { DataGrid } from "@mui/x-data-grid";
import { makeStyles } from "@material-ui/styles";

import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from "@fluentui/react";

const columns: any[] = [
  { field: "id", headerName: "ID", width: 90, hide: true },
  // {
  //   field: "empName",
  //   headerName: "Employee Name",
  //   width: 200,
  //   editable: false,
  //   headerClassName: "super-app-theme--header",
  //   headerAlign: "left",
  // },
  {
    field: "givenName",
    headerName: "First Name",
    width: 200,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
  {
    field: "surname",
    headerName: "Last Name",
    width: 200,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
  {
    field: "title",
    headerName: "Title",
    width: 200,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
  {
    field: "Email",
    headerName: "Email",
    width: 300,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },

  {
    field: "Zone",
    headerName: "Zone",
    width: 150,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },

  {
    field: "Ext",
    headerName: "Ext",
    width: 150,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
  {
    field: "mobile",
    headerName: "Mobile Number",
    width: 150,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
  {
    field: "dept",
    headerName: "Department",
    width: 200,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
  {
    field: "subDep",
    headerName: "Sub Department",
    width: 200,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
  {
    field: "Manager",
    headerName: "Manager",
    width: 150,
    editable: false,
    headerClassName: "super-app-theme--header",
    headerAlign: "left",
  },
];
const _columns: IColumn[] = [
  {
    key: "Column1",
    name: "First Name",
    fieldName: "givenName",
    minWidth: 200,
    maxWidth: 200,
  },
  {
    key: "Column2",
    name: "Last Name",
    fieldName: "surname",
    minWidth: 200,
    maxWidth: 200,
  },
  {
    key: "Column3",
    name: "Title",
    fieldName: "title",
    minWidth: 200,
    maxWidth: 200,
  },
  {
    key: "Column4",
    name: "Email",
    fieldName: "Email",
    minWidth: 300,
    maxWidth: 300,
  },
  {
    key: "Column5",
    name: "Zone",
    fieldName: "Zone",
    minWidth: 150,
    maxWidth: 150,
  },
  {
    key: "Column6",
    name: "Ext",
    fieldName: "Ext",
    minWidth: 150,
    maxWidth: 150,
  },
  {
    key: "Column7",
    name: "Mobile Number",
    fieldName: "mobile",
    minWidth: 150,
    maxWidth: 150,
  },
  {
    key: "Column8",
    name: "Department",
    fieldName: "dept",
    minWidth: 200,
    maxWidth: 200,
  },
  {
    key: "Column9",
    name: "Sub Department",
    fieldName: "subDep",
    minWidth: 200,
    maxWidth: 200,
  },
  {
    key: "Column10",
    name: "Manager",
    fieldName: "Manager",
    minWidth: 200,
    maxWidth: 200,
  },
];

export default function FluentUIDB(props) {
  const [data, setData] = React.useState([]);

  React.useEffect(() => {
    let rows = [];

    props.items.forEach((item, index) => {
      rows.push({
        id: item.ID,
        empName: item.text,
        title: item.jobTitle,
        givenName: item.givenName,
        surname: item.surname,
        Email: item.Email,
        dept: item.department,
        subDep: item.Dept,
        Zone: item.Zone,
        Manager: item.manager ? item.manager.displayName : "",
        Ext: item.Ext,
        mobile: item.mobilePhone,
      });
    });
    setData([...rows]);
  }, [props.items]);
  return (
    <div style={{ overflow: "auto", width: "100%" }}>
      <DetailsList
        items={data}
        columns={_columns}
        styles={{
          root: {
            ".ms-DetailsRow-cell": {
              height: 40,
            },
          },
        }}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
      />
    </div>
  );
}
