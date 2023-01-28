import * as React from "react";
import { DataGrid } from "@mui/x-data-grid";
import { makeStyles } from "@material-ui/styles";

const columns: any = [
  { field: "id", headerName: "ID", width: 90, hide: true },
  {
    field: "empName",
    headerName: "Employee Name",
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

let rows = [];

export default function MaterialDBNew(props) {
  console.log(props);

  rows = [];
  for (let i = 0; i < props.items.length; i++) {
      rows.push({
        id: props.items[i].ID,
        empName: props.items[i].text,
        title: props.items[i].jobTitle,
        Email: props.items[i].Email,
        subDep: props.items[i].Dept,
        Zone: props.items[i].Zone,
        Manager: props.items[i].manager
          ? props.items[i].manager.displayName
          : "",
        Ext: props.items[i].Ext,
        mobile: props.items[i].mobilePhone,
      });
  }
  return (
    <div style={{ height: 400, width: "100%" }}>
      <DataGrid
        rows={rows}
        columns={columns}
        pageSize={10}
        rowsPerPageOptions={[5]}
        // checkboxSelection
        disableSelectionOnClick
      />
    </div>
  );
}
