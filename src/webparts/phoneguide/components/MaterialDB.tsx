import * as React from "react";
import { DataGrid } from "@mui/x-data-grid";
import { makeStyles } from '@material-ui/styles';

const columns:any = [
  { field: "id", headerName: "ID", width: 90, hide: true },
  {
    field: "Name",
    headerName: "First name",
    width: 150,
    editable: false,
    headerClassName: 'super-app-theme--header',
    headerAlign: 'center'
  },
  {
    field: "Email",
    headerName: "Email",
    width: 150,
    editable: false,
    headerClassName: 'super-app-theme--header',
    headerAlign: 'center'
  },
  {
    field: "Department",
    headerName: "Department",
    width: 150,
    editable: false,
    headerClassName: 'super-app-theme--header',
    headerAlign: 'center'
  },
  {
    field: "Zone",
    headerName: "Zone",
    width: 150,
    editable: false,
    headerClassName: 'super-app-theme--header',
    headerAlign: 'center'
  },
  {
    field: "Dept",
    headerName: "Dept",
    width: 150,
    editable: false,
    headerClassName: 'super-app-theme--header',
    headerAlign: 'center'
  },
];

let rows = [];

export default function MaterialDB(props) {
  console.log(props);

  rows = [];
  for (let i = 0; i < props.items.length; i++) {
    if (props.Department == props.items[i].department)
      rows.push({
        id: props.items[i].ID,
        Name: props.items[i].text,
        Email: props.items[i].Email,
        Department: props.items[i].department,
        Zone: props.items[i].Zone,
        Dept: props.items[i].Dept,
      });
  }
  return (
    <div style={{ height: 400, width: "100%" }}>
      <DataGrid
        rows={rows}
        columns={columns}
        pageSize={5}
        rowsPerPageOptions={[5]}
        checkboxSelection
        disableSelectionOnClick
        
      />
    </div>
  );
}
