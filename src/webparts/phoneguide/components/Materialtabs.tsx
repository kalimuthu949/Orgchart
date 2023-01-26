import * as React from "react";
import * as PropTypes from "prop-types";
import { makeStyles, useTheme } from "@material-ui/core/styles";
import AppBar from "@material-ui/core/AppBar";
import Tabs from "@material-ui/core/Tabs";
import Tab from "@material-ui/core/Tab";
import Typography from "@material-ui/core/Typography";
import Box from "@material-ui/core/Box";
import MaterialDB from "./MaterialDB";
import { graph } from "@pnp/graph/presets/all";

import SPServices from "./SPServices";
import "../assets/Css/Phoneguide.css";
function TabPanel(props) {
  const { children, value, index, ...other } = props;

  return (
    <div
      role="tabpanel"
      hidden={value !== index}
      id={`simple-tabpanel-${index}`}
      aria-labelledby={`simple-tab-${index}`}
      {...other}
    >
      {value === index && (
        <Box p={3}>
          <Typography>{children}</Typography>
        </Box>
      )}
    </div>
  );
}

TabPanel.propTypes = {
  children: PropTypes.node,
  index: PropTypes.any.isRequired,
  value: PropTypes.any.isRequired,
};

function a11yProps(index) {
  return {
    id: `simple-tab-${index}`,
    "aria-controls": `simple-tabpanel-${index}`,
  };
}

const useStyles = makeStyles((theme) => ({
  root: {
    flexGrow: 1,
    backgroundColor: theme.palette.background.paper,
  },
}));

export default function MaterialDtabs() {
  const classes = useStyles();
  const [value, setValue] = React.useState(0);
  const [peopleList, setPeopleList] = React.useState([]);
  const [department, setdepartment] = React.useState([]);

  const handleChange = (event, newValue) => {
    setValue(newValue);
  };

  function removeDuplicates(arr) {
    return arr.filter((item, index) => arr.indexOf(item) === index);
  }

  React.useEffect(function () {
    getalluserssp();
    // getallusersgraph();
  }, []);

  async function getalluserssp() {
    SPServices.SPReadItems({
      Listname: "EmployeeDetails",
      Select: "*,Employee/Title,Employee/Id,Employee/EMail",
      Expand: "Employee",
    })
      .then((items: any) => {
        getallusersgraph(items);
        // console.log(items);
      })
      .catch(function (error) {
        console.log(error);
      });
  }

  async function getallusersgraph(userData) {
    await graph.users
      .select("department,mail,id,displayName,jobTitle,mobilePhone,manager,ext")
      .expand("manager")
      .top(999)
      .get()
      .then(function (data) {
        console.log(data);
        const users = [];
        let depts = [];
        for (let i = 0; i < data.length; i++) {
          let filteredArr = [];
          for (let j = 0; j < userData.length; j++) {
            let user = userData[j];
            if (user.EmployeeId && user.Employee.EMail == data[i].mail) {
              filteredArr.push(user);
            }
          }

          users.push({
            imageUrl:
              "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
            isValid: true,
            Email: data[i].mail,
            ID: data[i].id,
            key: i,
            text: data[i].displayName,
            jobTitle: data[i].jobTitle,
            mobilePhone: data[i].mobilePhone,
            department: data[i].department,
            Zone: filteredArr.length > 0 ? filteredArr[0].Zone : "",
            Dept:
              filteredArr.length > 0
                ? filteredArr[0].SubDepartments.join(", ")
                : "",
            manager: data[i].manager ? data[i].manager : null,
            Ext: filteredArr.length > 0 ? filteredArr[0].Ext : "",
          });

          if (data[i].department) depts.push(data[i].department);
          depts = removeDuplicates(depts);
        }
        console.log(users);
        setdepartment([...depts]);
        setPeopleList([...users]);
      })
      .catch(function (error) {
        console.log(error);
      });
  }

  return peopleList.length > 0 ? (
    <div className="clsMaterialtab">
      <div className={classes.root}>
        <AppBar position="static" className="clsTabs">
          <Tabs
            variant="scrollable"
            scrollButtons="auto"
            value={value}
            onChange={handleChange}
            aria-label="simple tabs example"
          >
            {department.map(function (item, index) {
              let name = item;
              return <Tab label={name} {...a11yProps(index)} />;
            })}
          </Tabs>
        </AppBar>
        {department.map(function (item, index) {
          return (
            <TabPanel value={value} index={index}>
              <MaterialDB Department={item} items={peopleList} />
            </TabPanel>
          );
        })}
      </div>
    </div>
  ) : null;
}
