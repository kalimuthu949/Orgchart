import * as React from "react";
import * as PropTypes from "prop-types";
import { makeStyles, useTheme } from "@material-ui/core/styles";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import AppBar from "@material-ui/core/AppBar";
import Tabs from "@material-ui/core/Tabs";
import Tab from "@material-ui/core/Tab";
import Typography from "@material-ui/core/Typography";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import {
  IPersonaProps,
  Persona,
  PersonaSize,
} from "@fluentui/react/lib/Persona";
import {
  NormalPeoplePicker,
  ValidationState,
} from "@fluentui/react/lib/Pickers";
import Box from "@material-ui/core/Box";
import MaterialDBNew from "./MaterialDBNew";
import { graph } from "@pnp/graph/presets/all";
import { Dropdown, IDropdownStyles } from "@fluentui/react/lib/Dropdown";
import SPServices from "./SPServices";
import "../assets/Css/Phoneguide.css";
import { IIconProps } from "@fluentui/react";
import { ThemeProvider, PartialTheme } from "@fluentui/react/lib/Theme";
import NewPivot from "./NewPivot";
import { useState } from "react";
import { Icon } from "@fluentui/react";
//Filter functionality
let listitems = []; //glb array which is having the all user details from sharepoint list
let graphuserdetails = []; //glb array which is having the all user details from grpah
const myTheme: PartialTheme = {
  palette: {
    themePrimary: "#03606a",
    themeLighterAlt: "#eff8f9",
    themeLighter: "#c3e4e7",
    themeLight: "#95cdd3",
    themeTertiary: "#459da6",
    themeSecondary: "#12727d",
    themeDarkAlt: "#035760",
    themeDark: "#024a51",
    themeDarker: "#02363c",
    neutralLighterAlt: "#faf9f8",
    neutralLighter: "#f3f2f1",
    neutralLight: "#edebe9",
    neutralQuaternaryAlt: "#e1dfdd",
    neutralQuaternary: "#d0d0d0",
    neutralTertiaryAlt: "#c8c6c4",
    neutralTertiary: "#a19f9d",
    neutralSecondary: "#605e5c",
    neutralPrimaryAlt: "#3b3a39",
    neutralPrimary: "#323130",
    neutralDark: "#201f1e",
    black: "#000000",
    white: "#ffffff",
  },
};
const syncIcon: IIconProps = { iconName: "Sync" };
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
  const [delayResults, setDelayResults] = React.useState(false);
  const [value, setValue] = React.useState(0);
  const [allusers, setallusers] = React.useState([]);
  const [masterPeopleList,setMasterPeopleList]=React.useState([]); //which is used to store the users from graph and sharepoint list as well dropdown filter.
  const [peopleList, setPeopleList] = React.useState([]); //which is used to store the users from graph and sharepoint list as well dropdown filter.
  const [alldepartment, setalldepartment] = React.useState([]); //which is used to bind tabs.
  const [loader, setloader] = React.useState(false);
  const [mostRecentlyUsed, setMostRecentlyUsed] = React.useState<
    IPersonaProps[]
  >([]);

  //For filter dropdowns
  const [zones, setzones] = React.useState([]);
  const [titles, settitles] = React.useState([]);
  const [selectedusers, setselectedusers] = React.useState([]);

  //For Filters
  const [empname, setempname] = React.useState("");
  const [zone, setzone] = React.useState([]);
  const [title, settitle] = React.useState([]);
  const [dept, setdept] = React.useState([]);
  const [isPG, setIsPG] = useState(true);
  const handleChange = (event, newValue) => {
    setValue(newValue);
  };

  function removeDuplicatesfromarray(arr) {
    return arr.filter((item, index) => arr.indexOf(item) === index);
  }

  React.useEffect(function () {
    setloader(true);
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
        listitems = items;
        getallusersgraph(items);
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
      });
  }

  async function getallusersgraph(userData) {
    await graph.users
      .select("department,mail,id,displayName,jobTitle,mobilePhone,manager,ext")
      .expand("manager")
      .top(999)
      .get()
      .then(function (data) {
        const users = [];

        let depts = [];
        let arrzones = [];
        let arrTitles = [];

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
            Email: data[i].mail?data[i].mail:'',
            ID: data[i].id,
            key: i,
            text: data[i].displayName,
            jobTitle: data[i].jobTitle?data[i].jobTitle:'',
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
          if (data[i].jobTitle) arrTitles.push(data[i].jobTitle);

          let zonename = filteredArr.length > 0 ? filteredArr[0].Zone : "";
          if (zonename) arrzones.push(zonename);
        }
        
        graphuserdetails = users;

        depts = removeDuplicatesfromarray(depts);
        arrzones = removeDuplicatesfromarray(arrzones);
        arrTitles = removeDuplicatesfromarray(arrTitles);

        let statedept = [];
        for (let i = 0; i < depts.length; i++) {
          statedept.push({ key: depts[i], text: depts[i] });
        }

        let statezones = [];
        for (let i = 0; i < arrzones.length; i++) {
          statezones.push({ key: arrzones[i], text: arrzones[i] });
        }

        let statetitles = [];
        for (let i = 0; i < arrTitles.length; i++) {
          statetitles.push({ key: arrTitles[i], text: arrTitles[i] });
        }

        setalldepartment([...statedept]);
        setzones([...statezones]);
        settitles([...statetitles]);
        setallusers([...users]);
        setMasterPeopleList([...users]);
        setPeopleList([...users]);
        setloader(false);
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
      });
  }

  async function filtervalues(useremail, userdept, usertitle) {

    let tempPeopleList=[...masterPeopleList];

    if(useremail){
      tempPeopleList=tempPeopleList.filter((_user)=>_user.Email==useremail)
    }

    if(userdept.length>0){
      tempPeopleList=tempPeopleList.filter((_user)=>userdept.some((_dept)=>_dept==_user.department))
    }

    if(usertitle.length>0){
      tempPeopleList=tempPeopleList.filter((_user)=>usertitle.some((_title)=>_title==_user.jobTitle))
    }

    setPeopleList([...tempPeopleList])
  }

  const onFilterChanged = (
    filterText: string,
    currentPersonas: IPersonaProps[],
    limitResults?: number
  ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (filterText) {
      let filteredPersonas: IPersonaProps[] = filterPersonasByText(filterText);

      filteredPersonas = removeDuplicates(filteredPersonas, currentPersonas);
      filteredPersonas = limitResults
        ? filteredPersonas.slice(0, limitResults)
        : filteredPersonas;
      return filterPromise(filteredPersonas);
    } else {
      return [];
    }
  };

  const filterPersonasByText = (filterText: string): IPersonaProps[] => {
    return allusers.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };

  const filterPromise = (
    personasToReturn: IPersonaProps[]
  ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    if (delayResults) {
      return convertResultsToPromise(personasToReturn);
    } else {
      return personasToReturn;
    }
  };

  const returnMostRecentlyUsed = (
    currentPersonas: IPersonaProps[]
  ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    return filterPromise(removeDuplicates(mostRecentlyUsed, currentPersonas));
  };

  function doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  }

  function removeDuplicates(
    personas: IPersonaProps[],
    possibleDupes: IPersonaProps[]
  ) {
    return personas.filter(
      (persona) => !listContainsPersona(persona, possibleDupes)
    );
  }

  function listContainsPersona(
    persona: IPersonaProps,
    personas: IPersonaProps[]
  ) {
    if (!personas || !personas.length || personas.length === 0) {
      return false;
    }
    return personas.filter((item) => item.text === persona.text).length > 0;
  }

  function convertResultsToPromise(
    results: IPersonaProps[]
  ): Promise<IPersonaProps[]> {
    return new Promise<IPersonaProps[]>((resolve, reject) =>
      setTimeout(() => resolve(results), 2000)
    );
  }

  function getTextFromItem(persona: IPersonaProps): string {
    return persona.text as string;
  }

  function validateInput(input: string): ValidationState {
    if (input.indexOf("@") !== -1) {
      return ValidationState.valid;
    } else if (input.length > 1) {
      return ValidationState.warning;
    } else {
      return ValidationState.invalid;
    }
  }

  return (
    <>
      <div className="innerToggleSection">
        {/* <PrimaryButton
            iconProps={{ iconName: "ManagerSelfService" }}
            text="Department"
          /> */}
        <button
          className="innerToggle"
          onClick={() => {
            setIsPG(!isPG);
          }}
          style={{ background: isPG ? "#ffe49f" : "#D8F4F7" }}
        >
          <Icon
            iconName={isPG ? "ManagerSelfService" : "Phone"}
            style={{ marginRight: 6 }}
          />
          {isPG ? "Department" : "Phone Guide"}
        </button>
      </div>
      {isPG ? (
        <ThemeProvider theme={myTheme}>
          {loader ? (
            <div className="spinnerBackground">
              <Spinner className="clsSpinner" size={SpinnerSize.large} />
            </div>
          ) : (
            <></>
          )}
          <div className="clsMaterialtab">
            <div className="clsFilters">
              <div className="clsFilterdropdowns">
                <NormalPeoplePicker
                  onResolveSuggestions={onFilterChanged}
                  getTextFromItem={getTextFromItem}
                  className={"ms-PeoplePicker"}
                  key={"normal"}
                  inputProps={{ placeholder: "Search User" }}
                  onValidateInput={validateInput}
                  selectedItems={selectedusers}
                  selectionAriaLabel={"Selected contacts"}
                  removeButtonAriaLabel={"Remove"}
                  resolveDelay={300}
                  itemLimit={1}
                  onChange={(data) => {
                    if (data.length > 0) {
                      setempname(data[0]["Email"]);
                      setselectedusers(data);
                      filtervalues(data[0]["Email"], zone, title);
                    } else {
                      setempname("");
                      setselectedusers([]);
                      filtervalues("", zone, title);
                    }
                  }}
                />
              </div>
              <div className="clsFilterdropdowns">
                <Dropdown
                  multiSelect
                  placeholder="Select Title"
                  options={titles}
                  selectedKeys={title}
                  onChange={(event, option, index) => {
                    console.log(option);
                    let tempTitle=title;
                    if (option) {
                      tempTitle = option.selected? [...tempTitle, option.key as string, ] : tempTitle.filter((key) => key !== option.key)
                    }

                    settitle(tempTitle);
                    filtervalues(empname, zone, tempTitle);
                  }}
                />
              </div>
              <div className="clsFilterdropdowns">
                <Dropdown
                  multiSelect
                  placeholder="Select Department"
                  options={alldepartment}
                  selectedKeys={dept}
                  onChange={(event, option, index) => {
                    console.log(option);
                    let tempDept=dept;
                    if (option) {
                      tempDept = option.selected? [...tempDept, option.key as string, ] : tempDept.filter((key) => key !== option.key)
                    }

                    setdept(tempDept);
                    filtervalues(empname, tempDept, title);
                  }}
                />
              </div>
              {/* <div className="clsFilterdropdowns">
                <Dropdown
                  multiSelect
                  placeholder="Select Zone"
                  options={zones}
                  selectedKeys={zone}
                  onChange={(event, option, index) => {
                    console.log(option);
                    let tempZone=zone;
                    if (option) {
                      tempZone = option.selected? [...tempZone, option.key as string, ] : tempZone.filter((key) => key !== option.key)
                    }

                    setzone(tempZone);
                    filtervalues(empname, tempZone, title);
                  }}
                />
              </div> */}
              
              <div className="clsFilterdropdowns">
                <PrimaryButton
                  iconProps={syncIcon}
                  onClick={() => {
                    setempname("");
                    setzone([]);
                    settitle([]);
                    setselectedusers([]);
                    setdept([]);
                    setPeopleList([...masterPeopleList])
                  }}
                />
              </div>
            </div>
            <div className={classes.root}>
              {/* <AppBar position="static" className="clsTabs">
          <Tabs
            variant="scrollable"
            scrollButtons="auto"
            value={value}
            onChange={handleChange}
            aria-label="simple tabs example"
          >
            {department.map(function (item, index) 
            {
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
        })} */}
              <MaterialDBNew Department={""} items={peopleList} />
            </div>
          </div>
        </ThemeProvider>
      ) : (
        <NewPivot />
      )}
    </>
  );
}
