import * as React from "react";
import { useState, useEffect } from "react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { MSGraphClient } from "@microsoft/sp-http";
import "../../phoneguide/assets/Css/Balkan.scss";
import "../assets/Css/org.css";
import SPServices from "./SPServices";
// import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
// import { Dropdown, IDropdownStyles } from "@fluentui/react/lib/Dropdown";
// import {
//   IPersonaProps,
//   Persona,
//   PersonaSize,
// } from "@fluentui/react/lib/Persona";
// import {
//   NormalPeoplePicker,
//   ValidationState,
// } from "@fluentui/react/lib/Pickers";

import {
  Dropdown,
  IDropdownStyles,
  Spinner,
  SpinnerSize,
  IPersonaProps,
  Persona,
  PersonaSize,
  NormalPeoplePicker,
  ValidationState,
  Label,
  ILabelStyles,
  Icon,
  mergeStyleSets,
} from "@fluentui/react";
import { Users } from "@pnp/graph/users";
import { func } from "prop-types";
// import orgChartJs from "../assets/Js/orgchart.js";
declare var OrgChart: any;
// let orgJS = "../assets/Js/orgchart.js";
var chart: any;
let alldatafromAD = [];
let allNodeData = [];
import ProdData from "./ProdData";
import Autocomplete from "@material-ui/lab/Autocomplete";
import TextField from "@material-ui/core/TextField";

export default function BalkanChart(props) {
  var testing = ProdData.cmd();
  var proddata = JSON.parse(testing);
  const [departmentConfigData, setDepartmentConfigData] = React.useState([]);
  const [departments, setdepartments] = React.useState([]);
  const [loader, setloader] = React.useState(true);
  const [userCount, setUserCount] = useState("");
  const [filterKeys, setFilterKeys] = React.useState({
    department: "All Host Healthcare",
    peoplePicker: [],
  });

  const [delayResults, setDelayResults] = React.useState(false);
  const [isPickerDisabled, setIsPickerDisabled] = React.useState(false);
  const [showSecondaryText, setShowSecondaryText] = React.useState(false);
  const [mostRecentlyUsed, setMostRecentlyUsed] = React.useState<
    IPersonaProps[]
  >([]);
  const [peopleList, setPeopleList] = React.useState([]);

  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 },
    root: {
      height: 100,
      selectors: {
        ".ms-Dropdown-title": {
          height: 37,
          paddingTop: 4,
        },
        ".ms-Dropdown-caretDownWrapper": {
          top: 3,
        },
      },
    },
  };

  const iconStyles = mergeStyleSets({
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 20,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 2,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
  });
  useEffect(() => {
    getDepartmentConfigData();
    // getallusersgraph();
  }, []);

  async function getnextitems(skiptoken) {
    await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          .select(
            "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities"
          )
          .expand("manager")
          .top(999)
          .skipToken(skiptoken)
          .get()
          .then(function (data) {
            let condition: boolean;
            for (let i = 0; i < data.value.length; i++) {
              let userIdentity = data.value[i].identities[0].issuer;
              let userPrinName = data.value[i].userPrincipalName
                ? data.value[i].userPrincipalName
                : "";
              if (!props.propertyPaneProps.propertyToggle) {
                if (userIdentity) {
                  if (
                    userIdentity.toLowerCase() == props.tenEmail &&
                    !userPrinName.includes("#EXT#")
                  )
                    alldatafromAD.push(data.value[i]);
                }
              } else {
                alldatafromAD.push(data.value[i]);
              }
            }

            let strtoken = "";
            if (data["@odata.nextLink"]) {
              strtoken = data["@odata.nextLink"].split("skipToken=")[1];
              getnextitems(data["@odata.nextLink"].split("skipToken=")[1]);
            } else {
              loadChart(alldatafromAD);
            }
          })
          .catch(function (error) {
            console.log(error);
          });
      });
  }

  async function getallusersgraph() {
    alldatafromAD = [];
    await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          .select(
            "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones,officeLocation,identities"
          )
          .expand("manager")
          .top(999)
          .get()
          .then(function (data) {
            let condition: boolean = false;
            for (let i = 0; i < data.value.length; i++) {
              let userIdentity = data.value[i].identities[0].issuer;
              let userPrinName = data.value[i].userPrincipalName
                ? data.value[i].userPrincipalName
                : "";
              if (!props.propertyPaneProps.propertyToggle) {
                if (userIdentity) {
                  if (
                    userIdentity.toLowerCase() == props.tenEmail &&
                    !userPrinName.includes("#EXT#")
                  )
                    alldatafromAD.push(data.value[i]);
                }
              } else {
                alldatafromAD.push(data.value[i]);
              }
            }

            let strtoken = "";
            if (data["@odata.nextLink"]) {
              strtoken = data["@odata.nextLink"].split("skiptoken=")[1];
              getnextitems(data["@odata.nextLink"].split("skiptoken=")[1]);
            } else {
              loadChart(alldatafromAD);
            }
          })
          .catch(function (error) {
            console.log(error);
          });
      });
  }

  async function getDepartmentConfigData() {
    SPServices.SPReadItems({
      Listname: "DepartmentConfigList",
    })
      .then((data: any) => {
        let _deptConfigData = [];
        for (let i = 0; i < data.length; i++) {
          _deptConfigData.push({
            ID: data[i].ID,
            Department: data[i].Department ? data[i].Department.trim() : "",
            Position: data[i].Position ? data[i].Position.trim() : "",
          });
        }

        setDepartmentConfigData([..._deptConfigData]);
        getEmployeeDetails();
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
      });
  }

  function removeDuplicatesfromarray(arr) {
    return arr.filter((item, index) => arr.indexOf(item) === index);
  }

  /* start for peoplepicker */
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
    return peopleList.filter((item) =>
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
    //return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    return text.toLowerCase().includes(filterText);
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
  /* End for peoplepicker */

  function LoadFilteredChartData(userData) {
    setloader(true);

    let _nodeData = [];
    let parentData = [];
    let childData = [];
    let pidnull = true;
    if (userData.length > 0) {
      for (let i = 0; i < userData.length; i++) {
        parentData = [];
        childData = [];
        parentData = allNodeData.filter(
          (_people) => _people.email == userData[i].email
        );

        childData = allNodeData.filter(
          (_people) => _people.pid == parentData[0].id
        );

        _nodeData = [..._nodeData, ...parentData, ...childData];

        while (childData.length > 0) {
          let tempChildData = [];
          pidnull = false;
          for (let i = 0; i < childData.length; i++) {
            tempChildData = [
              ...tempChildData,
              ...allNodeData.filter(
                (_people) => _people.pid == childData[i].id
              ),
            ];
          }
          childData = tempChildData;
          _nodeData = [..._nodeData, ...childData];
        }
      }
    } else {
      _nodeData = [...allNodeData];
      pidnull = false;
    }

    _nodeData = _nodeData.filter(
      (value, index, self) => index === self.findIndex((t) => t.id === value.id)
    );

    try {
      OrgChart.templates.myTemplate = Object.assign(
        {},
        OrgChart.templates.olivia
      );
      OrgChart.templates.myTemplate.field_0 =
        '<text data-width="130" data-text-overflow="multiline" style="font-size: 15px;width: 10px;font-weight:600;" fill="#03606a" x="100" y="25">{val}</text>';
      OrgChart.templates.myTemplate.field_1 =
        '<text data-width="130" data-text-overflow="multiline" style="font-size: 13px;width: 10px;" fill="#757575" x="100" y="70" text-anchor="middle">{val}</text>';

      chart = new OrgChart(document.getElementById("OrgChart"), {
        // collapse: {
        //   level: 1,
        //   allChildren: true,
        // },
        layout: OrgChart.treeRightOffset,
        // scaleInitial: 1,
        enableSearch: false,
        template: "myTemplate",
        showXScroll: OrgChart.scroll.none,
        showYScroll: OrgChart.scroll.none,
        mouseScrool: OrgChart.action.scroll,
        nodeBinding: {
          field_0: "name",
          field_1: "title",
          img_0: "img",
        },
        nodes: _nodeData,
        editForm: {
          generateElementsFromFields: false,
          elements: [
            { type: "textbox", label: "Name", binding: "name" },
            { type: "textbox", label: "Job Title", binding: "title" },
            { type: "textbox", label: "Email", binding: "email" },
            { type: "textbox", label: "Contact", binding: "Mobile Phone" },
            { type: "textbox", label: "Department", binding: "department" },
            { type: "textbox", label: "Manager", binding: "Manager" },
            { type: "textbox", label: "Zone", binding: "Zone" },
          ],
        },
      });
      // OrgChart.scroll.smooth = 2;
      // OrgChart.scroll.speed = 20;
      chart.on("expcollclick", function (sender, collapse, id, ids) {
        if (!collapse) {
          sender.expand(id, ids, function () {
            sender.center(id);
          });

          return false;
        }
      });
    } catch (e) {
      console.log(e);
    }

    setUserCount(
      _nodeData.length <= 9 && pidnull ? "one" : ""
      // : userData.length == 1 && _nodeData.length > 1
      // ? "many"
      // : "all"
    );
    // let selectedSVG = document.querySelector("#OrgChart svg");
    // selectedSVG.setAttribute("height", "auto");
    setloader(false);
  }

  const getEmployeeDetails = () => {
    SPServices.SPReadItems({
      Listname: "EmployeeGroupDetails",
      Select: "*,Manager/Title,Manager/Id,Manager/EMail",
      Expand: "Manager",
    })
      .then((data: any) => {
        let employeeArr = [];
        for (const item of data) {
          if (item.UserPrincipalName) {
            employeeArr.push({
              mail: item.UserPrincipalName,
              id: item.Title,
              displayName: [item.FirstName, item.LastName].join(" "),
              userPrincipalName: item.UserPrincipalName,
              jobTitle: item.JobTitle,
              givenName: item.FirstName,
              LastName: item.LastName,
              businessPhones: item.PhoneNumber
                ? item.PhoneNumber.split(",")
                : [],
              department: item.Department,
              officeLocation: item.Zone,
              manager: item.ManagerId ? item.Manager.Title : "",
              managerAzureId: item.ManagerId ? item.ManagerAzureId : "",
            });
          }
        }
        //console.log(employeeArr);
        loadChart(employeeArr);
      })
      .catch((error) => {
        console.log(error);
        setloader(false);
      });
  };

  function loadChart(data) {
    //console.log(JSON.stringify(data));
    data = proddata;
    const users = [];
    let arrdepartments = [];
    let arrDeptswithkey = [];
    let crntUserData = [];

    let nodeData = [];
    for (var i = 0; i < data.length; i++) {
      var loggedUserEmail = props.userEmail;
      //var loggedUserEmail="EPC@hosthealthcare.com";
      if (data[i].userPrincipalName == loggedUserEmail) {
        crntUserData.push({
          imageUrl:
            "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
          isValid: true,
          email: data[i].userPrincipalName,
          ID: data[i].id,
          key: i,
          text: data[i].displayName,
          surname: data[i].LastName,
          jobTitle: data[i].jobTitle,
          mobilePhone:
            data[i].businessPhones.length > 0 ? data[i].businessPhones[0] : [], //data[i].mobilePhone,
          department: data[i].department ? data[i].department.trim() : "",
          Zone: data[i].officeLocation ? data[i].officeLocation : "",
        });
      }

      if (data[i].userPrincipalName) {
        users.push({
          imageUrl:
            "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
          isValid: true,
          email: data[i].userPrincipalName,
          ID: data[i].id,
          key: i,
          text: data[i].displayName,
          surname: data[i].LastName,
          jobTitle: data[i].jobTitle,
          mobilePhone:
            data[i].businessPhones.length > 0 ? data[i].businessPhones[0] : [], //data[i].mobilePhone,
          department: data[i].department ? data[i].department.trim() : "",
          Zone: data[i].officeLocation ? data[i].officeLocation : "",
        });

        try {
          nodeData.push({
            id: data[i].id,
            // pid: data[i].manager.id,
            // ["Manager"]: data[i].manager.displayName,
            pid: data[i].managerAzureId,
            surname: data[i].LastName,
            ["Manager"]: data[i].manager,
            name: data[i].displayName,
            title: data[i].jobTitle ? data[i].jobTitle : "N/A",
            department: data[i].department ? data[i].department.trim() : "N/A",
            email: data[i].userPrincipalName
              ? data[i].userPrincipalName
              : "N/A",
            ["Zone"]: data[i].officeLocation ? data[i].officeLocation : "N/A",
            ["Mobile Phone"]:
              data[i].businessPhones.length > 0
                ? data[i].businessPhones[0]
                : "N/A",
            img:
              "/_layouts/15/userphoto.aspx?size=L&username=" +
              data[i].userPrincipalName,

            // ["Testing"]: <div>Hi World</div>,
          });
        } catch (e) {
          nodeData.push({
            id: data[i].id,
            pid: null,
            surname: data[i].LastName,
            ["Manager"]: "N/A",
            name: data[i].displayName,
            title: data[i].jobTitle ? data[i].jobTitle : "N/A",
            department: data[i].department ? data[i].department.trim() : "N/A",
            email: data[i].userPrincipalName
              ? data[i].userPrincipalName
              : "N/A",
            ["Zone"]: data[i].officeLocation ? data[i].officeLocation : "N/A",
            ["Mobile Phone"]:
              data[i].businessPhones.length > 0
                ? data[i].businessPhones[0]
                : "N/A",
            img:
              "/_layouts/15/userphoto.aspx?size=L&username=" +
              data[i].userPrincipalName,
            // ["Testing"]: <div>Hi World</div>,
          });
        }
      }
      if (data[i].department) arrdepartments.push(data[i].department.trim());
    }

    //console.log(arrdepartments);
    arrdepartments = removeDuplicatesfromarray(arrdepartments);
    //console.log(arrdepartments);

    arrdepartments = arrdepartments.sort();

    for (var i = 0; i < arrdepartments.length; i++) {
      arrDeptswithkey.push({
        key: arrdepartments[i],
        text: arrdepartments[i],
      });
    }

    arrDeptswithkey.unshift({
      key: "All Host Healthcare",
      text: "All Host Healthcare",
    });

    setdepartments([...arrDeptswithkey]);
    setPeopleList([...users]);

    allNodeData = nodeData;

    //console.log(JSON.stringify(allNodeData));

    SPComponentLoader.loadScript(
      props.URL + "/SiteAssets/OrgJS/orgchart.js"
    ).then(() => {
      OrgChart.templates.myTemplate = Object.assign(
        {},
        OrgChart.templates.olivia
      );
      OrgChart.templates.myTemplate.field_0 =
        '<text data-width="130" data-text-overflow="multiline" style="font-size: 16px;width: 10px;font-weight:600;" fill="#03606a" x="100" y="35">{val}</text>';
      OrgChart.templates.myTemplate.field_1 =
        '<text data-width="130" data-text-overflow="multiline" style="font-size: 14px;width: 10px;" fill="#757575" x="100" y="70" text-anchor="middle">{val}</text>';

      chart = new OrgChart(document.getElementById("OrgChart"), {
        // collapse: {
        //   level: 1,
        //   allChildren: true,
        // },
        layout: OrgChart.treeRightOffset,
        scaleInitial: 1,
        enableSearch: false,
        template: "myTemplate",
        showXScroll: OrgChart.scroll.visible,
        showYScroll: OrgChart.scroll.visible,
        mouseScrool: OrgChart.action.scroll,
        nodeBinding: {
          field_0: "name",
          field_1: "title",
          img_0: "img",
        },
        nodes: [],
        editForm: {
          generateElementsFromFields: false,
          elements: [
            { type: "textbox", label: "Name", binding: "name" },
            { type: "textbox", label: "Job Title", binding: "title" },
            { type: "textbox", label: "Email", binding: "email" },
            { type: "textbox", label: "Contact", binding: "Mobile Phone" },
            { type: "textbox", label: "Department", binding: "department" },
            { type: "textbox", label: "Manager", binding: "Manager" },
            { type: "textbox", label: "Zone", binding: "Zone" },
          ],
        },
      });
      OrgChart.scroll.smooth = 2;
      OrgChart.scroll.speed = 50;
      filterKeys.peoplePicker = crntUserData;
      filterKeys.department = "All Host Healthcare";
      chart.on("expcollclick", function (sender, collapse, id, ids) {
        if (!collapse) {
          sender.expand(id, ids, function () {
            sender.center(id);
          });

          return false;
        }
      });

      setTimeout(() => {
        LoadFilteredChartData([...crntUserData]);
        setloader(false);
      }, 2000);
    });
  }

  function filterList(_filterKeys) {
    let _filteredData = [];
    let filteredNodeData = [];

    let _allNodeData = [...allNodeData];

    if (_filterKeys.department != "All Host Healthcare") {
      _filteredData = departmentConfigData.filter(
        (user) => user.Department == _filterKeys.department
      );

      if (_filteredData.length > 0) {
        let positions = [];
        positions = _filteredData[0].Position
          ? _filteredData[0].Position.toLowerCase().split(";")
          : [];
        filteredNodeData = _allNodeData.filter(
          (_data) =>
            _data.department == _filterKeys.department &&
            _data.title &&
            //_data.title == _filteredData[0].Position
            //positions.includes(_data.title)
            positions.indexOf(_data.title.toLowerCase()) !== -1
        );

        if (filteredNodeData.length == 0) {
          filteredNodeData = _allNodeData.filter(
            (_data) => _data.department == _filterKeys.department && _data.name
          );
        }
      } else {
        filteredNodeData = _allNodeData.filter(
          (_data) => _data.department == _filterKeys.department && _data.name
        );
      }

      if (_filterKeys.department == "Recruitment")
        filteredNodeData = filteredNodeData.splice(6, 1);

      LoadFilteredChartData(filteredNodeData);
    } else {
      LoadFilteredChartData([]);
    }

    setFilterKeys({ ..._filterKeys });
  }

  const sortFunction = (a, b, key) => {
    if (a[key] < b[key]) {
      return -1;
    }
    if (a[key] > b[key]) {
      return 1;
    }
    return 0;
  };

  return (
    <div>
      {loader ? (
        <div className="spinnerBackground">
          <Spinner className="clsSpinner" size={SpinnerSize.large} />
        </div>
      ) : (
        <></>
      )}
      <div className="searchDiv">
        <div className="clsDropplussearch">
          <div style={{ marginRight: 10 }}>
            <Autocomplete
              title="Search User"
              aria-label="Search User"
              id="combo-box-demo"
              options={peopleList}
              placeholder="Search User"
              value={
                filterKeys.peoplePicker.length > 0
                  ? filterKeys.peoplePicker[0]
                  : {}
              }
              //defaultValue={filterKeys.peoplePicker.length>0?filterKeys.peoplePicker[0]:{}}
              getOptionLabel={(option) => option.text}
              onChange={(event, data) => {
                let newData = [];
                if (data) {
                  newData.push(data);
                  filterKeys.peoplePicker = newData;
                } else {
                  filterKeys.peoplePicker = [];
                }

                filterKeys.department = "All Host Healthcare";
                setFilterKeys({ ...filterKeys });
                LoadFilteredChartData(newData);
              }}
              style={{ width: 300 }}
              renderInput={(params) => (
                <TextField {...params} label="Search User" variant="outlined" />
              )}
            />

            {/* <NormalPeoplePicker
              onResolveSuggestions={onFilterChanged}
              getTextFromItem={getTextFromItem}
              className={"ms-PeoplePicker"}
              key={"normal"}
              inputProps={{ placeholder: "Search User" }}
              onValidateInput={validateInput}
              selectionAriaLabel={"Selected contacts"}
              removeButtonAriaLabel={"Remove"}
              resolveDelay={300}
              itemLimit={1}
              disabled={isPickerDisabled}
              selectedItems={filterKeys.peoplePicker}
              onChange={(data: any) => {
                filterKeys.peoplePicker = data;
                filterKeys.department = "All Host Healthcare";
                setFilterKeys({ ...filterKeys });
                LoadFilteredChartData(data);
              }} 
            />*/}
          </div>
          <div className="clsDeptDrpDown" style={{ marginRight: 10 }}>
            <Dropdown
              placeholder="All Host Healthcare"
              selectedKey={filterKeys.department}
              options={departments}
              styles={dropdownStyles}
              onChange={(event, option: any) => {
                filterKeys.peoplePicker = [];
                filterKeys.department = option.key;
                filterList(filterKeys);
              }}
            />
          </div>
          {true ? (
            <div className="clsDeptCount">
              <Label>
                Department User Count :{" "}
                {
                  allNodeData.filter((_nodeData) => {
                    return _nodeData.department == filterKeys.department;
                  }).length
                }
              </Label>
            </div>
          ) : null}
          <div>
            <Icon
              iconName="Refresh"
              title="Click to reset"
              className={iconStyles.refresh}
              onClick={() => {
                filterKeys.peoplePicker = [];
                filterKeys.department = "All";
                filterList(filterKeys);
              }}
            />
          </div>
        </div>
      </div>
      <div id={`OrgChart`} data-count={userCount}></div>
    </div>
  );
}
