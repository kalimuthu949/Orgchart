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

export default function BalkanChart(props) {
  const [departmentConfigData, setDepartmentConfigData] = React.useState([]);
  const [departments, setdepartments] = React.useState([]);
  const [loader, setloader] = React.useState(true);
  const [userCount, setUserCount] = useState("");
  const [filterKeys, setFilterKeys] = React.useState({
    department: "Select",
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
    root: { height: 100 },
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
    getallusersgraph();
    getDepartmentConfigData();
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
                  if (userIdentity.toLowerCase() ==props.tenEmail && !userPrinName.includes("#EXT#"))
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
                  if (userIdentity.toLowerCase() ==props.tenEmail &&!userPrinName.includes("#EXT#"))
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
            Department: data[i].Department,
            Position: data[i].Position,
          });
        }

        setDepartmentConfigData([..._deptConfigData]);
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
      chart = new OrgChart(document.getElementById("OrgChart"), {
        // collapse: {
        //   level: 1,
        //   allChildren: true,
        // },
        layout: OrgChart.treeRightOffset,
        scaleInitial: 1,
        enableSearch: false,
        template: "olivia",
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
      OrgChart.scroll.smooth = 2;
      OrgChart.scroll.speed = 20;
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

  function loadChart(data) {
    const users = [];
    let arrdepartments = [];
    let arrDeptswithkey = [];
    let crntUserData = [];

    let nodeData = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i].userPrincipalName == props.userEmail) {
        crntUserData.push({
          imageUrl:
            "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
          isValid: true,
          email: data[i].userPrincipalName,
          ID: data[i].id,
          key: i,
          text: data[i].displayName,
          jobTitle: data[i].jobTitle,
          mobilePhone:
            data[i].businessPhones.length > 0 ? data[i].businessPhones[0] : [], //data[i].mobilePhone,
          department: data[i].department,
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
          jobTitle: data[i].jobTitle,
          mobilePhone:
            data[i].businessPhones.length > 0 ? data[i].businessPhones[0] : [], //data[i].mobilePhone,
          department: data[i].department,
          Zone: data[i].officeLocation ? data[i].officeLocation : "",
        });

        try {
          nodeData.push({
            id: data[i].id,
            pid: data[i].manager.id,
            ["Manager"]: data[i].manager.displayName,
            name: data[i].displayName,
            title: data[i].jobTitle ? data[i].jobTitle : "N/A",
            department: data[i].department ? data[i].department : "N/A",
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
            ["Manager"]: "N/A",
            name: data[i].displayName,
            title: data[i].jobTitle ? data[i].jobTitle : "N/A",
            department: data[i].department ? data[i].department : "N/A",
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

    console.log(arrdepartments);
    arrdepartments = removeDuplicatesfromarray(arrdepartments);
    console.log(arrdepartments);

    arrdepartments = arrdepartments.sort();

    for (var i = 0; i < arrdepartments.length; i++) {
      arrDeptswithkey.push({
        key: arrdepartments[i],
        text: arrdepartments[i],
      });
    }

    arrDeptswithkey.unshift({
      key: "Select",
      text: "Select",
    });

    setdepartments([...arrDeptswithkey]);
    setPeopleList([...users]);

    allNodeData = nodeData;
    SPComponentLoader.loadScript(
      props.URL + "/SiteAssets/OrgJS/orgchart.js"
    ).then(() => {
      chart = new OrgChart(document.getElementById("OrgChart"), {
        // collapse: {
        //   level: 1,
        //   allChildren: true,
        // },
        layout: OrgChart.treeRightOffset,
        scaleInitial: 1,
        enableSearch: false,
        template: "olivia",
        showXScroll: OrgChart.scroll.visible,
        showYScroll: OrgChart.scroll.visible,
        mouseScrool: OrgChart.action.scroll,
        nodeBinding: {
          field_0: "name",
          field_1: "title",
          img_0: "img",
        },
        nodes: nodeData,
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
      filterKeys.department = "Select";
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

    if (_filterKeys.department != "Select") {
      _filteredData = departmentConfigData.filter(
        (user) => user.Department == _filterKeys.department
      );

      if (_filteredData.length > 0) {
        filteredNodeData = _allNodeData.filter(
          (_data) =>
            _data.department == _filterKeys.department &&
            _data.title &&
            _data.title == _filteredData[0].Position
        );

        if (filteredNodeData.length == 0) {
          filteredNodeData = _allNodeData.filter(
            (_data) => _data.department == _filterKeys.department
          );
        }
      } else {
        filteredNodeData = _allNodeData.filter(
          (_data) => _data.department == _filterKeys.department
        );
      }
      LoadFilteredChartData(filteredNodeData);
    } else {
      LoadFilteredChartData([]);
    }

    setFilterKeys({ ..._filterKeys });
  }

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
            <NormalPeoplePicker
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
                filterKeys.department = "Select";
                setFilterKeys({ ...filterKeys });

                LoadFilteredChartData(data);
              }}
            />
          </div>
          <div className="clsDeptDrpDown" style={{ marginRight: 10 }}>
            <Dropdown
              placeholder="Select department"
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
