import * as React from "react";
import { useState, useEffect } from "react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { MSGraphClient } from "@microsoft/sp-http";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import "../../phoneguide/assets/Css/Balkan.scss";
import { Dropdown, IDropdownStyles } from "@fluentui/react/lib/Dropdown";
import SPServices from "./SPServices";
import {
  IPersonaProps,
  Persona,
  PersonaSize,
} from "@fluentui/react/lib/Persona";
import {
  NormalPeoplePicker,
  ValidationState,
} from "@fluentui/react/lib/Pickers";
import { Users } from "@pnp/graph/users";
import { func } from "prop-types";
// import orgChartJs from "../assets/Js/orgchart.js";
declare var OrgChart: any;
// let orgJS = "../assets/Js/orgchart.js";
var chart: any;
let alldatafromAD = [];
let allNodeData = [];

export default function BalkanChart(props) {
  const [departmentConfigData,setDepartmentConfigData]=React.useState([]);
  const [departments, setdepartments] = React.useState([]);
  const [loader, setloader] = React.useState(true);

  const [filterKeys,setFilterKeys]=React.useState({
    department:'Select',
    peoplePicker:[]
  })

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
            "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,businessPhones"
          )
          .top(999)
          .skipToken(skiptoken)
          .get()
          .then(function (data) {
            for (let i = 0; i < data.value.length; i++) {
              if (data.value[i].userType != "Guest")
                alldatafromAD.push(data.value[i]);
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
            "department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones"
          )
          .expand("manager")
          .top(999)
          .get()
          .then(function (data) {
            console.log(data);
            for (let i = 0; i < data.value.length; i++) {
              if (data.value[i].userType != "Guest")
                alldatafromAD.push(data.value[i]);
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
      Select: "*,Employee/Title,Employee/Id,Employee/EMail",
      Expand: "Employee",
    })
      .then((data:any) => {
        let _deptConfigData=[];
        for(let i=0;i<data.length;i++){
          if(data[i].EmployeeId){
            _deptConfigData.push({
              ID:data[i].ID,
              Employee:{
                EmployeeId:data[i].EmployeeId,
                EmployeeName:data[i].Employee.Title,
                EmployeeEmail:data[i].Employee.EMail,
              },
              Department:data[i].Department,
              Position:data[i].Position,
            })
          }
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

    if(userData.length>0){
      for(let i=0;i<userData.length;i++){
        parentData = [];
        childData = [];
        parentData = allNodeData.filter(
          (_people) => _people.email.trim() == userData[i].email.trim()
        );
    
        childData = allNodeData.filter(
          (_people) => _people.pid == parentData[0].id
        );
    
        _nodeData = [..._nodeData,...parentData, ...childData];
    
        while (childData.length > 0) {
          let tempChildData = [];
          for (let i = 0; i < childData.length; i++) {
            tempChildData = [
              ...allNodeData.filter((_people) => _people.pid == childData[i].id),
            ];
          }
          childData = tempChildData;
          _nodeData = [..._nodeData, ...childData];
        }
      }
    }else{
      _nodeData=[...allNodeData]
    }

    chart = new OrgChart(document.getElementById("OrgChart"), {
      template: "olivia",
      layout: OrgChart.mixed,
      showXScroll: OrgChart.scroll.visible,
      showYScroll: OrgChart.scroll.visible,
      mouseScrool: OrgChart.action.scroll,
      enableSearch: false,
      nodeBinding: {
        field_0: "name",
        field_1: "title",
        img_0: "img",
      },
      nodes: _nodeData,
    });

    setloader(false);
  }

  function loadChart(data) {
    const users = [];
    let arrdepartments = [];
    let arrDeptswithkey = [];

    let nodeData = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i].department) arrdepartments.push(data[i].department);

      if (data[i].userType != "Guest") {
        users.push({
          imageUrl:
            "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
          isValid: true,
          email: data[i].mail,
          ID: data[i].id,
          key: i,
          text: data[i].displayName,
          jobTitle: data[i].jobTitle,
          mobilePhone:
            data[i].businessPhones.length > 0 ? data[i].businessPhones[0] : [], //data[i].mobilePhone,
          department: data[i].department,
          officeLocation: data[i].officeLocation,
        });
      }

      try {
        nodeData.push({
          id: data[i].id,
          pid: data[i].manager.id,
          name: data[i].displayName,
          title: data[i].jobTitle,
          department:data[i].department,
          email: data[i].userPrincipalName,
          img:
            "/_layouts/15/userphoto.aspx?size=L&username=" +
            data[i].userPrincipalName,
        });
      } catch (e) {
        nodeData.push({
          id: data[i].id,
          pid:null,
          name: data[i].displayName,
          title: data[i].jobTitle,
          department:data[i].department,
          email: data[i].userPrincipalName,
          img:
            "/_layouts/15/userphoto.aspx?size=L&username=" +
            data[i].userPrincipalName,
        });
      }
    }

    arrdepartments = removeDuplicatesfromarray(arrdepartments);

    for (var i = 0; i < arrdepartments.length; i++) {
      arrDeptswithkey.push({
        key: arrdepartments[i],
        text: arrdepartments[i],
      });
    }

    arrDeptswithkey.unshift({
      key:'Select',
      text:'Select'
    })
    setdepartments([...arrDeptswithkey]);
    setPeopleList([...users]);

    allNodeData = nodeData;
    SPComponentLoader.loadScript(
      props.URL + "/SiteAssets/Test/orgchart.js"
    ).then(
      // SPComponentLoader.loadScript(orgJS).then(
      function () {
        chart = new OrgChart(document.getElementById("OrgChart"), {
          template: "olivia",
          layout: OrgChart.mixed,
          showXScroll: OrgChart.scroll.visible,
          showYScroll: OrgChart.scroll.visible,
          mouseScrool: OrgChart.action.scroll,
          enableSearch: false,
          nodeBinding: {
            field_0: "name",
            field_1: "title",
            img_0: "img",
          },
          nodes: nodeData,
          // nodes: [
          // { id: "1", name: "Jack Hill", title: "Chairman and CEO", email: "amber@domain.com", img: "https://balkangraph.com/js/img/1.jpg" },
          // { id: "2", pid: "1", name: "Lexie Cole", title: "QA Lead", email: "ava@domain.com", img: "https://balkangraph.com/js/img/2.jpg" },
          // { id: "3", pid: "1", name: "Janae Barrett", title: "Technical Director", img: "https://balkangraph.com/js/img/3.jpg" },
          // { id: "4", pid: "1", name: "Aaliyah Webb", title: "Manager", email: "jay@domain.com", img: "https://balkangraph.com/js/img/4.jpg" },
          // { id: "5", pid: "2", name: "Elliot Ross", title: "QA", img: "https://balkangraph.com/js/img/5.jpg" },
          // { id: "6", pid: "2", name: "Anahi Gordon", title: "QA", img: "https://balkangraph.com/js/img/6.jpg" },
          // { id: "7", pid: "2", name: "Knox Macias", title: "QA", img: "https://balkangraph.com/js/img/7.jpg" },
          // { id: "8", pid: "3", name: "Nash Ingram", title: ".NET Team Lead", email: "kohen@domain.com", img: "https://balkangraph.com/js/img/8.jpg" },
          // { id: "9", pid: "3", name: "Sage Barnett", title: "JS Team Lead", img: "https://balkangraph.com/js/img/9.jpg" },
          // { id: "10", pid: "8", name: "Alice Gray", title: "Programmer", img: "https://balkangraph.com/js/img/10.jpg" },
          // { id: "11", pid: "8", name: "Anne Ewing", title: "Programmer", img: "https://balkangraph.com/js/img/11.jpg" },
          // { id: "12", pid: "9", name: "Reuben Mcleod", title: "Programmer", img: "https://balkangraph.com/js/img/12.jpg" },
          // { id: "13", pid: "9", name: "Ariel Wiley", title: "Programmer", img: "https://balkangraph.com/js/img/13.jpg" },
          // { id: "14", pid: "4", name: "Lucas West", title: "Marketer", img: "https://balkangraph.com/js/img/14.jpg" },
          // { id: "15", pid: "4", name: "Adan Travis", title: "Designer", img: "https://balkangraph.com/js/img/15.jpg" },
          // { id: "16", pid: "4", name: "Alex Snider", title: "Sales Manager", img: "https://balkangraph.com/js/img/16.jpg" }
          // ]
        });
      }
    );

    setloader(false);
  }
  
  function filterList(_filterKeys){
    let _filteredData = [];
    let filteredNodeData=[];

    let _allNodeData=[...allNodeData];

    if(_filterKeys.department !="Select"){
      _filteredData=departmentConfigData.filter((user)=>user.Department==_filterKeys.department);

      if( _filteredData.length > 0){
        filteredNodeData = _allNodeData.filter((_data)=>_data.department==_filterKeys.department && _data.title == _filteredData[0].Position);
      }else{
        filteredNodeData = _allNodeData.filter((_data)=>_data.department==_filterKeys.department);
      }
      LoadFilteredChartData(filteredNodeData);
    }else{
      LoadFilteredChartData([])
    }

    setFilterKeys({..._filterKeys})
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
          <div className="clsDeptCount">
            <label>
              <b>Department User Count</b> : {0}
            </label>
          </div>
          <div className="clsDeptDrpDown">
          <Dropdown
            placeholder="Select department"
            selectedKey={filterKeys.department}
            options={departments}
            styles={dropdownStyles}
            onChange={(event, option: any) => {
              filterKeys.peoplePicker=[];
              filterKeys.department=option.key;
              filterList(filterKeys);
            }}
          />
      </div>
        </div>
      </div>
      
      <div>
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
          onChange={(data:any) => {
            filterKeys.peoplePicker=data;
            filterKeys.department='Select';
            setFilterKeys({...filterKeys});

            LoadFilteredChartData(data);
          }}
        />
      </div>
      <div id="OrgChart"></div>
    </div>
  );
}
