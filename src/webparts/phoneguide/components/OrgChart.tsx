import * as React from "react";
import { useEffect } from "react";
import "../assets/Css/org.css";
import {
  IPersonaProps,
  Persona,
  PersonaSize,
} from "@fluentui/react/lib/Persona";
import {
  NormalPeoplePicker,
  ValidationState,
} from "@fluentui/react/lib/Pickers";
import { IPhoneguideProps } from "./IPhoneguideProps";
import { Stack } from "@fluentui/react/lib/Stack";
import { graph } from "@pnp/graph/presets/all";
import { Icon } from "@fluentui/react/lib/Icon";
// import { Icon } from "office-ui-fabric-react/lib/Icon";
import "../../../../node_modules/office-ui-fabric-react/dist/css/fabric.min.css";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from "./Phoneguide.module.scss"
import { Dropdown, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import SPServices from "./SPServices";
import {sp} from "@pnp/sp/presets/all";
initializeIcons();
const MyIcon = () => <Icon iconName="CompassNW" />;
const Manager = [];
const Reportees = [];

import {
  HoverCard,
  IHoverCard,
  IPlainCardProps,
  HoverCardType,
  ThemeProvider,
  DefaultButton,
  mergeStyleSets,
} from "@fluentui/react";
import { StylesContext } from "@material-ui/styles";

const classNames = mergeStyleSets({
  plainCard: {
    width: 300,
    height: 400,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  target: {
    fontWeight: "600",
    display: "inline-block",
    border: "1px dashed #605e5c",
    padding: 5,
    borderRadius: 2,
  },
});

let userID = "";
export const OrgChart: React.FunctionComponent<IPhoneguideProps> = (
  props: IPhoneguideProps
) => {
  const [delayResults, setDelayResults] = React.useState(false);
  const [isPickerDisabled, setIsPickerDisabled] = React.useState(false);
  const [showSecondaryText, setShowSecondaryText] = React.useState(false);
  const [mostRecentlyUsed, setMostRecentlyUsed] = React.useState<
    IPersonaProps[]
  >([]);
  const [peopleList, setPeopleList] = React.useState([]);
  const [alluserdata,setalluserdata]=React.useState([]);
  const [ManagerList, setManagerList] = React.useState(Manager);
  const [ReporteeList, setReporteeList] = React.useState(Reportees);
  const [SelectedPerson, setSelectedPerson] = React.useState([]);
  const[SelectedPersonManager, setSelectedPersonManager] = React.useState("");
  const [userdatafromsharepoint,setuserdatafromsharepoint]=React.useState([]);
  const [CallLink, setCallLink] = React.useState("#");
  const [chatlink, setchatlink] = React.useState("#");
  const [loader,setloader]=React.useState(true);
  const [departments,setdepartments]=React.useState([]);
  const [selecteddeprt,setselecteddeprt]=React.useState("");
  const [deptuserscount,setdeptuserscount]=React.useState(0);
  const [userzonefromuserprofile,setuserzonefromuserprofile]=React.useState("");


  const departDrpdownoptions = [
    { key: 'A', text: 'Option a' },
  ];

  const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 }, root: { height: 100 } };

  const hoverCard = React.useRef<IHoverCard>(null);
  const instantDismissCard = (): void => {
    if (hoverCard.current) {
      hoverCard.current.dismiss();
    }
  };

  const onRenderPlainCard = (): JSX.Element => {
    console.log(hoverCard.current["props"].itemID);
    return (
      <div className={classNames.plainCard}>
        <div>{hoverCard.current["props"].itemID.email}</div>
        <Persona
          className="treeview-person"
          {...hoverCard.current["props"].itemID}
          size={PersonaSize.size48}
        />
        <div>
          <Icon iconName="Chat" />
          <Icon iconName="Phone" />
        </div>
        <div>
          <label>Contact</label>
          <div>
            <Icon iconName="Mail" />
            <label>{hoverCard.current["props"].itemID.email}</label>
          </div>
          <div>
            <Icon iconName="Phone" />
            <label>{hoverCard.current["props"].itemID.mobilePhone}</label>
          </div>
        </div>
      </div>
    );
  };
  const plainCardProps: IPlainCardProps = {
    onRenderPlainCard: onRenderPlainCard,
  };
  const onCardHide = (): void => {
    console.log("I am now hidden");
  };

  function ShowPopup() {
    const element = document.getElementById("myPopup");
    element.classList.add("visible");
  }

  function HidePopup() {
    const element = document.getElementById("myPopup");
    element.classList.remove("visible");
  }

  useEffect(() => {
    setloader(true);
    getalluserssp();
    getcurrentuser();
    getallusers();
  }, []);

  async function getcurrentuser() {
    await graph.me
      .select('mail,id,displayName,jobTitle,mobilePhone,department,officeLocation,businessPhones')
      .get()
      .then(function (data) {
        const cnrtUserDetails = [];
        cnrtUserDetails.push({
          imageUrl: "/_layouts/15/userphoto.aspx?size=L&username=" + data.mail,
          isValid: true,
          Email: data.mail,
          ID: data.id,
          key: 0,
          text: data.displayName,
          jobTitle: data.jobTitle,
          mobilePhone: data.businessPhones.length>0?data.businessPhones[0]:[],//data.mobilePhone,
          department:data.department,
          officeLocation:data.officeLocation,
        });
        if(data.department)
        {
           setselecteddeprt(data.department);
           getdeptcount(data.department);
        }
        else
        {
          setselecteddeprt("");
        }

        setManagerList([...cnrtUserDetails]);
        
        getDirectreports(data.id);
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
      });
  }

  async function getdeptcount(dept) {
    await graph.users
      .top(999)
      .select('mail,id,displayName,jobTitle,mobilePhone,department,officeLocation,businessPhones')
      .get()
      .then(function (data) 
      {
        console.log(data);
        const users = [];
        let countofusers=0;
        for (let i = 0; i < data.length; i++) 
        {
          if(dept==data[i].department)
          {
            countofusers=countofusers+1;
          }
        }

        setdeptuserscount(countofusers);

        
      })
      .catch(function (error) 
      {
        console.log(error);
      });
  }

  async function getusersfromselecteddept(dept) {
    await graph.users
      .top(999)
      .select('mail,id,displayName,jobTitle,mobilePhone,department,officeLocation,businessPhones')
      .get()
      .then(function (data) {
        console.log(data);
        const users = [];
        for (let i = 0; i < data.length; i++) 
        {
          
          if(dept==data[i].department){
          users.push({
            imageUrl:
              "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
            isValid: true,
            Email: data[i].mail,
            ID: data[i].id,
            key: i,
            text: data[i].displayName,
            jobTitle: data[i].jobTitle,
            mobilePhone: data[i].businessPhones.length>0?data[i].businessPhones[0]:[],//data[i].mobilePhone,
            department:data[i].department,
            officeLocation:data[i].officeLocation,
          });
          }
        }
        setloader(false);
        setdeptuserscount(users.length);
        if(users.length>0)
        {
          getSelecteduser([users[0]])
        }
        
      })
      .catch(function (error) {
        console.log(error);
      });
  }


  async function getalluserssp() {
    SPServices.SPReadItems({
      Listname: "EmployeeDetails",
      Select: "*,Employee/Title,Employee/Id,Employee/EMail",
      Expand: "Employee",
    })
      .then((items: any) => {
        setalluserdata([...items]);
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
      });
  }

  async function getallusers() {
    await graph.users
      .top(999)
      .select('mail,id,displayName,jobTitle,mobilePhone,department,officeLocation,userType,businessPhones')
      .get()
      .then(function (data) {
        console.log(data);
        const users = [];
        let arrdepartments=[];
        let arrDeptswithkey=[];
        console.log("length"+data.length);
        for (let i = 0; i < data.length; i++) 
        {
          if (props.context.pageContext.user.email == data[i].mail) {
            userID = data[i].id;
          }

          if( data[i].department)
          arrdepartments.push(data[i].department);

          if(data[i].userType!="Guest"){
          users.push({
            imageUrl:
              "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
            isValid: true,
            Email: data[i].mail,
            ID: data[i].id,
            key: i,
            text: data[i].displayName,
            jobTitle: data[i].jobTitle,
            mobilePhone: data[i].businessPhones.length>0?data[i].businessPhones[0]:[],//data[i].mobilePhone,
            department:data[i].department,
            officeLocation:data[i].officeLocation,

          });
          }

          if(i==data.length-1)
          {
              console.log("start")
          }
        }

        console.log("end");
        arrdepartments=removeDuplicatesfromarray(arrdepartments);

        for(var i=0;i<arrdepartments.length;i++)
        {
          arrDeptswithkey.push({ key: arrdepartments[i], text: arrdepartments[i]})
        }
        setdepartments([...arrDeptswithkey]);
        setPeopleList([...users]);
        setloader(false);
      })
      .catch(function (error) {
        console.log(error);
      });
  }

  function removeDuplicatesfromarray(arr) {
    return arr.filter((item, index) => arr.indexOf(item) === index);
  }

  async function getManagerforcard(userID,userEmail) 
  {
    setloader(true);

    let testdata=[];
    for(let i=0;i<alluserdata.length;i++)
    {
      if(alluserdata[i].Employee.EMail)
      {
        testdata.push(alluserdata[i]);
        setuserdatafromsharepoint(alluserdata[i])
        break;
      }
    }

    // const loginName = "i:0#.f|membership|"+userEmail;
    // const propertyName = "SPS-TimeZone";
    // const property = await sp.profiles.getUserProfilePropertyFor(loginName, propertyName).then(function (data: any) 
    // {
    //   setuserzonefromuserprofile(data)
    // }).catch(function (error) 
    // {
    //   console.log(error);
    //   setloader(false);
    // })

    await graph.users
      .getById(userID)
      .select('mail,id,displayName,jobTitle,mobilePhone,department,officeLocation,businessPhones')
      .manager()
      .then(function (data: any) {
        if (data) {
          setSelectedPersonManager(data.displayName);
          ShowPopup();
          setloader(false);
        }
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
        ShowPopup();
        setSelectedPersonManager("");
      });
  }

  async function getManager(userID) {
    setloader(true);
    await graph.users
      .getById(userID)
      .select('mail,id,displayName,jobTitle,mobilePhone,department,officeLocation,businessPhones')
      .manager()
      .then(function (data: any) {
        if (data) {
          const userdetails = [];
          userdetails.push({
            imageUrl:
              "/_layouts/15/userphoto.aspx?size=L&username=" + data.mail,
            ID: data.id,
            Manager: "",
            Email: data.mail,
            text: data.displayName,
            jobTitle: data.jobTitle,
            mobilePhone: data.businessPhones.length>0?data.businessPhones[0]:[],//data.mobilePhone,
            department:data.department,
            officeLocation:data.officeLocation,
          });
          getSelecteduser(userdetails);
        }
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
      });
  }

  async function getDirectreports(userID) {
    await graph.users
      .getById(userID)
      .select('mail,id,displayName,jobTitle,mobilePhone,department,officeLocation,businessPhones')
      .directReports()
      .then(function (data: any) {
        const directreports: any = [];
        for (let i = 0; i < data.length; i++) {
          directreports.push({
            imageUrl:
              "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
            ID: data[i].id,
            Email: data[i].mail,
            text: data[i].displayName,
            Manager: "",
            jobTitle: data[i].jobTitle,
            mobilePhone: data[i].businessPhones.length>0?data[i].businessPhones[0]:[],//data[i].mobilePhone,
            department:data[i].department,
            officeLocation:data[i].officeLocation,
          });
        }
        setReporteeList([...directreports]);
        setloader(false);
      })
      .catch(function (error) {
        console.log(error);
        setloader(false);
      });
  }

  async function getSelecteduser(userDetails) {
    const users = [];
    if (userDetails.length > 0) {
      for (let i = 0; i < peopleList.length; i++) {
        if (peopleList[i].ID == userDetails[0].ID) 
        {
          users.push({
            imageUrl:
              "/_layouts/15/userphoto.aspx?size=L&username=" +
              peopleList[i].Email,
            ID: peopleList[i].ID,
            Manager: "",
            Email: peopleList[i].Email,
            text: peopleList[i].text,
            jobTitle: peopleList[i].jobTitle,
            mobilePhone: peopleList[i].mobilePhone.length>0?peopleList[i].mobilePhone:"",//peopleList[i].mobilePhone,
            department:peopleList[i].department,
            officeLocation:peopleList[i].officeLocation,
          });
        }
      }
      if(users[0].department)
      {
        setselecteddeprt(users[0].department);
      }
      else
      {
        setselecteddeprt("");
      }
      setManagerList([...users]);
      await getDirectreports(userDetails[0].ID);
    } else {
      getcurrentuser();
    }
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



  return (
    <div>
      {loader?<div className="spinnerBackground"><Spinner className="clsSpinner" size={SpinnerSize.large} /></div>:<></>}
      <div className="searchDiv">
        <div
          className="clsBack"
          onClick={() => {
            getManager(ManagerList[0].ID);
            HidePopup();
          }}
        >
          <li>
            <Icon
              style={{ color: "#03606a" }}
              iconName="NavigateBack"
              title="Back"
            />
          </li>
        </div>
        <div className="clsDropplussearch">
        <div className="clsDeptCount">
          <label><b>Department User Count</b> : {deptuserscount}</label>
        </div>
        <div className="clsDeptDrpDown">
            <Dropdown
            placeholder="Select department"
            selectedKey={selecteddeprt}
            options={departments}
            styles={dropdownStyles}
            onChange={(event,option:any,index)=>{
                setselecteddeprt(option.key);
                setSelectedPerson([]);
                getusersfromselecteddept(option.key);
                getdeptcount(option.key)
            }}
          />
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
            onChange={(data) => {
              setloader(true);
              getSelecteduser(data);
              HidePopup();
            }}
          />
        </div>
        </div>
      </div>

      <div className="App">
        <div className="Manager">
          {ManagerList.map(function (item, key) {
            return (
              <div key={item.ID} className="treeview-parent">
                <Stack className="clsPersons">
                  <div className="treeview-stack">
                    <div className="treeview-content-top">
                      <div className="treeview-content-inner">
                        <div>
                          <a href="#" className="treeview-link">
                            <Persona
                              className="treeview-person"
                              {...item}
                              size={PersonaSize.size40}
                              secondaryText={item.jobTitle}
                              onClick={() => {
                                const userDetails = [];
                                userDetails.push(item);
                                setloader(true);
                                getSelecteduser(userDetails);
                                HidePopup();
                              }}
                            />
                          </a>
                        </div>
                        <div>
                          <a href="#" className="icon-link">
                            <Icon
                              iconName="ChevronRight"
                              onClick={() => {
                                const ItemID = item.ID;
                                setSelectedPerson([{ ...item }]);
                                setCallLink('https://teams.microsoft.com/l/call/0/0?users='+item.Email);
                                setchatlink('https://teams.microsoft.com/l/chat/0/0?users='+item.Email);
                                getManagerforcard(item.ID,item.Email);
                                
                              }}
                            />
                          </a>
                        </div>
                      </div>
                    </div>
                  </div>
                </Stack>
              </div>
            );
          })}
        </div>
        <div className="Reportees">
          <div className="Reportees-bg">
            <label>
              No. of reporting person to{" "}
              <b>{ManagerList.length > 0 ? ManagerList[0].text : ""}</b> (
              {ReporteeList.length})
            </label>
            <div className="Reportees-box">
              {ReporteeList.map((item, key) => {
                return (
                  <div key={item.ID} className="treeview-parent">
                    <Stack className="clsPersons">
                      <div className="treeview-stack">
                        <div className="treeview-content-top">
                          <div className="treeview-content-inner">
                            <div>
                              <a href="#" className="treeview-link">
                                <Persona
                                  className="treeview-person"
                                  {...item}
                                  secondaryText={item.jobTitle}
                                  size={PersonaSize.size40}
                                  onClick={() => {
                                    const userDetails = [];
                                    userDetails.push(item);
                                    setloader(true);
                                    getSelecteduser(userDetails);
                                    HidePopup();
                                  }}
                                />
                              </a>
                            </div>
                            <div>
                              <a href="#" className="icon-link">
                                <Icon
                                  iconName="ChevronRight"
                                  onClick={() => {
                                    const ItemID = item.ID;
                                    setSelectedPerson([{ ...item }]);
                                    setCallLink('https://teams.microsoft.com/l/call/0/0?users='+item.Email);
                                    setchatlink('https://teams.microsoft.com/l/chat/0/0?users='+item.Email);
                                    getManagerforcard(item.ID,item.Email);
                                  }}
                                />
                              </a>
                            </div>
                          </div>
                        </div>
                      </div>
                    </Stack>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
      {SelectedPerson.length > 0 ? (
        <div id="myPopup" className="clsPopup">
          <div className="clsFirstDiv">
            <div className="clsPersonDetails">
              <Persona
                className="treeview-person"
                {...SelectedPerson[0]}
                size={PersonaSize.size56}
              />
            </div>
            <div className="clsClose">
              <li>
                <Icon
                  iconName="ChromeClose"
                  onClick={() => {
                    setSelectedPerson([]);
                    HidePopup();
                  }}
                />
              </li>
            </div>
          </div>
          <div className="clsIcons">
            <li>
              <a href={chatlink} target="_blank" rel="noopener noreferrer">
                <Icon iconName="Chat" title="Chat" />
              </a>
            </li>
            <li>
              <a
                href={"mailto:" + SelectedPerson[0].Email}
                target="_blank"
                rel="noopener noreferrer"
              >
                <Icon iconName="Mail" title="Mail" />
              </a>
            </li>
            <li>
              <a href={CallLink} target="_blank" rel="noopener noreferrer">
                <Icon iconName="Phone" title="Phone" />
              </a>
            </li>
          </div>
          <div className="clsEmail">
            <h3>
              <b>Email</b>
            </h3>
            <div>{SelectedPerson[0].Email}</div>
          </div>
          <div className="clsContacts">
            <h3>
              <b>Contact</b>
            </h3>
            <div>
              {userdatafromsharepoint?(userdatafromsharepoint['Ext']?userdatafromsharepoint['Ext']:""):""}
              {" "}
              {SelectedPerson[0].mobilePhone
                ? SelectedPerson[0].mobilePhone
                : "N/A"}
            </div>
          </div>
          <div className="clsEmail">
            <h3>
              <b>JobTitle</b>
            </h3>
            <div>{SelectedPerson[0].jobTitle?SelectedPerson[0].jobTitle:"N/A"}</div>
          </div>
          <div className="clsEmail">
            <h3>
              <b>Department</b>
            </h3>
            <div>{SelectedPerson[0].department?SelectedPerson[0].department:"N/A"}</div>
          </div>
          <div className="clsEmail">
            <h3>
              <b>Manager</b>
            </h3>
            <div>{SelectedPersonManager?SelectedPersonManager:"N/A"}</div>
          </div>
          <div className="clsEmail">
            <h3>
              <b>Zone</b>
            </h3>
            {/* <div>{userdatafromsharepoint?(userdatafromsharepoint['Zone']?userdatafromsharepoint['Zone']:"N/A"):"N/A"}</div> */}
            {/* <div>{userzonefromuserprofile}</div> */}
            <div>{SelectedPerson[0].officeLocation?SelectedPerson[0].officeLocation:"N/A"}</div>
          </div>
        </div>
      ) : (
        <div id="myPopup" className="clsPopup">
          <div className="clsFirstDiv">
            <div className="clsPersonDetails">
              <Persona
                className="treeview-person"
                {...[]}
                size={PersonaSize.size56}
              />
            </div>
            <div className="clsClose">
              <li>
                <Icon
                  iconName="ChromeClose"
                  onClick={() => {
                    setSelectedPerson([]);
                    HidePopup();
                  }}
                />
              </li>
            </div>
          </div>
          <div className="clsIcons">
            <li>
              <Icon iconName="Chat" title="Chat" />
            </li>
            <li>
              <Icon iconName="Mail" title="Mail" />
            </li>
            <li>
              <Icon iconName="Phone" title="Phone" />
            </li>
          </div>
          <div className="clsEmail">
            <h3>
              <b>Email</b>
            </h3>
            <div>N/A</div>
          </div>
          <div className="clsContacts">
            <h3>
              <b>Contact</b>
            </h3>
            <div>N/A</div>
          </div>
        </div>
      )}
    </div>
  );
};
