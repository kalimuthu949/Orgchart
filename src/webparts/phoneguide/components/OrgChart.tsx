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
  const [ManagerList, setManagerList] = React.useState(Manager);
  const [ReporteeList, setReporteeList] = React.useState(Reportees);
  const [SelectedPerson, setSelectedPerson] = React.useState([]);
  const [CallLink, setCallLink] = React.useState("#");

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
    getcurrentuser();
    getallusers();
  }, []);

  async function getcurrentuser() {
    await graph.me
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
          mobilePhone: data.mobilePhone,
        });
        setManagerList([...cnrtUserDetails]);
        getDirectreports(data.id);
      })
      .catch(function (error) {
        console.log(error);
      });
  }
  async function getallusers() {
    await graph.users
      .top(999)
      .get()
      .then(function (data) {
        console.log(data);
        const users = [];

        for (let i = 0; i < data.length; i++) {
          if (props.context.pageContext.user.email == data[i].mail) {
            userID = data[i].id;
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
          });
        }
        console.log(users);
        setPeopleList([...users]);
      })
      .catch(function (error) {
        console.log(error);
      });
  }
  async function getManager(userID) {
    await graph.users
      .getById(userID)
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
            mobilePhone: data.mobilePhone,
          });
          getSelecteduser(userdetails);
        }
      })
      .catch(function (error) {
        console.log(error);
      });
  }

  async function getDirectreports(userID) {
    await graph.users
      .getById(userID)
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
            Manager: userID,
            jobTitle: data[i].jobTitle,
            mobilePhone: data[i].mobilePhone,
          });
        }
        setReporteeList([...directreports]);
      })
      .catch(function (error) {
        console.log(error);
      });
  }

  async function getSelecteduser(userDetails) {
    const users = [];
    if (userDetails.length > 0) {
      for (let i = 0; i < peopleList.length; i++) {
        if (peopleList[i].ID == userDetails[0].ID) {
          users.push({
            imageUrl:
              "/_layouts/15/userphoto.aspx?size=L&username=" +
              peopleList[i].Email,
            ID: peopleList[i].ID,
            Manager: "",
            Email: peopleList[i].Email,
            text: peopleList[i].text,
            jobTitle: peopleList[i].jobTitle,
            mobilePhone: peopleList[i].mobilePhone,
          });
        }
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
              getSelecteduser(data);
              HidePopup();
            }}
          />
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
                                setCallLink(
                                  "https://teams.microsoft.com/_#/conversations/19:" +
                                    userID +
                                    "_" +
                                    ItemID +
                                    "@unq.gbl.spaces?ctx=chat"
                                );
                                ShowPopup();
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
                                    setCallLink(
                                      "https://teams.microsoft.com/_#/conversations/19:" +
                                        userID +
                                        "_" +
                                        ItemID +
                                        "@unq.gbl.spaces?ctx=chat"
                                    );
                                    ShowPopup();
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
              <a href={CallLink} target="_blank" rel="noopener noreferrer">
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
              {SelectedPerson[0].mobilePhone
                ? SelectedPerson[0].mobilePhone
                : "N/A"}
            </div>
          </div>
        </div>
      ) : (
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
            <div>Kalimuthu@chandrudemo.onmicrosoft.com</div>
          </div>
          <div className="clsContacts">
            <h3>
              <b>Contact</b>
            </h3>
            <div>9942757885</div>
          </div>
        </div>
      )}
    </div>
  );
};
