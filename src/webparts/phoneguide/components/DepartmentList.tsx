import * as React from "react";
import { useState, useEffect } from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Label,
  ILabelStyles,
  Dropdown,
  IDropdownStyles,
  IDropdownOption,
  IColumn,
  mergeStyleSets,
  Spinner,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPeopleObj } from "./ISPServicesProps";
import Pagination from "office-ui-fabric-react-pagination";

import Services from "./SPServices";
import CommonServices from "./CommonServices";

import EditIcon from "@material-ui/icons/Edit";
import * as moment from "moment";
import { CircularProgress } from "@material-ui/core";

interface IProps {
  context: any;
  spContext: WebPartContext;
  graphContext: any;
  peopleList: IPeopleObj[];
  filterKeys: {};
  // FilterChoices: any;
  EditPageNavigate: any;
}

let totalPageItems: number = 10;
let sortData: any[] = [];

const DepartmentList = (props: IProps): JSX.Element => {
  let loggedUserName: string = props.spContext.pageContext.user.displayName;
  let loggedUserEmail: string = props.spContext.pageContext.user.email;
  let allPeoples = props.peopleList;
  let ThisWeekNumber = moment().isoWeek();
  let NextWeekNumber = moment().add(1, "week").isoWeek();
  let LastWeekNumber = moment().subtract(1, "week").isoWeek();
  let ThisMonthNumber = moment().month();

  // Variable-Declaration Starts
  const _columns: IColumn[] = [
    {
      key: "Column1",
      name: "Service Number",
      fieldName: "JobNumber",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column2",
      name: "Consumer",
      fieldName: "clientName",
      minWidth: 100,
      maxWidth: 150,

      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column3",
      name: "Consumer Address",
      fieldName: "clientAddress",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Contractor",
      fieldName: "ContractorName",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column5",
      name: "Provider",
      fieldName: "ProviderName",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column6",
      name: "Int. Staff Assigned",
      fieldName: "ACEManagerName",
      minWidth: 100,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column7",
      name: "Status",
      fieldName: "Status",
      minWidth: 200,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Status == "Not started" ? (
            <div className={statusDesign.NotStarted}>{item.Status}</div>
          ) : item.Status == "Job Completed & invoiced" ? (
            <div className={statusDesign.InvoiceCompleted}>{item.Status}</div>
          ) : item.Status == "Booked and confirmed with client" ? (
            <div className={statusDesign.BookedandConfirmed}>{item.Status}</div>
          ) : item.Status == "Declined" ? (
            <div className={statusDesign.Declined}>{item.Status}</div>
          ) : (
            ""
          )}
        </>
      ),
    },
    {
      key: "Column8",
      name: "Service Type",
      fieldName: "ServiceType",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column9",
      name: "Booking date",
      fieldName: "BookingDate",
      minWidth: 100,
      maxWidth: 100,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => moment(item.BookingDate).format("DD/MM/YYYY"),
    },
    {
      key: "Column10",
      name: "Action",
      fieldName: "Action",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => {
        return (
          <p style={{ margin: "0" }}>
            <EditIcon
              style={{ color: "#4194C5", cursor: "pointer" }}
              onClick={() => props.EditPageNavigate(item.ID)}
            />
          </p>
        );
      },
    },
  ];

  // Variable-Declaration Ends
  // Style-Section Starts
  const statusDesign = mergeStyleSets({
    NotStarted: [
      {
        backgroundColor: "rgb(241,236,187,100%)",
        padding: "5px 10px",
        borderRadius: "15px",
        width: "180px",
        textAlign: "center",
        margin: "0",
      },
    ],
    BookedandConfirmed: [
      {
        backgroundColor: "rgb(65,148,197,30%)",
        padding: "5px 10px",
        borderRadius: "15px",
        width: "180px",
        textAlign: "center",
        margin: "0",
      },
    ],
    InvoiceCompleted: [
      {
        backgroundColor: "rgb(88,214,68,35%)",
        padding: "5px 10px",
        width: "180px",
        borderRadius: "15px",
        textAlign: "center",
        margin: "0",
      },
    ],
    Declined: [
      {
        backgroundColor: "#ffcccb",
        padding: "5px 10px",
        width: "180px",
        borderRadius: "15px",
        textAlign: "center",
        margin: "0",
      },
    ],
  });
  // Style-Section Ends
  // State-Declaration Starts
  const [masterData, setMasterData] = useState<any[]>([]);
  const [data, setData] = useState<any[]>([]);
  const [displayData, setDisplayData] = useState<any[]>([]);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [columns, setColumns] = useState<IColumn[]>(_columns);
  const [tempPeople, setTempPeople] = useState<any[]>([]);
  const [loader, setLoader] = useState<boolean>(true);
  // State-Declaration Ends

  // Function-Declaration Starts
  /* get Client Details function */
  const getClientDetails = async () => {
    let arrClientDetails: any[] = [];
    await Services.SPReadItems({
      Listname: "Client Details",
      Select: "*, Provider/ID, Provider/Title",
      Expand: "Provider",
    })
      .then((rec: any) => {
        if (rec.length > 0) {
          rec.forEach((data) => {
            arrClientDetails.push({
              Title: data.Title ? data.Title : "",
              Email: data.Email ? data.Email : "",
              PhoneNumber: data.PhoneNumber ? data.PhoneNumber : "",
              Address: data.Address ? data.Address : "",
              NOK: data.NOK ? data.NOK : "",
              NOKName: data.NOKName ? data.NOKName : "",
              NOKPhoneNumber: data.NOKPhoneNumber ? data.NOKPhoneNumber : "",
              ProviderId: data.ProviderId ? data.ProviderId : null,
              Provider: data.ProviderId ? data.Provider.Title : "",
              ID: data.ID,
            });
          });
        }
        getItems(arrClientDetails);
      })
      .catch((error: any) => {
        console.log(error);
      });
  };

  /* get ServiceArchive Details function */
  const getItems = async (arrClientDetails: any[]) => {
    let drpFilterChoices = {
      serviceType: [{ key: "All", text: "All" }],
      contractor: [{ key: "All", text: "All" }],
      provider: [{ key: "All", text: "All" }],
      status: [{ key: "All", text: "All" }],
      jobView: [
        { key: "All", text: "All" },
        { key: "Last week", text: "Last week" },
        { key: "This week", text: "This week" },
        { key: "Next week", text: "Next week" },
        { key: "This month", text: "This month" },
      ],
      showAll: [
        { key: "All", text: "All" },
        { key: "Mine", text: "Mine" },
      ],
    };
    await Services.SPReadItems({
      Listname: "ServiceArchive",
      Select:
        "*, Client/ID, Client/Title, Client/PhoneNumber, Provider/ID, Provider/Title, Contractor/ID, Contractor/Title,ACEManager/Title",
      Expand: "Client, Provider, Contractor, ACEManager",
      Topcount: 5000,
      Orderbydecorasc: false,
    }).then((items: any) => {
      let _data = [];
      if (items.length > 0) {
        items.forEach(async (item, index) => {
          _data.push({
            ID: item["ID"],
            JobNumber: item["JobNumber"] ? item["JobNumber"] : "",
            clientId: item["ClientId"] ? item["ClientId"] : null,
            clientName: item["ClientId"] ? item["Client"]["Title"] : "",
            clientAddress: item["ClientId"]
              ? arrClientDetails.filter((e) => e.ID == item["ClientId"])
                  .length > 0
                ? arrClientDetails.filter((e) => e.ID == item["ClientId"])[0]
                    .Address
                : ""
              : "",
            ContractorId: item["ContractorId"] ? item["ContractorId"] : null,
            ContractorName: item["ContractorId"]
              ? item["Contractor"]["Title"]
              : "",
            ProviderId: item["ProviderId"] ? item["ProviderId"] : null,
            ProviderName: item["ProviderId"] ? item["Provider"]["Title"] : "",
            ACEManager: item["ACEManagerId"] ? item["ACEManagerId"] : null,
            ACEManagerName: item["ACEManagerId"]
              ? item["ACEManager"]["Title"]
              : "",
            ACEManagerEmail: item["ACEManagerId"]
              ? item["ACEManager"]["EMail"]
              : null,
            Status: item["Status"] ? item["Status"] : "",
            BookingDate: item["BookingDate"],
            ServiceType: item["ServiceType"] ? item["ServiceType"] : "",
          });
        });
      }
      setMasterData([..._data]);
      filterFunction([..._data]);
    });
  };

  const filterFunction = (data) => {
    let tempArr = data;
    if (props.filterKeys["serviceType"] != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.ServiceType == props.filterKeys["serviceType"];
      });
    }
    if (props.filterKeys["contractor"] != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.ContractorName == props.filterKeys["contractor"];
      });
    }
    if (props.filterKeys["provider"] != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.ProviderName == props.filterKeys["provider"];
      });
    }
    if (props.filterKeys["status"] != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Status == props.filterKeys["status"];
      });
    }
    if (props.filterKeys["search"] != "") {
      tempArr = tempArr.filter((arr) => {
        return (
          arr.JobNumber &&
          arr.clientName &&
          arr.clientAddress &&
          arr.ContractorName &&
          arr.ProviderName &&
          arr.ACEManagerName &&
          arr.Status &&
          arr.ServiceType
        );
      });

      tempArr = tempArr.filter((arr) => {
        return (
          arr.JobNumber.toLowerCase().includes(
            props.filterKeys["search"].toLowerCase()
          ) ||
          arr.clientName
            .toLowerCase()
            .includes(props.filterKeys["search"].toLowerCase()) ||
          arr.clientAddress
            .toLowerCase()
            .includes(props.filterKeys["search"].toLowerCase()) ||
          arr.ContractorName.toLowerCase().includes(
            props.filterKeys["search"].toLowerCase()
          ) ||
          arr.ProviderName.toLowerCase().includes(
            props.filterKeys["search"].toLowerCase()
          ) ||
          arr.ACEManagerName.toLowerCase().includes(
            props.filterKeys["search"].toLowerCase()
          ) ||
          arr.Status.toLowerCase().includes(
            props.filterKeys["search"].toLowerCase()
          ) ||
          arr.ServiceType.toLowerCase().includes(
            props.filterKeys["search"].toLowerCase()
          )
        );
      });
    }

    setData([...tempArr]);
    sortData = tempArr;
    let paginatedData = CommonServices.paginateFunction(
      totalPageItems,
      currentPage,
      tempArr
    );
    setDisplayData([...paginatedData.displayitems]);
    setCurrentPage(paginatedData.currentPage);
    setLoader(false);
  };

  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    let sortedData = CommonServices.detailsListColumnSortingFunction(
      ev,
      column,
      _columns,
      sortData
    );

    let paginatedData = CommonServices.paginateFunction(
      totalPageItems,
      currentPage,
      sortedData
    );

    setData([...sortedData]);
    setDisplayData([...paginatedData.displayitems]);
    setCurrentPage(paginatedData.currentPage);
  };

  const GetUserDetails = (filterText: any): IPeopleObj[] => {
    let result: any = allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string): boolean => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };
  // Function-Declaration Ends

  useEffect(() => {
    setLoader(true);
    getClientDetails();
  }, [props.filterKeys]);

  return (
    <>
      {loader ? (
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
            height: "100vh",
            width: "100%",
          }}
        >
          <CircularProgress style={{ color: "#4194c5" }} />
        </div>
      ) : (
        <div>
          <div>
            <DetailsList
              items={displayData}
              columns={columns}
              styles={{
                root: {
                  ".ms-DetailsRow-cell": {
                    display: "flex",
                    alignItems: "center",
                  },
                },
              }}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
            />
          </div>
          {data.length > 0 ? (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                margin: "10px 0",
              }}
            >
              <Pagination
                currentPage={currentPage}
                totalPages={
                  data.length > 0 ? Math.ceil(data.length / totalPageItems) : 1
                }
                onChange={(page: number) => {
                  let paginatedData = CommonServices.paginateFunction(
                    totalPageItems,
                    page,
                    data
                  );
                  setDisplayData([...paginatedData.displayitems]);
                  setCurrentPage(paginatedData.currentPage);
                }}
              />
            </div>
          ) : (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                marginTop: 15,
              }}
            >
              <Label style={{ color: "#000" }}>No data found !!!</Label>
            </div>
          )}
        </div>
      )}
    </>
  );
};

export default DepartmentList;