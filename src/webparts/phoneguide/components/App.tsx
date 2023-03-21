import * as React from "react";

import { OrgChart } from "./OrgChart";
import MaterialDtabs from "./Materialtabs";
import NewPivot from "./NewPivot";
import "../assets/Css/App.scss";
import { useState } from "react";
import BalkanChart from "./BalkanChart";
// import { initializeIcons } from "@uifabric/icons/fonts";
// initializeIcons("@uifabric/icons/fonts");
const App = (props) => {
  const [activeTab, setActiveTab] = useState("OrgChart");
  let tenantEmail = "chandrudemo.onmicrosoft.com";
  //let tenantEmail="hosthealthcare.onmicrosoft.com";
  return (
    <>
      <div className="headerAndTabSection">
        <div className="Title-section">
          <h2>
            {" "}
            {activeTab === "OrgChart" ? "Organization Chart" : "Employee Guide"}
          </h2>
        </div>
        <div className="Toggle-section">
          <button
            className={`${activeTab === "OrgChart" ? "Active" : ""}`}
            onClick={() => setActiveTab("OrgChart")}
          >
            Organization Chart
          </button>
          <button
            className={`${activeTab === "PhoneGuide" ? "Active" : ""}`}
            onClick={() => setActiveTab("PhoneGuide")}
          >
            Employee Guide
          </button>
          {/* <button
          className={`${activeTab === "Pivot" ? "Active" : ""}`}
          onClick={() => setActiveTab("Pivot")}
        >
          Pivot
        </button> */}
        </div>
      </div>
      <div>
        {activeTab === "OrgChart" ? (
          //<OrgChart context={props.context} />
          <BalkanChart
            propertyPaneProps={props.propertyPaneProps}
            context={props.context}
            URL={props.URL}
            userEmail={props.context.pageContext.user.email}
            tenEmail={tenantEmail}
          />
        ) : activeTab === "PhoneGuide" ? (
          <MaterialDtabs
            propertyPaneProps={props.propertyPaneProps}
            context={props.context}
            tenEmail={tenantEmail}
          />
        ) : (
          ""
        )}
      </div>
    </>
  );
};
export default App;
