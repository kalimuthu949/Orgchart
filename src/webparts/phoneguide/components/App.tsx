import * as React from "react";

import { OrgChart } from "./OrgChart";
import MaterialDtabs from "./Materialtabs";
import NewPivot from "./NewPivot";
import "../assets/Css/App.scss";
import { useState } from "react";
// import { initializeIcons } from "@uifabric/icons/fonts";
// initializeIcons("@uifabric/icons/fonts");
const App = (props) => {
  const [activeTab, setActiveTab] = useState("PhoneGuide");
  return (
    <>
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
          Phone Guide
        </button>
        {/* <button
          className={`${activeTab === "Pivot" ? "Active" : ""}`}
          onClick={() => setActiveTab("Pivot")}
        >
          Pivot
        </button> */}
      </div>
      <div>
        {activeTab === "OrgChart" ? (
          <OrgChart context={props.context} />
        ) : activeTab === "PhoneGuide" ? (
          <MaterialDtabs />
        ) : (
          ""
        )}
      </div>
    </>
  );
};
export default App;
