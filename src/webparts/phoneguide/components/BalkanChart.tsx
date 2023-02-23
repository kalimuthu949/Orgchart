import * as React from 'react';
import { useState, useEffect } from 'react'
import { SPComponentLoader } from '@microsoft/sp-loader';
import { MSGraphClient } from "@microsoft/sp-http";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
declare var OrgChart:any;
var chart:any;
let alldatafromAD=[];
export default function BalkanChart(props)
{
    
    const [loader,setloader]=React.useState(true);
    useEffect(()=>
    {
        
        getallusersgraph();
    });


    async function getnextitems(skiptoken)
    {
      await props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
            .api("users")
            .select("department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,businessPhones")
            .top(999)
            .skipToken(skiptoken)
            .get()
            .then(function (data) 
            {
              for(let i=0;i<data.value.length;i++)
              {
                if(data.value[i].userType!="Guest")  
                alldatafromAD.push(data.value[i]);
              }
  
              let strtoken='';
              if(data["@odata.nextLink"])
              {
                strtoken=data["@odata.nextLink"].split("skipToken=")[1];
                getnextitems(data["@odata.nextLink"].split("skipToken=")[1]);
              }
              else
              {
                loadChart(alldatafromAD);
              }
            })
        .catch(function (error) 
        { 
            console.log(error);
          
        })
      })
    }
  
    async function getallusersgraph() 
    {
      alldatafromAD=[];
      await props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
            .api("users")
            .select("department,mail,id,displayName,jobTitle,mobilePhone,manager,ext,givenName,surname,userPrincipalName,userType,businessPhones")
            .expand("manager")
            .top(999)
            .get()
            .then(function (data) 
            {
              console.log(data);
              for(let i=0;i<data.value.length;i++)
              {
                if(data.value[i].userType!="Guest")
                alldatafromAD.push(data.value[i]);
              }
  
              let strtoken='';
              if(data["@odata.nextLink"])
              {
                strtoken=data["@odata.nextLink"].split("skiptoken=")[1];
                getnextitems(data["@odata.nextLink"].split("skiptoken=")[1]);
              }
              else
              {
                
                loadChart(alldatafromAD);
              }
            })
            .catch(function (error) 
            { 
              console.log(error)
            })
      })
    }

function loadChart(data)
{
    
    
    let nodeData=[];
    for(var i=0;i<data.length;i++)
    {
        try{
        nodeData.push({ 
            id: data[i].id,
            pid: data[i].manager.id, 
            name: data[i].displayName, 
            title: data[i].jobTitle, 
            email: data[i].userPrincipalName, 
            img: "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].userPrincipalName 
        });
    }
        catch(e)
        {
            nodeData.push({ 
                id: data[i].id,
                name: data[i].displayName, 
                title: data[i].jobTitle, 
                email: data[i].userPrincipalName, 
                img: "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].userPrincipalName 
            });
        }
    }
    SPComponentLoader.loadScript("../../SiteAssets/Test/orgchart.js").then(function()
    {
            chart = new OrgChart(document.getElementById("OrgChart"), {
            template: "olivia",
            layout: OrgChart.mixed,
            showXScroll: OrgChart.scroll.visible,
            showYScroll: OrgChart.scroll.visible,
            mouseScrool: OrgChart.action.zoom,
            enableSearch: false,
            nodeBinding: {
            field_0: "name",
            field_1: "title",
            img_0: "img"
            },
            nodes:nodeData
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
    });

    setloader(false);
}

const preview = () => 
{
return OrgChart.pdfPrevUI.show(chart, {
format: 'A4'
});
};
    
    return(<div>
        {loader?<div className="spinnerBackground"><Spinner className="clsSpinner" size={SpinnerSize.large} /></div>:<></>}
        <div id='OrgChart'>
    </div></div>)
}
