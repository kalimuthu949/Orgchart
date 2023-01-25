import * as React from "react";
import { useState, useEffect } from "react";
import { IColumn } from "@fluentui/react";

import { IPaginationObj } from "./ICommonServicesProps";


const paginateFunction = (totalPageItems:number,pagenumber:number, data:any[]):IPaginationObj => {
    let resultObj:IPaginationObj=null
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      resultObj={
        displayitems: paginatedItems,
        currentPage:pagenumber,
      }
    } else {
        resultObj={
            displayitems: [],
            currentPage:1,
          }
    }

    return resultObj;
  };

  const detailsListColumnSortingFunction = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn,
    masterColumns: IColumn[],
    data:any[]
  ): any[] => {
    const _columns = masterColumns;
    const newColumns: IColumn[] = _columns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    let sortData = _copyAndSort(
      data,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );


    return sortData;
  };
  const _copyAndSort=<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[]=> {
    let key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  const validationFunction=(type:string,value:string):boolean=>{
    let res:boolean=false
    if(type=='numberOnlyString'){
      let _value:string=value.toString();

      return /^[0-9]+$|^$/.test(_value)
    }
    else if(type=='email'){
    const validRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/;

    return value.match(validRegex)?true:false;
    }

    return res;
  }


  export default {
    paginateFunction,
    detailsListColumnSortingFunction,
    validationFunction
  };