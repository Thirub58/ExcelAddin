/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { shareData } from "../functions/shareData";
import { hideBrandColumn } from "../functions/hideBrandColumn";
import { loadMasterData } from "../functions/loadMasterData";
import { officeAddinConstants } from "../officeAddinConstants";
import { lockTableHeaders } from "../functions/lockTableHeaders";
const setVisibility = (selector, hidden) =>
  hidden
    ? (document.querySelector(selector).style.display = "none")
    : (document.querySelector(selector).style.display = "block");

const { async } = require("regenerator-runtime");

/* global console, document, Excel, Office */

/**
 * a. Check if protected
 * if Yes, unprotect, hide, protect
 * a2 = if no, hide} context
 */
const hideWorksheet = async (context, name) => {
  const workbook = context.workbook;
  const worksheets = context.workbook.worksheets;
  const worksheet = worksheets.getItem(name);
  worksheet.visibility = Excel.SheetVisibility.hidden;
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("loadData").addEventListener("click", loadMasterData);
    document.getElementById("shareData").addEventListener("click", shareData);
    const runExcelAddin = (async () => {
      try {
        Excel.run(async (context) => {
          const worksheets = context.workbook.worksheets;
          worksheets.load("items");
          await context.sync();
          const requiredWorksheet = worksheets.getItemOrNullObject(officeAddinConstants.country_A);
          await context.sync();
          if (requiredWorksheet.isNullObject) {
            document.getElementById("message").textContent = "Please open a compatible excel file";
            setVisibility("#loadData", true);
            setVisibility("#shareData", true);
            return;
          }
          const workbook = context.workbook;
          workbook.load(officeAddinConstants.protectionMethod)
          await context.sync()
          if(workbook.protection.protected){
            workbook.protection.unprotect(officeAddinConstants.password)
            await context.sync()
            console.log("Unprotecting the Workbook")
          }
          const country_A=workbook.worksheets.getItem(officeAddinConstants.country_A)
          hideWorksheet(context, officeAddinConstants.productsData);
          hideWorksheet(context, officeAddinConstants.country_A);
          lockTableHeaders()
          const workings=context.workbook.worksheets.getItem("Workings")
          const range = country_A.getRange(officeAddinConstants.A1);
          range.load("values");
          await context.sync();
          if (range.values[0][0] != "ANZ") {
            console.log("Hide Brand Column");
            hideBrandColumn();
          }
           workbook.load(officeAddinConstants.protectionMethod)
           await context.sync()
           console.log(workbook.protection.protected)
          if (!workbook.protection.protected) {
            console.log("Protecting the workbook");
            workbook.protection.protect(officeAddinConstants.password);
            await context.sync()
          }
          console.log("Excel-addin Operation");
        });
      } catch (error) {
        console.log("Error message is :" + error);
      }
    })();
  }
});
