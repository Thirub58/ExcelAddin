import { officeAddinConstants } from "../officeAddinConstants";
export const lockTableHeaders = () => {
  Excel.run(async function (context) {
    const worksheet = context.workbook.worksheets.getItem("Workings");
    worksheet.load(officeAddinConstants.protectionMethod)
    await context.sync()
    if(worksheet.protection.protected){
      console.log("UnProtecting the Worksheet in lock TableHeaders Function")
      worksheet.protection.unprotect(officeAddinConstants.password)
      await context.sync()
    }
    const tableHeaderRange=worksheet.getRange("A1:D1")
    tableHeaderRange.load(officeAddinConstants.cellProtection)
    await context.sync()
    tableHeaderRange.format.protection.locked=true
    await context.sync()
    worksheet.load(officeAddinConstants.protectionMethod)
    await context.sync()
    if(!worksheet.protection.protected){
      console.log("Protecting the WorkSheet in lockTable Headers")
      worksheet.protection.protect( {
        allowInsertRows: false,
        allowDeleteRows: false,
      },officeAddinConstants.password)
      await context.sync()
    }
    console.log("Locking the table Headers")
    await context.sync()
  });
};
