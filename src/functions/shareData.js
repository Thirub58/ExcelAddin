import { officeAddinConstants } from "../officeAddinConstants";

export const shareData=()=>{
    Excel.run(async function (context) {
        try {
         console.log("Executing the share data function")
          const workingsWorkSheet = context.workbook.worksheets.getItem(officeAddinConstants.workings);
          const workingsTable=workingsWorkSheet.tables.getItem(officeAddinConstants.workingSet)
          const tableRange=workingsTable.getRange().getUsedRange()
          tableRange.load(officeAddinConstants.values)
          await context.sync()
          console.log("The Data is ",JSON.stringify(tableRange.values))
          
        } catch (error) {
          console.log(error);
        }
      });
}