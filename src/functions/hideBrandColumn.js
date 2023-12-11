import { officeAddinConstants } from "../officeAddinConstants";

export const hideBrandColumn = () => {
  Excel.run(async function (context) {
    const currentWorksheet = context.workbook.worksheets.getItem(officeAddinConstants.workings);
    currentWorksheet.activate();
    const range = currentWorksheet.getRange("B:B");
    range.columnHidden = true;
    const protectionOptions = {
      allowFormatCells: false,
    };
    const password = officeAddinConstants.password;
    currentWorksheet.load(officeAddinConstants.protectionMethod);
    await context.sync();
    if (!currentWorksheet.protection.protected) {
      currentWorksheet.protection.protect(protectionOptions, password);
      console.log("Current worksheet is password protected");
    }
    await context.sync();
  });
};
