import { async } from "regenerator-runtime";
import { loadBrandAndPriceColumn } from "./loadBrandAndPriceColumn";
import { officeAddinConstants } from "../officeAddinConstants";
export const productNameValidation = () => {
  Excel.run(async (context) => {
    const productsDatasheet = context.workbook.worksheets.getItem(officeAddinConstants.productsData);
    const workingsSheet = context.workbook.worksheets.getItem(officeAddinConstants.workings);
    const productNameRange = workingsSheet.getRange("A2:A1048576");
    productNameRange.load("address");
    productNameRange.dataValidation.clear();
    await context.sync();
    productNameRange.dataValidation.clear();
    await context.sync();
    const productTitleRange = productsDatasheet.getRange("B:B").getUsedRange();
    productTitleRange.load(officeAddinConstants.values);
    await context.sync();
    productNameRange.dataValidation.rule = {
      list: {
        inCellDropDown: true,
        source: productTitleRange,
      },
    };
    await context.sync()
    workingsSheet.onChanged.add(async (args) => {
      loadBrandAndPriceColumn(args);
    });
  });
};
