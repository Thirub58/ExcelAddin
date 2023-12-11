import { discountedPriceValidation } from "./discountedpricevalidation";
import { officeAddinConstants } from "../officeAddinConstants";
import { formatRows } from "./formatRows";
const lockPreviouslyUsedRange = async (context, column) => {
  const sheet = context.workbook.worksheets.getItem(officeAddinConstants.workings);
  const previousUsedRange = sheet.getRange(column).getUsedRange();
  previousUsedRange.load("address");
  await context.sync();
  if (previousUsedRange.address.toString().length > 11) {
    const previousEndRange = previousUsedRange.address.toString().split("!")[1].split(":")[1];
    if (previousEndRange[0] == "B") {
      const rangeToBeLocked = sheet.getRange(`B2:${previousEndRange}`);
      return rangeToBeLocked;
    } else {
      const rangeToBeLocked = sheet.getRange(`C2:${previousEndRange}`);
      return rangeToBeLocked;
    }
  }
  return null;
};

const loadProductData = async (context, cellIndex, productTitleIndex) => {
  const productSheet = context.workbook.worksheets.getItem(officeAddinConstants.productsData);
  const productRange = productSheet.getRange(`${cellIndex}${productTitleIndex + 1}`);
  productRange.load("values");
  await context.sync();
  return productRange.values;
};
export const loadBrandAndPriceColumn = (args) => {
  if (!loadBrandAndPriceColumn.modifiedCellIndexArray) {
    loadBrandAndPriceColumn.modifiedCellIndexArray = [];
    console.log("The Array is declared");
  }
  const selectedRange = args.address;

  if (selectedRange[0] == "A") {
    Excel.run(async (context) => {
      try {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(selectedRange);
        sheet.load(officeAddinConstants.protectionMethod);
        await context.sync();
        if (sheet.protection.protected) {
          sheet.protection.unprotect(officeAddinConstants.password);
          await context.sync();
          console.log("The Sheet is unprotected");
        }
        const brandrangeToBeLocked = await lockPreviouslyUsedRange(context, "B:B");
        await context.sync();
        const pricerangeToBeLocked = await lockPreviouslyUsedRange(context, "C:C");
        await context.sync();
        const productSheet = context.workbook.worksheets.getItem(officeAddinConstants.productsData);
        const addedCellIndex = selectedRange.slice(1);
        const productTitleRange = productSheet.getRange("B:B").getUsedRange();
        productTitleRange.load(officeAddinConstants.values);
        await context.sync();
        const productTitles = productTitleRange.values;
        range.load(officeAddinConstants.values);
        await context.sync();
        let productTitleIndex = 0;
        for (let i = 0; i < productTitles.length; i++) {
          if (productTitles[i][0].toString() == range.values.toString()) {
            productTitleIndex = i;
            break;
          }
        }
        const productBrandData = await loadProductData(context, "H", productTitleIndex);
        await context.sync();
        const productPriceData = await loadProductData(context, "D", productTitleIndex);
        await context.sync();
        const brandColumnWorkingsRange = sheet.getRange(`B${addedCellIndex}`);
        brandColumnWorkingsRange.values = productBrandData;
        const priceColumnWorkingsRange = sheet.getRange(`C${addedCellIndex}`);
        priceColumnWorkingsRange.values = productPriceData;

        //pushing the Brand and Price Values cell Range to array
        loadBrandAndPriceColumn.modifiedCellIndexArray.push([`B${addedCellIndex}`, `C${addedCellIndex}`]);

        const sheetRange = sheet.getRange();
        sheetRange.load(officeAddinConstants.cellProtection);
        await context.sync();

        sheetRange.format.protection.locked = false;
        await context.sync();
        const tableHeaderRange = sheet.getRange("A1:D1");
        tableHeaderRange.load(officeAddinConstants.cellProtection);
        await context.sync();
        tableHeaderRange.format.protection.locked = true;
        await context.sync();
        if (brandrangeToBeLocked != null && pricerangeToBeLocked != null) {
          brandrangeToBeLocked.format.protection.locked = true;
          pricerangeToBeLocked.format.protection.locked = true;
          await context.sync();
        }
        //Iterating through the array and locking the cells
        for (let i = 0; i < loadBrandAndPriceColumn.modifiedCellIndexArray.length; i++) {
          const brandRange = sheet.getRange(loadBrandAndPriceColumn.modifiedCellIndexArray[i][0]);
          const priceRange = sheet.getRange(loadBrandAndPriceColumn.modifiedCellIndexArray[i][1]);
          brandRange.load(officeAddinConstants.cellProtection);
          priceRange.load(officeAddinConstants.cellProtection);
          await context.sync();
          brandRange.format.protection.locked = true;
          priceRange.format.protection.locked = true;
          await context.sync();
        }

        sheet.load(officeAddinConstants.protectionMethod);
        await context.sync();
        if (!sheet.protection.protected) {
          sheet.protection.protect(
            {
              allowInsertRows: false,
              allowDeleteRows: false,
              allowFormatRows: false,
            },
            officeAddinConstants.password
          );
          console.log("The Sheet is protected");
        }
        discountedPriceValidation();
        await context.sync();
      } catch (error) {
        console.error("Error Message:", error.message);
      }
    });
  }
  if (selectedRange[0] == "D") {
   formatRows(args)
  }
};
