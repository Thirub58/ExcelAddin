import { officeAddinConstants } from "../officeAddinConstants";
import { productNameValidation } from "./productNameValidation";
import { lockTableHeaders } from "./lockTableHeaders";
export const loadMasterData = () => {
  Excel.run(async function (context) {
    try {
      const workbook = context.workbook;
      workbook.protection.unprotect(officeAddinConstants.password);
      console.log("The WorkBook is Unprotected");
      lockTableHeaders()
      const ProductDataWorksheet = workbook.worksheets.getItem(officeAddinConstants.productsData);
      const response = await fetch("https://dummyjson.com/products");
      const responseData = await response.json();
      const productsData = responseData.products.map((product) => {
        const newProduct = { ...product };
        delete newProduct.images;
        return Object.values(newProduct);
      });
      if (productsData.length > 0) {
        const rowCount = productsData.length;
        const endColumn = String.fromCharCode(65 + productsData[0].length - 1);
        const range = ProductDataWorksheet.getRange(`A1:${endColumn}${rowCount}`);
        range.values = productsData;
        await context.sync();
        productNameValidation();
        workbook.load("protection/protected")
        await context.sync()
        if(!workbook.protection.protected){
          console.log("Protecting the workbook")
          workbook.protection.protect(officeAddinConstants.password)
        }
        await context.sync();
        return;
      }
      console.log("Products Data length is Zero");
    } catch (error) {
      console.log("The Error in LoadMasterData is ", error.code);
    }
  });
};
