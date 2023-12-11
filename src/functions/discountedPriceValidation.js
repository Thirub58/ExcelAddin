
export const discountedPriceValidation = () => {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const discountedPriceRange = sheet.getRange("D:D");
    discountedPriceRange.dataValidation.clear();
    await context.sync();
    discountedPriceRange.dataValidation.rule = {
      wholeNumber: {
        formula1: "=C1",
        operator: Excel.DataValidationOperator.lessThan,
        format: {
          fill: {
              color: "red" // Background color: red
          }
      }
      },
    };

    discountedPriceRange.dataValidation.errorAlert = {
      message: "Discounted Price of the product must be less than Actual Price of the Product.",
      showAlert: true,
      style: Excel.DataValidationAlertStyle.stop,
      title: "Invalid Value",
    };
      
    await context.sync();
  });
};
