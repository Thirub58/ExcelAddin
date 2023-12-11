// import { officeAddinConstants } from "../officeAddinConstants";
// export const formatRows = (args) => {
//   Excel.run(async (context) => {
//     try {
//       const sheet = context.workbook.worksheets.getActiveWorksheet();
//       const range = sheet.getRange(args.address);
//       range.load("values");
//       range.load("address");
//       await context.sync();
//       const cellIndex = range.address.toString().split("!")[1].slice(1);
//       const discountedPrice = sheet.getRange(`D${cellIndex}`);
//       const actualPrice = sheet.getRange(`C${cellIndex}`);
//       discountedPrice.load("values");
//       actualPrice.load("values");
//       await context.sync();
//       if (discountedPrice.values[0][0] != "") {
//         console.log(discountedPrice.values[0][0]);
//         const discountedPercentage = Math.round((discountedPrice.values[0][0] / actualPrice.values[0][0]) * 100);
//         await context.sync();
//         console.log("The Discounted Percentage is", discountedPercentage);
//         const entireRowRange = sheet.getRange(`A${cellIndex}:D${cellIndex}`);
//         sheet.load(officeAddinConstants.protectionMethod);
//         await context.sync();
//         if (sheet.protection.protected) {
//           console.log("Unprotecting the Sheet to Format the cells");
//           sheet.protection.unprotect(officeAddinConstants.password);
//           await context.sync();
//         }
//         if (discountedPercentage > 80 && discountedPercentage <= 90) {
//           entireRowRange.format.fill.color = "#FFB6C1";
//           entireRowRange.format.font.color = "#000000";
//         } else if (discountedPercentage > 70 && discountedPercentage <= 80) {
//           entireRowRange.format.fill.color = "#8B0000";
//           entireRowRange.format.font.color = "#FFFFFF";
//         } else if (discountedPercentage <= 70) {
//           entireRowRange.format.fill.color = "#FF0000";
//           entireRowRange.format.font.color = "#FFFFFF";
//         } else {
//           console.log("Discounted Price is greater than 90");
//         }
//         await context.sync()
//         sheet.load(officeAddinConstants.protectionMethod);
//         await context.sync();
//         if (!sheet.protection.protected) {
//           sheet.protection.protect(
//             {
//               allowInsertRows: false,
//               allowDeleteRows: false,
//               allowFormatRows: false,
//             },
//             officeAddinConstants.password
//           );
//           console.log("The Sheet is Protected Now");
//           await context.sync();
//         }
//         entireRowRange.load("values");
//         await context.sync();
//         console.log("The Entire Row Values is", entireRowRange.values);
//       }
//     } catch (error) {
//       console.log("Some Error in the Application:", error.message);
//     }
//   });
// };
