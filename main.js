const Excel = require("exceljs");
const ifsc = require("ifsc");
const workbook = new Excel.Workbook();

const bank_list = {};
const valid_list = {};

workbook.xlsx.readFile("sample.xlsx").then(async function () {
  const worksheet = workbook.getWorksheet(1);
  const total = worksheet.rowCount;
  let count = 0,
    valid = 0,
    bar = "█";

  for (let i = 1; i <= total; i++) {
    const row = worksheet.getRow(i);
    const code = row.getCell(1).value;

    if (!(code in valid_list)) valid_list[code] = ifsc.validate(code);

    if (valid_list[code]) {
      if (!(code in bank_list)) bank_list[code] = await ifsc.fetchDetails(code);

      details = bank_list[code];
      row.getCell(2).value = details.BANK;
      row.getCell(3).value = details.BRANCH;
      valid++;
    } else {
      row.getCell(2).value = "Invalid IFSC";
      row.getCell(2).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0000" },
      };
    }
    count++;
    if (count % 20 === 0) bar += "█";
    console.clear();
    console.log(`\n${bar} \nProcessed ${count} out of ${total}`);
  }

  console.log(`\nValid: ${valid}\nInvalid: ${total - valid}`);
  console.log("\nDONE...");
  return workbook.xlsx.writeFile("output.xlsx");
});
