const Excel = require("exceljs");
const workbook = new Excel.Workbook();

let emails = {};
async function parse(oldFile, newFile) {
  const currentWorkbook = await workbook.xlsx.readFile(oldFile);
  const worksheet = currentWorkbook.getWorksheet(1);
  const col = worksheet.getColumn("A");

  col.eachCell((cell, index) => {
    if (index < 2) {
      return;
    }

    const value = cell.value.toString().trim();
    if (!value) {
      return;
    }
    for (const email of value.split(",")) {
      emails[email] = email;
    }
  });

  const newWorkbook = new Excel.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("Emails");
  newWorksheet.columns = [{ header: "Email с сайта компании", key: "email" }];

  for (const email in emails) {
    await newWorksheet.addRow({ email });
  }

  await newWorkbook.xlsx.writeFile(newFile);
}

const argv = require("minimist")(process.argv.slice(2));

parse(argv.old || "addresses-list.xlsx", argv.new || "new.xlsx");
