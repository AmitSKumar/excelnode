const express = require("express");

const excelJs = require("exceljs");

const fs = require("fs");
const app = express();
const PORT = 4000;

app.get("/export", async (req, res) => {
  try {
    // create instance of exceljs and workbook
    let workbook = new excelJs.Workbook();
    // create sheet
    const sheet = workbook.addWorksheet("employee");
    workbook.created = new Date();
    //add columnaname to sheet
    sheet.columns = [
      { header: "ID", key: "id", width: 2 },
      { header: "First Name", key: "first_name", width: 25 },
      { header: "Last Name", key: "last_name", width: 25 },
      { header: "Email", key: "email", width: 35 },
      { header: "Gender", key: "gender", width: 25 },
    ];
    let object = JSON.parse(fs.readFileSync("data.json", "utf-8"));
    await object?.map((value) => {
      sheet.addRow({
        id: value.id,
        first_name: value.first_name,
        last_name: value.last_name,
        email: value.email,
        gender: value.gender,
      });
    });
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename= " + "employee.xlsx"
    );

    workbook.xlsx.write(res);
  } catch (error) {
    console.log(error);
  }
});

app.listen(PORT, () => {
  console.log("server listening to 4000");
});
