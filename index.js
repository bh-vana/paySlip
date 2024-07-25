const multer = require("multer");
const xlsx = require("xlsx");
const express = require("express");
const bodyParser = require("body-parser");
const archiver = require("archiver");
const htmlDocx = require("html-docx-js");

const app = express();
const port = 3000;
let fileName;
let data;
let foundUser;

const ds = multer.diskStorage({
  destination: "import/",
  filename: (req, file, cb) => {
    fileName = file.originalname;
    cb(null, file.originalname);
  },
});

const a = multer({
  storage: ds,
});

function convertIntoJSON() {
  const workbook = xlsx.readFile(`import/${fileName}`);

  // Assume data is in the first sheet
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert worksheet to JSON
  const jsonData = xlsx.utils.sheet_to_json(worksheet);

  // Write JSON data

  return JSON.stringify(jsonData);
}
app.use(express.json());
app.use(express.static("public"));
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.render("index.ejs");
});

app.post("/upload", a.single("excelFile"), (req, res) => {
  if (!req.file) {
    return res.status(400).send("No file uploaded.");
  }

  let jsons = convertIntoJSON();
  data = JSON.parse(jsons);
  const {
    ecode,
    ename,
    designation,
    doj,
    estatus,
    stmtMonth,
    basicPay,
    houseRent,
    cityCompensatory,
    travel,
    food,
    performance,
    profTax,
    incomeTax,
    pF,
    ESI,
    leaves,
    others,
    grossPay,
    deductions,
    netPay,
  } = data[0];
  if (
    !ecode ||
    !ename ||
    !designation ||
    !doj ||
    !estatus ||
    !stmtMonth ||
    !basicPay ||
    !houseRent ||
    !cityCompensatory ||
    !travel ||
    !food ||
    !performance ||
    !profTax ||
    !incomeTax ||
    !pF ||
    !ESI ||
    !leaves ||
    !others ||
    !grossPay ||
    !deductions ||
    !netPay
  ) {
    const reasonObj = {
      ecode,
      ename,
      designation,
      doj,
      estatus,
      stmtMonth,
      basicPay,
      houseRent,
      cityCompensatory,
      travel,
      food,
      performance,
      profTax,
      incomeTax,
      pF,
      ESI,
      leaves,
      others,
      grossPay,
      deductions,
      netPay,
    };
    res.render("error.ejs", reasonObj);
  } else {
    res.render("search.ejs");
  }
});

function convertIntoDate(dateToBeConverted) {
  // Serial number from Excel
  var serialNumber = dateToBeConverted;

  // Number of milliseconds in a day
  var millisecondsPerDay = 24 * 60 * 60 * 1000;

  // Base date in milliseconds in Excel (December 30, 1899)
  var baseDateMilliseconds = new Date(1899, 11, 30).getTime();

  // Calculate the milliseconds offset for the serial number
  var serialNumberOffsetMilliseconds = serialNumber * millisecondsPerDay;

  // Calculate the resulting date in milliseconds
  var resultingDateMilliseconds =
    baseDateMilliseconds + serialNumberOffsetMilliseconds;

  // Create a new Date object with the resulting date
  var resultingDate = new Date(resultingDateMilliseconds);
  var cdy = resultingDate.getFullYear();
  var cdm = resultingDate.getMonth() + 1;
  var cdd = resultingDate.getDate();
  if (cdm < 10) {
    cdm = "0" + cdm;
  }
  if (cdd < 10) {
    cdd = "0" + cdd;
  }
  var convertedDate = cdd + "-" + cdm + "-" + cdy;
  return convertedDate;
}

app.post("/search-emp", (req, res) => {
  const { ecode } = req.body;
  const singleUser = data.find((emp) => emp.ecode == ecode);
  foundUser = { ...singleUser };
  foundUser.doj = convertIntoDate(foundUser.doj);
  if (singleUser) res.render("search.ejs", foundUser);
  else res.render("search.ejs", { notfound: "No Employees Found !!!" });
});

app.get("/download", async (req, res) => {
  try {
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", "attachment; ");

    const archive = archiver("zip", {
      zlib: { level: 9 },
    });

    archive.on("error", function (err) {
      throw err;
    });

    // Pipe the archive data to the response
    archive.pipe(res);

    const {
      ecode,
      ename,
      designation,
      netPay,
      deductions,
      grossPay,
      others,
      leaves,
      ESI,
      pF,
      incomeTax,
      profTax,
      performance,
      food,
      travel,
      cityCompensatory,
      houseRent,
      basicPay,
      stmtMonth,
      estatus,
      doj,
    } = foundUser;
    const htmlContent = `<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Salary Statement</title>
            <style>
                *{
                    margin: 0;
                    padding: 0;
                }
                body {
                    font-family: Arial, sans-serif;
                    margin: 20px;
                }
                .container {
                    width: 100%;
                    max-width: 800px;
                    margin: 0 auto;
                }
                .header {
                    text-align: center;
                }
        
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                table, th, td {
                    border: 1px solid #000;
                }
                th, td {
                    padding: 2px;
                    text-align: left;
                }
                .signature {
                    margin-top: 20px;
                }
                .th2{
                    text-align: center;
                }
             td{
                    width: 100px;
                } 
                th,td{
                    font-size:small;
                }
                #tb1{
                    margin-bottom: 10px;
                }
            </style>
        </head>
        <body>
        <div class="container">
        <div class="header">
            <h3>SYMBIOSYS TECHNOLOGIES</h3>
            <p style="margin-bottom: 10px;">Plot No 1&2,Hill no-2,IT Park,<br>
            Rushikonda, Visakhapatnam-45<br>
            Ph: 2550369, 2595657</p>
            <u><h4>SALARY STATEMENT FOR THE MONTH OF JANUARY 2024</h4></u>
        </div>
        <div id="tb1">
            <table>
                <tr>
                    <th>Employee Code</th>
                    <td>${ecode}</td>
                    <th>Date of Joining</th>
                    <td>${doj}</td>
                </tr>
                <tr>
                    <th>Employee Name</th>
                    <td>${ename}</td>
                    <th>Employment Status</th>
                    <td>${estatus}</td>
                </tr>
                <tr>
                    <th>Designation</th>
                    <td>${designation}</td>
                    <th>Statement for the month</th>
                    <td>${stmtMonth}</td>
                </tr>
            </table>
        </div>

        <div id="tb2">
            <table>
                <thead>
                    <tr>
                        <th colspan="2" class="th2">Classified Income</th>
                        <th colspan="2" class="th2">Deductions</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>Basic Pay</td>
                        <td>${basicPay}</td>
                        <td>Professional Tax</td>
                        <td>${profTax}</td>
                    </tr>
                    <tr>
                        <td>House Rent Allowance</td>
                        <td>${houseRent}</td>
                        <td>Income Tax</td>
                        <td>${incomeTax}</td>
                    </tr>
                    <tr>
                        <td>City Compensatory Allowance</td>
                        <td>${cityCompensatory}</td>
                        <td>Provident Fund</td>
                        <td>${pF}</td>
                    </tr>
                    <tr>
                        <td>Travel Allowance</td>
                        <td>${travel}</td>
                        <td>ESI</td>
                        <td>${ESI}</td>
                    </tr>
                    <tr>
                        <td>Food Allowance</td>
                        <td>${food}</td>
                        <td>Leaves - Loss of Pay</td> 
                        <td>${leaves}</td>
                    </tr>
                    <tr>
                        <td>Performance Incentives</td>
                        <td>${performance}</td>
                        <td>Others</td>
                        <td>${others}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    
        <table style="margin-top: 10px;">
            <tr>
                <th class="th2">GROSS PAY</th>
                <th class="th2">DEDUCTIONS</th>
                <th class="th2">NET PAY</th>
            </tr>
            <tr>
                <td class="th2">${grossPay}</td>
                <td class="th2">${deductions}</td>
                <td class="th2">${netPay}</td>
            </tr>
        </table>

        <div class="signature">
            <p>AUTHORISED SIGNATORY</p>
            <p>Durgaaprasadh,<br>H.R Executive</p>
        </div>

        <div class="footer">
            <p>We request you to verify employment details with our office on email: hr@symbiosystech.com (+91-0891-2550369)</p>
        </div>
    </div>
        </body>
        </html>
        `;

    const docxContent = htmlDocx.asBlob(htmlContent);
    const buffer = Buffer.from(await docxContent.arrayBuffer()); // Convert Blob to Buffer
    archive.append(buffer, { name: `document1.docx` });

    await archive.finalize();
  } catch (error) {
    console.error("Error generating documents:", error);
    res.status(500).send("Internal Server Error");
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});



//check at readme.md