const axios = require("axios");
const ExcelJS = require("exceljs");
const moment = require("moment");
const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");
require('dotenv').config();

const PRIVATE_TOKEN = process.env.PRIVATE_TOKEN;
const TAIGA_USERNAME = process.env.TAIGA_USERNAME;
const TAIGA_PASSWORD = process.env.TAIGA_PASSWORD;
const TAIGA_AUTH_TOKEN = process.env.TAIGA_AUTH_TOKEN;

const [startDateInput, endDateInput] = process.argv.slice(2);

const startDate = moment(startDateInput, "DD.MM.YYYY");
const endDate = moment(endDateInput, "DD.MM.YYYY");

if (!startDate.isValid() || !endDate.isValid()) {
  logErrorAndExit(
    "Даты не заданы или заданы неверно. Убедитесь, что вы передаете даты в формате DD.MM.YYYY"
  );
}

const data = {
  type: "normal",
  username: TAIGA_USERNAME,
  password: TAIGA_PASSWORD,
};

main();

async function main() {
  try {
    if (!TAIGA_AUTH_TOKEN) {
      const authToken = await authenticateAndGetTokenFromTaiga(data);
      updateEnvFileWithAuthTokenFromTaiga(authToken);
    }
    await getMergeRequestsAndWriteToExcel();
  } catch (error) {
    logError(error);
  }
}

async function authenticateAndGetTokenFromTaiga(data) {
  try {
    const response = await axios({
      method: "post",
      url: "https://api.taiga.io/api/v1/auth",
      data: data,
    });

    const authToken = response.data.auth_token;

    if (!authToken) {
      logErrorAndExit("Error: Incorrect username and/or password supplied");
    }

    console.log("auth_token is", authToken);
    return authToken;
  } catch (error) {
    logError(error);
  }
}

function updateEnvFileWithAuthTokenFromTaiga(authToken) {
  const envConfig = dotenv.parse(fs.readFileSync(".env")); // read .env file
  envConfig.TAIGA_AUTH_TOKEN = authToken; // update auth token

  let envContent = "";
  for (let key in envConfig) {
    if (envConfig.hasOwnProperty(key)) {
      envContent += `${key}=${envConfig[key]}\n`;
    }
  }

  fs.writeFile(path.join(__dirname, ".env"), envContent, (err) => {
    if (err) {
      logError("Failed to write to .env: " + err);
    } else {
      console.log("Auth token written to .env");
    }
  });
}

async function getMergeRequestsAndWriteToExcel() {
  try {
    const response = await axios({
      method: "get",
      url: "https://gitlab.byhelp.ru/api/v4/projects/6/merge_requests",
      headers: {
        "PRIVATE-TOKEN": PRIVATE_TOKEN,
      },
      params: {
        state: "merged",
        per_page: 100,
        created_after: startDate.format("YYYY-MM-DDTHH:mm:ssZ"),
        created_before: endDate.format("YYYY-MM-DDTHH:mm:ssZ"),
      },
    });

    await writeMergeRequestsToExcel(response);
  } catch (error) {
    logError(error);
  }
}

async function writeMergeRequestsToExcel(response) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Merge Requests");

  worksheet.columns = [
    { header: "Дата", key: "date" },
    { header: "Исполнитель", key: "executor" },
    { header: "Ветка", key: "branch" },
  ];

  const filteredData = response.data.filter((item) => {
    const updatedAt = moment(item.merged_at);
    return (
      updatedAt.isSameOrAfter(startDate) && updatedAt.isSameOrBefore(endDate)
    );
  });

  filteredData.forEach((item) => {
    worksheet.addRow({
      date: moment(item.merged_at).format("DD.MM.YYYY HH:mm:ss"),
      executor: item.author.name,
      branch: item.source_branch,
    });
  });

  try {
    await workbook.xlsx.writeFile("output.xlsx");
  } catch (error) {
    logError(error);
  }
}

async function logError(error) {
  const errorMsg = error.toString();
  console.error(errorMsg);
  try {
    await fs.promises.appendFile(
      "error.log",
      `${new Date().toISOString()} - ${errorMsg}\n`
    );
  } catch (err) {
    console.error("Failed to write to error.log:", err);
  }
}

function logErrorAndExit(message) {
  logError(message);
  process.exit(1);
}
