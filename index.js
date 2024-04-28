const axios = require("axios");
const ExcelJS = require("exceljs");
const moment = require("moment");
const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");å
require("dotenv").config();

const GITLAB_HOST = process.env.GITLAB_HOST;
const GITLAB_PROJECT_ID = process.env.GITLAB_PROJECT_ID;
const PRIVATE_TOKEN = process.env.PRIVATE_TOKEN;
const TAIGA_URL = process.env.TAIGA_URL;
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
  const envConfig = dotenv.parse(fs.readFileSync(".env"));
  envConfig.TAIGA_AUTH_TOKEN = authToken;

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
      url: `${GITLAB_HOST}api/v4/projects/${GITLAB_PROJECT_ID}/merge_requests`,
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
    { header: "Дата", key: "merged_at" },
    { header: "MR создан", key: "created_at" },
    { header: "Исполнитель", key: "executor" },
    { header: "Ветка", key: "branch" },
    { header: "Тип задачи", key: "task_type" },
    { header: "Ссылка на задачу", key: "task_url" },
  ];

  const filteredData = response.data.filter((item) => {
    const updatedAt = moment(item.merged_at);
    return (
      updatedAt.isSameOrAfter(startDate) && updatedAt.isSameOrBefore(endDate)
    );
  });

  // Группировка данных по исполнителю
  const groupedData = filteredData.reduce((groups, item) => {
    const key = item.author.name;
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(item);
    return groups;
  }, {});

  Object.keys(groupedData).forEach(async (executor) => {
    // Сортировка данных в группе по дате в убывающем порядке
    const sortedData = groupedData[executor].sort(
      (a, b) => new Date(b.merged_at) - new Date(a.merged_at)
    );

    // Добавление данных в Excel
    for (const item of sortedData) {
      const link = `${TAIGA_URL}${formatBranchNameToTaigaUrlEnd(item.source_branch)}`;
      worksheet.addRow({
        merged_at: moment(item.merged_at).format("DD.MM.YYYY HH:mm:ss"),
        created_at: moment(item.created_at).format("DD.MM.YYYY HH:mm:ss"),
        executor: item.author.name,
        branch: item.source_branch,
        task_type: getTaskType(item.source_branch),
        task_url: `=HYPERLINK("${link}")`,
      });
    }

    // Добавление пустой строки между группами
    worksheet.addRow({});
  });

  try {
    await workbook.xlsx.writeFile("output.xlsx");
  } catch (error) {
    logError(error);
  }
}

async function getBranchCreationDate(branchName) {
  try {
    const response = await axios.get(
      `${GITLAB_HOST}api/v4/projects/${GITLAB_PROJECT_ID}/repository/commits?ref_name=${branchName}`
    );
    const commits = response.data;
    const firstCommit = commits[commits.length - 1];
    return firstCommit.created_at;
  } catch (error) {
    console.error(error);
  }
}

function getTaskType(branch) {
  if (branch.includes("feat")) {
    return "Feature";
  } else if (branch.includes("bug")) {
    return "Bug";
  } else if (branch.includes("hotfix")) {
    return "Hotfix";
  } else if (branch.includes("us")) {
    return "Userstory";
  } else {
    return "Unknown";
  }
}

function formatBranchNameToTaigaUrlEnd(branchName) {
  const branch = branchName.replace("-", "/");
  const [sourceTaskType, taskNumber] = branch.split("/");
  let taskType = "";
  if (sourceTaskType.includes("bug") || branch.includes("hotfix")) {
    taskType = "issue";
  } else if (branch.includes("feat")) {
    taskType = "task";
  } else if (branch.includes("us")) {
    taskType = "us";
  } else {
    taskType = "";
  }
  return `${taskType}/${taskNumber}`;
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
