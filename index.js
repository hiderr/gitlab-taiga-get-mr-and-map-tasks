require("dotenv").config();
const axios = require("axios");
const ExcelJS = require("exceljs");
const moment = require("moment");

const PRIVATE_TOKEN = process.env.PRIVATE_TOKEN;
const TAIGA_USERNAME = process.env.TAIGA_USERNAME;
const TAIGA_PASSWORD = process.env.TAIGA_PASSWORD;

const [startDateInput, endDateInput] = process.argv.slice(2);

const startDate = moment(startDateInput, "DD.MM.YYYY");
const endDate = moment(endDateInput, "DD.MM.YYYY");

if (!startDate.isValid() || !endDate.isValid()) {
  console.error(
    "Даты не заданы или заданы неверно. Убедитесь, что вы передаете даты в формате DD.MM.YYYY"
  );
  process.exit(1);
}

const data = {
  type: "normal",
  username: TAIGA_USERNAME,
  password: TAIGA_PASSWORD,
};

axios({
  method: "post",
  url: "https://api.taiga.io/api/v1/auth",
  data: data,
})
  .then((response) => {
    const authToken = response.data.auth_token;

    if (!authToken) {
      console.error("Error: Incorrect username and/or password supplied");
      process.exit(1);
    } else {
      console.log("auth_token is", authToken);
      // Proceed to use API calls as desired
    }
  })
  .catch((error) => {
    console.error(error);
  });

axios({
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
})
  .then((response) => {
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

    return workbook.xlsx.writeFile("output.xlsx");
  })
  .catch((error) => {
    console.error(error);
  });
