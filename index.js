const express = require("express");
const path = require("path");
const puppeteer = require("puppeteer");
const fs = require("fs");
const xlsx = require("xlsx");

const peserta = require("./json/data.json");

process.setMaxListeners(0);

const app = express();

const port = process.env.PORT || 5000;

app.use(express.json());
app.use("/template", express.static(path.join(__dirname, "/template")));

app.get("/excel", (req, res) => {
  const wb = xlsx.readFile("test.xlsx", { cellDates: true });
  const ws = wb.Sheets["response"];
  const data = xlsx.utils.sheet_to_json(ws);

  fs.writeFileSync(`${__dirname}/json/data.json`, JSON.stringify(data), {
    encoding: "utf-8",
  });

  res.json({
    status: "success",
    data,
  });
});

app.get("/generate", (req, res) => {
  peserta.forEach((data) => {
    (async function () {
      try {
        const browser = await puppeteer.launch();
        const page = await browser.newPage();

        const content = fs
          .readFileSync(`${__dirname}/template/index.html`, "utf-8")
          .replace("{{NAME}}", data["Nama"]);

        console.log(data["Nama"]);

        await page.setContent(content);
        await page.emulateMediaType("screen");
        await page.pdf({
          path: `pdfs/${data["Nama"]}.pdf`,
          width: "4.7in",
          height: "3.3781in",
          printBackground: true,
        });

        await browser.close();
      } catch (e) {
        console.log(e);
      }
    })();
  });
});

app.listen(port, () => {
  console.log(`listening to port ${port}...`);
});
