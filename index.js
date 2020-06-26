const express = require("express");
const path = require("path");
const puppeteer = require("puppeteer");
const fs = require("fs");
const xlsx = require("xlsx");

const peserta = require("./json/data.json");

// process.setMaxListeners(0);

const app = express();

const port = process.env.PORT || 5000;

app.use(express.json());
app.use("/template", express.static(path.join(__dirname, "/template")));

// di browser hit localhost:5000/excel buat bikin file json dari excel test.xlsx
// nama-nama yang dari excel bakal keluar di folder /json/data.json

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

// Fungsi generate pdf, template sertifnya pake html css bisa liat di folder /template/index.html
// Buat template sertifnya pake jpg/png file design yg udh jadi, dgn kolom namanya di kosongin, nanti programnya bakal ngisi kekosongan pake Nama dari html css
// Pengaturan posisi Namanya liat di /template/index.html di bagian

const generatePdf = async (data) => {
  try {
    const browser = await puppeteer.launch({
      executablePath:
        "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe", // ini pake path google chrome tergantung nginstallnya dimana, kalo pas install node_modules nya ke install chromium, options ini di hapus aja.
    });

    const content = fs
      .readFileSync(`${__dirname}/template/index.html`, "utf-8")
      .replace("{{NAME}}", data["Nama"]);

    const page = await browser.newPage();

    await page.setContent(content);
    await page.emulateMediaType("screen");
    await page.goto("data:text/html," + content, {
      waitUntil: "networkidle2",
    });
    await page.pdf({
      path: `pdfs/${data["Nama"]}.pdf`,
      width: "17.78in", // ini lebar dari sertifnya, rekomen unitnya inchi liat lebarnya bisa pake AI / CDR dari design yg dikasih designer
      height: "12.56in", // ini heightnya sama kayak lebar rinciannya
      printBackground: true,
    });

    await browser.close();
    console.log(data["Nama"]);
  } catch (e) {
    console.log(e);
  }
};

// hit localhost:5000/generate untuk bikin pdf sertifnya dan enjoyy!!!

app.get("/generate", (req, res) => {
  let count = 1;
  (async () => {
    for (let i = 0; i < peserta.length; i++) {
      try {
        await generatePdf(peserta[i]);
        console.log(`Creating Progress ${(count / peserta.length) * 100}%`);
        count++;
      } catch (e) {
        console.log(e);
      }
    }
  })().then(() => {
    res.json({
      status: "success",
    });
  });
});

app.listen(port, () => {
  console.log(`listening to port ${port}...`);
});
