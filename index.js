import xlsx from "xlsx";
import axios from "axios";
import * as cheerio from "cheerio";
import * as puppeteer from "puppeteer";
import { SingleBar, Presets } from "cli-progress";
import { google } from "googleapis";
import dotenv from "dotenv";

dotenv.config();
let browser;

const { CX, API_KEY } = process.env;

const getFirstGoogleResult = async (query) => {
  try {
    const customsearch = google.customsearch("v1");
    const res = await customsearch.cse.list({
      cx: CX,
      q: query,
      key: API_KEY,
      num: 1,
    });

    if (res.data.items && res.data.items.length > 0) {
      return res.data.items[0].link;
    } else {
      return null;
    }
  } catch (error) {
    console.error(`Error fetching Google results: ${error.message}`);
    return null;
  }
};

const loadExcelFile = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return xlsx.utils.sheet_to_json(sheet);
};

const saveToExcel = (data, filePath) => {
  const worksheet = xlsx.utils.json_to_sheet(data);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Results");
  xlsx.writeFile(workbook, filePath);
};

const extractContentBySectionTitle = async ($, sectionTitle) => {
  const section = $(`.section-title:contains('${sectionTitle}')`).closest(
    ".detail-section"
  );

  if (!section.length) {
    return null;
  }

  return section
    .find(".section-body")
    .text()
    .trim()
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
};

const scrapeFromPage = async (url, type) => {
  try {
    if (type === "encyKorea") {
      const response = await axios.get(url, {
        headers: {
          "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        },
      });

      const $ = cheerio.load(response.data);

      return {
        summary: $("#cm_def .text-detail").text().trim() || "No content found.",
        contents:
          (await extractContentBySectionTitle($, "내용")) ||
          (await extractContentBySectionTitle($, "변천")) ||
          $("#cm_smry .text-detail").text().trim() ||
          (await extractContentBySectionTitle($, "형태")) ||
          "No content found.",
      };
    }

    if (type === "visitKorea") {
      const page = await browser.newPage();
      await page.goto(url, { waitUntil: "networkidle2" });
      const html = await page.content();

      const $ = cheerio.load(html);
      await page.close();

      return {
        summary:
          $("#contents .titTypeWrap").text().trim() || "No content found",
        contents:
          $("#detailGo .wrap_contView .area_txtView .inr_wrap .inr p")
            .text()
            .trim() || "No content found",
      };
    }
  } catch (error) {
    console.error(`Error fetching page: ${error.message}`);
    return `Error fetching page: ${error.message}`;
  }
};

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const processHeritageSites = async () => {
  let browser;
  const results = [];
  try {
    browser = await puppeteer.launch();
    const inputFilePath = "./modified_heritage_site_list.xlsx";
    const outputFilePath = "./heritage_google_results.xlsx";
    const heritageData = loadExcelFile(inputFilePath);

    const progressBar = new SingleBar({}, Presets.shades_classic);
    progressBar.start(heritageData.length, 0);

    for (const site of heritageData) {
      await delay(3000);

      const poiName = site["POI_NM"];
      const location = site["SIGNGU_NM"];

      const encyKorea = await scrapeFromPage(
        await getFirstGoogleResult(`${location} ${poiName}`),
        "encyKorea"
      );

      // const visitKorea = await scrapeFromPage(
      //   await getFirstGoogleResult(`${location} ${poiName}`, "visitKorea"),
      //   "visitKorea"
      // );

      progressBar.increment(); // 진행 바 업데이트

      results.push({
        Name: poiName,
        Location: location,
        encyKoreaSummary: encyKorea.summary,
        encyKoreaContents: encyKorea.contents,
      });
    }

    progressBar.stop();
  } catch (error) {
    console.error(error);
  } finally {
    const outputFilePath = "./heritage_google_results.xlsx";
    saveToExcel(results, outputFilePath);
    console.log(`Results saved to ${outputFilePath}`);
    if (browser) {
      await browser.close();
    }
  }
};

processHeritageSites();
