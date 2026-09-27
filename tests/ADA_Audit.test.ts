import { test, expect } from '@playwright/test';
import * as XLSX from 'xlsx';
import lighthouse from 'lighthouse';
import * as chromeLauncher from 'chrome-launcher';
import * as fs from 'fs';
import * as path from 'path';

async function runLighthouse(url: string, reportName: string) {
  const chrome = await chromeLauncher.launch({ chromeFlags: ['--headless'] });
  try {
    const options: Parameters<typeof lighthouse>[1] = {
      logLevel: 'info',
      output: 'html',
      onlyCategories: ['performance', 'accessibility', 'best-practices'],
      port: chrome.port
    };

    const runnerResult = await lighthouse(url, options);

    if (!runnerResult) {
      throw new Error(`Lighthouse did not return a result for ${url}`);
    }

    const reportHtml = runnerResult.report;
    const reportPath = path.join('lighthouse-reports', `${reportName}.html`);

    fs.writeFileSync(reportPath, Array.isArray(reportHtml) ? reportHtml.join('') : reportHtml);
    console.log(`✔ Lighthouse report saved: ${reportPath}`);
  } finally {
    await chrome.kill();
  }
}

test('Generate Lighthouse reports from Excel URLs', async ({ browser }) => {
  // Load Excel file
  const workbook = XLSX.readFile('urls.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);

  // Ensure output folder exists
  if (!fs.existsSync('lighthouse-reports')) {
    fs.mkdirSync('lighthouse-reports');
  }

  for (const row of rows) {
    const url = row.URL || row.url || row.link;

    if (typeof url !== 'string' || !url) {
      console.log('Skipping row without URL');
      continue;
    }

    console.log(`🔎 Checking: ${url}`);

    const context = await browser.newContext();
    const page = await context.newPage();

    try {
      await page.goto(url, { timeout: 30000 });
      await expect(page).toHaveURL(url);
      console.log(`✔ Loaded: ${url}`);

      const reportName = url.replace(/[^a-z0-9]/gi, '_').toLowerCase();
      await runLighthouse(url, reportName);

    } catch (err) {
      console.error(`❌ Failed to load ${url}:`, err);
    }

    await context.close();
  }
});
