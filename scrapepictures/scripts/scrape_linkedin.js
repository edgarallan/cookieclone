#!/usr/bin/env node
const fs = require('fs');
const puppeteer = require('puppeteer');

(async () => {
  const DIRECTORY_URL = 'https://developers.google.com/community/experts/directory?hl=it&specialization=android';
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });
  const page = await browser.newPage();
  await page.goto(DIRECTORY_URL, { waitUntil: 'networkidle2' });
  // scroll until all entries are loaded
  let prevCount = 0;
  while (true) {
    const count = await page.$$eval(
      'a.profile-card-button.button.button-white',
      els => els.length
    );
    if (count === prevCount) break;
    prevCount = count;
    await page.evaluate(() => window.scrollBy(0, window.innerHeight));
    await page.waitFor(1000);
  }
  const profileUrls = await page.$$eval(
    'a.profile-card-button.button.button-white',
    els => els.map(a => a.href + (a.href.includes('?') ? '' : '?hl=it'))
  );
  const results = [];
  for (const url of profileUrls) {
    try {
      const detailPage = await browser.newPage();
      await detailPage.goto(url, { waitUntil: 'networkidle2' });
      // give time for client-side render
      await detailPage.waitFor(2000);
      const finalUrl = detailPage.url();
      const slugMatch = finalUrl.match(/\/experts\/people\/([^?]+)/);
      const slug = slugMatch ? slugMatch[1] : '';
      const linkedinLinks = await detailPage.$$eval(
        'a[href^="https://www.linkedin.com"]',
        els => els.map(e => e.href)
      );
      const linkedin = linkedinLinks.length ? linkedinLinks[0] : '';
      results.push({ slug, linkedin });
      await detailPage.close();
    } catch (e) {
      console.error('Error fetching', url, e.message);
    }
  }
  await browser.close();
  // merge into CSV
  const csv = fs.readFileSync('speakers.csv', 'utf8').trim().split(/\r?\n/);
  const [header, ...rows] = csv;
  const out = [header + ',linkedin_url'];
  const map = new Map(results.map(r => [r.slug, r.linkedin]));
  rows.forEach(line => {
    const parts = line.split(',');
    const slug = parts
      .slice(0, 2)
      .map(s => s.toLowerCase().replace(/\s+/g, '-'))
      .join('-');
    parts.push(map.get(slug) || '');
    out.push(parts.join(','));
  });
  fs.writeFileSync('speakers.csv', out.join('\n'));
  console.log('speakers.csv updated with LinkedIn URLs.');
})();
