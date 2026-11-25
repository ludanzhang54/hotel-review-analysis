// Turn the report page into an A4 PDF (when the user wants a PDF): node scripts/to_pdf.js <absolute path of report.html> <output.pdf>
// Needs playwright and Chromium; "Evidence" and "Why" are all expanded; scaled to 0.68, desktop layout.
const { chromium } = require('playwright');
(async () => {
  const [src, out] = process.argv.slice(2);
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1100, height: 1400 }, colorScheme: 'light' });
  await p.goto('file://' + src, { waitUntil: 'networkidle' });
  await p.evaluate(() => {
    document.querySelectorAll('details').forEach(d => d.open = true);  // expand Evidence and Why
    const s = document.createElement('style');
    s.textContent = `@page { size: A4; margin: 12mm 10mm; }
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .chart, .actions > li, .glance section, .donuts section, .kpi, .watch, .stats, .fhead { break-inside: avoid; } .cover { break-after: avoid; }
      h2 { break-after: avoid; } summary { list-style: none; } summary::-webkit-details-marker { display: none; }`;
    document.head.appendChild(s);
  });
  await p.pdf({ path: out, format: 'A4', printBackground: true, scale: 0.68, margin: { top: '10mm', bottom: '10mm', left: '10mm', right: '10mm' } });
  await b.close();
})();
