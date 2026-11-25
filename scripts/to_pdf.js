// 把报告网页转成 A4 PDF（用户要 PDF 时用）：node scripts/to_pdf.js <报告.html 的绝对路径> <输出.pdf>
// 需要 playwright 和 Chromium；"依据""为什么"会全部展开；按 0.68 缩放，用电脑版排版。
const { chromium } = require('playwright');
(async () => {
  const [src, out] = process.argv.slice(2);
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1100, height: 1400 }, colorScheme: 'light' });
  await p.goto('file://' + src, { waitUntil: 'networkidle' });
  await p.evaluate(() => {
    document.querySelectorAll('details').forEach(d => d.open = true);  // 依据、为什么全部展开
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
