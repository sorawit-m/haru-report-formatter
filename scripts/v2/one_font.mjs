// ฟอนต์เดียวทั้งเล่ม (ยกเว้นปก และ Symbol ของหัวข้อย่อย) + โค้ดที่เล็กกว่า 14pt ขยายเป็น 14pt + เลขหน้าชิดขวา
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { coverEnd } from '../lib/wordxml.mjs';
const [f, FONT = 'TH SarabunPSK', MINCODE = '28'] = process.argv.slice(2);
const d = loadDocx(f); const R = `<w:rFonts w:ascii="${FONT}" w:eastAsia="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>`;
const CODE = /w:ascii="(?:Consolas|Courier New|Courier|Menlo|Monaco|Lucida Console|Source Code Pro)"/;
const cnt = { fonts: 0, codeSize: 0, pageRight: 0 };
const fixFonts = s => s.replace(/<w:rFonts [^>]*\/>/g, m => { if (/w:ascii="(?:Symbol|Wingdings)/.test(m)) return m; cnt.fonts++; return R; });
// code runs: bump size before swapping the font away (so they can still be recognised)
const bumpCode = s => s.replace(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g, r => {
  if (!CODE.test(r)) return r;
  return r.replace(/<w:sz w:val="(\d+)"\/>/, (m, v) => +v < +MINCODE ? (cnt.codeSize++, `<w:sz w:val="${MINCODE}"/>`) : m).replace(/<w:szCs w:val="(\d+)"\/>/, (m, v) => +v < +MINCODE ? `<w:szCs w:val="${MINCODE}"/>` : m);
});
for (const e of d.entries) {
  const n = e.name;
  if (n === 'word/document.xml') {
    let x = d.get(n); const b = x.indexOf('<w:body>'); const ce = coverEnd(x);
    d.set(n, x.slice(0, ce) + fixFonts(bumpCode(x.slice(ce))));
  } else if (/^word\/(styles|numbering|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n)) {
    let x = d.get(n);
    if (n === 'word/styles.xml') x = x.replace(/<w:style\b[\s\S]*?<\/w:style>/g, s => CODE.test(s) ? s.replace(/<w:sz w:val="(\d+)"\/>/, (m, v) => +v < +MINCODE ? `<w:sz w:val="${MINCODE}"/>` : m).replace(/<w:szCs w:val="(\d+)"\/>/, (m, v) => +v < +MINCODE ? `<w:szCs w:val="${MINCODE}"/>` : m) : s);
    x = fixFonts(bumpCode(x));
    if (/^word\/footer\d*\.xml$/.test(n)) x = x.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
      if (!/PAGE/.test(p)) return p; cnt.pageRight++;
      if (/<w:jc w:val="[^"]+"\/>/.test(p)) return p.replace(/<w:jc w:val="[^"]+"\/>/, '<w:jc w:val="right"/>');
      return /<w:pPr>/.test(p) ? p.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (a, i) => { const k = i.search(/<w:rPr\b/); return '<w:pPr>' + (k < 0 ? i + '<w:jc w:val="right"/>' : i.slice(0, k) + '<w:jc w:val="right"/>' + i.slice(k)) + '</w:pPr>'; }) : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:jc w:val="right"/></w:pPr>');
    });
    d.set(n, x);
  } else if (n === 'word/theme/theme1.xml') {
    d.set(n, d.get(n).replace(/(<a:(?:majorFont|minorFont)>[\s\S]*?<a:latin typeface=")[^"]*(")/g, `$1${FONT}$2`).replace(/(<a:(?:majorFont|minorFont)>[\s\S]*?<a:cs typeface=")[^"]*(")/g, `$1${FONT}$2`).replace(/<a:font script="Thai" typeface="[^"]*"\/>/g, `<a:font script="Thai" typeface="${FONT}"/>`));
  }
}
saveDocx(d, f); console.log(JSON.stringify(cnt));
