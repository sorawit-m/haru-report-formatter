import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); let boxes = 0, kn = 0, kl = 0, paras = 0, cs = 0, emptyRuns = 0;
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) { const t = m[0]; if (((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length !== 1) continue; boxes++; paras += (t.match(/<w:p[ >]/g) || []).length; kn += (t.match(/<w:keepNext\/>/g) || []).length; kl += (t.match(/<w:keepLines\/>/g) || []).length; cs += (t.match(/<w:cantSplit/g) || []).length; }
// runs of 2+ empty paragraphs in a row outside tables
const ems = x.match(/(?:<w:p [^>]*>(?:<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<\/w:pPr>)?<\/w:p>){2,}/g) || [];
console.log({ boxes, paras, keepNext: kn, keepLines: kl, cantSplit: cs, emptyParaRuns: ems.length });
