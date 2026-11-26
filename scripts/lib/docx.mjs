// Minimal, dependency-free .docx (zip + XML) helpers. Works on Windows and macOS (Node 18+).
import zlib from 'node:zlib';
import fs from 'node:fs';

// ---------------- ZIP ----------------
const crcTable = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(buf) { let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }

export function readZip(file) {
  const buf = fs.readFileSync(file);
  let eocd = buf.length - 22;
  while (eocd >= 0 && buf.readUInt32LE(eocd) !== 0x06054b50) eocd--;
  if (eocd < 0) throw new Error('not a zip/docx file');
  const count = buf.readUInt16LE(eocd + 10); let off = buf.readUInt32LE(eocd + 16);
  const entries = [];
  for (let i = 0; i < count; i++) {
    const method = buf.readUInt16LE(off + 10), csize = buf.readUInt32LE(off + 20);
    const nlen = buf.readUInt16LE(off + 28), xlen = buf.readUInt16LE(off + 30), clen = buf.readUInt16LE(off + 32);
    const lho = buf.readUInt32LE(off + 42); const name = buf.toString('utf8', off + 46, off + 46 + nlen);
    const lnlen = buf.readUInt16LE(lho + 26), lxlen = buf.readUInt16LE(lho + 28);
    const raw = buf.subarray(lho + 30 + lnlen + lxlen, lho + 30 + lnlen + lxlen + csize);
    const data = method === 8 ? zlib.inflateRawSync(raw) : Buffer.from(raw);
    entries.push({ name, data });
    off += 46 + nlen + xlen + clen;
  }
  return entries;
}

export function writeZip(file, entries) {
  const parts = [], central = []; let off = 0;
  for (const e of entries) {
    const name = Buffer.from(e.name, 'utf8'); const crc = crc32(e.data);
    const store = /\.(png|jpe?g|gif|emf|wmf|bin)$/i.test(e.name);
    const comp = store ? e.data : zlib.deflateRawSync(e.data, { level: 6 });
    const h = Buffer.alloc(30); h.writeUInt32LE(0x04034b50, 0); h.writeUInt16LE(20, 4); h.writeUInt16LE(0x0800, 6);
    h.writeUInt16LE(store ? 0 : 8, 8); h.writeUInt32LE(0, 10); h.writeUInt32LE(crc, 14); h.writeUInt32LE(comp.length, 18);
    h.writeUInt32LE(e.data.length, 22); h.writeUInt16LE(name.length, 26); h.writeUInt16LE(0, 28);
    parts.push(h, name, comp);
    const c = Buffer.alloc(46); c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(0x0800, 8);
    c.writeUInt16LE(store ? 0 : 8, 10); c.writeUInt32LE(0, 12); c.writeUInt32LE(crc, 16); c.writeUInt32LE(comp.length, 20);
    c.writeUInt32LE(e.data.length, 24); c.writeUInt16LE(name.length, 28); c.writeUInt32LE(off, 42);
    central.push(c, name);
    off += 30 + name.length + comp.length;
  }
  const cd = Buffer.concat(central);
  const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(off, 16);
  fs.writeFileSync(file, Buffer.concat([...parts, cd, end]));
}

// ---------------- XML (lossless, raw text/attributes kept as-is) ----------------
export function parseXml(s) {
  const root = { type: 'root', children: [] }; const stack = [root]; let i = 0;
  const top = () => stack[stack.length - 1];
  while (i < s.length) {
    if (s[i] === '<') {
      if (s.startsWith('<?', i)) { const j = s.indexOf('?>', i); top().children.push({ type: 'raw', text: s.slice(i, j + 2) }); i = j + 2; continue; }
      if (s.startsWith('<!--', i)) { const j = s.indexOf('-->', i); top().children.push({ type: 'raw', text: s.slice(i, j + 3) }); i = j + 3; continue; }
      if (s.startsWith('<![CDATA[', i)) { const j = s.indexOf(']]>', i); top().children.push({ type: 'raw', text: s.slice(i, j + 3) }); i = j + 3; continue; }
      if (s[i + 1] === '/') { const j = s.indexOf('>', i); stack.pop(); i = j + 1; continue; }
      let j = i + 1, q = null;
      while (j < s.length) { const ch = s[j]; if (q) { if (ch === q) q = null; } else if (ch === '"' || ch === "'") q = ch; else if (ch === '>') break; j++; }
      let body = s.slice(i + 1, j); const selfClose = body.endsWith('/'); if (selfClose) body = body.slice(0, -1);
      const m = body.match(/^([^\s]+)/); const name = m[1]; const attrs = [];
      const re = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g; let a;
      while ((a = re.exec(body.slice(name.length)))) attrs.push([a[1], a[3] !== undefined ? a[3] : a[4]]);
      const el = { type: 'el', name, attrs, children: [], parent: top() };
      top().children.push(el); if (!selfClose) stack.push(el); i = j + 1;
    } else {
      const j = s.indexOf('<', i); const t = s.slice(i, j < 0 ? s.length : j);
      top().children.push({ type: 'text', text: t }); i = j < 0 ? s.length : j;
    }
  }
  return root;
}
export function serialize(n) {
  if (n.type === 'root') return n.children.map(serialize).join('');
  if (n.type === 'text' || n.type === 'raw') return n.text;
  const at = n.attrs.map(([k, v]) => ` ${k}="${v}"`).join('');
  if (!n.children.length) return `<${n.name}${at}/>`;
  return `<${n.name}${at}>${n.children.map(serialize).join('')}</${n.name}>`;
}
// helpers
export const els = (n, name) => (n.children || []).filter(c => c.type === 'el' && (!name || c.name === name));
export const el = (n, name) => els(n, name)[0];
export function all(n, name, out = [], stopAt = null) {
  for (const c of n.children || []) { if (c.type !== 'el') continue; if (c.name === name) out.push(c); if (stopAt && c.name === stopAt) continue; all(c, name, out, stopAt); }
  return out;
}
export const attr = (n, k) => { const a = n && n.attrs.find(x => x[0] === k); return a ? a[1] : undefined; };
export function setAttr(n, k, v) { const a = n.attrs.find(x => x[0] === k); if (a) a[1] = v; else n.attrs.push([k, v]); }
export function mk(name, attrs = [], children = []) { const e = { type: 'el', name, attrs, children, parent: null }; children.forEach(c => c.parent = e); return e; }
export function replaceWith(old, news) { const p = old.parent; const i = p.children.indexOf(old); p.children.splice(i, 1, ...news); news.forEach(n => n.parent = p); }
export function remove(n) { const p = n.parent; p.children.splice(p.children.indexOf(n), 1); }
export function insertBefore(ref, n) { const p = ref.parent; p.children.splice(p.children.indexOf(ref), 0, n); n.parent = p; }
export function textOf(n) { return all(n, 'w:t').map(t => t.children.map(c => c.text || '').join('')).join(''); }
export function ensure(parent, name, first = true) { let c = el(parent, name); if (!c) { c = mk(name); c.parent = parent; if (first) parent.children.unshift(c); else parent.children.push(c); } return c; }

export function loadDocx(file) { const entries = readZip(file); return { entries, get(name) { const e = entries.find(x => x.name === name); return e ? e.data.toString('utf8') : null; }, set(name, str) { const e = entries.find(x => x.name === name); const b = Buffer.from(str, 'utf8'); if (e) e.data = b; else entries.push({ name, data: b }); } }; }
export function saveDocx(doc, file) { writeZip(file, doc.entries); }
