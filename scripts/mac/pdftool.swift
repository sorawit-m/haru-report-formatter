// pdftool.swift — เครื่องมือ PDF บน Mac (แทน scripts/windows/pdf-to-png.ps1 + v2/pdf_check.ps1) ใช้ PDFKit ที่มากับ macOS
// ใช้ผ่าน scripts/mac/pdftool.sh (คอมไพล์ครั้งแรกแล้วเก็บไว้ใน ~/.cache/haru)
//   png   <pdf> <outdir> [width=400] [from=1] [to=all]   -> p001.png …
//   check <pdf> <outdir> [ขอบขวา ซม.]                   -> ตรวจทุกหน้า (รูปแบบผลเหมือน pdf_check.ps1)
//   text  <pdf>                                          -> JSON: ข้อความแต่ละหน้า ["หน้า 1", "หน้า 2", …]
import Foundation
import PDFKit
import AppKit

func fail(_ m: String) -> Never { FileHandle.standardError.write((m + "\n").data(using: .utf8)!); exit(1) }
let a = CommandLine.arguments
guard a.count >= 3 else { fail("usage: pdftool png|check|text <pdf> …") }
guard let doc = PDFDocument(url: URL(fileURLWithPath: a[2])) else { fail("open pdf failed: \(a[2])") }

struct Bitmap { let w: Int; let h: Int; let px: [UInt8]
  @inline(__always) func rgb(_ x: Int, _ y: Int) -> (Int, Int, Int) { let i = (y * w + x) * 4; return (Int(px[i]), Int(px[i + 1]), Int(px[i + 2])) } }

func render(_ i: Int, width: Int) -> Bitmap {
  let p = doc.page(at: i)!; let r = p.bounds(for: .mediaBox)
  let rot = p.rotation % 180 != 0
  let pw = rot ? r.height : r.width, ph = rot ? r.width : r.height
  let w = width, h = max(1, Int((CGFloat(width) * ph / pw).rounded()))
  var px = [UInt8](repeating: 255, count: w * h * 4)
  px.withUnsafeMutableBytes { buf in
    let ctx = CGContext(data: buf.baseAddress, width: w, height: h, bitsPerComponent: 8, bytesPerRow: w * 4,
                        space: CGColorSpaceCreateDeviceRGB(), bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1)); ctx.fill(CGRect(x: 0, y: 0, width: w, height: h))
    ctx.scaleBy(x: CGFloat(w) / pw, y: CGFloat(h) / ph)
    p.draw(with: .mediaBox, to: ctx)
  }
  // CGContext origin = bottom-left; buffer row 0 = top row already (bitmap memory is top-down)
  return Bitmap(w: w, h: h, px: px)
}

func savePNG(_ b: Bitmap, _ path: String) {
  var px = b.px
  let rep = px.withUnsafeMutableBytes { buf -> NSBitmapImageRep in
    var plane: UnsafeMutablePointer<UInt8>? = buf.baseAddress!.assumingMemoryBound(to: UInt8.self)
    return NSBitmapImageRep(bitmapDataPlanes: &plane, pixelsWide: b.w, pixelsHigh: b.h, bitsPerSample: 8, samplesPerPixel: 4,
                            hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: b.w * 4, bitsPerPixel: 32)!.copy() as! NSBitmapImageRep
  }
  try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: path))
}

let n = doc.pageCount
switch a[1] {
case "png":
  let out = a[3]; let width = a.count > 4 ? Int(a[4]) ?? 400 : 400
  let from = a.count > 5 ? max(1, Int(a[5]) ?? 1) : 1; let to = a.count > 6 ? min(n, Int(a[6]) ?? n) : n
  try? FileManager.default.createDirectory(atPath: out, withIntermediateDirectories: true)
  print("pages=\(n)")
  if from <= to { for i in (from - 1)..<to { savePNG(render(i, width: width), String(format: "%@/p%03d.png", out, i + 1)) } }

case "text":
  var pages: [String] = []
  for i in 0..<n { pages.append(doc.page(at: i)?.string ?? "") }
  let d = try! JSONSerialization.data(withJSONObject: pages, options: [])
  FileHandle.standardOutput.write(d)

case "check":
  // ตรรกะเดียวกับ v2/pdf_check.ps1 (ภาพกว้าง 400 px) + 2 อย่างที่กันผลหลอก:
  //  - แถวที่มีหมึกตรงกันเกือบทุกหน้า (เส้น/ข้อความหัว-ท้ายกระดาษ) ไม่นับเป็นเนื้อหา
  //  - ขอบขวาใช้ค่าจริงของเล่ม: check <pdf> <outdir> [ขอบขวา ซม. = 1.6]
  let out = a[3]; let rmCm = a.count > 4 ? Double(a[4]) ?? 1.6 : 1.6
  try? FileManager.default.removeItem(atPath: out)
  try? FileManager.default.createDirectory(atPath: out, withIntermediateDirectories: true)
  func inkRows(_ b: Bitmap) -> [Bool] {
    var rows: [Bool] = []
    for y in Int(Double(b.h) * 0.11)..<Int(Double(b.h) * 0.90) {
      var ink = false; var x = 30
      while x < b.w - 30 { let (r, g, bl) = b.rgb(x, y); if r < 251 || g < 251 || bl < 251 { ink = true; break }; x += 2 }
      rows.append(ink)
    }
    return rows
  }
  // ลายเซ็นของแถว: แถวที่เหมือนกันทุกพิกเซลในหลายหน้า = หัว/ท้ายกระดาษ (ข้อความเนื้อหาไม่ซ้ำกันแบบนี้)
  func rowSig(_ b: Bitmap, _ y: Int) -> UInt64 {
    // ไม่นับ 20% ขวาสุด (เลขหน้าชิดขวาเปลี่ยนทุกหน้า)
    var h: UInt64 = 14695981039346656037; var x = 0
    while x < b.w * 4 / 5 {
      let (r, g, bl) = b.rgb(x, y)
      let v: UInt64 = UInt64((r >> 3) << 10) | UInt64((g >> 3) << 5) | UInt64(bl >> 3)
      h = (h ^ v) &* 1099511628211; x += 1
    }
    return h
  }
  // รอบแรก: หาแถวหัว-ท้ายกระดาษ แยกตามขนาดหน้า (แนวตั้ง/แนวนอน)
  var bmps: [Bitmap] = []; var allRows: [[Bool]] = []
  for i in 0..<n { let b = render(i, width: 400); bmps.append(b); allRows.append(inkRows(b)); savePNG(b, String(format: "%@/p%03d.png", out, i + 1)) }
  var furniture: [String: Set<Int>] = [:]
  var groups: [String: [Int]] = [:]; for i in 2..<max(2, n) { groups["\(bmps[i].w)x\(bmps[i].h)", default: []].append(i) }
  for (key, idx) in groups where idx.count >= 4 {
    let len = allRows[idx[0]].count; let y0 = Int(Double(bmps[idx[0]].h) * 0.11); var set = Set<Int>()
    for r in 0..<len {
      var tally: [UInt64: Int] = [:]
      for i in idx where allRows[i].count > r && allRows[i][r] { tally[rowSig(bmps[i], y0 + r), default: 0] += 1 }
      if let top = tally.values.max(), Double(top) >= Double(idx.count) * 0.7 { set.insert(r) }
    }
    // ขยาย 1 แถวรอบแถวหัว-ท้ายกระดาษ (ขอบเส้นที่ลบรอยหยัก ไม่เหมือนกันทุกพิกเซล)
    for r in Array(set) { for d in [-1, 1] where r + d >= 0 && r + d < len { set.insert(r + d) } }
    furniture[key] = set
  }
  var empty: [Int] = [], endHdr: [Int] = [], overR: [Int] = [], gaps: [Int] = [], short: [Int] = []
  var starts: [Int: Int] = [:]
  for i in 0..<n {
    let b = bmps[i]; let pg = i + 1; let W = b.w, H = b.h
    let landscape = W > H
    var rows = allRows[i]
    if let fs = furniture["\(W)x\(H)"] { for r in fs where r < rows.count { rows[r] = false } }
    let first = rows.firstIndex(of: true) ?? -1, last = rows.lastIndex(of: true) ?? -1
    if pg > 2 {
      if last < 0 || Double(last) < Double(rows.count) * 0.05 { empty.append(pg) }
      else if Double(last) < Double(rows.count) * 0.25 { short.append(pg) }
      var run = 0, best = 0
      if last >= 0 { for k in max(0, first)...last { if !rows[k] { run += 1; best = max(best, run) } else { run = 0 } } }
      if best > 30 { gaps.append(pg) }
      starts[pg] = first
      let yl = Int(Double(H) * 0.11) + last; var dark = 0, blue = 0, tint = 0, tot = 0
      if last >= 0 {
        for y in max(0, yl - 12)...min(H - 1, yl) { var x = 30; while x < W - 30 { let (r, _, bl) = b.rgb(x, y); tot += 1
          if r < 150 { dark += 1; if bl > r + 40 { blue += 1 } }; if bl > r + 15 && r > 150 { tint += 1 }; x += 2 } }
      }
      if (dark > 10 && Double(blue) / Double(dark) > 0.6) || Double(tint) / Double(max(1, tot)) > 0.4 { endHdr.append(pg) }
      let pr = doc.page(at: i)!.bounds(for: .mediaBox); let rot = doc.page(at: i)!.rotation % 180 != 0
      let pageWcm = Double(rot ? pr.height : pr.width) / 72 * 2.54
      let mr = rmCm / pageWcm; let xr = Int(Double(W) * (1 - mr)) + 3; var hit = 0
      var y = Int(Double(H) * 0.12)
      let fs = furniture["\(W)x\(H)"] ?? []
      while y < Int(Double(H) * 0.92) { let r = y - Int(Double(H) * 0.11)
        if !fs.contains(r) && xr < W - 2 { for x in xr..<(W - 2) { if b.rgb(x, y).0 < 150 { hit += 1 } } }; y += 2 }
      if hit > 5 { overR.append(pg) }
    }
    _ = landscape
  }
  var cnt: [Int: Int] = [:]; for v in starts.values { cnt[v, default: 0] += 1 }
  let mode = cnt.max { $0.value < $1.value || ($0.value == $1.value && $0.key > $1.key) }?.key ?? 0
  let low = starts.filter { $0.value > mode + 20 }.keys.sorted()
  func j(_ x: [Int]) -> String { x.map(String.init).joined(separator: ",") }
  print("pages: \(n)")
  print("blank: \(j(empty))")
  print("heading/table header at page end: \(j(endHdr))")
  print("past right margin: \(j(overR))")
  print("big white gap mid-page: \(j(gaps))")
  print("starts lower than normal: \(j(low))")
  print("less than 25% used: \(j(short))")

case "sheet":
  // ภาพรวมหลายหน้าในภาพเดียว (ไล่ดูทั้งเล่ม): sheet <pdf> <outprefix> [ต่อภาพ=28] [คอลัมน์=7] -> <outprefix>-1.png …
  let outp = a[3]; let per = a.count > 4 ? Int(a[4]) ?? 28 : 28; let cols = a.count > 5 ? Int(a[5]) ?? 7 : 7
  let cw = 220, chh = 320, lab = 18; var s = 0, sheetNo = 0
  while s < n {
    let cnt = min(per, n - s); let rows = (cnt + cols - 1) / cols; let W = cols * cw, H = rows * (chh + lab)
    var px = [UInt8](repeating: 160, count: W * H * 4)
    for k in 0..<cnt {
      let b = render(s + k, width: cw - 10); let c = k % cols, r = k / cols
      let ox = c * cw + 5, oy = r * (chh + lab) + lab
      for y in 0..<min(b.h, chh - 4) { for x in 0..<b.w { let si = (y * b.w + x) * 4, di = ((oy + y) * W + ox + x) * 4
        px[di] = b.px[si]; px[di + 1] = b.px[si + 1]; px[di + 2] = b.px[si + 2]; px[di + 3] = 255 } }
    }
    sheetNo += 1; let path = "\(outp)-\(sheetNo).png"
    var bm = Bitmap(w: W, h: H, px: px)
    // เลขหน้าเหนือภาพแต่ละหน้า
    let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: W, pixelsHigh: H, bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: W * 4, bitsPerPixel: 32)!
    bm.px.withUnsafeBytes { memcpy(rep.bitmapData!, $0.baseAddress!, W * H * 4) }
    NSGraphicsContext.saveGraphicsState(); NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
    for k in 0..<cnt { let c = k % cols, r = k / cols
      let y = H - (r * (chh + lab)) - lab + 2
      NSString(string: "\(s + k + 1)").draw(at: NSPoint(x: c * cw + 8, y: y), withAttributes: [.font: NSFont.boldSystemFont(ofSize: 13), .foregroundColor: NSColor.black]) }
    NSGraphicsContext.restoreGraphicsState()
    try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: path))
    print(path); s += cnt
  }

default: fail("unknown command \(a[1])")
}
