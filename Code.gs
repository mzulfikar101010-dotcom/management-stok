/**
 * STOK JAJAN — Jembatan Google Sheets
 * ------------------------------------------------
 * Cara pakai:
 * 1. Buat Google Sheet baru (boleh kosong).
 * 2. Buka menu Extensions > Apps Script.
 * 3. Hapus semua kode default, ganti dengan seluruh isi file ini.
 * 4. Klik Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Klik Deploy, salin URL Web App yang muncul.
 * 6. Tempel URL itu ke bagian SHEET_API_URL di file HTML aplikasi.
 *
 * Tab yang dipakai (dibuat otomatis kalau belum ada):
 * - Produk      : Nama | Stok | StokKritis (default 30 kalau dikosongkan)
 * - PaketItems  : NamaPaket | NamaProduk | Qty
 * - Riwayat     : ID | Waktu | Tipe | Keterangan | Detail
 * - Resi        : ID | Resi | OrderID | Waktu | Tipe | NamaProdukAtauPaket |
 *                 Qty | Status | Keterangan
 *   Kolom Status bisa diedit MANUAL langsung di spreadsheet, misalnya
 *   diganti jadi "Selesai" kalau uang COD sudah cair. Status lain yang
 *   dipakai otomatis oleh aplikasi: "Belum Packing", "Scan Packing", "Retur".
 *
 * TIPS IMPORT DATA LAMA:
 * Tinggal buka tab "Produk", isi kolom A (Nama) dan B (Stok) dengan
 * data lama kamu (boleh copy-paste dari spreadsheet lama). Begitu
 * aplikasi dibuka/di-refresh, data itu otomatis muncul di app.
 */

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = readAllData(ss);
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:'Payload tidak valid'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  writeAllData(ss, body);
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function clearBody_(sheet, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
  }
}

function readAllData(ss) {
  const prodSheet = getOrCreateSheet_(ss, 'Produk', ['Nama', 'Stok', 'StokKritis']);
  const pakSheet  = getOrCreateSheet_(ss, 'PaketItems', ['NamaPaket', 'NamaProduk', 'Qty']);
  const riwSheet  = getOrCreateSheet_(ss, 'Riwayat', ['ID', 'Waktu', 'Tipe', 'Keterangan', 'Detail']);
  const resiSheet = getOrCreateSheet_(ss, 'Resi', ['ID', 'Resi', 'OrderID', 'Waktu', 'Tipe', 'NamaProdukAtauPaket', 'Qty', 'Status', 'Keterangan']);

  // Produk
  const products = [];
  const prodValues = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodValues.length; i++) {
    const row = prodValues[i];
    const name = String(row[0] || '').trim();
    if (!name) continue;
    const critRaw = row[2];
    const critical = (critRaw === '' || critRaw === null || critRaw === undefined) ? 30 : (Number(critRaw) || 0);
    products.push({ id: name, name: name, stock: Number(row[1]) || 0, criticalStock: critical });
  }

  // Paket (dikelompokkan dari baris-baris PaketItems)
  const pkgMap = {};
  const pakValues = pakSheet.getDataRange().getValues();
  for (let i = 1; i < pakValues.length; i++) {
    const row = pakValues[i];
    const pname = String(row[0] || '').trim();
    if (!pname) continue;
    if (!pkgMap[pname]) pkgMap[pname] = { id: pname, name: pname, items: [] };
    const prodName = String(row[1] || '').trim();
    const qty = Number(row[2]) || 0;
    if (prodName && qty > 0) {
      pkgMap[pname].items.push({ productId: prodName, qty: qty });
    }
  }
  const packages = Object.keys(pkgMap).map(k => pkgMap[k]);

  // Riwayat
  const movements = [];
  const riwValues = riwSheet.getDataRange().getValues();
  for (let i = 1; i < riwValues.length; i++) {
    const row = riwValues[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    let changes = [];
    try { changes = JSON.parse(row[4] || '[]'); } catch (err) { changes = []; }
    const ts = Date.parse(row[1]) || Date.now();
    movements.push({ id: id, ts: ts, type: String(row[2] || ''), note: String(row[3] || ''), changes: changes });
  }
  movements.sort((a, b) => b.ts - a.ts);

  // Resi
  const resi = [];
  const resiValues = resiSheet.getDataRange().getValues();
  for (let i = 1; i < resiValues.length; i++) {
    const row = resiValues[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    const ts = Date.parse(row[3]) || Date.now();
    resi.push({
      id: id,
      resi: String(row[1] || ''),
      orderId: String(row[2] || ''),
      ts: ts,
      matchType: String(row[4] || ''),
      name: String(row[5] || ''),
      qty: Number(row[6]) || 0,
      status: String(row[7] || 'Belum Packing'),
      ket: String(row[8] || '')
    });
  }
  resi.sort((a, b) => b.ts - a.ts);

  return { products: products, packages: packages, movements: movements, resi: resi };
}

function writeAllData(ss, data) {
  const prodSheet = getOrCreateSheet_(ss, 'Produk', ['Nama', 'Stok', 'StokKritis']);
  const pakSheet  = getOrCreateSheet_(ss, 'PaketItems', ['NamaPaket', 'NamaProduk', 'Qty']);
  const riwSheet  = getOrCreateSheet_(ss, 'Riwayat', ['ID', 'Waktu', 'Tipe', 'Keterangan', 'Detail']);
  const resiSheet = getOrCreateSheet_(ss, 'Resi', ['ID', 'Resi', 'OrderID', 'Waktu', 'Tipe', 'NamaProdukAtauPaket', 'Qty', 'Status', 'Keterangan']);

  // Produk
  clearBody_(prodSheet, 3);
  const prodRows = (data.products || []).map(p => [p.name, p.stock, (p.criticalStock === undefined || p.criticalStock === null) ? 30 : p.criticalStock]);
  if (prodRows.length) prodSheet.getRange(2, 1, prodRows.length, 3).setValues(prodRows);

  // PaketItems
  clearBody_(pakSheet, 3);
  const pakRows = [];
  (data.packages || []).forEach(pkg => {
    (pkg.items || []).forEach(it => {
      pakRows.push([pkg.name, it.productId, it.qty]);
    });
  });
  if (pakRows.length) pakSheet.getRange(2, 1, pakRows.length, 3).setValues(pakRows);

  // Riwayat (dibatasi 1000 terbaru)
  clearBody_(riwSheet, 5);
  const mv = (data.movements || []).slice(0, 1000);
  const riwRows = mv.map(m => [
    m.id,
    new Date(m.ts).toISOString(),
    m.type,
    m.note,
    JSON.stringify(m.changes || [])
  ]);
  if (riwRows.length) riwSheet.getRange(2, 1, riwRows.length, 5).setValues(riwRows);

  // Resi (dibatasi 2000 terbaru)
  clearBody_(resiSheet, 9);
  const rs = (data.resi || []).slice(0, 2000);
  const resiRows = rs.map(r => [
    r.id,
    r.resi,
    r.orderId || '',
    new Date(r.ts || Date.now()).toISOString(),
    r.matchType || '',
    r.name || '',
    r.qty || 0,
    r.status || 'Belum Packing',
    r.ket || ''
  ]);
  if (resiRows.length) resiSheet.getRange(2, 1, resiRows.length, 9).setValues(resiRows);
}
