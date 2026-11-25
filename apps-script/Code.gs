function doGet(e) {
  const sheet = SpreadsheetApp.openById('1Nz745kUTwFhXCzyAKOVZ3vaUs8Jkvydt67JEZ1PkAuU').getSheetById(0);
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const colors = range.getBackgrounds();
  const headers = values[0];
  const nameCol = headers.findIndex(h => h.indexOf('บัฟบ้าน') !== -1);
  const codeCols = headers.map((h,i) => h.indexOf('เลขบ้าน') !== -1 ? i : -1).filter(i => i >= 0);
  if (nameCol < 0 || !codeCols.length) throw new Error('ไม่พบหัวคอลัมน์บัฟบ้านหรือเลขบ้าน');
  const buffs = [];
  for (let r = 1; r < values.length; r++) {
    const name = values[r][nameCol].trim();
    if (!name) continue;
    const codes = [];
    for (const c of codeCols) {
      const hex = colors[r][c].replace('#','');
      const red = parseInt(hex.slice(0,2),16), green = parseInt(hex.slice(2,4),16), blue = parseInt(hex.slice(4,6),16);
      const value = values[r][c].trim();
      if (green > red + 8 && green > blue + 8 && /^\d{7}(?:\s*\(Lv\s*\d{1,2}\))?$/i.test(value)) codes.push(value.replace(/\s+/g,''));
    }
    if (codes.length) buffs.push({name:name, codes:[...new Set(codes)]});
  }
  const json = JSON.stringify({buffs:buffs, updatedAt:new Date().toISOString()}).replace(/</g,'\\u003c');
  if (e && e.parameter.callback === 'sssFoodBuffs') {
    return ContentService.createTextOutput('sssFoodBuffs(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
