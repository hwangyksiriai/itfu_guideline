/**
 * 잇퓨 26년 10월 오프라인(무신사 매장 방문) 캠페인 신청 → 구글시트 첫 번째 시트에 기록
 * 시트: https://docs.google.com/spreadsheets/d/1B9X3vL9fVaAbv0Arl-vRrASzZmIhTlgV1uZgeixD_pc
 *
 * 머리글 이름 기준으로 값을 넣기 때문에, 이미 머리글이 있는 시트에서도
 * 없는 컬럼(우편번호·주소 등)은 오른쪽 끝에 자동으로 추가됩니다.
 */
const COLUMNS = [
  ['제출일시', d => Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')],
  ['고료', d => d.price],
  ['등급', d => d.tier],
  ['방문 지점', d => d.branch],
  ['이름', d => d.name],
  ['인스타그램', d => d.instagram],
  ['휴대폰', d => String(d.phone || '')],
  ['이메일', d => d.email],
  ['우편번호', d => String(d.zipcode || '')],
  ['주소', d => d.address],
  ['요청사항', d => d.note],
  ['캠페인', d => d.campaign],
  ['페이지', d => d.page]
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // 머리글 확인 — 없으면 만들고, 빠진 컬럼은 오른쪽에 추가
    let headers = sheet.getLastRow() === 0
      ? []
      : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    COLUMNS.forEach(([name]) => { if (headers.indexOf(name) === -1) headers.push(name); });
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    const row = headers.map(h => {
      const col = COLUMNS.find(([name]) => name === h);
      const v = col ? col[1](data) : '';
      return v == null ? '' : v;
    });

    // 휴대폰·우편번호 앞자리 0이 사라지지 않도록 텍스트 형식으로 기록
    const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length);
    range.setNumberFormat('@');
    range.setValues([row]);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
