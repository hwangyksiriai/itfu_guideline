/**
 * 잇퓨 26년 10월 오프라인(무신사 매장 방문) 캠페인 신청 → 구글시트 첫 번째 시트에 기록
 * 시트: https://docs.google.com/spreadsheets/d/1B9X3vL9fVaAbv0Arl-vRrASzZmIhTlgV1uZgeixD_pc
 */
const HEADERS = ['제출일시', '고료', '등급', '방문 지점', '이름', '인스타그램', '휴대폰', '이메일', '요청사항', '캠페인', '페이지'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const row = [
      Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      data.price || '',
      data.tier || '',
      data.branch || '',
      data.name || '',
      data.instagram || '',
      String(data.phone || ''),
      data.email || '',
      data.note || '',
      data.campaign || '',
      data.page || ''
    ];

    // 휴대폰 앞자리 0이 사라지지 않도록 새 행을 텍스트 형식으로 지정한 뒤 기록
    const next = sheet.getLastRow() + 1;
    const range = sheet.getRange(next, 1, 1, row.length);
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
