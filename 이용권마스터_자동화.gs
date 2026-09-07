/**
 * 학교인쇄통 이용권 마스터 — 자동화 스크립트 (v2 · 결제·증빙 추가)
 * 1) H열이 「확인」으로 바뀌면 → 이용권 코드 메일 자동 발송, N열 「발송」
 * 2) 매일 오전 8시 → 상태가 「만료임박」인데 O열이 「미발송」이면 만료 알림 메일 발송, O열 「발송」
 * 3) Q열(결제방법)이 「계좌이체」로 입력되면 → 견적서 PDF 자동 메일 발송, T열 「발송」
 * 4) 매일 오전 8시 → 세금계산서 미발행(R열 세금계산서 + H열 확인 + U열 빈칸) 목록을 관리자에게 메일
 *
 * 최초 설치: setupTriggers 1회 실행 → setupPaymentColumns 1회 실행(Q~U열·설정·메일문안 추가)
 */

const SHEET_MASTER = '이용권마스터';
const SHEET_SETTING = '설정';
const SHEET_MAIL = '메일문안';

// 열 번호 (1부터)
const COL = { NO:1, DATE:2, SCHOOL:3, TEACHER:4, EMAIL:5, PHONE:6, TYPE:7, CONFIRM:8,
              CONFIRM_DATE:9, CODE:10, EXPIRE:11, DAYS:12, STATUS:13, CODE_SENT:14, EXPIRE_SENT:15, MEMO:16,
              PAY:17, DOC:18, BIZNO:19, QUOTE_SENT:20, DOC_ISSUED:21 };
const LAST_COL = 21;

/** 최초 1회 실행 — 트리거 2개 설치 */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailyExpiryCheck').timeBased().everyDays(1).atHour(8).create();
  Logger.log('트리거 설치 완료: H열 확인→코드 메일 / 매일 08시 만료 알림');
}

/** 1회 실행 — 결제·증빙 열(Q~U), 설정 회사정보, 견적서 메일문안 추가 */
function setupPaymentColumns() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_MASTER);
  if (sh.getMaxColumns() < LAST_COL) sh.insertColumnsAfter(sh.getMaxColumns(), LAST_COL - sh.getMaxColumns());
  const heads = ['결제방법', '증빙종류', '사업자등록번호', '견적서발송', '증빙발행'];
  const hr = sh.getRange(1, COL.PAY, 1, 5);
  if (!String(hr.getValues()[0][0]).trim()) {
    hr.setValues([heads]).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  }
  const n = Math.max(sh.getMaxRows() - 1, 150);
  sh.getRange(2, COL.PAY, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['계좌이체', '법인카드', '개인카드', '휴대폰', '무료'], true).setAllowInvalid(false).build());
  sh.getRange(2, COL.DOC, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['세금계산서', '현금영수증', '카드전표', '없음'], true).setAllowInvalid(false).build());
  sh.getRange(2, COL.DOC_ISSUED, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['발행', '자동', '해당없음'], true).setAllowInvalid(true).build());
  sh.getRange(2, COL.BIZNO, n, 1).setNumberFormat('@');
  sh.setColumnWidths(COL.PAY, 5, 110);

  // 설정 탭 회사 정보 (E4~F10)
  const st = ss.getSheetByName(SHEET_SETTING);
  const rows = [
    ['회사명(견적서)', '메가플러스(주) 학교인쇄통'],
    ['사업자등록번호', '(입력)'],
    ['대표자', '김정수'],
    ['주소', '(입력)'],
    ['입금계좌', '(입력) 예: 국민은행 000-00-0000 메가플러스(주)'],
    ['담당자 연락처', '02-2269-1100 · tong2269@naver.com'],
    ['견적 유효기간(일)', 30],
  ];
  if (!String(st.getRange('E4').getValue()).trim()) {
    st.getRange('E3').setValue('견적서 회사 정보 (노란 칸 입력)').setFontWeight('bold');
    st.getRange(4, 5, rows.length, 2).setValues(rows);
    st.getRange(4, 6, rows.length, 1).setBackground('#FFF2CC');
    st.setColumnWidth(5, 150); st.setColumnWidth(6, 320);
  }

  // 메일문안 탭 견적서 메일 (A10~B11)
  const ml = ss.getSheetByName(SHEET_MAIL);
  if (!String(ml.getRange('B10').getValue()).trim()) {
    ml.getRange('A10:B11').setValues([
      ['견적서 메일 제목', '[학교인쇄통] {학교명} 디지털자료실 연간 이용권 견적서'],
      ['견적서 메일 본문', '{담당교사} 선생님, 안녕하세요. 학교인쇄통입니다.\n\n요청하신 {학교명} 디지털자료실 연간 이용권 견적서를 첨부해 드립니다.\n\n· 금액: {가격} (부가세 포함)\n· 입금계좌: {입금계좌}\n· 견적 유효기간: {유효기간}까지\n\n입금이 확인되면 학교 코드를 이메일로 보내드리고, 세금계산서는 홈택스로 발행해 드립니다.\n행정실 처리에 필요한 서류(사업자등록증·통장사본)가 있으면 회신 주세요.\n\n감사합니다.\n학교인쇄통 드림 · 02-2269-1100'],
    ]);
    ml.getRange('B10:B11').setBackground('#FFF2CC').setWrap(true);
  }
  Logger.log('결제·증빙 열 설치 완료: Q~U열, 설정 E4~F10, 메일문안 B10~B11. 설정 탭 노란 칸(사업자번호·주소·계좌) 입력 필요');
}

/** 시트 편집 트리거 */
function onEditInstalled(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_MASTER) return;
  const row = e.range.getRow(), col = e.range.getColumn();
  if (row < 2) return;

  // H열 「확인」 → 코드 메일
  if (col === COL.CONFIRM && String(e.value).trim() === '확인') {
    const dateCell = sh.getRange(row, COL.CONFIRM_DATE);
    if (!dateCell.getValue()) dateCell.setValue(new Date());
    SpreadsheetApp.flush();
    if (sh.getRange(row, COL.CODE_SENT).getValue() !== '발송') sendCodeMail(sh, row);
    return;
  }
  // Q열 「계좌이체」 → 견적서 메일
  if (col === COL.PAY && String(e.value).trim() === '계좌이체') {
    if (sh.getRange(row, COL.QUOTE_SENT).getValue() === '발송') return;
    sendQuoteMail(sh, row);
  }
}

/** 코드 발급 메일 */
function sendCodeMail(sh, row) {
  const v = sh.getRange(row, 1, 1, LAST_COL).getValues()[0];
  const email = String(v[COL.EMAIL - 1]).trim();
  const code = String(v[COL.CODE - 1]).trim();
  if (!email || !code) return;
  const set = getSettings();
  const tpl = getMailTemplates();
  const map = baseMap(v, set);
  map['{코드}'] = code;
  GmailApp.sendEmail(email, fill(tpl.codeSubject, map), fill(tpl.codeBody, map), { name: '학교인쇄통', cc: set.adminMail });
  sh.getRange(row, COL.CODE_SENT).setValue('발송');
  sh.getRange(row, COL.MEMO).setValue(appendMemo(v[COL.MEMO - 1], '코드메일 ' + fmt(new Date())));
}

/** 견적서 PDF 메일 (계좌이체 학교) */
function sendQuoteMail(sh, row) {
  const v = sh.getRange(row, 1, 1, LAST_COL).getValues()[0];
  const email = String(v[COL.EMAIL - 1]).trim();
  if (!email) return;
  const set = getSettings();
  const tpl = getMailTemplates();
  const map = baseMap(v, set);
  const pdf = buildQuotePdf(v, set, row);
  GmailApp.sendEmail(email, fill(tpl.quoteSubject, map), fill(tpl.quoteBody, map),
    { name: '학교인쇄통', cc: set.adminMail, attachments: [pdf] });
  sh.getRange(row, COL.QUOTE_SENT).setValue('발송');
  sh.getRange(row, COL.MEMO).setValue(appendMemo(v[COL.MEMO - 1], '견적서 ' + fmt(new Date())));
}

/** 견적서 PDF 생성 (HTML → PDF) */
function buildQuotePdf(v, set, row) {
  const today = new Date();
  const valid = new Date(today.getTime() + (Number(set.validDays) || 30) * 86400000);
  const total = Number(set.price) || 0;
  const supply = Math.round(total / 1.1);
  const vat = total - supply;
  const no = 'ST-Q-' + Utilities.formatDate(today, 'Asia/Seoul', 'yyyyMMdd') + '-' + String(row).padStart(4, '0');
  const won = n => fmtNum(n) + '원';
  const esc = t => String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const html = `<html><head><meta charset="utf-8"><style>
    body{font-family:'Noto Sans KR','Malgun Gothic',sans-serif;color:#1F2430;font-size:12px;margin:36px}
    h1{font-size:26px;letter-spacing:.4em;text-align:center;margin:0 0 6px;color:#1B2A4A}
    .sub{text-align:center;color:#666;margin-bottom:24px;font-size:11px}
    table{border-collapse:collapse;width:100%} td,th{border:1px solid #BBB;padding:7px 9px;vertical-align:top}
    th{background:#EEF1F6;text-align:left;width:22%;font-weight:700;white-space:nowrap}
    .two{width:100%;border:0} .two>tbody>tr>td{border:0;padding:0 8px 0 0;width:50%;vertical-align:top}
    .items th{text-align:center;width:auto;white-space:nowrap} .items td.r{text-align:right;white-space:nowrap}
    .total{background:#FBF3D9;font-weight:700}
    .note{margin-top:18px;font-size:11px;color:#444;line-height:1.7}
    .stamp{margin-top:28px;text-align:right;font-size:13px;font-weight:700;color:#1B2A4A}
  </style></head><body>
  <h1>견 적 서</h1><div class="sub">견적번호 ${no} · 작성일 ${fmt(today)} · 유효기간 ${fmt(valid)}</div>
  <table class="two"><tr><td>
    <table><tr><th>수 신</th><td>${esc(v[COL.SCHOOL-1])}</td></tr>
      <tr><th>담 당</th><td>${esc(v[COL.TEACHER-1])} 선생님</td></tr>
      <tr><th>연락처</th><td>${esc(v[COL.PHONE-1])}<br>${esc(v[COL.EMAIL-1])}</td></tr>
      <tr><th>사업자번호</th><td>${esc(v[COL.BIZNO-1])}</td></tr></table>
    </td><td>
    <table><tr><th>공급자</th><td>${esc(set.company)}</td></tr>
      <tr><th>사업자번호</th><td>${esc(set.bizNo)}</td></tr>
      <tr><th>대표자</th><td>${esc(set.ceo)}</td></tr>
      <tr><th>주 소</th><td>${esc(set.address)}</td></tr>
      <tr><th>연락처</th><td>${esc(set.contact)}</td></tr></table>
  </td></tr></table>
  <p style="margin:16px 0 8px">아래와 같이 견적합니다.</p>
  <table class="items"><tr><th>품 목</th><th>내 용</th><th>수량</th><th>공급가액</th><th>세 액</th><th>합 계</th></tr>
    <tr><td>디지털자료실 학교 연간 이용권</td><td>학교 코드 1개 · 선생님 ${esc(set.accounts)}명 등록 · ${esc(set.months)}개월 · 자료 무제한 다운로드 · 인쇄비 ${Math.round(set.discount*100)}% 할인</td><td class="r">1</td><td class="r">${won(supply)}</td><td class="r">${won(vat)}</td><td class="r">${won(total)}</td></tr>
    <tr class="total"><td colspan="3">합계 (부가세 포함)</td><td class="r">${won(supply)}</td><td class="r">${won(vat)}</td><td class="r">${won(total)}</td></tr></table>
  <div class="note">· 입금계좌: <b>${esc(set.account)}</b><br>· 입금 확인 후 1영업일 내 학교 코드를 이메일로 발급하며, 세금계산서는 홈택스로 전자발행합니다.<br>· 이용 기간은 코드 발급일부터 ${esc(set.months)}개월입니다.<br>· 문의: ${esc(set.contact)}</div>
  <div class="stamp">${esc(set.company)} &nbsp;(인)</div>
  </body></html>`;
  return Utilities.newBlob(html, MimeType.HTML, 'quote.html').getAs(MimeType.PDF)
    .setName('견적서_' + String(v[COL.SCHOOL-1]).replace(/\s+/g, '') + '_' + Utilities.formatDate(today, 'Asia/Seoul', 'yyyyMMdd') + '.pdf');
}

/** 매일 오전 8시 — 만료임박 알림 + 세금계산서 미발행 목록 */
function dailyExpiryCheck() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MASTER);
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(2, 1, last - 1, LAST_COL).getValues();
  const set = getSettings();
  const tpl = getMailTemplates();
  let sent = 0;
  const pending = [];

  data.forEach((v, i) => {
    const row = i + 2;
    // 세금계산서 미발행
    if (String(v[COL.DOC - 1]) === '세금계산서' && String(v[COL.CONFIRM - 1]).trim() === '확인' && !String(v[COL.DOC_ISSUED - 1]).trim()) {
      pending.push(`${row}행 ${v[COL.SCHOOL - 1]} / ${v[COL.TEACHER - 1]} / 사업자번호 ${v[COL.BIZNO - 1] || '(없음)'} / 확인일 ${fmt(v[COL.CONFIRM_DATE - 1])}`);
    }
    // 만료 알림
    if (String(v[COL.STATUS - 1]) !== '만료임박') return;
    if (String(v[COL.EXPIRE_SENT - 1]) === '발송') return;
    const email = String(v[COL.EMAIL - 1]).trim();
    if (!email) return;
    const map = baseMap(v, set);
    map['{코드}'] = v[COL.CODE - 1];
    GmailApp.sendEmail(email, fill(tpl.expSubject, map), fill(tpl.expBody, map), { name: '학교인쇄통', cc: set.adminMail });
    sh.getRange(row, COL.EXPIRE_SENT).setValue('발송');
    sh.getRange(row, COL.MEMO).setValue(appendMemo(v[COL.MEMO - 1], '만료알림 ' + fmt(new Date())));
    sent++;
  });

  if (set.adminMail && (sent > 0 || pending.length > 0)) {
    let body = '';
    if (sent > 0) body += '· 만료 알림 발송 ' + sent + '건 (O열 확인)\n\n';
    if (pending.length) body += '· 세금계산서 미발행 ' + pending.length + '건 — 홈택스 발행 후 U열에 「발행」 입력\n' + pending.join('\n') + '\n\n';
    body += SpreadsheetApp.getActive().getUrl();
    GmailApp.sendEmail(set.adminMail, '[이용권 마스터] ' + fmt(new Date()) + ' 만료 알림 ' + sent + '건 · 계산서 미발행 ' + pending.length + '건', body);
  }
}

/** 설정 탭 읽기 */
function getSettings() {
  const s = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTING);
  const g = a => s.getRange(a).getValue();
  return {
    price: g('B4'), discount: g('B5'), months: g('B6'), alertDays: g('B7'), prefix: g('B8'), accounts: g('B9'),
    adminMail: String(g('B10')).trim(),
    company: g('F4'), bizNo: g('F5'), ceo: g('F6'), address: g('F7'), account: g('F8'), contact: g('F9'), validDays: g('F10'),
  };
}

/** 메일문안 탭 읽기 */
function getMailTemplates() {
  const s = SpreadsheetApp.getActive().getSheetByName(SHEET_MAIL);
  return {
    codeSubject: s.getRange('B4').getValue(), codeBody: s.getRange('B5').getValue(),
    expSubject: s.getRange('B7').getValue(), expBody: s.getRange('B8').getValue(),
    quoteSubject: s.getRange('B10').getValue(), quoteBody: s.getRange('B11').getValue(),
  };
}

function baseMap(v, set) {
  const valid = new Date(Date.now() + (Number(set.validDays) || 30) * 86400000);
  return {
    '{학교명}': v[COL.SCHOOL - 1], '{담당교사}': v[COL.TEACHER - 1], '{코드}': v[COL.CODE - 1],
    '{만료일}': fmt(v[COL.EXPIRE - 1]), '{계정수}': set.accounts,
    '{가격}': fmtNum(set.price) + '원', '{할인율}': Math.round(set.discount * 100) + '%',
    '{입금계좌}': set.account, '{유효기간}': fmt(valid), '{사업자번호}': v[COL.BIZNO - 1],
  };
}
function fill(text, map) {
  let out = String(text);
  Object.keys(map).forEach(k => { out = out.split(k).join(map[k] == null ? '' : map[k]); });
  return out;
}
function fmtNum(n) { return String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function fmt(d) {
  if (!(d instanceof Date)) return String(d || '');
  return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
}
function appendMemo(old, add) {
  old = String(old || '').trim();
  return old ? old + ' / ' + add : add;
}

/** 테스트용: 선택한 행에 코드 메일 강제 발송 */
function testSendSelectedRow() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_MASTER) { Logger.log('이용권마스터 탭에서 행을 선택하세요'); return; }
  sendCodeMail(sh, sh.getActiveRange().getRow());
  Logger.log('발송 완료 (N열 확인)');
}
/** 테스트용: 선택한 행에 견적서 강제 발송 */
function testSendQuoteSelectedRow() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_MASTER) { Logger.log('이용권마스터 탭에서 행을 선택하세요'); return; }
  sendQuoteMail(sh, sh.getActiveRange().getRow());
  Logger.log('견적서 발송 완료 (T열 확인)');
}
