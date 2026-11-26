// 授業計画のブック（学校アカウント）に付ける GAS。空のスプレッドシートの「拡張機能 → Apps Script」に貼る。
// カレンダーの Code.gs とは別のブックに貼る。生徒のデータはこのブックに入れない（成績のブックから引くのは課題の一覧だけ）。
// 使い方は README の「授業計画のブック」。

const PLAN_TZ = 'Asia/Tokyo';
const HOLIDAY_CAL = 'ja.japanese#holiday@group.v.calendar.google.com';
const DAY_KINDS = ['考査', '短縮', '休み', '行事'];
const NO_CLASS = ['考査', '休み'];          // この種類の日は授業の行を作らない
const STATES = ['中止', '自習', '振替'];
const W_NAMES = ['日', '月', '火', '水', '木', '金', '土'];

const TABS = {
  '設定': ['項目', '値', 'メモ'],
  '時間割': ['校時', '月', '火', '水', '木', '金'],
  '日': ['始まり', '終わり', '種類', 'メモ'],
  '授業': ['日付', '曜', '校時', 'クラス', '授業名', '状態', '実施ID', '単元', '内容・進捗'],
  '単元': ['単元', 'クラス', '予定の回数', '割り当て済み', '済み', 'メモ'],
  '課題': [],
  '進み具合': [],
  '予定表': [],
};
const SUBJECT_MARK = { 'e': 'eスポーツ', '美': '美術' };   // 「選択3e」の最後の字 → 授業名

function onOpen() {
  SpreadsheetApp.getUi().createMenu('授業計画')
    .addItem('最初の用意（シートと見出しを作る）', 'setup')
    .addItem('授業の行を足す', 'makeLessons')
    .addItem('課題の引き込みを作り直す', 'refreshTasks')
    .addItem('進み具合・予定表の色を作り直す（単元を足したら）', 'refreshViews')
    .addToUi();
}

// シートが無ければ作り、見出し・入力の決まり・式を置く。あるシートの中身は消さない
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach((name, i) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, i);
    const head = TABS[name];
    if (head.length) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1f2a5c').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });
  ss.getSheets().forEach(sh => { if (/^(シート|Sheet)\d+$/.test(sh.getName()) && sh.getLastRow() === 0) ss.deleteSheet(sh); });

  const set = ss.getSheetByName('設定');
  if (!set.getRange('A2').getValue()) {
    set.getRange('A2:C12').setValues([
      ['期間の始まり', '', '授業の行を作る最初の日（1 学期の始業の日など）'],
      ['期間の終わり', '', '授業の行を作る最後の日（3 学期の授業の終わりなど）。考査・長期の休みは日のシートに書く'],
      ['', '', ''],
      ['時間割の切り替え', 'シートの名前', 'その日から使う時間割のシート。上から日付の順に。いちばん上の日付より前は、いちばん上のシート'],
      ['', '時間割', '例: 1 学期の始業の日'],
      ['', '', '例: 3 学期の始業の日 / 時間割_3学期'],
      ['', '', ''],
      ['', '', ''],
      ['', '', ''],
      ['課題を引くブック', 'シートの名前', '下の行に、成績のブックの URL と課題管理のシートの名前（空なら 課題管理）を 1 行ずつ'],
      ['', '', ''],
    ]);
    set.getRange('A5:C5').setFontWeight('bold');
    set.getRange('A11:C11').setFontWeight('bold');
    set.getRange('B2:B3').setNumberFormat('yyyy-mm-dd');
    set.getRange('A6:A10').setNumberFormat('yyyy-mm-dd');
  }
  set.setColumnWidth(1, 160).setColumnWidth(2, 320).setColumnWidth(3, 420);

  const day = ss.getSheetByName('日');
  day.getRange('A2:B').setNumberFormat('yyyy-mm-dd (ddd)').setDataValidation(null);  // 前の形（B 列が種類）のプルダウンを消す
  day.getRange('C2:C').setDataValidation(list_(DAY_KINDS));

  const les = ss.getSheetByName('授業');
  les.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  les.getRange('F2:F').setDataValidation(list_(STATES));
  les.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('課題').getRange('E2:E'), true).setAllowInvalid(true).build());
  les.getRange('H2:H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('単元').getRange('A2:A'), true).setAllowInvalid(true).build());
  // 同じクラス・同じ単元の何回目か（中止を除く。行は日付順に並べてある前提）
  les.getRange('J1').setFormula('={"回";ARRAYFORMULA(IF((H2:H="")+(F2:F="中止"),"",COUNTIFS(D2:D&"|"&H2:H,D2:D&"|"&H2:H,ROW(H2:H),"<="&ROW(H2:H),F2:F,"<>中止")))}');
  les.setColumnWidth(9, 360);

  const uni = ss.getSheetByName('単元');
  // 割り当て済み: 授業のシートでこの単元が書かれた行の数。済み: そのうち今日までの行
  // クラスは「31,32,33」のようにまとめて書いてよい。クラスごとに「31: 3 / 32: 4」の形で出す
  const per = (cond) => 'MAP(A2:A,B2:B,LAMBDA(u,c,IF(u="","",JOIN(" / ",MAP(SPLIT(c&"",",，、 "),LAMBDA(x,x&": "&COUNTIFS(授業!H2:H,u,授業!D2:D,x,授業!F2:F,"<>中止"' + cond + ')))))))';
  uni.getRange('D1').setFormula('={"割り当て済み";' + per('') + '}');
  uni.getRange('E1').setFormula('={"済み";' + per(',授業!A2:A,"<="&TODAY()') + '}');

  const tt = ss.getSheetByName('時間割');
  if (!tt.getRange('A2').getValue()) tt.getRange('A2:A7').setValues([['1'], ['2'], ['3'], ['4'], ['5'], ['6']]);

  refreshTasks();
  buildViews_(ss);
  SpreadsheetApp.getActive().toast('シートと見出しを用意しました。設定のシートに期間と成績のブックを書いてください', '授業計画', 8);
}

// 設定のシートに書いた成績のブックから、課題管理の生徒のデータでない列だけを引く
// 課題名・対象クラス・提出期限・状態・実施ID・期（課題管理の A・C・D・I・J・K 列）
function refreshTasks() {
  const ss = SpreadsheetApp.getActive();
  const set = ss.getSheetByName('設定');
  const rows = section_(set, '課題を引くブック').map(r => [String(r[0]).trim(), String(r[1]).trim()]).filter(r => /^https:\/\//.test(r[0]));
  const sh = ss.getSheetByName('課題');
  sh.getRange('A1:F1').setValues([['課題名', '対象クラス', '提出期限', '状態', '実施ID', '期']]).setFontWeight('bold').setBackground('#1f2a5c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A2:F').clearContent();
  if (!rows.length) { sh.getRange('A2').setValue('（設定のシートの「課題を引くブック」の下に、成績のブックの URL を書いて「課題の引き込みを作り直す」）'); return; }
  const parts = rows.map(([url, tab]) =>
    'QUERY(IMPORTRANGE("' + url + '","' + (tab || '課題管理') + '!A2:K"),"select Col1,Col3,Col4,Col9,Col10,Col11 where Col1 is not null and Col10 is not null",0)');
  sh.getRange('A2').setFormula('=' + (parts.length === 1 ? parts[0] : 'VSTACK(' + parts.join(',') + ')'));
  sh.getRange('C2:C').setNumberFormat('yyyy-mm-dd (ddd)');
}

// 時間割と日のシートから、期間の平日の授業の行を足す。ある行（日付・校時・クラスが同じ）は触らない。
// 祝日と、日のシートで考査・休みの日は作らない。最後に日付・校時の順に並べ直す
function makeLessons() {
  const ss = SpreadsheetApp.getActive();
  const set = ss.getSheetByName('設定');
  const from = set.getRange('B2').getValue(), to = set.getRange('B3').getValue();
  if (!(from instanceof Date) || !(to instanceof Date) || to < from) throw new Error('設定のシートの B2・B3 に期間の始まりと終わりの日付を入れてください');

  // 時間割の切り替え: [使い始める日, シート]。日付の無い行はいちばん初めから
  const order = {};
  const switches = section_(set, '時間割の切り替え')
    .filter(r => String(r[1]).trim())
    .map(r => ({ from: r[0] instanceof Date ? ymd_(r[0]) : '', slots: readTimetable_(ss, String(r[1]).trim(), order) }))
    .sort((a, b) => a.from < b.from ? -1 : a.from > b.from ? 1 : 0);
  if (!switches.length) switches.push({ from: '', slots: readTimetable_(ss, '時間割', order) });
  const slotsOn = k => { let s = switches[0].slots; switches.forEach(w => { if (w.from <= k) s = w.slots; }); return s; };

  const skip = {};
  ss.getSheetByName('日').getDataRange().getValues().slice(1).forEach(r => {
    if (!(r[0] instanceof Date) || NO_CLASS.indexOf(r[2]) < 0) return;
    const last = r[1] instanceof Date ? r[1] : r[0];
    for (let d = new Date(r[0]); d <= last; d.setDate(d.getDate() + 1)) skip[ymd_(d)] = true;
  });
  const end = new Date(to.getFullYear(), to.getMonth(), to.getDate() + 1);
  const hol = CalendarApp.getCalendarById(HOLIDAY_CAL);
  if (!hol) ss.toast('祝日のカレンダー（日本の祝日）が読めないので、祝日にも行を作ります。Google カレンダーで「日本の祝日」を登録すると読めます', '授業計画', 10);
  if (hol) hol.getEvents(from, end).forEach(ev => { if (ev.isAllDayEvent()) skip[ymd_(ev.getAllDayStartDate())] = true; });

  const sh = ss.getSheetByName('授業');
  const last = sh.getLastRow();
  const old = last > 1 ? sh.getRange(2, 1, last - 1, 9).getValues().filter(r => r[0] !== '') : [];
  const have = {};
  old.forEach(r => {           // 前の版でクラスを読めなかった行（「o31美術1」など）を直す
    if (r[3] === '' && r[4] !== '') { const [c, sub] = splitClass_(String(r[4])); if (c) { r[3] = c; r[4] = sub; } }
  });
  old.forEach(r => { have[(r[0] instanceof Date ? ymd_(r[0]) : r[0]) + '|' + r[2] + '|' + r[3]] = true; });

  let added = 0;
  for (let d = new Date(from.getFullYear(), from.getMonth(), from.getDate()); d < end; d.setDate(d.getDate() + 1)) {
    const k = ymd_(d);
    if (skip[k]) continue;
    slotsOn(k).forEach(([wd, period, cls, subject]) => {
      if (wd !== d.getDay() || have[k + '|' + period + '|' + cls]) return;
      old.push([new Date(d), W_NAMES[wd], period, cls, subject, '', '', '', '']);
      have[k + '|' + period + '|' + cls] = true;
      added++;
    });
  }
  old.sort((a, b) => (a[0] - b[0]) || ((order[a[2]] ?? 99) - (order[b[2]] ?? 99)));
  if (last > 1) sh.getRange(2, 1, last - 1, 9).clearContent();   // 途中の空の行を詰めるので、前の範囲を消してから書く
  if (old.length) sh.getRange(2, 1, old.length, 9).setValues(old);
  ss.toast(added + ' 行を足しました（全部で ' + old.length + ' 行）', '授業計画', 6);
}

// 設定のシートで、A 列がその見出しの行の下から、A・B 列が両方空の行までを返す
function section_(set, title) {
  const v = set.getDataRange().getValues();
  const i = v.findIndex(r => String(r[0]).trim() === title);
  if (i < 0) return [];
  const out = [];
  for (let j = i + 1; j < v.length; j++) {
    if (v[j][0] === '' && v[j][1] === '') break;
    out.push(v[j]);
  }
  return out;
}

// 時間割のシートを [曜の番号, 校時, クラス, 授業名] の並びにする。order に校時の並びを入れる
function readTimetable_(ss, name, order) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('時間割のシートが無い: ' + name);
  const tt = sh.getDataRange().getDisplayValues();
  const head = tt[0];
  const slots = [];
  tt.slice(1).forEach((r, i) => {
    if (!r[0]) return;
    if (!(r[0] in order)) order[r[0]] = i;
    for (let c = 1; c < head.length; c++) {
      const v = String(r[c] || '').trim();
      const wd = W_NAMES.indexOf(head[c]);
      if (!v || wd < 0) continue;
      const [cls, subject] = splitClass_(v);
      slots.push([wd, r[0], cls, subject]);
    }
  });
  return slots;
}

// 時間割のマスを [クラス, 授業名] に分ける。「31 美術1」「o31美術1」「o31 美術1」→ ['31', '美術1']。
// 頭の英字（担当の印など）は落とす。数字で始まらないときは、空白があれば前をクラス（「選択3年 美術」）、無ければクラスは空で全体を授業名にする
function splitClass_(v) {
  v = String(v).trim();
  const sp = v.match(/^(\S+)[ 　]+(.+)$/);
  const head = sp ? sp[1] : v;
  const sel = head.match(/^[A-Za-z]*(選択(\d)(\S*))$/);   // 「選択3e」→ 3 年の eスポーツ。クラスは「選択3e」のまま
  if (sel) return [sel[1], sp ? sp[2] : (SUBJECT_MARK[sel[3]] || sel[3] || '選択')];
  const m = head.match(/^[A-Za-z]*(\d+)(.*)$/);
  if (!m) return sp ? [sp[1], sp[2]] : ['', v];
  const rest = (m[2] + (sp ? ' ' + sp[2] : '')).trim();
  return rest ? [m[1], rest] : ['', v];
}

// 見る用のシート 2 枚。中身は式なので、授業・単元のシートを書けば勝手に変わる。色の決まりだけ GAS が置く
// 進み具合: 縦に単元、横にクラス。マスは「済み/予定の回数」（済み = 今日までに、その単元を書いた中止でない授業の数）
// 予定表: 縦にクラス、横に授業のある日付。マスはその日のその単元（中止は「中止」、単元が空なら「・」）
const SEP = '",，、 "';
function buildViews_(ss) {
  const prog = ss.getSheetByName('進み具合');
  prog.getRange('A1').setFormula(
    '=LET(u,FILTER(単元!A2:C,単元!A2:A<>""),' +
    'cls,SORT(UNIQUE(TOCOL(SPLIT(TEXTJOIN(",",TRUE,INDEX(u,,2)&""),' + SEP + '),3))),' +
    'body,MAKEARRAY(ROWS(u),ROWS(cls),LAMBDA(i,j,LET(cl,INDEX(cls,j,1),' +
    'IF(ISNUMBER(MATCH(cl&"",SPLIT(INDEX(u,i,2)&"",' + SEP + ')&"",0)),' +
    'COUNTIFS(授業!H2:H,INDEX(u,i,1),授業!D2:D,cl,授業!F2:F,"<>中止",授業!A2:A,"<="&TODAY())&"/"&N(INDEX(u,i,3)),"")))),' +
    'VSTACK(HSTACK("単元",TRANSPOSE(cls)),HSTACK(INDEX(u,,1),body)))');
  prog.setFrozenRows(1); prog.setFrozenColumns(1);
  prog.getRange('1:1').setFontWeight('bold');
  prog.setColumnWidth(1, 180);
  prog.getRange('B1:Z').setHorizontalAlignment('center');
  const area = prog.getRange('B2:Z200');
  const d = 'VALUE(REGEXEXTRACT(B2&"","^(\\d+)/"))', pl = 'VALUE(REGEXEXTRACT(B2&"","/(\\d+)$"))';
  const rule = (f, bg) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=IFERROR(' + f + ',FALSE)').setBackground(bg).setRanges([area]).build();
  prog.setConditionalFormatRules([
    rule(d + '>' + pl, '#f8c58c'),                              // 予定より多い: オレンジ
    rule('AND(' + pl + '>0,' + d + '=' + pl + ')', '#6fbf73'),  // 予定どおり終わった: 濃い緑
    rule(d + '>0', '#d4ecd5'),                                  // 途中: 薄い緑
  ]);

  const cal = ss.getSheetByName('予定表');
  cal.getRange('A1').setFormula(
    '=LET(dt,TOROW(SORT(UNIQUE(FILTER(授業!A2:A,授業!A2:A<>""))),1),' +
    'cls,SORT(UNIQUE(FILTER(授業!D2:D,授業!D2:D<>""))),' +
    'body,MAKEARRAY(ROWS(cls),COLUMNS(dt),LAMBDA(i,j,LET(f,FILTER(IF(授業!F2:F="中止","中止",IF(授業!H2:H="","・",授業!H2:H)),授業!A2:A=INDEX(dt,1,j),授業!D2:D=INDEX(cls,i,1)),' +
    'IFERROR(TEXTJOIN("・",TRUE,UNIQUE(f)),"")))),' +
    'VSTACK(HSTACK("クラス",dt),HSTACK(cls,body)))');
  if (cal.getMaxColumns() < CAL_COLS) cal.insertColumnsAfter(cal.getMaxColumns(), CAL_COLS - cal.getMaxColumns());   // 1 年分の授業の日が横に並ぶので広げておく
  cal.setFrozenRows(1); cal.setFrozenColumns(1);
  cal.getRange(1, 2, 1, CAL_COLS - 1).setNumberFormat('m/d').setFontWeight('bold');
  cal.getRange('A1:A').setFontWeight('bold');
  cal.setColumnWidths(2, CAL_COLS - 1, 56);
  cal.getRange(2, 2, 49, CAL_COLS - 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setFontSize(9);
  refreshViews();
}

// 予定表は、単元のシートの単元ごとに色を塗る（単元を足したら作り直す）
const CAL_COLS = 260;
const PALETTE = ['#cfe2f3', '#d9ead3', '#fff2cc', '#f4cccc', '#d9d2e9', '#fce5cd', '#d0e0e3', '#ead1dc', '#c9daf8', '#e6b8af'];
function refreshViews() {
  const ss = SpreadsheetApp.getActive();
  const cal = ss.getSheetByName('予定表');
  if (!cal) return;
  const units = ss.getSheetByName('単元').getRange('A2:A').getDisplayValues().map(r => r[0].trim()).filter(Boolean);
  const area = cal.getRange(2, 2, 49, CAL_COLS - 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('中止').setFontColor('#999999').setRanges([area]).build(),
  ];
  [...new Set(units)].forEach((u, i) => {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(u).setBackground(PALETTE[i % PALETTE.length]).setRanges([area]).build());
  });
  cal.setConditionalFormatRules(rules);
}

function list_(items) { return SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build(); }
function ymd_(d) { return Utilities.formatDate(d, PLAN_TZ, 'yyyy-MM-dd'); }
