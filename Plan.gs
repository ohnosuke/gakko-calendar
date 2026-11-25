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
  '日': ['日付', '種類', 'メモ'],
  '授業': ['日付', '曜', '校時', 'クラス', '授業名', '状態', '実施ID', '単元', '内容・進捗'],
  '単元': ['単元', 'クラス', '予定の回数', '割り当て済み', '済み', 'メモ'],
  '課題': [],
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('授業計画')
    .addItem('最初の用意（タブと見出しを作る）', 'setup')
    .addItem('授業の行を足す', 'makeLessons')
    .addItem('課題の引き込みを作り直す', 'refreshTasks')
    .addToUi();
}

// タブが無ければ作り、見出し・入力の決まり・式を置く。あるタブの中身は消さない
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
    set.getRange('A2:C5').setValues([
      ['期間の始まり', '', '授業の行を作る最初の日'],
      ['期間の終わり', '', '授業の行を作る最後の日'],
      ['', '', ''],
      ['課題を引くブック', 'タブの名前', '下の行に、成績のブックの URL と課題管理のタブの名前を 1 行ずつ'],
    ]);
    set.getRange('A5:C5').setFontWeight('bold');
    set.getRange('B2:B3').setNumberFormat('yyyy-mm-dd');
  }
  set.setColumnWidth(1, 160).setColumnWidth(2, 320).setColumnWidth(3, 420);

  const day = ss.getSheetByName('日');
  day.getRange('A2:A').setNumberFormat('yyyy-mm-dd (ddd)');
  day.getRange('B2:B').setDataValidation(list_(DAY_KINDS));

  const les = ss.getSheetByName('授業');
  les.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  les.getRange('F2:F').setDataValidation(list_(STATES));
  les.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('課題').getRange('E2:E'), true).setAllowInvalid(true).build());
  les.getRange('H2:H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('単元').getRange('A2:A'), true).setAllowInvalid(true).build());
  // 同じクラス・同じ単元の何回目か（中止を除く。行は日付順に並べてある前提）
  les.getRange('J1').setFormula('={"回";ARRAYFORMULA(IF((H2:H="")+(F2:F="中止"),"",COUNTIFS(D2:D&"|"&H2:H,D2:D&"|"&H2:H,ROW(H2:H),"<="&ROW(H2:H),F2:F,"<>中止")))}');
  les.setColumnWidth(9, 360);

  const uni = ss.getSheetByName('単元');
  // 割り当て済み: 授業のタブでこの単元が書かれた行の数。済み: そのうち今日までの行
  uni.getRange('D1').setFormula('={"割り当て済み";ARRAYFORMULA(IF(A2:A="","",COUNTIFS(授業!H2:H,A2:A,授業!D2:D,B2:B,授業!F2:F,"<>中止")))}');
  uni.getRange('E1').setFormula('={"済み";ARRAYFORMULA(IF(A2:A="","",COUNTIFS(授業!H2:H,A2:A,授業!D2:D,B2:B,授業!F2:F,"<>中止",授業!A2:A,"<="&TODAY())))}');

  const tt = ss.getSheetByName('時間割');
  if (!tt.getRange('A2').getValue()) tt.getRange('A2:A7').setValues([['1'], ['2'], ['3'], ['4'], ['5'], ['6']]);

  refreshTasks();
  SpreadsheetApp.getActive().toast('タブと見出しを用意しました。設定のタブに期間と成績のブックを書いてください', '授業計画', 8);
}

// 設定のタブに書いた成績のブックから、課題管理の生徒のデータでない列だけを引く
// 課題名・対象クラス・提出期限・状態・実施ID・期（課題管理の A・C・D・I・J・K 列）
function refreshTasks() {
  const ss = SpreadsheetApp.getActive();
  const set = ss.getSheetByName('設定');
  const rows = set.getRange('A6:B' + Math.max(6, set.getLastRow())).getDisplayValues().filter(r => /^https:\/\//.test(r[0]));
  const sh = ss.getSheetByName('課題');
  sh.getRange('A1:F1').setValues([['課題名', '対象クラス', '提出期限', '状態', '実施ID', '期']]).setFontWeight('bold').setBackground('#1f2a5c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A2:F').clearContent();
  if (!rows.length) { sh.getRange('A2').setValue('（設定のタブの 6 行目から下に、成績のブックの URL を書いて「課題の引き込みを作り直す」）'); return; }
  const parts = rows.map(([url, tab]) =>
    'QUERY(IMPORTRANGE("' + url + '","' + (tab || '課題管理') + '!A2:K"),"select Col1,Col3,Col4,Col9,Col10,Col11 where Col1 is not null and Col10 is not null",0)');
  sh.getRange('A2').setFormula('=' + (parts.length === 1 ? parts[0] : 'VSTACK(' + parts.join(',') + ')'));
  sh.getRange('C2:C').setNumberFormat('yyyy-mm-dd (ddd)');
}

// 時間割と日のタブから、期間の平日の授業の行を足す。ある行（日付・校時・クラスが同じ）は触らない。
// 祝日と、日のタブで考査・休みの日は作らない。最後に日付・校時の順に並べ直す
function makeLessons() {
  const ss = SpreadsheetApp.getActive();
  const set = ss.getSheetByName('設定');
  const from = set.getRange('B2').getValue(), to = set.getRange('B3').getValue();
  if (!(from instanceof Date) || !(to instanceof Date) || to < from) throw new Error('設定のタブの B2・B3 に期間の始まりと終わりの日付を入れてください');

  const tt = ss.getSheetByName('時間割').getDataRange().getDisplayValues();
  const head = tt[0];
  const order = {};
  const slots = [];                       // [曜の番号, 校時, クラス, 授業名]
  tt.slice(1).forEach((r, i) => {
    if (!r[0]) return;
    order[r[0]] = i;
    for (let c = 1; c < head.length; c++) {
      const v = String(r[c] || '').trim();
      const wd = W_NAMES.indexOf(head[c]);
      if (!v || wd < 0) continue;
      const m = v.match(/^(\S+)[ 　](.+)$/);
      slots.push([wd, r[0], m ? m[1] : '', m ? m[2] : v]);
    }
  });

  const skip = {};
  ss.getSheetByName('日').getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] instanceof Date && NO_CLASS.indexOf(r[1]) >= 0) skip[ymd_(r[0])] = true;
  });
  const end = new Date(to.getFullYear(), to.getMonth(), to.getDate() + 1);
  const hol = CalendarApp.getCalendarById(HOLIDAY_CAL);
  if (hol) hol.getEvents(from, end).forEach(ev => { if (ev.isAllDayEvent()) skip[ymd_(ev.getAllDayStartDate())] = true; });

  const sh = ss.getSheetByName('授業');
  const last = sh.getLastRow();
  const old = last > 1 ? sh.getRange(2, 1, last - 1, 9).getValues().filter(r => r[0] !== '') : [];
  const have = {};
  old.forEach(r => { have[(r[0] instanceof Date ? ymd_(r[0]) : r[0]) + '|' + r[2] + '|' + r[3]] = true; });

  let added = 0;
  for (let d = new Date(from.getFullYear(), from.getMonth(), from.getDate()); d < end; d.setDate(d.getDate() + 1)) {
    const k = ymd_(d);
    if (skip[k]) continue;
    slots.forEach(([wd, period, cls, subject]) => {
      if (wd !== d.getDay() || have[k + '|' + period + '|' + cls]) return;
      old.push([new Date(d), W_NAMES[wd], period, cls, subject, '', '', '', '']);
      have[k + '|' + period + '|' + cls] = true;
      added++;
    });
  }
  old.sort((a, b) => (a[0] - b[0]) || ((order[a[2]] ?? 99) - (order[b[2]] ?? 99)));
  if (old.length) sh.getRange(2, 1, old.length, 9).setValues(old);
  ss.toast(added + ' 行を足しました（全部で ' + old.length + ' 行）', '授業計画', 6);
}

function list_(items) { return SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build(); }
function ymd_(d) { return Utilities.formatDate(d, PLAN_TZ, 'yyyy-MM-dd'); }
