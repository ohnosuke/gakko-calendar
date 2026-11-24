// gakko-calendar の GAS。時間割のスプレッドシートの「拡張機能 → Apps Script」に貼る。
// 合言葉（TOKEN）はスクリプトのプロパティに入れる（makeToken）。ここにも git にも書かない。
// カレンダー ID は「カレンダー」シートに置く。ここに書き込まない。

const TZ = 'Asia/Tokyo';
const MAX_WEEKS = 8;
const DAY = { '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6 };

function doGet(e) {
  const p = (e && e.parameter) || {};
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || p.token !== token) return json_({ error: 'unauthorized' });

  const weeks = Math.min(Math.max(parseInt(p.weeks, 10) || 3, 1), MAX_WEEKS);
  const start = p.start ? parseDate_(p.start) : sundayOf_(new Date());
  if (!start) return json_({ error: 'bad start' });

  const cache = CacheService.getScriptCache();
  const key = 'v2_' + fmt_(start, 'yyyyMMdd') + '_' + weeks;
  const hit = p.fresh === '1' ? null : cache.get(key);
  if (hit) return text_(hit);

  const end = new Date(start);
  end.setDate(end.getDate() + weeks * 7);
  const warnings = [];
  const data = {
    generated: fmt_(new Date(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
    start: fmt_(start, 'yyyy-MM-dd'),
    days: weeks * 7,
    periods: readPeriods_(),
    meetings: readMeetings_(),
    timetable: readTimetable_(),
    events: readEvents_(start, end, warnings),
    warnings: warnings,
  };
  const out = JSON.stringify(data);
  if (out.length < 90000) cache.put(key, out, 300);
  return text_(out);
}

// 時程: 校時 / 開始 / 終了。その右に「考査 開始」「考査 終了」のような列の組を足すと、
// 学校の終日の予定にその言葉（考査）が入る日は、その時刻を使う
function readPeriods_() {
  const all = sheet_('時程').getDataRange().getDisplayValues();
  const head = all[0] || [];
  const alts = {};
  head.forEach((h, c) => {
    const m = String(h).trim().match(/^(\S+)[ 　]+(開始|終了)$/);
    if (!m || c < 3) return;
    alts[m[1]] = alts[m[1]] || {};
    alts[m[1]][m[2] === '開始' ? 's' : 'e'] = c;
  });
  return all.slice(1)
    .filter(r => r[0])
    .map(r => {
      const p = { name: r[0], start: r[1], end: r[2] };
      Object.keys(alts).forEach(k => {
        const a = alts[k];
        const st = a.s != null ? r[a.s] : '', en = a.e != null ? r[a.e] : '';
        if (st || en) { p.alt = p.alt || {}; p.alt[k] = { start: st, end: en }; }
      });
      return p;
    });
}

// 会議: 名前 / 開始 / 終了。シートが無ければ空
function readMeetings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('会議');
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues().slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({ name: String(r[0]).trim(), start: r[1], end: r[2] }));
}

function readTimetable_() {
  const rows = sheet_('時間割').getDataRange().getDisplayValues();
  const head = rows[0];
  const list = [];
  rows.slice(1).forEach(r => {
    for (let c = 1; c < head.length; c++) {
      const v = (r[c] || '').trim();
      if (!v || !(head[c] in DAY)) continue;
      const m = v.match(/^(\S+)[ 　](.+)$/);
      list.push({
        day: DAY[head[c]],
        period: r[0],
        cls: m ? m[1] : '',
        subject: m ? m[2] : v,
      });
    }
  });
  return list;
}

function readEvents_(start, end, warnings) {
  const rows = sheet_('カレンダー').getDataRange().getDisplayValues().slice(1);
  const list = [];
  rows.forEach(r => {
    const [id, label, mode] = r;
    if (!id) return;
    const cal = CalendarApp.getCalendarById(id);
    if (!cal) { warnings.push('読めないカレンダー: ' + label); return; }
    cal.getEvents(start, end).forEach(ev => {
      try {
        if (ev.getMyStatus() === CalendarApp.GuestStatus.NO) return;
      } catch (err) {}
      const allDay = ev.isAllDayEvent();
      list.push({
        cal: label,
        title: mode === '予定あり' ? '予定あり' : ev.getTitle(),
        allDay: allDay,
        start: allDay ? fmt_(ev.getAllDayStartDate(), 'yyyy-MM-dd')
                      : fmt_(ev.getStartTime(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
        end: allDay ? fmt_(ev.getAllDayEndDate(), 'yyyy-MM-dd')
                    : fmt_(ev.getEndTime(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
      });
    });
  });
  return list;
}

function makeToken() {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN', t);
  console.log('合言葉: ' + t);
}

function testRun() {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  const data = JSON.parse(doGet({ parameter: { token: token, fresh: '1' } }).getContent());
  if (data.error) { console.log('エラー: ' + data.error); return; }
  console.log(`${data.start} から ${data.days} 日 / 時程 ${data.periods.length} 行（別の時程: ${[...new Set(data.periods.flatMap(p => Object.keys(p.alt || {})))].join('・') || 'なし'}） / 会議 ${data.meetings.length} 件 / 時間割 ${data.timetable.length} コマ / 予定 ${data.events.length} 件`);
  if (data.warnings.length) console.log('注意: ' + data.warnings.join(' / '));
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('シートが無い: ' + name);
  return sh;
}
function sundayOf_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() - d.getDay()); }
function parseDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f); }
function json_(o) { return text_(JSON.stringify(o)); }
function text_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON); }
