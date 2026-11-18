/**
 * 棚在庫_記録（Google Apps Script）
 *
 * スマホの「担当棚カメラ」（https://mocchihonpo.github.io/qr-shogo/tana/）専用の中継です。
 * 担当者の棚に置いてある品物（いずれ出ていく預かり品）を、スプレッドシート「棚在庫」に記録します。
 *
 *   ・写真1枚 → AIがラベルのマシンNoを全部読む
 *   ・バーコード／QR → そのまま記録
 *   ・同じ品物をまた読んだら、新しい行は作らず「置き場所（担当・段）」を書き換える
 *   ・出ていった品物は「出た」にする（行は消さない）
 *   ・マシンNoが提出用出荷リスト台帳にあれば、形名・得意先・出荷日・台帳の担当者を自動で入れる
 *
 * 使い方（最初の1回だけ）
 *  1. script.google.com で「新しいプロジェクト」を作り、名前を「棚在庫_記録」にする
 *  2. このコードを全部貼り付けて保存
 *  3. 左の歯車「プロジェクトの設定」→ 一番下「スクリプト プロパティ」→「プロパティを追加」
 *       プロパティ：ANTHROPIC_API_KEY　　値：中継スクリプトに入れているのと同じキー
 *  4. 上の関数の選択で setup を選んで「実行」→ 許可を求められたら許可する
 *     実行ログに「棚在庫」のURLが出ればOK（マイドライブにスプレッドシートができます）
 *  5. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員
 *     →「デプロイ」。出てきたURL（…/exec）をコピー
 *  6. 携帯の担当棚カメラの「設定」に、URLと合言葉を入れて「保存」→「つながるか試す」
 *
 * コードを書き換えたときは「デプロイを管理」→ 鉛筆 →「新バージョン」→「デプロイ」（URLは変わりません）
 */

const KEY       = 'mocchi2026';                                  // 合言葉（携帯の設定と同じ）
const LEDGER_ID = '1m5SU6OK972G4e-I-5VmbVyhsxtmbRgGuP1XkDWOTpN0'; // 提出用出荷リスト_履歴
const MODEL     = 'claude-sonnet-5-5';
const PHOTO_DIR = '棚在庫_写真';

const HEAD = ['登録日時', '担当者番号', '担当者名', '段', 'マシンNo', '形名', '得意先名', '出荷日', '状態', '出た日時', '読み方', '写真', '台帳の担当者', '更新日時'];
const TAN_HEAD = ['担当者番号', '担当者名', '3D棚ID', 'メモ'];
const ON = '棚にある', OUT = '出た';

/* ---------- 準備 ---------- */
function setup() {
  const ss = book_();
  Logger.log('棚在庫：' + ss.getUrl());
  Logger.log('APIキー：' + (PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') ? 'あり' : 'まだ入っていません（写真で読むには必要）'));
}
function book_() {
  const P = PropertiesService.getScriptProperties();
  let id = P.getProperty('BOOK_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('棚在庫');
    P.setProperty('BOOK_ID', ss.getId());
    const sh = ss.getSheets()[0]; sh.setName('棚在庫');
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#1d2140').setFontColor('#f2c94c');
    sh.setFrozenRows(1);
    const t = ss.insertSheet('担当者');
    t.getRange(1, 1, 1, TAN_HEAD.length).setValues([TAN_HEAD]).setFontWeight('bold').setBackground('#1d2140').setFontColor('#f2c94c');
    t.setFrozenRows(1);
  }
  return ss;
}
function sheet_(name) { return book_().getSheetByName(name); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'); }

// マシンNoの比べ方：英数字の先頭8文字（KKE-B5729-01 と KKEB5729 は同じ品物）
function mbase_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 8); }
function tidy_(s) {
  const t = String(s || '').toUpperCase().normalize('NFKC').replace(/[‐－ー—–\s]+/g, '-');
  const m = t.match(/([A-Z0-9]{3})-?([A-Z][A-Z0-9]\d{3})(?:-?(\d{2,3}))?/);
  return m ? m[1] + '-' + m[2] + (m[3] ? '-' + m[3] : '') : t;
}

/* ---------- 台帳から形名などを引く ---------- */
function ledger_() {
  const sh = SpreadsheetApp.openById(LEDGER_ID).getSheetByName('出荷リスト');
  const v = sh.getDataRange().getValues(), h = v[0].map(String);
  const c = k => h.indexOf(k);
  const iM = c('マシンNo'), iS = c('標準品'), iT = c('特殊品'), iC = c('得意先名'), iD = c('出荷日'), iP = c('担当者');
  const map = {};
  for (let r = 1; r < v.length; r++) {
    const b = mbase_(v[r][iM]); if (!b) continue;
    map[b] = { kata: String(v[r][iS] || v[r][iT] || ''), cust: String(v[r][iC] || ''), day: String(v[r][iD] || '').replace(/\.0$/, ''), tan: String(v[r][iP] || '').replace(/\.0$/, '') };
  }
  return map;
}
function enrich_(list) {
  let L = {}; try { L = ledger_(); } catch (e) {}
  return list.map(m => { const x = L[mbase_(m)] || {}; return { mno: tidy_(m), kata: x.kata || '', cust: x.cust || '', day: x.day || '', ltan: x.tan || '' }; });
}

/* ---------- 写真をAIで読む ---------- */
function ocr_(b64) {
  const keyA = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!keyA) throw new Error('スクリプト プロパティに ANTHROPIC_API_KEY がありません');
  const prompt = 'これは倉庫の棚の写真です。箱や品物に貼ってあるラベルから、三菱電機のマシンNoをすべて読み取ってください。\n' +
    'マシンNoは「英数字3文字-英数字5文字」の形です（例：KKE-B5729、K52-M8536-01）。後ろに -01 などが付くことがあります。\n' +
    'バーコードの数字、注文番号（例 311-4175）、形名（例 RTC-05FA3、CP30-BA）は含めないでください。\n' +
    '読めたものを重複なしで、次のJSONだけで答えてください：{"codes":["KKE-B5729","KKE-B5733"]}　見つからなければ {"codes":[]}';
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': keyA, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: MODEL, max_tokens: 1500, messages: [{ role: 'user', content: [
      { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
      { type: 'text', text: prompt }] }] })
  });
  const j = JSON.parse(res.getContentText());
  if (j.error) throw new Error('AI：' + (j.error.message || j.error.type));
  const text = (j.content || []).map(x => x.text || '').join('');
  const m = text.match(/\{[\s\S]*\}/);
  const codes = m ? (JSON.parse(m[0]).codes || []) : [];
  const seen = {}, outL = [];
  codes.forEach(c => { const t = tidy_(c); const b = mbase_(t); if (/^[A-Z0-9]{3}[A-Z][A-Z0-9]\d{3}$/.test(b) && !seen[b]) { seen[b] = 1; outL.push(t); } });
  return outL;
}
function savePhoto_(b64, tan, dan) {
  const root = DriveApp.getRootFolder(), it = root.getFoldersByName(PHOTO_DIR);
  const dir = it.hasNext() ? it.next() : root.createFolder(PHOTO_DIR);
  const name = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '_' + (tan || '担当不明') + '_' + (dan || '') + '段.jpg';
  return dir.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name)).getUrl();
}

/* ---------- 記録 ---------- */
function rows_() {
  const sh = sheet_('棚在庫'), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, HEAD.length).getValues();
}
function saveItems_(p) {
  const sh = sheet_('棚在庫'), v = rows_(), t = now_();
  const col = k => HEAD.indexOf(k);
  const idx = {}; v.forEach((r, i) => { if (r[col('状態')] === ON) idx[mbase_(r[col('マシンNo')])] = i; });
  const items = enrich_((p.items || []).map(x => x.mno));
  const add = []; let moved = 0, same = 0;
  items.forEach((x, k) => {
    const src = p.items[k] || {};
    const b = mbase_(x.mno), i = idx[b];
    if (i != null) {
      const r = v[i];
      if (String(r[col('担当者番号')]) === String(p.tan) && String(r[col('段')]) === String(p.dan || '')) { same++; return; }
      const row = i + 2;
      sh.getRange(row, col('担当者番号') + 1, 1, 2).setValues([[p.tan, p.name || '']]);
      sh.getRange(row, col('段') + 1).setValue(p.dan || '');
      sh.getRange(row, col('更新日時') + 1).setValue(t);
      moved++; return;
    }
    idx[b] = -1;
    add.push([t, p.tan, p.name || '', p.dan || '', x.mno, x.kata, x.cust, x.day, ON, '', src.how || '', src.photo || p.photo || '', x.ltan, t]);
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, HEAD.length).setValues(add);
  if (p.tan) saveTan_(p.tan, p.name);
  return { added: add.length, moved: moved, same: same, items: items };
}
function markOut_(mnos) {
  const sh = sheet_('棚在庫'), v = rows_(), t = now_();
  const col = k => HEAD.indexOf(k), want = {}, exact = {};
  (mnos || []).forEach(m => { exact[String(m)] = 1; if (!/^伝票写真/.test(m)) want[mbase_(m)] = 1; });
  let n = 0;
  v.forEach((r, i) => {
    const m = String(r[col('マシンNo')]);
    if (r[col('状態')] === ON && (exact[m] || (!/^伝票写真/.test(m) && want[mbase_(m)]))) {
      sh.getRange(i + 2, col('状態') + 1, 1, 2).setValues([[OUT, t]]);
      sh.getRange(i + 2, col('更新日時') + 1).setValue(t);
      n++;
    }
  });
  return n;
}
function saveTan_(tan, name) {
  const sh = sheet_('担当者'), n = sh.getLastRow();
  const v = n < 2 ? [] : sh.getRange(2, 1, n - 1, 2).getValues();
  const i = v.findIndex(r => String(r[0]) === String(tan));
  if (i < 0) sh.appendRow([tan, name || '', '', '']);
  else if (name && String(v[i][1]) !== String(name)) sh.getRange(i + 2, 2).setValue(name);
}
function tanList_() {
  const sh = sheet_('担当者'), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, 3).getValues().map(r => ({ tan: String(r[0]), name: String(r[1]), tana3d: String(r[2]) }));
}
function listOn_(tan) {
  const col = k => HEAD.indexOf(k);
  return rows_().filter(r => r[col('状態')] === ON && (!tan || String(r[col('担当者番号')]) === String(tan)))
    .map(r => ({ how: String(r[col('読み方')]), photo: String(r[col('写真')]), tan: String(r[col('担当者番号')]), name: String(r[col('担当者名')]), dan: String(r[col('段')]), mno: String(r[col('マシンNo')]),
      kata: String(r[col('形名')]), cust: String(r[col('得意先名')]), day: String(r[col('出荷日')]), at: String(r[col('更新日時')] || r[col('登録日時')]), ltan: String(r[col('台帳の担当者')]) }));
}

// 読めない伝票は、写真のまま1行にする（マシンNo欄は「伝票写真-日時」）
function savePhotoRow_(p) {
  const sh = sheet_('棚在庫'), t = now_();
  const url = savePhoto_(p.data, p.tan, p.dan);
  const id = '伝票写真-' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  sh.appendRow([t, p.tan, p.name || '', p.dan || '', id, '（伝票写真）', p.memo || '', '', ON, '', '伝票写真', url, '', t]);
  if (p.tan) saveTan_(p.tan, p.name);
  return { id: id, photo: url };
}

/* ---------- 入口 ---------- */
function doGet(e) {
  try {
    const q = e.parameter || {};
    if ((q.key || '') !== KEY) return out_({ ok: false, error: '合言葉が違います' });
    if (q.action === 'list') return out_({ ok: true, rows: listOn_(q.tan || ''), tans: tanList_() });
    return out_({ ok: true, book: book_().getUrl(), tans: tanList_() });
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.key !== KEY) return out_({ ok: false, error: '合言葉が違います' });
    if (p.action === 'ocr') {
      const codes = ocr_(p.data);
      const photo = p.keepPhoto ? savePhoto_(p.data, p.tan, p.dan) : '';
      return out_({ ok: true, items: enrich_(codes), photo: photo });
    }
    if (p.action === 'save') return out_(Object.assign({ ok: true }, saveItems_(p)));
    if (p.action === 'photo') return out_(Object.assign({ ok: true }, savePhotoRow_(p)));
    if (p.action === 'out') return out_({ ok: true, n: markOut_(p.mnos) });
    if (p.action === 'info') return out_({ ok: true, items: enrich_(p.mnos || []) });
    return out_({ ok: false, error: '知らない操作です' });
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}
