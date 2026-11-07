/**
 * 転送ボックス_受け取り（Google Apps Script）
 *
 * 転送ボックス（https://mocchihonpo.github.io/qr-shogo/okuru/）専用の中継です。
 * 入荷品カメラとは別のスクリプト・別のフォルダを使います。
 *
 *   携帯 → Drive：マイドライブ「転送ボックス」/「携帯から」に保存
 *   Drive → 携帯：マイドライブ「転送ボックス」/「携帯へ」に入れたファイルを携帯で開ける
 *   （フォルダが無ければ最初に自動で作ります）
 *
 * 使い方
 *  1. script.google.com で「新しいプロジェクト」を作り、名前を「転送ボックス_受け取り」にする
 *  2. このコードを全部貼り付けて保存
 *  3. 上の関数の選択で testFolder を選んで「実行」→ 許可を求められたら許可する
 *     実行ログに「転送ボックス/携帯から」「転送ボックス/携帯へ」と出ればOK
 *  4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員
 *     → 「デプロイ」。出てきたURL（…/exec）をコピー
 *  5. 携帯の転送ボックスの「設定」に、URLと合言葉を入れて「保存」→「つながるか試す」
 *
 * コードを書き換えたときは「デプロイを管理」→ 鉛筆 →「新バージョン」→「デプロイ」で公開し直す
 */

const KEY = 'mocchi2026';        // 合言葉（携帯の設定と同じにする。変えてもOK）
const TOP = '転送ボックス';        // マイドライブ直下に作る親フォルダ
const SEND = '携帯から';           // 携帯から送ったものが入る
const RECV = '携帯へ';             // ここに入れたものを携帯で開ける
const MAX_GET = 25 * 1024 * 1024;  // 携帯へ渡す1ファイルの上限（25MB）

function sub_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function top_() { return sub_(DriveApp.getRootFolder(), TOP); }
function sendDir_() { return sub_(top_(), SEND); }
function recvDir_() { return sub_(top_(), RECV); }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function list_(dir, n) {
  const a = [];
  const it = dir.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    a.push({ id: f.getId(), name: f.getName(), size: f.getSize(), mime: f.getMimeType(), at: f.getLastUpdated().getTime() });
  }
  a.sort((x, y) => y.at - x.at);
  return a.slice(0, n);
}

// 携帯から受け取って保存
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.key !== KEY) return out_({ ok: false, error: '合言葉が違います' });
    if (!p.name) return out_({ ok: false, error: '名前が空です' });
    const name = String(p.name).replace(/[\\\/:*?"<>|]/g, '_');
    const dir = sendDir_();
    // 再送のとき、同じ名前がもうあれば保存し直さない
    const same = dir.getFilesByName(name);
    if (same.hasNext()) return out_({ ok: true, id: same.next().getId(), name: name, dup: true });
    let blob;
    if (p.text != null) blob = Utilities.newBlob(String(p.text), 'text/plain', name);
    else if (p.data) blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || 'application/octet-stream', name);
    else return out_({ ok: false, error: '中身が空です' });
    const file = dir.createFile(blob);
    return out_({ ok: true, id: file.getId(), name: name });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// つながるか試す／「携帯へ」の一覧／1ファイルを渡す
function doGet(e) {
  try {
    const q = e.parameter || {};
    if ((q.key || '') !== KEY) return out_({ ok: false, error: '合言葉が違います' });
    if (q.action === 'list') {
      return out_({ ok: true, recv: list_(recvDir_(), 50), sent: list_(sendDir_(), 20) });
    }
    if (q.action === 'get') {
      const f = DriveApp.getFileById(q.id);
      // 転送ボックスのフォルダの中のファイルだけ渡す
      const ps = f.getParents(); let inside = false;
      const ids = [recvDir_().getId(), sendDir_().getId()];
      while (ps.hasNext()) if (ids.indexOf(ps.next().getId()) >= 0) inside = true;
      if (!inside) return out_({ ok: false, error: '転送ボックスのファイルではありません' });
      let blob = f.getBlob();
      const mime = f.getMimeType();
      // Googleドキュメント・スプレッドシートはPDFにして渡す
      if (mime.indexOf('application/vnd.google-apps') === 0) blob = f.getAs('application/pdf');
      if (blob.getBytes().length > MAX_GET) return out_({ ok: false, error: '大きすぎて携帯に渡せません（25MBまで）' });
      return out_({ ok: true, name: blob.getName() || f.getName(), mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) });
    }
    return out_({ ok: true, send: TOP + '/' + SEND, recv: TOP + '/' + RECV });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// 最初に一度だけ実行（許可を出すため）
function testFolder() {
  Logger.log(TOP + '/' + sendDir_().getName());
  Logger.log(TOP + '/' + recvDir_().getName());
}
