/**
 * 入荷品カメラ_写真受け取り（Google Apps Script）
 *
 * 入荷品カメラ（https://mocchihonpo.github.io/qr-shogo/nyuka/）で撮った写真を
 * Google Drive の「納入品」フォルダに保存するだけの中継です。
 *
 * 使い方
 *  1. script.google.com で「新しいプロジェクト」を作り、このコードを全部貼り付けて保存
 *  2. 上の関数の選択で testFolder を選んで「実行」→ 許可を求められたら許可する
 *     実行ログに送り先のフォルダ名が出ればOK
 *  3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員
 *     → 「デプロイ」。出てきたURL（…/exec）をコピー
 *  4. 携帯の入荷品カメラの「設定」に、URLと合言葉を入れて「保存」→「つながるか試す」
 *
 * コードを書き換えたときは「デプロイを管理」→ 鉛筆 →「新バージョン」→「デプロイ」で公開し直す
 */

const KEY = 'mocchi2026';      // 合言葉（携帯の設定と同じにする）
const FOLDER_NAME = '納入品';   // 送り先フォルダの名前（入荷MasterSheetがあるフォルダ）
const FOLDER_ID = '1nfKxlCdTNgKTxVjGDoaPPYNy2clB5Gjb'; // 「納入品」フォルダのID（入っていればこちらが優先）

function folder_() {
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!it.hasNext()) throw new Error('「' + FOLDER_NAME + '」フォルダが見つかりません');
  const f = it.next();
  if (it.hasNext()) throw new Error('「' + FOLDER_NAME + '」が複数あります。FOLDER_ID にフォルダIDを入れてください');
  return f;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 写真を受け取って保存
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.key !== KEY) return out_({ ok: false, error: '合言葉が違います' });
    if (!p.data || !p.name) return out_({ ok: false, error: '写真が空です' });
    const name = String(p.name).replace(/[\\\/:*?"<>|]/g, '_');
    const dir = folder_();
    // 同じ名前がもうあれば（再送のとき）保存し直さない
    const same = dir.getFilesByName(name);
    if (same.hasNext()) return out_({ ok: true, id: same.next().getId(), name: name, dup: true });
    const blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || 'image/jpeg', name);
    const file = dir.createFile(blob);
    if (p.memo) file.setDescription(String(p.memo));
    return out_({ ok: true, id: file.getId(), name: name });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// 「つながるか試す」用
function doGet(e) {
  try {
    if ((e.parameter.key || '') !== KEY) return out_({ ok: false, error: '合言葉が違います' });
    return out_({ ok: true, folder: folder_().getName() });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// 最初に一度だけ実行して、許可とフォルダの確認をする
function testFolder() {
  const f = folder_();
  Logger.log('送り先：' + f.getName() + '（' + f.getId() + '）');
}
