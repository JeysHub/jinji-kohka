/**
 * 人事考課表：スタッフの自己評価を受け取り、担当者のメールへ送るスクリプト
 * （Google Apps Script。あなたのGoogleアカウントで動きます）
 *
 * ▼ 最初に下の2つだけ書き換えてください
 */
const TO = 'ここに送信先メールアドレス';   // 例：リハ室のメールアドレス（複数ならカンマ区切り）
const PASSCODE = 'ここに合言葉';           // スタッフに口頭や紙で伝える合言葉

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (String(req.passcode || '').trim() !== PASSCODE) return out({ ok: false, error: 'passcode' });
    const r = req.rec;
    if (!r || r.format !== 'jinji-kohka-3' || !r.name) return out({ ok: false, error: 'データ形式が正しくありません' });

    const name = String(r.name).slice(0, 50);
    const scores = (r.self && r.self.scores) || [];
    const total = scores.reduce((a, b) => a + (Number(b) || 0), 0);
    const filled = scores.filter(v => v).length;
    const fname = '考課_' + name.replace(/[\\/:*?"<>|\s]+/g, '') + '.json';
    const body = [
      '人事考課（自己評価）の提出がありました。',
      '',
      '氏名：' + name,
      '所属：' + (r.dept || ''),
      '自己評価合計：' + total + ' / 75点（' + filled + '/15項目）',
      '',
      '添付ファイルを保存し、考課ツールの「一次考課者」画面に読み込んでください。',
    ].join('\n');

    MailApp.sendEmail({
      to: TO,
      subject: '【人事考課 自己評価】' + name,
      body: body,
      attachments: [Utilities.newBlob(JSON.stringify(r, null, 1), 'application/json', fname)],
    });
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
