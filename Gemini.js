/*************************************************
 * Gemini.gs
 * Gemini OCR・領収書解析
 *************************************************/


function analyzeReceiptImageWithGemini(base64Data, mimeType) {
  if (!base64Data) throw new Error('領収書画像がありません。');

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません。');

  const model = getGeminiModel();
  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const prompt = getReceiptAnalysisPrompt();

  const payload = {
    contents: [
      {
        parts: [
          { text: prompt },
          {
            inlineData: {
              mimeType: mimeType || 'image/jpeg',
              data: base64Data
            }
          }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: 'application/json'
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseText = response.getContentText();
  const json = JSON.parse(responseText);
  const text = json.candidates?.[0]?.content?.parts?.[0]?.text;

  if (!text) {
    throw new Error(responseText);
  }

  return normalizeGeminiReceiptResult(JSON.parse(text));
}

function normalizeGeminiReceiptResult(result) {
  result = result || {};

  // 支払方法は、判定不能なら現金扱い。
  if (!result.paymentMethod || result.paymentMethod === '不明') {
    result.paymentMethod = '現金';
  }

  // キー表記ゆれを吸収。
  result.invoiceNumber = result.invoiceNumber || result.invoice_number || '';
  result.invoiceJudgement = result.invoiceJudgement || result.invoice_judgement || '';
  result.taxRate = result.taxRate || result.tax_rate || '';
  result.taxAmount = result.taxAmount || result.tax_amount || '';
  result.invoiceNote = result.invoiceNote || result.invoice_note || '';

  return result;
}

function analyzeReceiptWithGemini(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const mimeType = blob.getContentType();
  const base64Data = Utilities.base64Encode(blob.getBytes());

  return analyzeReceiptImageWithGemini(base64Data, mimeType);
}

function getReceiptAnalysisPrompt() {
  return `
あなたは日本の経費精算担当者です。
添付画像の領収書を読み取り、以下のJSONだけを返してください。

{
  "date": "YYYY-MM-DD",
  "vendor": "店舗名または取引先名",
  "amount": 数値,
  "category": "接待交際費|旅費交通費|車両費|通信費|消耗品費|その他",
  "paymentMethod": "現金|クレジット|電子マネー|QR決済|不明",
  "invoiceNumber": "Tから始まる13桁の登録番号。無ければ空文字",
  "invoiceJudgement": "登録番号あり|登録番号なし|不明",
  "taxRate": "10%|8%|非課税|不明",
  "taxAmount": 数値または空文字,
  "taxEstimated": trueまたはfalse,
  "taxNote": "税率・消費税額の判断理由"
}

ルール:
- 金額は税込合計と思われる金額。
- 読み取れない項目は空文字。
- 登録番号は T から始まる13桁の番号を抽出する。
- 登録番号が確認できる場合 invoiceJudgement は「適格」。
- 登録番号が見つからない場合 invoiceJudgement は「登録番号なし」。
- 登録番号の有無が画像品質などで判断できない場合 invoiceJudgement は「不明」。
- 「クレジット」「カード支払」「カード支払金額」「VISA」「JCB」「MASTER」「AMEX」「UFJニコス」などがあれば paymentMethod は「クレジット」。
- 「現金」「投入現金」「釣銭」「お釣り」などがあれば paymentMethod は「現金」。
- 支払方法を判定できない場合は「現金」。
- JSON以外の文章は返さない。
登録番号について：
- 登録番号は必ず「T」+ 数字13桁です。
- 領収書の「登録番号：」の直後に記載されている番号を最優先で読み取ってください。
- 登録番号の下にあるレシート番号、取引番号、端末番号、承認番号などは登録番号として扱わないでください。
- 数字の順序を推測で並び替えないでください。
- 1文字でも不鮮明な場合は invoiceJudgement を「不明」にしてください。
- 登録番号が読み取れた場合でも、13桁でない場合は空文字にしてください。
税率・消費税額について：
- 領収書に税率または消費税額が明記されている場合は、その値を優先してください。
- 税率や消費税額が記載されていない場合でも、税込金額から推定してください。
- 飲食店、居酒屋、レストラン、駐車場、ガソリン、宿泊、サービス業は原則10%として推定してください。
- 食品スーパー、食品販売、弁当、持ち帰り食品と思われる場合は8%の可能性があります。
- 判断できない場合は taxRate を「不明」、taxAmount を空文字にしてください。
- 推定で計算した場合は taxEstimated を true にしてください。
- 領収書に明記されていた場合は taxEstimated を false にしてください。
- 消費税額は、税込金額 ÷ 1.10 × 0.10 または 税込金額 ÷ 1.08 × 0.08 で計算し、小数点以下は四捨五入してください。
`;
}

function testGeminiText() {

  const apiKey =
    PropertiesService.getScriptProperties()
      .getProperty("GEMINI_API_KEY");

  const model = getGeminiModel();

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    apiKey;

  const payload = {
    contents: [
      {
        parts: [
          {
            text: "日本語で『Gemini接続成功』とだけ返してください。"
          }
        ]
      }
    ]
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });

  Logger.log(response.getContentText());

}