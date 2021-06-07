URL = "https://www.nhk.jp/p/pitagora/ts/WLQ76PGNW2/schedule/?area=120"; // スケジュール表のURL
APIKEY = "xx-xxxxx-xxxxx-xxxxx-xxxxx-xxxxx"; // PhantomJs CloudのAPI Key
SHEET_NAME = "ピタゴラスイッチ";  // スプレッドシートのシート名

function writePitagora() {

  // システム日付を取得
  const currentDate = new Date();

  const url = "https://phantomjscloud.com/api/browser/v2/" + [APIKEY] + "/?request=%7Burl:%22" + URL + "%22,renderType:%22html%22,outputAsJson:true%7D";
  console.log(url);

  const response = UrlFetchApp.fetch(url).getContentText();
  const json = JSON.parse(response);
  const html = json["content"]["data"];

  const table = Parser.data(html).from('<table class="series-contents">').to('</table>').build();
  const tbody = Parser.data(table).from('<tbody>').to('</tbody>').build();
  const rows = Parser.data(tbody).from('<tr class="">').to('</tr>').iterate();
  
  rows.forEach((key, index) => {
    //console.log(rows[index]);

    // タイトル
    const title = Parser.data(rows[index]).from('<td class="name">').to('</td>').build()
      .replace('<!---->', '').trim();
  
    // 日時
    const dates = Parser.data(rows[index]).from('<td class="date">').to('</td>').build().split('<br>');
    const date = dates[1].trim();
    const time = dates[2].trim();

    // 内容
    const description = Parser.data(rows[index]).from('<td class="description">').to('</td>').build()
      .trim();

    // 書き込み  
    appendSpreadRow(convertDate(currentDate, date), time, title, description);
  });

  // 重複削除
  removeSpreadDuplicates();
}

// M月D日（ddd）を日付に変換
function convertDate(currentDate, value) {
  let year = currentDate.getFullYear();
  const month = value.match(/\d+月/)[0].replace("月", "");
  const day = value.match(/\d+日/)[0].replace("日", "");
  // システム日付が12月で、出力月が1の時は1年足す
  if (currentDate.getMonth() + 1 == 12 &&
      month == 1) {
    year++;
  }
  return new Date(year, month - 1, day);
}

// 最終行に追加 (日/時間/タイトル/内容)
function appendSpreadRow(day, time, title, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_NAME);
  sheet.appendRow([day, time, title, description]);
}

// 重複行を削除
function removeSpreadDuplicates(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_NAME);
  sheet.getRange("A:D").removeDuplicates([1, 2]); // 日と時間が重複している行を削除
}
