const SHEET_NAME_USER_LIST = "名前";
const SHEET_NAME_TO_RECORD = "支払い(記載)";
const SHEET_NAME_TO_READ = "支払い金額(OnlyRead)";
const userNameSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USER_LIST);
const NAME_LIST = userNameSheet.getRange(`A2:A${userNameSheet.getLastRow()}`).getValues().flat();

function alphaToNum(alphabet) {
  if (typeof alphabet !== "string" && alphabet !== alphabet.toUpperCase())
    return null; // 文字列でない or 大文字でない => null
  let initalValue = 0, idx = 0;
  // 26進数(AA = 27)のように扱い、末尾の文字をidx=0にするためreverse()を使用する
  return alphabet.split("").reverse().reduce((accumulator, currentValue) => {
      let code = currentValue.charCodeAt(0);
      // (アルファベットの数字表現)*(26進数における位による計算), idxを次のループ用にインクリメントする必要があるが、return で返しているので、後置インクリメントを使用する
      return accumulator + (code - 64) * 26 ** idx++; 
    },initalValue
  );
}

function resetSheet() {
  const sheetToRecord = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_RECORD);
  const sheetToRead = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_READ);
  // 記録用のシートを初期化
  sheet.getRange("A2:Z100").clearDataValidations();
  sheet.getRange("A2:Z100").clearContent();
  // 名前部分のリセット
  sheet.getRange("F1:Z1").clear();
  sheet.getRange(1, 6, 1, NAME_LIST.length).setValues([NAME_LIST]);
  // 数式の更新
  for (let row = 2; row <= 100; row++) {
    const cell = sheet.getRange(`D${row}`);
    cell.setFormula(
      `=IF(B${row}="全員",C${row}/${NAME_LIST.length},(C${row}/COUNTA(SPLIT(B${row},","))))`
    );
  }
  // プルダウンの再生成
  resetPullDown();
  

  // 確認用のシートを初期化
  sheet.getRange(`A2:F${sheet.getLastRow()}`).clear();
  // 各名前ごとの合計を枠組みを作成する
  const header = [[ "", "支払い金額", "かかった金額", "支払い済み", "受け取り済み", "受け取り or 支払い"]];
  sheet.getRange(`A1:F1`).setValues(header);
  sheet.getRange(`A2:A${NAME_LIST.length + 1}`).setValues(NAME_LIST.map((name) => [name]));
  sheet.getRange(`A${NAME_LIST.length + 1 + 2}`).setValue("詳細");
  // 枠線の追加
  sheet.getRange(`A1:F${NAME_LIST.length + 1}`).setBorder(true, true, true, true, true, true);


}
function resetPullDown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_RECORD);
  // 支払い者のプルダウンの作成
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(NAME_LIST).build();
  let cell = sheet.getRange("A2:A100");
  cell.setDataValidation(rule);
  // 対象者のプルダウンの作成
  const targetValue = ["全員"].concat(NAME_LIST);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(targetValue).build();
  cell = sheet.getRange("B2:B100");
  cell.setDataValidation(rule);
}

function onEdit(e) {
  const TARGET_COL_IDX = [alphaToNum("A"), alphaToNum("B")];
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const target = sheet.getRange(`B${range.getRow()}`).getValue(); // 対象者を取得
  const sheetName = sheet.getName();

  // 対象のシートの編集かどうか
  if (sheetName !== SHEET_NAME_TO_RECORD) return;
  // 複数範囲の時は除外
  if (range.getNumRows() !== 1 && range.getNumColumns() !== 1) return;
  // 内容を変更する関数の実行
  if (TARGET_COL_IDX.includes(range.getColumn()))
    changeCheckBox(sheet, range, target);
  calculationMoney();
}

function changeCheckBox(sheet, range, target) {
  if (target == "") return;
  let row = range.getRow();
  let check_box_range = sheet.getRange(row, 6, 1, NAME_LIST.length); // F列からNAME_LISTの長さの範囲
  let save_origin_checkbox_data = check_box_range.getValues();
  // 初期化
  check_box_range.clearDataValidations();
  check_box_range.clearContent();
  if (target === undefined) return;
  let name_list = target.split(",").map((name) => name.trim());

  for (var i = 0; i < name_list.length; i++) {
    let name = name_list[i];
    if (name == "全員") {
      check_box_range.insertCheckboxes();
      break;
    } else {
      let index = NAME_LIST.indexOf(name);
      sheet.getRange(row, 6 + index).insertCheckboxes(); // 6+index => "F"+index
    }
  }
  // 元の値をsetする(その部分がcheckboxの時のみ)
  for (let row_idx = 0; row_idx < save_origin_checkbox_data.length; row_idx++) {
    for (
      let col_idx = 0;
      col_idx < save_origin_checkbox_data[row_idx].length;
      col_idx++
    ) {
      let range = sheet.getRange(row + row_idx, 6 + col_idx);
      if (
        range.getValue() !== "" &&
        save_origin_checkbox_data[row_idx][col_idx]
      ) {
        // checkboxである and 元の値がTureである時
        range.setValue("True");
      }
    }
  }
  let payer = sheet.getRange(`A${row}`).getValue().trim();
  if (name_list.includes(payer) || name_list.includes("全員")) {
    let index = NAME_LIST.indexOf(payer);
    if (index == -1) return;
    sheet.getRange(row, 6 + index).setValue("True");
  }
}

function calculationMoney() {
  // {立替者:{支払った金額(paidAmount):, かかった金額(totalCost):, 返済済み金額(repaidAmount):, 受け取り済み金額(receivedAmount):, 差額(受け取る(+) or 支払う(-), difference):}
  // 冗長であるがキーを増やし、わかりやすさ重視
  let payDictByName = {};
  NAME_LIST.forEach((name) => {
    payDictByName[name] = {
      paidAmount: 0,
      totalCost: 0,
      repaidAmount: 0,
      receivedAmount: 0,
      difference: 0,
    };
  });

  const sheetToRecord  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_RECORD);
  const sheetToRead = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_READ);
  const range = sheetToRecord.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN);
  const lastRow = range.getLastRow();
  const values = sheetToRecord.getRange(2, 1, lastRow - 1, alphaToNum("E") + NAME_LIST.length).getValues();
  // 計算する
  for (let row_idx = 0; row_idx < values.length; row_idx++) {
    const payer = values[row_idx][0];
    let targets = values[row_idx][1].split(",");
    const amount = values[row_idx][2];
    const amountPerPerson = values[row_idx][3];
    if (targets == "全員") targets = NAME_LIST;
    else if (targets == "") continue;
    // 支払った金額に加算する
    payDictByName[payer]["paidAmount"] += amount;
    // 合計を求める
    for (let i = 0; i < targets.length; i++) {
      let targetName = targets[i].trim();
      // かかった金額に加算する
      payDictByName[targetName]["totalCost"] += amountPerPerson;
      // targetに支払い済みマークがついているかを取得する
      let name_index = NAME_LIST.indexOf(targetName);
      let checkbox_value = values[row_idx][5 + name_index]; // 5はcheckboxの始まるIndex
      // チェックがついていれば返済済み金額、受け取り済み金額に加算する
      if (checkbox_value) {
        payDictByName[targetName]["repaidAmount"] += amountPerPerson;
        payDictByName[payer]["receivedAmount"] += amountPerPerson;
      }
    }
  }
  // 差額(受け取る or 支払う)を出す
  NAME_LIST.forEach((name) => {
    // 差額 = 支払い - 負担 + 返済済み - 受け取り済み
    payDictByName[name]["difference"] = payDictByName[name]["paidAmount"] - payDictByName[name]["totalCost"] + payDictByName[name]["repaidAmount"] - payDictByName[name]["receivedAmount"];
  });
  // 記述する
  // paymentForViewTable = [[名前,支払った金額, かかった金額, 返済済み金額, 受け取り済み金額, 差額]]
  let paymentForViewTable = Object.entries(payDictByName).map(([key, value]) => [key, ...Object.values(value)]);
  // [0]: 名前,[1]: 差額
  let receiveAmountByName = [];
  let repayAmountByName = [];
  Object.entries(payDictByName).forEach(([name, value]) => {
    let diffAmount = value["difference"];
    if (diffAmount > 0) receiveAmountByName.push([name, diffAmount]);
    else if (diffAmount < 0) repayAmountByName.push([name, diffAmount]);
  });
  receiveAmountByName.sort((a, b)=> b[1] - a[1]); // 降順ソート(絶対値が大きい順にする)
  repayAmountByName.sort((a, b)=> a[1] - b[1]); // 昇順ソート(絶対値が大きい順にする)
  let transactions = [];
  for (let i = 0; i < receiveAmountByName.length; i++) {
    const receiveName = receiveAmountByName[i][0];
    let receiveAmount = receiveAmountByName[i][1];
    for (let j = 0; j < repayAmountByName.length; j++) {
      const repayName = repayAmountByName[j][0];
      let repayAmount = repayAmountByName[j][1];
      if (receiveAmount === 0 || repayAmount === 0) continue;
      // 誰かに支払う分だけ残りの受け取り金額、支払い金額を減らす
      let minAmountReveiveOrRepay = Math.min(receiveAmount, Math.abs(repayAmount)); // 正の数で返す
      transactions.push([`${repayName} -> ${receiveName}から${Math.round(minAmountReveiveOrRepay)}円支払う`,]); // スプレッドシートに書き込むので2次元配列
      // 元のデータも変数のデータも両方変える
      receiveAmountByName[i][1] -= minAmountReveiveOrRepay;
      receiveAmount -= minAmountReveiveOrRepay;
      repayAmountByName[j][1] += minAmountReveiveOrRepay;
      repayAmount -= minAmountReveiveOrRepay;
    }
  }
  sheetToRead.getRange(`A2:F${NAME_LIST.length + 1}`).setValues(paymentForViewTable);
  sheetToRead.getRange(`A${NAME_LIST.length + 1 + 3}:A${NAME_LIST.length + 1 + 3 + transactions.length - 1}`).setValues(transactions);
}
