const SHEET_NAME_USER_LIST = "名前";
const SHEET_NAME_TO_RECORD = "支払い(記載)";
const SHEET_NAME_TO_READ = "支払い金額(OnlyRead)";
const userNameSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USER_LIST);
const NAME_LIST = userNameSheet.getRange(`A2:A${userNameSheet.getLastRow()}`).getValues().flat();

/**
 * 名前のリストが重複している場合に警告する
 * @return {boolean} 重複しているかどうか
 */
function isDuplicateName(){
  const setNameList = new Set(NAME_LIST);
  if(setNameList.size !== NAME_LIST.length){
    Browser.msgBox("名前が重複しています");
    return true;
  }
  return false;
}


/**
 * アルファベット -> 数字(A->1, AB -> 28)
 * @params {string} alphabet 変換するアルファベット
 * @return {number} アルファベットを数字で示した値 
 * 
 */
function alphaToNum(alphabet) {
  // 文字列でない or 大文字でない => null
  if (typeof alphabet !== "string" && alphabet !== alphabet.toUpperCase())return null; 
  /**
   * @type {number} initalValue 初期値を示す
   * @type {number} idx 何桁目のアルファベットかを示す
   */
  const initalValue = 0;
  let idx = 0;
  // 26進数(AA = 27)のように扱い、末尾の文字をidx=0にするためreverse()を使用する
  return alphabet.split("").reverse().reduce((accumulator, currentValue) => {
      const code = currentValue.charCodeAt(0);
      // (アルファベットの数字表現)*(26進数における位による計算), idxを次のループ用にインクリメントする必要があるが、return で返しているので、後置インクリメントを使用する
      return accumulator + (code - 64) * 26 ** idx++; 
    },initalValue
  );
}

/**
 * シートをリセットする
 * 基本的に名前が変更になった場合のみ使用する
 */
function resetSheet() {
  if(isDuplicateName()) return;
    /**
   * @type {sheet} sheetToRecord 記録用のシート
   * @type {sheet} sheetToRead   確認用のシート
   */
  const sheetToRecord = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_RECORD);
  const sheetToRead   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_READ);
  // 記録用のシートを初期化
  sheetToRecord.getRange("A2:Z100").clearDataValidations();
  sheetToRecord.getRange("A2:Z100").clearContent();
  // 名前部分のリセット
  sheetToRecord.getRange("F1:Z1").clear();
  sheetToRecord.getRange(1, 6, 1, NAME_LIST.length).setValues([NAME_LIST]);
  // 数式の更新
  for (let row = 2; row <= 100; row++) {
    const cell = sheetToRecord.getRange(`D${row}`);
    cell.setFormula(
      `=IF(B${row}="全員",C${row}/${NAME_LIST.length},(C${row}/COUNTA(SPLIT(B${row},","))))`
    );
  }
  // プルダウンの再生成
  resetPullDown();
  

  // 確認用のシートを初期化
  sheetToRead.getRange(`A2:F${sheetToRead.getLastRow()}`).clear();
  // 各名前ごとの合計を枠組みを作成する
  const header = [[ "", "支払い金額", "かかった金額", "支払い済み", "受け取り済み", "受け取り or 支払い"]];
  sheetToRead.getRange(`A1:F1`).setValues(header);
  sheetToRead.getRange(`A2:A${NAME_LIST.length + 1}`).setValues(NAME_LIST.map((name) => [name]));
  sheetToRead.getRange(`A${NAME_LIST.length + 1 + 2}`).setValue("詳細");
  // 枠線の追加
  sheetToRead.getRange(`A1:F${NAME_LIST.length + 1}`).setBorder(true, true, true, true, true, true);


}

/**
 * プルダウン部分をリセットする
 * プルダウンには名前("全員"を含む)を選択できるようにする
 * ※現状ではプルダウンの複数選択をコードから選択することは不可能
 */
function resetPullDown() {
  /**
   * @type {sheet} sheet プルダウンをセットするシート(記録用シート)
   * @type {rule}  rule  プルダウンのルール
   * @type {range} range プルダウンを適応する範囲
   */
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_RECORD);
  // 支払い者のプルダウンの作成
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(NAME_LIST).build();
  let range = sheet.getRange("A2:A100");
  range.setDataValidation(rule);
  // 対象者のプルダウンの作成
  rule = SpreadsheetApp.newDataValidation().requireValueInList(["全員",...NAME_LIST]).build();
  range = sheet.getRange("B2:B100");
  range.setDataValidation(rule);
}

/**
 * 変更があった際に実行するトリガー
 */
function onEdit(e) {
  if(isDuplicateName()) return;
  /**
   * @type {[number]} colIdxForCheckBox チェックボックスを変更するかどうか決める列
   * @type {sheet}    sheet          選択されたシート
   * @type {range}    range          選択された範囲
   */
  const colIdxForCheckBox = [alphaToNum("A"), alphaToNum("B")];
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  

  // 対象のシートの編集かどうか
  if (sheet.getName() !== SHEET_NAME_TO_RECORD) return;
  // 複数範囲の時は除外
  if (range.getNumRows() !== 1 && range.getNumColumns() !== 1) return;
  // 支払い者, 対象者が変更された場合のみチェックボックスの追加, 削除を行う
  /**
   * @type {[name]}   target         選択された行の対象者の名前
   */
  const targets = sheet.getRange(`B${range.getRow()}`).getValue(); // 対象者を取得
  if (colIdxForCheckBox.includes(range.getColumn())) changeCheckBox(sheet, range.getRow(), targets);
  
  // 金額を計算する
  calculationMoney();
}

/**
 * @params {sheet}  sheet   変更するシート
 * @params {row}  row       変更する行
 * @params {[name]} targets 対象者の名前
 */
function changeCheckBox(sheet, row, targets) {
  if (targets == "") return;
  /**
   * @type {range}     check_box_range           選択された行のチェックボックスの範囲
   * @type {[boolean]} save_origin_checkbox_data 選択された行のチェックボックスの値 
   */
  const check_box_range = sheet.getRange(row, 6, 1, NAME_LIST.length); // F列からNAME_LISTの長さの範囲
  const save_origin_checkbox_data = check_box_range.getValues();
  // 初期化
  check_box_range.clearDataValidations();
  check_box_range.clearContent();
  if (targets === undefined) return;
  /**
   * @type {[string]} nameList 対象者の名前一覧(複数の場合もあり)
   */
  const nameList = targets.split(",").map((name) => name.trim()); // trim()で空白がある場合には削除
  // チェックボックスを追加する
  for (var i = 0; i < nameList.length; i++) {
    const name = nameList[i];
    if (name == "全員") {
      check_box_range.insertCheckboxes();
      break;
    } else {
      const nameIdx = NAME_LIST.indexOf(name);
      sheet.getRange(row, alphaToNum("F") + nameIdx).insertCheckboxes();
    }
  }
  // チェックボックスの初期化する前の値をセットする
  for (let rowIdx = 0; rowIdx < save_origin_checkbox_data.length; rowIdx++) {
    for (let colIdx = 0;colIdx < save_origin_checkbox_data[rowIdx].length;colIdx++) {
      const range = sheet.getRange(row + rowIdx, 6 + colIdx);
      // checkboxである and 元の値がTureである時
      if (range.getValue() !== "" && save_origin_checkbox_data[rowIdx][colIdx]) range.setValue("True");
    }
  }
  /**
   * @type {string} payer 支払い者
   */
  const payer = sheet.getRange(`A${row}`).getValue().trim();
  if (nameList.includes(payer) || nameList.includes("全員")) {
    const nameIdx = NAME_LIST.indexOf(payer);
    if (nameIdx == -1) return;
    sheet.getRange(row, alphaToNum("F") + nameIdx).setValue("True");
  }
}

/**
 * 支払いを計算してSHEET_NAME_TO_READに記載する
 */
function calculationMoney() {
  /**
   * @type {{string: {string: number}}} payDictByName 立替者:{支払った金額(paidAmount):, かかった金額(totalCost):, 返済済み金額(repaidAmount):, 受け取り済み金額(receivedAmount):, 差額(受け取る(+) or 支払う(-), difference):}  冗長であるがキーを増やし、わかりやすさ重視
   */
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
  /**
   * @type {sheet}             sheetToRecord 記録用のシート
   * @type {sheet}             sheetToRead   確認用のシート
   * @type {[[string|number]]} values        名前、金額の入った値
   */
  const sheetToRecord = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_RECORD);
  const sheetToRead   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_READ);
  const values        = sheetToRecord.getRange(2, 1, sheetToRead.getLastRow() - 1, alphaToNum("D") + NAME_LIST.length).getValues();
  // 計算する
  for (let row_idx = 0; row_idx < values.length; row_idx++) {
    /**
     * @type {string} payer           支払い者
     * @type {string} targets         対象者(複数の場合もあり)
     * @type {number} amount          支払った金額
     * @type {number} amountPerPerson 一人当たりの金額
     */
    const payer = values[row_idx][0];
    const targets = values[row_idx][1].split(",");
    const amount = values[row_idx][2];
    const amountPerPerson = values[row_idx][3];
    if (targets == "全員") targets = NAME_LIST;
    else if (targets == "") continue;
    // 支払った金額に加算する
    payDictByName[payer]["paidAmount"] += amount;
    // 合計を求める
    for (let i = 0; i < targets.length; i++) {
      /**
       * @type {string}  対象者の名前
       * @type {string}  NAME_LISTの対象者の番号
       * @type {boolean} 支払ったかどうか
       */
      const targetName = targets[i].trim();
      const name_index = NAME_LIST.indexOf(targetName);
      const checkbox_value = values[row_idx][alphaToNum("F") + name_index]; // targetに支払い済みマークがついているかを取得する
      // かかった金額に加算する
      payDictByName[targetName]["totalCost"] += amountPerPerson;
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
  /**
   * @type {[string]} paymentForViewTable　表示する用の表([[名前,支払った金額, かかった金額, 返済済み金額, 受け取り済み金額, 差額]])
   */
  const paymentForViewTable = Object.entries(payDictByName).map(([key, value]) => [key, ...Object.values(value)]);
  /**
   * @type{[string]} receiveAmountByName 受け取る人([0]: 名前,[1]: 差額)
   * @type{[string]} repayAmountByName   返済者([0]: 名前,[1]: 差額)
   */
  let receiveAmountByName = [];
  let repayAmountByName = [];
  Object.entries(payDictByName).forEach(([name, value]) => {
    const diffAmount = value["difference"];
    if (diffAmount > 0) receiveAmountByName.push([name, diffAmount]);
    else if (diffAmount < 0) repayAmountByName.push([name, diffAmount]);
  });
  receiveAmountByName.sort((a, b)=> b[1] - a[1]); // 降順ソート(絶対値が大きい順にする)
  repayAmountByName.sort((a, b)=> a[1] - b[1]); // 昇順ソート(絶対値が大きい順にする)
  /**
   * @type {[string]} transactions 誰が誰にいくら支払うかの詳細を記録する配列
   */
  let transactions = []; // スプレッドシートに書き込むので2次元配列にする
  for (let i = 0; i < receiveAmountByName.length; i++) {
    for (let j = 0; j < repayAmountByName.length; j++) {
      /**
       * @type {string} receiveName   受け取る人の名前
       * @type {number} receiveAmount 受け取る額
       * @type {string} repayName     返済人の名前
       * @type {number} repayAmount   返済額
       */
      const receiveName = receiveAmountByName[i][0];
      const receiveAmount = receiveAmountByName[i][1];
      const repayName = repayAmountByName[j][0];
      const repayAmount = repayAmountByName[j][1];
      if (receiveAmount === 0 || repayAmount === 0) continue;
      // 誰かに支払う分だけ残りの受け取り金額、支払い金額を減らす
      /**
       * @type {number} minAmountReveiveOrRepay 受け取る額と返済額の小さい方(正の数)
       */
      const minAmountReveiveOrRepay = Math.min(receiveAmount, Math.abs(repayAmount));
      transactions.push([`${repayName} -> ${receiveName}から${Math.round(minAmountReveiveOrRepay)}円支払う`,]);
      receiveAmountByName[i][1] -= minAmountReveiveOrRepay;
      repayAmountByName[j][1] += minAmountReveiveOrRepay;
    }
  }
  sheetToRead.getRange(`A2:F${NAME_LIST.length + 1}`).setValues(paymentForViewTable);
  if(transactions.length !== 0) sheetToRead.getRange(`A${NAME_LIST.length + 1 + 3}:A${NAME_LIST.length + 1 + 3 + transactions.length - 1}`).setValues(transactions);
}
