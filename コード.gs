const SHEET_NAME_USER_LIST = "名前"
const SHEET_NAME_TO_READ = "支払い(記載)"
const SHEET_NAME_TO_WRITE = "支払い金額(OnlyRead)"
const userNameSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USER_LIST);
const NAME_LIST = userNameSheet.getRange(`A2:A${userNameSheet.getLastRow()}`).getValues().flat();

function alphaToNum(alphabet){
  if(typeof(alphabet) !== "string"  && alphabet !== alphabet.toUpperCase()) return null; // 文字列でない or 大文字でない => null

  let initalValue = 0, idx = 0;
  return alphabet.split('').reverse().reduce( // 26進数(AA = 27)のように扱い、末尾の文字をidx=0にするためreverse()を使用する
    (accumulator, currentValue) => {
    let code = currentValue.charCodeAt(0);
    return accumulator + (code - 64)*(26**idx++); // (アルファベットの数字表現)*(26進数における位による計算), idxを次のループ用にインクリメントする必要があるが、return で返しているので、後置インクリメントを使用する
    },initalValue
  );
}



function resetSheet(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_READ);
  // 初期化
  sheet.getRange("A2:Z100").clearDataValidations();
  sheet.getRange("A2:Z100").clearContent();
  // 名前部分のリセット
  sheet.getRange("F1:Z1").clear()
  sheet.getRange(1,6,1,NAME_LIST.length).setValues([NAME_LIST]);
  // 数式の更新
  for(let row=2; row<=100;row++){
    const cell = sheet.getRange(`D${row}`);
    cell.setFormula(`=IF(B${row}="全員",C${row}/${NAME_LIST.length},(C${row}/COUNTA(SPLIT(B${row},","))))`);
  }
  // プルダウンの再生成
  resetPullDown();
  // SHEET_NAME_TO_WRITEの再生成
  resetTotalToWrite();
}
function resetPullDown(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_READ);
  // 支払い者のプルダウンの作成
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(NAME_LIST).build();
  let cell = sheet.getRange('A2:A100');
  cell.setDataValidation(rule);
  // 対象者のプルダウンの作成
  const targetValue = ["全員"].concat(NAME_LIST);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(targetValue).build();
  cell = sheet.getRange('B2:B100');
  cell.setDataValidation(rule);
}

function resetTotalToWrite(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_WRITE);
  // 初期化
  sheet.getRange(`A2:G${sheet.getLastRow()}`).clear()
  // 各名前ごとの合計を枠組みを作成する
  let arrayToWriteForEachName = [];
  NAME_LIST.forEach((name, idx) => {
    arrayToWriteForEachName.push([name,"",""]); // 誰用なのかを示す行
    NAME_LIST.forEach(name2 => {
      if(name2 === name) return;
      arrayToWriteForEachName.push(["",name2,0]); // 誰から受け取るのかの名前を示す行
    })
    // 必要な情報を追加する(常に同じ位置であるため後から追加する)
    arrayToWriteForEachName[idx*NAME_LIST.length+1][0] = "受け取り金額(残り)";
    arrayToWriteForEachName[idx*NAME_LIST.length+2][0] = 0;
    sheet.getRange(`A2:C${arrayToWriteForEachName.length + 1}`).setValues(arrayToWriteForEachName)// 2(初期値)+配列の長さ-1(2行目から始まっているため)
  })
  // 全合計を枠組みを作成する
  sheet.getRange(`F2:G${NAME_LIST.length + 1}`).setValues(NAME_LIST.map(name=>[name,0])) // 2(初期値)+NAME_LIST.length-1(2行目から始まっているため)
  // 枠線の追加
  sheet.getRange(`F2:G${NAME_LIST.length + 1}`).setBorder(true, true, true, true, true, true);
}

function onEdit(e){
  const TARGET_COL_IDX = [alphaToNum("A"),alphaToNum("B")]
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const target = sheet.getRange(`B${range.getRow()}`).getValue(); // 対象者を取得
  const sheetName = sheet.getName();

  // 対象のシートの編集かどうか
  if(sheetName !== SHEET_NAME_TO_READ) return;
  // 複数範囲の時は除外
  if(range.getNumRows() !== 1 && range.getNumColumns() !== 1) return; 
  // 内容を変更する関数の実行
  if(TARGET_COL_IDX.includes(range.getColumn())) changeCheckBox(sheet, range, target);
  calculationMoney();
}


function changeCheckBox(sheet, range, target){
  if(target == "") return;
  let row = range.getRow();
  let check_box_range = sheet.getRange(row, 6, 1, NAME_LIST.length) // F列からNAME_LISTの長さの範囲
  let save_origin_checkbox_data = check_box_range.getValues();
  // 初期化
  check_box_range.clearDataValidations();
  check_box_range.clearContent();
  if(target === undefined)return;
  let name_list = target.split(",").map((name) => name.trim());

  for(var i=0; i < name_list.length; i++){
    let name = name_list[i];
    if(name == "全員"){
      check_box_range.insertCheckboxes();
      break;
    }else{
      let index = NAME_LIST.indexOf(name);
      sheet.getRange(row,6+index).insertCheckboxes(); // 6+index => "F"+index
    }
  }
  // 元の値をsetする(その部分がcheckboxの時のみ)
  for(let row_idx = 0; row_idx < save_origin_checkbox_data.length;row_idx++){
    for(let col_idx = 0; col_idx < save_origin_checkbox_data[row_idx].length;col_idx++){
      let range = sheet.getRange(row+row_idx,6+col_idx)
      if(range.getValue() !== "" && save_origin_checkbox_data[row_idx][col_idx]){ // checkboxである and 元の値がTureである時
        range.setValue("True");
      }
    }
  }
  let payer = sheet.getRange(`A${row}`).getValue().trim();
  if(name_list.includes(payer) || name_list.includes("全員")){
    let index = NAME_LIST.indexOf(payer);
    if(index == -1) return;
    sheet.getRange(row,6+index).setValue("True");
  }
}


function calculationMoney(){
  // {立替者:{支払い者:[0,0],...}}このような辞書を作る([0,0] = [支払い済み, 支払い合計], 立替者 = 支払い者の場合も作る)
  let payDictByName = {}
  for(let i = 0; i<NAME_LIST.length; i++){
    let name = NAME_LIST[i];
    payDictByName[name] = {};
    for(let j = 0; j<NAME_LIST.length; j++){
      name_2 = NAME_LIST[j];
      payDictByName[name][name_2] = [0,0]
    }
  }
  
  const sheetToRead = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_READ);
  const sheetToWrite = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TO_WRITE);
  const range = sheetToRead.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN);
  const lastRow = range.getLastRow();
  const values = sheetToRead.getRange(2,1,lastRow-1,alphaToNum("E")+NAME_LIST.length).getValues();
  for(let row_idx = 0; row_idx < values.length; row_idx++){
    const payer = values[row_idx][0];
    let targets = values[row_idx][1].split(",");
    const amount = values[row_idx][2];
    const amountPerPerson = values[row_idx][3];
    if(targets == "全員"){
     targets = NAME_LIST; 
    }else if (targets == "") continue;
    // 自分に足す
    payDictByName[payer][payer][0] += amountPerPerson
    payDictByName[payer][payer][0] += amountPerPerson
    // 合計を求める
    for(let i=0; i < targets.length; i++){
      let name = targets[i].trim();
      let name_index = NAME_LIST.indexOf(name);
      let checkbox_value = values[row_idx][5 + name_index] // 5はcheckboxの始まるIndex
      payDictByName[payer][name][1] += amountPerPerson
      if(checkbox_value){
        payDictByName[payer][name][0] += amountPerPerson
      }
    }
  }
  // 記述する
  let paymentTotalByEach=[]
  let paymentTotalByName=NAME_LIST.map(()=>[0])
  Object.keys(payDictByName).forEach(payer => {
    paymentTotalByEach.push([""]); // 空白を作る
    let i=0;
    Object.keys(payDictByName[payer]).forEach(target => {
      if(payer !== target){
        let payment = payDictByName[payer][target][1] - payDictByName[payer][target][0];
        let payment_reverse = payDictByName[target][payer][1] - payDictByName[target][payer][0];
        if(payment >= payment_reverse){ // 支払い金額が多い人の方に合算する
          paymentTotalByEach.push([payment- payment_reverse]);
        }else{
          paymentTotalByEach.push([0]); // 支払いはない
        }
      }
      paymentTotalByName[i][0] += payDictByName[payer][target][1];
      i++;
    });
  });
  sheetToWrite.getRange(`C2:C${(NAME_LIST.length**2)+1}`).setValues(paymentTotalByEach);
  sheetToWrite.getRange(`G2:G${NAME_LIST.length+1}`).setValues(paymentTotalByName);
}