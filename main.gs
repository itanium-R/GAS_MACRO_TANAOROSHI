// シート更新時に発火
function onEdit(e){
  try{
    const ROW_ST    = 2; // 棚卸レコード開始行番号
    const COL_ST    = 2; // 棚卸レコード開始列番号
    const COL_LEN   = 5; // 棚卸レコード列長
    const LOCAT_COL = 3; // 棚卸レコード　保管場所用列番号
    const NUM_COL   = 4; // 棚卸レコード　数量用列番号
    const CHECK_COL = 5; // 棚卸レコード　1万円超チェックボックス用列番号
    const DEF_NUM   = 1; // 数量標準値
    
    var sht = e.source.getActiveSheet();
    
    // 棚卸シートでのみ実行
    if(sht.getSheetName() === "棚卸"){
      
      // レコード開始列に書き込みがあれば，数量に標準値を入力
      if(e.range.columnStart == COL_ST && e.range.columnEnd == COL_ST) {
        if(typeof(e.oldValue) == "undefined"){
          var rng = sht.getRange(e.range.rowStart, NUM_COL);
          if(rng.getValue() == "") rng.setValue(DEF_NUM);
        }
      } 
      
      // 1万円超チェックボックスにチェックが入っているものはハイライト
      if(e.range.columnStart == CHECK_COL && e.range.columnEnd == CHECK_COL) {
        var rng = sht.getRange(e.range.rowStart, COL_ST, 1, COL_LEN);
        
        if(e.value == "TRUE"){
          rng.setBackground("#AA0000");
          rng.setFontColor("#FFFFFF");
        }else{
          rng.setBackground("#FFFFFF");
          rng.setFontColor("#000000");
        }
      }
      
      // 保管場所列編集時に保管列入力候補更新
      if(e.range.columnStart == LOCAT_COL && e.range.columnEnd == LOCAT_COL) {
        var rowLen    = sht.getLastRow() - ROW_ST + 1;
        var rng       = sht.getRange(ROW_ST, LOCAT_COL, rowLen, 1);   
        var locatList = rng.getValues();
        for(var i = rowLen-1; i >= 0; i -= 1){
          locatList[i] = locatList[i][0];
        }
        
        var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(uniq(locatList), true)
        .build();  
        rng.setDataValidation(rule);
      }
    }
  }catch(e){
    Browser.msgBox(JSON.stringify(e));
  }
}

// cf) https://qiita.com/piroor/items/02885998c9f76f45bfa0
function uniq(array) {
  const knownElements = {};
  const uniquedArray = [];
  for (var i = 0, maxi = array.length; i < maxi; i++) {
    if (array[i] in knownElements)
      continue;
    uniquedArray.push(array[i]);
    knownElements[array[i]] = true;
  }
  return uniquedArray;
};
