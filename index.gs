/* 角色的物件 */
function person(name,attend,time,role,bopae,ring,ear,organ,weapon,remark) {
  this.name = name;
  this.attend = attend;
  this.time = time;
  this.role = role;
  this.bopae = bopae;
  this.ring = ring;
  this.ear = ear;
  this.organ = organ;
  this.weapon = weapon;
  this.remark = remark
}


function assign() {  
  var dataSS = SpreadsheetApp.getActiveSpreadsheet();
  //轉到目標表單
  var data = dataSS.getSheetByName("表單回應 1");
  //var showSheet = dataSS.getActiveSheet();
  //抓取資料
  data = data.getDataRange().getValues();
  //篩選出不重複的資料(抓最後一筆)
  data = uniqueArray(data);
  //找出參加的人
  data = findAttend(data);
  /*
  根據參加人數判定可以開幾團
  (綜合火力判定及人數判定)
  ※人數判定方針
  最適團數 = Math.floor(人數/10)
  
  ※火力判定方針
  武器,魂,攻擊力
  最適火力 = 坦+遠坦+指揮*1
  小降(刺.拳)*2 降(咒)*2 指揮(槍)*2
  */
  Logger.log("參加人數: " + data.length);
  
  var fire = 1; //簡易火力總計算 (分析最低火力限制)
  
  // 分析可以開幾團 if num_
  var best_group_num = Math.floor(data.length / 10); // 理想團數
  // 分析開在哪天
  /*
  1.分析每天可打人數
  2.分析只能打一天的人數
  3.將最少人可打的那天補齊
  How to 補齊?
  
  */
  // 分析每團人數
  // var groupMax = data.length / best_group_num;
  // 開團 團數為 best_group_num
  var groups = []; //開團
  // 
  
  
  
  
  
  
  
  
  
  
  
  
  setDefault();
  for (var x = 2; x <= 5;x++) {
    outPut(sortData(getData(data,"A" + x)),x-1);
  }
  for (var x =2; x<=11;x++) {
    outPut(getData(data,"D" + x),x+4);
  }
  Logger.getLog();
}

/* 新統計專用 */
function statics() {
  var runtimeCountStart = new Date(); //開始記錄runtime
  
  
  //定義目標
  var target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("新統計");
  
  var dataSS = SpreadsheetApp.getActiveSpreadsheet();
  //轉到目標表單
  var data = dataSS.getSheetByName("表單回應 1");
  //抓取資料
  data = data.getDataRange().getValues();
  //篩選出不重複的資料(抓最後一筆)
  data = uniqueArray(data);
  var total_num = data.length;
  //找出參加的人
  data = findAttend(data);
  var attend_num = data.length;
  //按照 星期分團
  var group1 = [];
  var group2 = [];
  var group3 = [];
  
  for (n=0;n<data.length;n++) {
    if (data[n][3].indexOf("星期五") != -1) group1.push(data[n]);
    if (data[n][3].indexOf("星期六") != -1) group2.push(data[n]);
    if (data[n][3].indexOf("星期日晚上") != -1) group3.push(data[n]);
  }
  
  /* 重設 */
  resetStat(target);
  
  /* 參加統計 */
  attendStat(target,group1,2);
  attendStat(target,group2,3);
  attendStat(target,group3,4);
  target.getRange("B5").setValue(total_num - attend_num);
  target.getRange("B6").setValue(total_num);
  
  /* 職業統計 */
  roleStat(target,group1,"E");
  roleStat(target,group2,"F");
  roleStat(target,group3,"G");
  
  /* 八卦統計 */
  bopaeStat(target,group1,0);
  bopaeStat(target,group2,1);
  bopaeStat(target,group3,2);
  
  /* 飾品統計 */
  accStat(target,group1,0);
  accStat(target,group2,1);
  accStat(target,group3,2);
  
  /* 器官統計 */
  organStat(target,group1,0);
  organStat(target,group2,1);
  organStat(target,group3,2);
  
  /* 火力統計 */
  fireStat(target,group1,0);
  fireStat(target,group2,1);
  fireStat(target,group3,2);
  
  
  //SpreadsheetApp.getUi().alert("資料已刷新完成" + "\n運行時間: " + runtimeCountStop(runtimeCountStart) + " 秒");
  
  Logger.log("資料已刷新完成" + "\n運行時間: " + runtimeCountStop(runtimeCountStart) + " 秒");
  Logger.getLog();

}

/* 重設統計 */
function resetStat(target) {
  var temp = ["B2:B6","E2:G11","J3:AI12","AL3:AS8","AV2:AZ4","AV7:AZ11"]; //存有變數資料的範圍
  for (n=0;n<temp.length;n++) {
    target.getRange(temp[n]).clear({commentsOnly: true, contentsOnly: true}); //對資料做清除
  }
}

/* 參加統計 */
function attendStat(target,data,k) {
  var item = target.getRange("B"+k);
  var tempText = "";
  item.setValue(data.length);
  for (n=0;n<data.length;n++) {
    tempText += (data[n][1] + "\n");
  }
  item.setNote(tempText);
}

/* 職業統計 */
function roleStat(target,data,k) {
  for (p=2;p<=11;p++) {
    var count = 0;
    var item = target.getRange("D"+p).getValue();
    var tempText = "";
    for (n=0;n<data.length;n++) {
      if(item == data[n][4]) {
        tempText += (data[n][1] + "\n");
        count++;
      }
    }
    target.getRange(k+p).setValue(count);
    target.getRange(k+p).setNote(tempText);
  }
}

/* 八卦統計 */
function bopaeStat(target,data,k) {
  for (p=3;p<=12;p++) { //職業
    for (b=1;b<=8;b++) { //八卦號碼
      var count = 0;
      var item = target.getRange("I"+p).getValue();
      var tempText = "";
      for (n=0;n<data.length;n++) {
        if (item == data[n][4] && data[n][5].toString().indexOf(b) != -1) {
          tempText += (data[n][1] + "\n");
          count++;
        }
      }
      if (k<=1) {
        target.getRange(String.fromCharCode("I".charCodeAt(0)+b+(k*9))+p).setValue(count);
        target.getRange(String.fromCharCode("I".charCodeAt(0)+b+(k*9))+p).setNote(tempText);
      } else if(k == 2) {
        target.getRange("A" + String.fromCharCode("A".charCodeAt(0)+b) + p).setValue(count);
        target.getRange("A" + String.fromCharCode("A".charCodeAt(0)+b) + p).setNote(tempText);
      }
    }
  }
}

/* 飾品統計 */
/* 此欄位資料有被 "切開" 因為經過 Z 和 AA 之間 */
function accStat(target,data,k) {
  //戒指
  for (l=3;l<=8;l++) { //飾品類型
    var item = target.getRange("AK" + l).getValue();
    var count = 0;
    var tempText = "";
    for (n=0;n<data.length;n++) { //走訪資料
      if (item == data[n][6]) {
        tempText += (data[n][1] +"\n");
        count++;
      }
    }
    target.getRange("A" + String.fromCharCode("K".charCodeAt(0) + 1 + (k*3)) + l).setValue(count);
    target.getRange("A" + String.fromCharCode("K".charCodeAt(0) + 1 + (k*3)) + l).setNote(tempText);
  }
  //耳環
  for (l=3;l<=8;l++) { //飾品類型
    var item = target.getRange("AK" + l).getValue();
    var count = 0;
    var tempText = "";
    for (n=0;n<data.length;n++) { //走訪資料
      if (item == data[n][7]) {
        tempText += (data[n][1] +"\n");
        count++;
      }
    }
    target.getRange("A" + String.fromCharCode("K".charCodeAt(0) + 2 + (k*3)) + l).setValue(count);
    target.getRange("A" + String.fromCharCode("K".charCodeAt(0) + 2 + (k*3)) + l).setNote(tempText);
  }
}

/* 器官統計 */
function organStat(target,data,k) {
  for (l=2;l<=4;l++) {
    var count = 0;
    var item = target.getRange("AU" + l).getValue();
    var tempText = "";
    for (n=0;n<data.length;n++) {
      if (item == data[n][8]) {
        tempText += (data[n][1] + "\n");
        count++;
      }
    }
    target.getRange("A" + String.fromCharCode("U".charCodeAt(0) + 1 +(k*2)) +l).setValue(count);
    target.getRange("A" + String.fromCharCode("U".charCodeAt(0) + 1 +(k*2)) + l).setNote(tempText);
  }
}

/* 火力統計 */
function fireStat(target,data,k) {
  for(l=7;l<=11;l++) {
    var count = 0;
    var item = target.getRange("AU" + l).getValue();
    var tempText = "";
    for (n=0;n<data.length;n++) {
      if (data[n][9].indexOf(item) != -1) {
        tempText += (data[n][1] + "\n");
        count++;
      }
    }
    target.getRange("A" + String.fromCharCode("U".charCodeAt(0) + 1 +(k*2)) +l).setValue(count);
    target.getRange("A" + String.fromCharCode("U".charCodeAt(0) + 1 +(k*2)) + l).setNote(tempText);
  }
}





function uniqueArray(array){
  array = array.reverse();
  var item = [];
  for(var i = 0; i < array.length; i++){
    if (!(isItemInArray(item,array[i]))) item.push(array[i]);
  }
  return item.reverse();
}

function isItemInArray(array, item) {
  for (var i = 0; i < array.length; i++) {
    if (array[i][1] == item[1]) return true;
  }
  return false;
}

function findAttend(array) {
  var item = [];
  for(n=1;n<array.length;++n){
    if (array[n][2] != "不參加") item.push(array[n]);
  }
  return item;
}


/* 
D 劍士2  拳士3  力士4  刺客5  氣功士6  燐劍士7  召喚師8  咒術師9  乾坤士10  槍擊士11
A 星期五2 星期六3 星期4 
*/
function showData(data,item) {
  var target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統計").getRange(item).getValue();
  for (n=0;n<data.length;n++) {
    if (item.indexOf("A") != -1) {
      if (data[n][3].indexOf(target) != -1) Logger.log(data[n][1] + data[n][4] + " "+ data[n][3]);
    } else if (item.indexOf("D") != -1) {
      if (data[n][4].indexOf(target) == 0) Logger.log(data[n][1] + data[n][4]);
    }
  }
}

function getData(data,item) {
  var target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("統計").getRange(item).getValue();
  var newData = [];
  for (n=0;n<data.length;n++) {
    if (item.indexOf("A") != -1) {
      if (data[n][3].indexOf(target) != -1) newData.push(data[n]);
    } else if (item.indexOf("D") != -1) {
      if (data[n][4].indexOf(target) == 0) newData.push(data[n]);
    }
  }
  return newData;
}

function setDefault() {
  var target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("團表");
  //清空
  var temp = ["A2:D100","F2:O100"];
  for (n=0;n<temp.length;n++) {
    target.getRange(temp[n]).clear({commentsOnly: true, contentsOnly: true});
  }
  target.getRange(temp[0]).setFontColor('Black');
  
  //target.getDataRange().clearContent();
  //預設
  /*
  target.getRange("A1").setValue("星期五");
  target.getRange("B1").setValue("星期六");
  target.getRange("C1").setValue("星期日下午");
  target.getRange("D1").setValue("星期日晚上");
  燐劍士	召喚師	咒術師	乾坤士	槍擊士
  target.getRange("E1").setValue("劍士");
  target.getRange("F1").setValue("拳士");
  target.getRange("G1").setValue("力士");
  target.getRange("H1").setValue("刺客");
  target.getRange("I1").setValue("氣功士");
  */
}

/* 根據可打天數排序 */
function sortData(data) {
  var newData = [];
  var max = 0;
  /* 找出 length 的最大值*/
  for (n=0;n<data.length;n++) {
    if (data[n][3].split(',').length > max) {
      max = data[n][3].split(',').length;
    }
  }
  /* 迴圈 1~max */
  for (p=1;p<=max;p++) {
    /* 走訪全部資料 */
    for (n=0;n<data.length;n++) {
      //符合就 push
      if (data[n][3].split(',').length == p) newData.push(data[n]);
    }
  }
  return newData;
}

/* 先輸出到 spreadsheet */
function outPut(data, column) {
  var target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("團表");
  if (column <= 4) {
    /* 左邊 */
    for (x=0;x<data.length;x++) {
      target.getRange(2+x, column).setValue(data[x][1]);
      target.getRange(2+x, column).setNote("職業: " + data[x][4] + "\n八卦: " + data[x][5] + "\n器官: " + data[x][8] + "\n戒指: " + data[x][6] + "\n耳環: " + data[x][7] + "\n武器: " + data[x][9] + "\n備註: " + data[x][10]);
    }
  } else {
    /* 右邊 */
    for (x=0;x<data.length;x++) {
      target.getRange(2+x, column).setValue(data[x][1]);
      target.getRange(2+x, column).setNote("時間: " + data[x][3]);
    }
  }
}


function runtimeCountStop(start) {
  var stop = new Date();
  var newRuntime = Number(stop) - Number(start);
  return newRuntime / 1000 ; //單位秒
}