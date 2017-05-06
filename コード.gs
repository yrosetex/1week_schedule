/*
プログラムが起動された日の次の日から1週間分のスケジュールサマリーをメールで送る。
WJを作成するためのプログラム。
2016/9/23：一応、作成完了
2016/9/25：初運用。gmail宛てに送られたメールをコピペすると、TABがスペースになってしまう。
　　→ノーツメール宛てにしてみたが、こんどはTABがなくなる。仕方がないので、gmail宛てで当面運用。本質的にはTEXTファイル生成しかダメか？
2017/5/6：Githubでのコード管理に移行のため、非公開情報をスクリプトプロパティに移動。
　 このスクリプトに必要なプロパティの管理は別途、Dropbox/Git_Repos/GAS_Properties/1week_schedule.txtで管理する。
*/

function myFunction() {
  sendSchedule() 
}

function sendSchedule() {
  /* 起点の日付のセッティング */
  var dt = new Date(); //日付を手動で指定する場合にはData(2017,9-1,1)と、月の数字をマイナス１する。
  dt.setDate (dt.getDate() + 1); //プログラムが起動された日の翌日から1週間分だから、dtを1日進める。
  var strBody = '\n' + '■来週のスケジュール' + '\n'; //メールの本文の最初の行をセット

  //起点日から1週間分のデータを取得
  for(var i=0;i<7;i++){

    //日付行の生成
    var dayOfTheWeek = '日月火水木金土'[dt.getDay()];
    strBody = strBody + dayOfTheWeek + ' ' + (dt.getMonth()+1) + '月' + dt.getDate() + '日' + '\n';

    //日本の祝日カレンダのイベント取得
    var strCalList = getCalList('ja.japanese#holiday@group.v.calendar.google.com',dt);
    if (strCalList.trim() != '予定なし'){
      strBody = strBody + strCalList;
    }

    //記念日カレンダのイベント取得
    var calID = PropertiesService.getScriptProperties().getProperty('CAL_ID_KINENBI');    
    var strCalList = getCalList(calID,dt);
    if (strCalList.trim() != '予定なし'){
      strBody = strBody + strCalList;
    }
        
    Utilities.sleep(1000);//連続してカレンダーを呼び出すと怒られるのでsleep      

    //園芸定期ToDoカレンダのイベント取得
    var calID = PropertiesService.getScriptProperties().getProperty('CAL_ID_ENGEI');    
    var strCalList = getCalList(calID,dt);
    if (strCalList.trim() != '予定なし'){
      strBody = strBody + strCalList;
    }
    
    //[定期ToDo]カレンダのイベント取得
    var calID = PropertiesService.getScriptProperties().getProperty('CAL_ID_TEKITODO');    
    var strCalList = getCalList(calID,dt);
    if (strCalList.trim() != '予定なし'){
      strBody = strBody + strCalList;
    }

    Utilities.sleep(1000);//連続してカレンダーを呼び出すと怒られるのでsleep      

    //[Main]カレンダのイベント取得
    var calID = PropertiesService.getScriptProperties().getProperty('CAL_ID_MAIN');    
    var strCalList = getCalList(calID,dt);
    if (strCalList.trim() != '予定なし'){
      strBody = strBody + strCalList;
    }
      
    dt.setDate (dt.getDate() + 1);
  }

  //Logger.log(strBody);
  var mailAddr = PropertiesService.getScriptProperties().getProperty('MAIL_ADDR');    
  MailApp.sendEmail(mailAddr, '1週間分スケジュール', strBody);
  
}

/*-------------------------------------------------------------------------------*/
/* 関数定義：時刻の表記をHH:mmに変更 */
function _HHmm(str){
  return Utilities.formatDate(str, 'JST', 'HH:mm');
}

/* 関数定義：カレンダと日付を引数で指定。スケジュールのリストを文字列で返す FIX@2016/9/23 10:47 */
function getCalList(calID,dt){
  var returnStr = '';
  var targetCal = CalendarApp.getCalendarById(calID); //指定されたIDのカレンダーを取得
  var evetsInTargetCal = targetCal.getEventsForDay(dt);　//カレンダーの本日のイベントを取得

  if (evetsInTargetCal.length == 0){
    returnStr='予定なし' + '\n';//イベントの数がゼロの場合は、「予定なし」と表示
  }else{
    for(var i=0;i<evetsInTargetCal.length;i++){
      var strTitle=evetsInTargetCal[i].getTitle(); //イベントのタイトル
      var strStart = _HHmm(evetsInTargetCal[i].getStartTime()); //イベントの開始時刻
      var strEnd = _HHmm(evetsInTargetCal[i].getEndTime()); //イベントの終了時刻
        if (strStart == strEnd){
          returnStr = returnStr + ' \t' +strTitle + '\n';
        }else{
          returnStr = returnStr + ' \t' + strStart + ' - ' + strEnd + '\t' +strTitle + '\n';
        }
    }
    }
return returnStr;  
}
