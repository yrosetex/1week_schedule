/*
プログラムが起動された日の次の日から1週間分のスケジュールサマリーをメールで送る。
WJを作成するためのプログラム。
2016/9/23：一応、作成完了
2016/9/25：初運用。gmail宛てに送られたメールをコピペすると、TABがスペースになってしまう。
　　→ノーツメール宛てにしてみたが、こんどはTABがなくなる。仕方がないので、gmail宛てで当面運用。本質的にはTEXTファイル生成しかダメか？
2017/5/6：Githubでのコード管理に移行のため、非公開情報をスクリプトプロパティに移動。
　 このスクリプトに必要なプロパティの管理は別途、Dropbox/Git_Repos/GAS_Properties/1week_schedule.txtで管理する。
2020/2/23：「F仕事」など、サマリーに入れたいカレンダーが増えたので、それに対応する改造を実施。ついでにfor...of構文を使ってコードをシンプル化。
2020/2/23：iOSのショートカットからも叩けるように、function doGet()も追加。手動で動かすときは、myFunction()を実行すればよいはず。
2023/1/21：T仕事OutlookのイベントをGカレンダー取り込みすることになったことの対応。
スクリプトプロパティに「T仕事OLイベントのIDを追加」
2024/7/4：起動されたら、次の日曜日を調べて、それを起点にスケジュールサマリーを作るように改造。github assistantがうまく動かない！！
*/

function myFunction() {
  sendSchedule() 
}

function doGet() {
  sendSchedule();
  
  var html = '';
  html += '<h1>ヘッダー</h1>';
  html += '<p>実行完了：1week_schedule.txt(1週間分の予定サマリーをメールする（WJ作成用）</p>';
  html += '<p>実行日付: ' + Utilities.formatDate(new Date(),"JST","yyyy/MM/dd"); + '</p>';
  return HtmlService.createHtmlOutput(html);
}

function sendSchedule() {
  /* 起点の日付のセッティング */
  //var dt = new Date(); //日付を手動で指定する場合にはData(2024,9-1,25)と、月の数字をマイナス１する。
  var dt = new Date(getNextSundayDate()); //次の日曜日を基準日にセット
  dt.setDate (dt.getDate() + 1); //プログラムが起動された日の翌日から1週間分だから、dtを1日進める。
  var strBody = '\n' + '■来週のスケジュール' + '\n'; //メールの本文の最初の行をセット
  
  //取得するカレンダーのリストのセッティング
  //スクリプトのプロパティとして設定してあるカレンダーIDを示すプロパティ名を列挙。
  var cal_id_prop_list = ['CAL_ID_KINENBI','CAL_ID_SHARE_W_FAMILY','CAL_ID_ENGEI','CAL_ID_TEKITODO','CAL_ID_MAIN','CAL_ID_FSHIGOTO','CAL_ID_FHAJIKYOUYU','CAL_ID_CSHIGOTO','CAL_ID_T-WORK'];  

  //起点日から1週間分のデータを取得
  for(var i=0;i<14;i++){

    //日付行の生成
    var dayOfTheWeek = '日月火水木金土'[dt.getDay()];
    strBody = strBody + dayOfTheWeek + ' ' + (dt.getMonth()+1) + '月' + dt.getDate() + '日' + '\n';

    //日本の祝日カレンダのイベント取得
    var strCalList = getCalList('ja.japanese#holiday@group.v.calendar.google.com',dt);
    if (strCalList.trim() != '予定なし'){
      strBody = strBody + strCalList;
    }
   
    //配列 cal_id_list に設定したカレンダーから情報を収集。
    for (var cal_id_prop of cal_id_prop_list) {
      var calID = PropertiesService.getScriptProperties().getProperty(cal_id_prop);    
      var strCalList = getCalList(calID,dt);
      if (strCalList.trim() != '予定なし'){
        strBody = strBody + strCalList;
      }
      Utilities.sleep(200);//連続してカレンダーを呼び出すと怒られるのでsleep      
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
  var evetsInTargetCal = targetCal.getEventsForDay(dt); //カレンダーの本日のイベントを取得

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

/* 関数定義：スクリプトが起動された日の次の日曜（当日が日曜の場合は当日）を返す関数*/
/* 
  Chat GPTに作ってもらった。プロンプトは下記の通り。
    下記内容の関数を作りたい。
    本日の曜日を取得し、曜日が日曜日の場合は、本日の日付を返す。
    本日の曜日が日曜日以外の場合は、次の日曜日の日付を返す。
*/
function getNextSundayDate() {
  // 本日の日付を取得
  const today = new Date();
  
  // 曜日を取得 (0: 日曜日, 1: 月曜日, ..., 6: 土曜日)
  const dayOfWeek = today.getDay();
  
  // 日曜日の場合は本日の日付を返す
  if (dayOfWeek === 0) {
    return today;
  }
  
  // 次の日曜日までの日数を計算
  const daysUntilNextSunday = 7 - dayOfWeek;
  
  // 次の日曜日の日付を計算
  const nextSunday = new Date();
  nextSunday.setDate(today.getDate() + daysUntilNextSunday);
  
  return nextSunday;
}
