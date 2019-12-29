var API_TOKEN = PropertiesService.getScriptProperties().getProperty('slack_api_token');
if (!API_TOKEN) {
    throw 'You should set "slack_api_token" property from [File] > [Project properties] > [Script properties]';
}
var FOLDER_ID = PropertiesService.getScriptProperties().getProperty('folder_id');
if (!FOLDER_ID) {
    throw 'You should set "folder_id" property from [File] > [Project properties] > [Script properties]';
}
if (typeof(Drive) === 'undefined') {
    throw 'You should turn on Drive API v2 from [Resources] > [Advanced Google services...]';
}

function FindOrCreateFolder(folder, folderName) 
{
  var itr = folder.getFoldersByName(folderName);
  if( itr.hasNext() )  {
    return itr.next();
  }
  var newFolder = folder.createFolder(folderName);
  newFolder.setName(folderName);
  return newFolder;
}

function FindOrCreateSpreadsheet(folder, fileName)
{
  var it = folder.getFilesByName(fileName);
  if (it.hasNext()) {
    var file = it.next();
    return SpreadsheetApp.openById(file.getId());
  }
  else {
    var ss = SpreadsheetApp.create(fileName);
    folder.addFile(DriveApp.getFileById(ss.getId()));
    return ss;
  }
}

// Slack 上にアップロードされたデータをダウンロード
function　DownloadData(url, folder, savefilePrefix)
{
  var options = {
    "headers": {'Authorization': 'Bearer '+ API_TOKEN}
  };
  var response = UrlFetchApp.fetch(url, options);
  var fileName = savefilePrefix + "_" + url.split('/').pop();
  var fileBlob = response.getBlob().setName(fileName);
  
  console.log("Download: " + url + "\n =>" + fileName);

  // もし同名ファイルがあったら削除してから新規に作成
  var itr = folder.getFilesByName(fileName);
  if( itr.hasNext() ) {
    folder.removeFile(itr.next());
  }
  return folder.createFile(fileBlob);
}

// Slack テキスト整形
function UnescapeMessageText(text, memberList) {
  return (text || '')
  .replace(/&lt;/g, '<')
  .replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"')
  .replace(/&amp;/g, '&')
  .replace(/<@(.+?)>/g, function ($0, userID) {
    var name = memberList[userID];
    return name ? "@" + name : $0;
  });
};  
        

  
// Slack へのアクセサ
var SlackAccessor = (function() {
  function SlackAccessor(apiToken) {
    this.APIToken = apiToken;
  }
  
  var MAX_HISTORY_PAGINATION = 10;
  var HISTORY_COUNT_PER_PAGE = 1000;

  var p = SlackAccessor.prototype;
  
  // API リクエスト
  p.requestAPI = function (path, params) {
    if (params === void 0) { params = {}; }
    var url = "https://slack.com/api/" + path + "?";
    var qparams = [("token=" + encodeURIComponent(this.APIToken))];
    for (var k in params) {
      qparams.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
    }
    url += qparams.join('&');
    
    console.log("==> GET " + url);
    
    var response = UrlFetchApp.fetch(url);
    var data = JSON.parse(response.getContentText());
    if (data.error) {
      throw "GET " + path + ": " + data.error;
    }
    return data;
  };
  
  // メンバーリスト取得
  p.requestMemberList = function () {
    var response = this.requestAPI('users.list');
    var memberNames = {};
    response.members.forEach(function (member) {
      memberNames[member.id] = member.name;
      console.log("memberNames[" + member.id + "] = " + member.name);
    });
    return memberNames;
  };
  
  // チャンネル情報取得
  p.requestChannelInfo = function() {
    var response = this.requestAPI('channels.list');
    response.channels.forEach(function (channel) {
      console.log("channel(id:" + channel.id + ") = " + channel.name);
    });
    return response.channels;
  };
  
  // 特定チャンネルのメッセージ取得
  p.requestMessages = function (channel, oldest) {
    var _this = this;
    if (oldest === void 0) { oldest = '1'; }
    
    var messages = [];
    var options = {};
    options['oldest'] = oldest;
    options['count'] = HISTORY_COUNT_PER_PAGE;
    options['channel'] = channel.id;
    
    var loadChannelHistory = function (oldest) {
      if (oldest) {
        options['oldest'] = oldest;
      }
      var response = _this.requestAPI('channels.history', options);
      messages = response.messages.concat(messages);
      return response;
    };
    
    var resp = loadChannelHistory();
    var page = 1;
    while (resp.has_more && page <= MAX_HISTORY_PAGINATION) {
      resp = loadChannelHistory(resp.messages[0].ts);
      page++;
    }
    console.log("channel(id:" + channel.id + ") = " + channel.name + " => loaded messages.");
    // 最新レコードを一番下にする
    return messages.reverse();
  };
  
  return SlackAccessor;
})();


// スプレッドシートへの操作
var SpreadsheetController = (function() {
  function SpreadsheetController(spreadsheet, folder) {
    this.ss = spreadsheet;
    this.folder = folder;
  }
  
  var COL_DATE = 1; // 日付・時間(タイムスタンプから読みやすい形式にしたもの)
  var COL_USER = 2; // ユーザ名 
  var COL_TEXT = 3; // テキスト内容
  var COL_URL = 4;  // URL
  var COL_LINK = 5; // ダウンロードファイルリンク
  var COL_TIME = 6; // 差分取得用に使用するタイムスタンプ
  var COL_JSON = 7; // 念の為取得した JSON をまるごと記述しておく列
  
  var COL_MAX = COL_JSON;  // COL 最大値
  
  var COL_WIDTH_DATE = 130;
  var COL_WIDTH_TEXT = 800;
  var COL_WIDTH_URL = 400;

  var p = SpreadsheetController.prototype;
  
  // シートを探してなかったら新規追加
  p.findOrCreateSheet = function (sheetName) {
    var sheet = null;
    var sheets = this.ss.getSheets();
    sheets.forEach(function (s) {
      var name = s.getName();
      if (name == sheetName) {
        sheet = s;
        return;
      }
    });
    if (sheet == null) {
      sheet = this.ss.insertSheet();
      sheet.setName(sheetName);
      // 各 Column の幅設定
      sheet.setColumnWidth(COL_DATE, COL_WIDTH_DATE);
      sheet.setColumnWidth(COL_TEXT, COL_WIDTH_TEXT);
      sheet.setColumnWidth(COL_URL,  COL_WIDTH_URL);
    }
    return sheet;
  };
  
  // チャンネルからシート名取得
  p.channelToSheetName = function (channel) {
    return channel.name + " (" + channel.id + ")";
  };
  
  // チャンネルごとのシートを取得
  p.getChannelSheet = function (channel) {
    var sheetName = this.channelToSheetName(channel);
    return this.findOrCreateSheet(sheetName);
  };
  
  // 最後に記録したタイムスタンプ取得
  p.getLastTimestamp = function (channel) {
    var sheet = this.getChannelSheet(channel);
    var lastRow = sheet.getLastRow();
    if(lastRow > 0) {
      return sheet.getRange(lastRow, COL_TIME).getValue();
    }
    return '1';
  };
  
  // ダウンロードフォルダの確保
  p.getDownloadFolder = function (channel) {
    var sheetName = this.channelToSheetName(channel);
    return FindOrCreateFolder(this.folder, sheetName);
  };
  
  // 取得したチャンネルのメッセージを保存する
  p.saveChannelHistory = function (channel, messages, memberList) {
    console.log("saveChannelHistory: " + this.channelToSheetName(channel));
    var _this = this;
    
    var sheet = this.getChannelSheet(channel);    
    var lastRow = sheet.getLastRow();
    var currentRow = lastRow + 1;
    
    // チャンネルごとにダウンロードフォルダを用意する
    var downloadFolder = this.getDownloadFolder(channel);
    
    var record = [];
    // メッセージ内容ごとに整形してスプレッドシートに書き込み
    messages.forEach(function (msg) {
      var date = new Date(+msg.ts * 1000);
      console.log("message: " + date);
      
      var row = [];
      
      // 日付
      var date = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      row[COL_DATE - 1] = date;
      // ユーザー名
      row[COL_USER - 1] = memberList[msg.user] || msg.username;
      // Slack テキスト整形
      row[COL_TEXT - 1] = UnescapeMessageText(msg.text, memberList);
      // アップロードファイル URL とダウンロード先 Drive の Viewer リンク
      var url = "";
      var alternateLink = "";
      if(msg.upload == true) {
        url = msg.files[0].url_private_download;
        // ダウンロードとダウンロード先
        var file = DownloadData(url, downloadFolder, date);
        var driveFile = Drive.Files.get(file.getId());
        alternateLink = driveFile.alternateLink;
      }
      row[COL_URL - 1] = url;
      row[COL_LINK - 1] = alternateLink;
      row[COL_TIME - 1] = msg.ts;
      // メッセージの JSON 形式
      row[COL_JSON - 1] = JSON.stringify(msg);
      
      record.push(row);
    });
    
    if (record.length > 0)
    {
      var range = sheet.insertRowsAfter(lastRow || 1, record.length)
                    .getRange(lastRow + 1, 1, record.length, COL_MAX);
      range.setValues(record);    
    }
    
  };
   
  return SpreadsheetController;
})();

function Run()
{
  var folder = FindOrCreateFolder(DriveApp.getFolderById(FOLDER_ID), "SlackLog");
  var ss = FindOrCreateSpreadsheet(folder, "LogData");
  
  var ssCtrl = new SpreadsheetController(ss, folder);
  var slack = new SlackAccessor(API_TOKEN);

  // メンバーリスト取得
  var memberList = slack.requestMemberList();
  // チャンネル情報取得
  var channelInfo = slack.requestChannelInfo(); 

  // チャンネルごとにメッセージ内容を取得 
  channelInfo.forEach(function (ch) {
    var timestamp = ssCtrl.getLastTimestamp(ch);
    var messages = slack.requestMessages(ch, timestamp);
    
    // ファイル保存
    ssCtrl.saveChannelHistory(ch, messages, memberList);
  });
}
