// 読み込み・更新対象の全シート名
const TARGET_SHEETS = [
  "hololive_original_song",
  "hololive_cover_song",
  "LocalVoid_weekly_best10",
  "VTuber_original_song",
  "anime_song",
  "General"
];

const NICONICO_SHEET_NAME = 'VOCALOID_niconico';

/**
 * 1. データの取得（GET）: 全シートの有効データを統合して返却
 */
function assertAllowedUser_() {
  const email = String(Session.getActiveUser().getEmail() || "")
    .trim()
    .toLowerCase();
  if (!email) {
    throw new Error("Google account authentication is required");
  }

  const configured =
    PropertiesService.getScriptProperties()
      .getProperty("ALLOWED_USER_EMAILS") || "";
  const allowedEmails = configured
    .split(",")
    .map(value => value.trim().toLowerCase())
    .filter(Boolean);

  if (!allowedEmails.includes(email)) {
    throw new Error("This Google account is not authorized");
  }
  return email;
}

function doGet(e) {
  const callback = String(e && e.parameter && e.parameter.callback || "");
  try {
    const response = doGetCore_(e);
    if (!callback) return response;
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      throw new Error("Invalid JSONP callback");
    }
    const payload = JSON.parse(response.getContent());
    if (Array.isArray(payload.data)) {
      payload.data.forEach(item => {
        if (item && typeof item === "object") delete item.comment;
      });
    }
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (err) {
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      const errorPayload = { status: "error", message: err.message || String(err) };
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(errorPayload) + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGetCore_(e) {
  assertAllowedUser_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const params = (e && e.parameter) || {};
  const action = params.action || "";

  // 1. 更新系アクション（JSONP / GET経由）のハンドリング
  if (action === "updateVideoId") {
    const youtubeId = String(params.youtubeId || "").trim();
    if (!youtubeId) throw new Error("youtubeId is required");
    return updateSongRow_(params, 4, youtubeId);
  }

  if (action === "saveCustomStartSec" || action === "saveStartSec") {
    const rawSec = params.customStartSec;
    let customStartSec = "";
    if (rawSec !== null && rawSec !== "" && rawSec !== undefined) {
      customStartSec = Number(rawSec);
      if (!Number.isFinite(customStartSec) || customStartSec < 0) {
        throw new Error("customStartSec must be a non-negative number or null");
      }
    }
    return updateSongRow_(params, 15, customStartSec);
  }

  if (action === "saveComment") {
    return updateSongRow_(params, 6, String(params.comment || ""));
  }

  if (action === "disableRow") {
    return updateSongRow_(params, 1, false);
  }

  if (action === "syncPlaylistAndGetQuiz") {
    const playlistId = params.playlistId;
    const genreTags = params.genreTags || "";
    const resultData = syncPlaylistToGeneralSheet(playlistId, genreTags);
    Logger.log("Result Count: " + resultData.length);
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: resultData
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "importPlaylist") {
    const result = importPlaylistToGeneralSheet(params.playlistId, params.genreTags || "");
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      result: result
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 2. 既存の取得系処理
  // 存在するすべてのシート名を取得
  const allSheetNames = sheets.map(sheet => sheet.getName());

  // クエリパラメータから取得対象のシート名リスト（カンマ区切り）を取得
  const requestedSheetsStr = params.sheets || "";
  const requestedSheets = requestedSheetsStr ? requestedSheetsStr.split(",") : [];

    /*
   * ===================================================
   * NicoNicoクイズ専用エンドポイント
   * ===================================================
   *
   * YouTubeのシート取得処理とは完全に分離。
   */
  
  if (action === 'niconicoQuiz') {

    try {

      const data =
        getNiconicoQuizData_();

      return ContentService
        .createTextOutput(
          JSON.stringify({
            status: 'success',
            source: 'niconico',
            sheetName: NICONICO_SHEET_NAME,
            count: data.length,
            data: data
          })
        )
        .setMimeType(
          ContentService.MimeType.JSON
        );

    } catch (error) {

      Logger.log(
        '[NicoNico ERROR] ' +
        error.message
      );

      return ContentService
        .createTextOutput(
          JSON.stringify({
            status: 'error',
            source: 'niconico',
            message:
              String(
                error.message ||
                error
              )
          })
        )
        .setMimeType(
          ContentService.MimeType.JSON
        );
    }
  }

  // requestedSheets が空の場合は、シート名一覧のみを高速に返却
  if (requestedSheets.length === 0 && action !== "fetchPlaylistItems") {
    return createJsonResponse({
      sheetNames: allSheetNames,
      data: []
    });
  }

  // 再生リストの直接取得リクエスト処理
  if (action === "fetchPlaylistItems") {
    const playlistId = e.parameter.playlistId;
    const items = getPlaylistItemsDirect(playlistId);
    return ContentService.createTextOutput(JSON.stringify({
      status: items.length > 0 ? "success" : "error",
      data: items
    })).setMimeType(ContentService.MimeType.JSON);
  }

  let fetchedData = [];

  // 指定されたシートのみデータを読み込む
  requestedSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return; // ヘッダーのみ、または空シートはスキップ

    // ヘッダー（1行目）を除外してデータオブジェクト化
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const isEnabled = row[0]; // A列: 有効フラグ

      if (isEnabled === true || isEnabled === "TRUE" || isEnabled === 1) {
        fetchedData.push({
          sheetName: sheetName,
          rowIndex: i + 1,
          artist: row[1] || "",
          publishedAt: row[2] || "",
          youtubeId: row[3] || "",
          videoIdList: row[4] ? String(row[4]).split(",").map(id => id.trim()) : [row[3]],
          comment: row[5] || "",
          videoTitle: row[6] || "",
          description: row[7] || "",
          views: Number(row[8] || 0),
          likes: Number(row[9] || 0),
          postDate: row[10] instanceof Date && !isNaN(row[10].getTime()) ? Utilities.formatDate(row[10], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd") : String(row[10] || "").slice(0, 10),
          duration: row[11] || "",
          channelId: row[12] || "",
          channelName: row[13] || "",
          customStartSec: row[14] !== "" && row[14] !== null ? Number(row[14]) : null,
          tags: row[15] ? String(row[15]).split(",").map(t => t.trim()) : []
        });
      }
    }
  });

  return createJsonResponse({
    sheetNames: allSheetNames,
    data: fetchedData
  });
}

/**
 * 2. データの更新（POST）: 特定シートの指定行を書き換え、またはプレイリストの取り込み
 */
function updateSongRow_(params, column, value) {
  const sheetName = String(params.sheetName || "").trim();
  const rowIndex = Number(params.rowIndex);

  if (!sheetName) {
    throw new Error("sheetName is required");
  }
  if (!Number.isInteger(rowIndex) || rowIndex < 2) {
    throw new Error("rowIndex must be a sheet row number >= 2");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Sheet not found: " + sheetName);
  }
  if (rowIndex > sheet.getLastRow()) {
    throw new Error("rowIndex is outside the sheet data");
  }

  sheet.getRange(rowIndex, column).setValue(value);
  return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    // 1. リクエスト内容をログに記録

    
    assertAllowedUser_(); const params = JSON.parse(e.postData.contents);
    assertAllowedUser_();
    const action = params.action;

    Logger.log("Action: " + action + ", PlaylistId: " + params.playlistId);

    // 画面からの単項目更新（代表ID=D列、指定秒数=O列、コメント=F列）
    if (action === "disableRow") {
      return updateSongRow_(params, 1, false);
    }

    // 画面側の通常保存リクエストにはactionが無いため、送信項目で判定する。
    if (!action) {
      if (Object.prototype.hasOwnProperty.call(params, "youtubeId")) {
        const youtubeId = String(params.youtubeId || "").trim();
        if (!youtubeId) throw new Error("youtubeId is required");
        return updateSongRow_(params, 4, youtubeId);
      }

      if (Object.prototype.hasOwnProperty.call(params, "customStartSec")) {
        const rawSec = params.customStartSec;
        let customStartSec = "";
        if (rawSec !== null && rawSec !== "") {
          customStartSec = Number(rawSec);
          if (!Number.isFinite(customStartSec) || customStartSec < 0) {
            throw new Error("customStartSec must be a non-negative number or null");
          }
        }
        return updateSongRow_(params, 15, customStartSec);
      }

      if (Object.prototype.hasOwnProperty.call(params, "comment")) {
        return updateSongRow_(params, 6, String(params.comment || ""));
      }

      throw new Error("Unsupported POST request");
    }

    // 再生リスト同期＆出題データ取得処理
    if (action === "syncPlaylistAndGetQuiz") {
      const playlistId = params.playlistId;
      const genreTags = params.genreTags || "";
      
      const resultData = syncPlaylistToGeneralSheet(playlistId, genreTags);
      Logger.log("Result Count: " + resultData.length);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        data: resultData
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "importPlaylist") {
      const result = importPlaylistToGeneralSheet(params.playlistId, params.genreTags || "");
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        result: result
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // アクションが一致しなかった場合
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Unknown action: " + action
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("doPost Error: " + err.message + "\nStack: " + err.stack);
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 3. プレイリストからの個別インポート（ジャンルタグ書き込み対応）
 */
function importPlaylistToSheet(playlistId, targetSheetName, genreTags) {
  return saveNewSongsFromPlaylistToDB(playlistId, targetSheetName, genreTags);
}

/**
 * 4. プレイリストの全曲を全DB（TARGET_SHEETS）と照合し、未登録の曲のみをDBへ追記
 */
function saveNewSongsFromPlaylistToDB(playlistId, targetSheetName, genreTags) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingVideoIds = new Set();

  // ① 全シートから既存の YouTube ID を抽出してセット化（重複チェック用）
  TARGET_SHEETS.forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3]) existingVideoIds.add(String(data[i][3]).trim());
    }
  });

  // 対象シートを取得（存在しなければ General シート）
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.getSheetByName("General");
  }
  if (!targetSheet) {
    throw new Error("追加先のシートが見つかりません: " + targetSheetName);
  }

  let addedCount = 0;
  let skippedCount = 0;
  let nextPageToken = "";

  // デフォルトタグの設定（入力タグが無ければデフォルト値）
  const tagToSave = genreTags ? genreTags : "プレイリスト追加";

  // ② ページネーションでプレイリストの全アイテムを巡回
  do {
    const playlistResponse = YouTube.PlaylistItems.list('snippet,contentDetails', {
      playlistId: playlistId,
      maxResults: 50,
      pageToken: nextPageToken
    });

    const items = playlistResponse.items || [];
    for (const item of items) {
      const videoId = item.contentDetails.videoId;

      if (existingVideoIds.has(videoId)) {
        skippedCount++;
      } else {
        // 未登録曲を末尾へ追記
        const publishedAt = item.contentDetails && item.contentDetails.videoPublishedAt ? item.contentDetails.videoPublishedAt.split("T")[0] : "";
        
        targetSheet.appendRow([
          true,                                         // A: 有効
          item.snippet.videoOwnerChannelTitle || "",    // B: 歌唱者
          "",           // C: 音源公開日
          videoId,                                      // D: 代表ID
          videoId,                                      // E: IDリスト
          "プレイリストから自動登録",                     // F: メモ
          item.snippet.title || "",                     // G: タイトル
          item.snippet.description || "",               // H: 説明文
          0, 0,                                         // I, J: 再生数, いいね数
          publishedAt,                                  // K: 投稿日
          "", "", "",                                   // L, M, N: 尺, ChID, Ch名
          null,                                         // O: ラントロ秒
          tagToSave                                     // P: タグ（指定ジャンルタグを書き込み）
        ]);

        existingVideoIds.add(videoId);
        addedCount++;
      }
    }

    nextPageToken = playlistResponse.nextPageToken;
  } while (nextPageToken);

  return { addedCount: addedCount, skippedCount: skippedCount };
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ==========================================
// 設定項目
// ==========================================

// YouTube APIキー（updateYouTubeDataのUrlFetchAppで使用）
function getRequiredScriptProperty_(name) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) {
    throw new Error('スクリプトプロパティが未設定です: ' + name);
  }
  return value;
}

// プレイリストID と 追加先シート名 のマッピング設定
const PLAYLIST_MAP = [
  { playlistId: 'PLZT1qIN_o6E4', targetSheetName: 'hololive_original_song' },
  { playlistId: 'PLMPMTEHaOUg4', targetSheetName: 'hololive_cover_song' }
];


// ==========================================
// 定期実行・手動バッチ処理関数
// ==========================================

/**
 * 全シートのD列（動画ID）を巡回し、YouTube APIから最新の統計情報（G〜N列）を一括更新します。
 */
function updateYouTubeData() {
  const apiKey = getRequiredScriptProperty_('YOUTUBE_API_KEY');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = new Date().getTime(); // 開始時刻を記録（4分30秒制限用）

  for (const sheetName of TARGET_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue; // ヘッダーのみの場合はスキップ

    // D列（4列目）の動画IDを取得
    const idRange = sheet.getRange(2, 4, lastRow - 1, 1);
    const rawVideoIds = idRange.getValues().flatMap(row => String(row[0]).trim());

    const validIds = [...new Set(rawVideoIds.filter(id => id !== ""))];
    if (validIds.length === 0) continue;

    const dataMap = {};
    const chunkSize = 50;

    // 50件ずつAPIへリクエスト
    for (let i = 0; i < validIds.length; i += chunkSize) {
      if (new Date().getTime() - startTime > 270000) {
        Logger.log(`実行時間制限に達したため、${sheetName} の処理途中で中断しました。次回実行時に続きが処理されます。`);
        return;
      }

      const chunk = validIds.slice(i, i + chunkSize);
      const url = `https://www.googleapis.com/youtube/v3/videos?part=snippet,statistics,contentDetails&id=${chunk.join(',')}&key=${encodeURIComponent(apiKey)}`;
      
      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const json = JSON.parse(response.getContentText());
        
        if (json.items) {
          json.items.forEach(item => {
            dataMap[item.id] = {
              title: item.snippet.title || '',
              description: item.snippet.description || '',
              viewCount: item.statistics.viewCount ? Number(item.statistics.viewCount) : 0,
              likeCount: item.statistics.likeCount ? Number(item.statistics.likeCount) : 0,
              publishedAt: item.snippet.publishedAt ? "" : '',
              duration: formatIsoDuration(item.contentDetails.duration),
              channelId: item.snippet.channelId || '',
              channelTitle: item.snippet.channelTitle || ''
            };
          });
        }
      } catch (e) {
        Logger.log(`API取得エラー (${sheetName}): ` + e.toString());
      }
    }

    // 元のD列の行順に合わせて書き込み用配列を作成
    const results = rawVideoIds.map(id => {
      if (!id) {
        return ['', '', '', '', '', '', '', ''];
      } else if (dataMap[id]) {
        const d = dataMap[id];
        return [
          d.title,
          d.description,
          d.viewCount,
          d.likeCount,
          d.publishedAt,
          d.duration,
          d.channelId,
          d.channelTitle
        ];
      } else {
        return null;
      }
    });

    const sheetValues = sheet.getRange(2, 7, results.length, 8).getValues();
    for (let i = 0; i < results.length; i++) {
      if (results[i] !== null) {
        sheetValues[i] = results[i];
      }
    }

    if (sheetValues.length > 0) {
      sheet.getRange(2, 7, sheetValues.length, 8).setValues(sheetValues);
      Logger.log(`${sheetName} シートの動画情報を更新しました。`);
    }
  }
}

/**
 * 各指定プレイリストから新着動画を取得し、全シートと重複チェックした上で
 * 指定されたそれぞれのシートへ自動配分して追記します。
 */
function fetchNewVideosOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 全対象シートから登録済み動画IDを一括取得
  const existingVideoIds = new Set();
  TARGET_SHEETS.forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // D列 (4列目)
      values.forEach(r => {
        const vid = String(r[0]).trim();
        if (vid) existingVideoIds.add(vid);
      });
    }
  });

  // 2. 各プレイリストの処理
  PLAYLIST_MAP.forEach(config => {
    const playlistId = config.playlistId;
    const targetSheetName = config.targetSheetName;
    const targetSheet = ss.getSheetByName(targetSheetName);

    if (!targetSheet) {
      Logger.log(`エラー: 対象シート '${targetSheetName}' が存在しません。スキップします。`);
      return;
    }

    const newItems = [];
    let nextPageToken = '';
    let duplicateCount = 0; // 連続重複カウント

    try {
      do {
        const response = YouTube.PlaylistItems.list('snippet,contentDetails', {
          playlistId: playlistId,
          maxResults: 50,
          pageToken: nextPageToken
        });

        const items = response.items || [];
        for (const item of items) {
          const videoId = item.contentDetails ? item.contentDetails.videoId : (item.snippet.resourceId ? item.snippet.resourceId.videoId : "");
          if (!videoId) continue;

          // 既に登録済みの場合はスキップ（※連続5件重複で新着探索を終了）
          if (existingVideoIds.has(videoId)) {
            duplicateCount++;
            if (duplicateCount >= 5) break; 
            continue;
          }

          duplicateCount = 0; // 新着が見つかったらカウントリセット
          newItems.push({
            videoId: videoId,
            addedAt: item.contentDetails && item.contentDetails.videoPublishedAt ? item.contentDetails.videoPublishedAt.substring(0, 10) : ''
          });

          existingVideoIds.add(videoId);
        }

        if (duplicateCount >= 5) break;
        nextPageToken = response.nextPageToken;
      } while (nextPageToken);

    } catch (e) {
      Logger.log(`プレイリスト取得エラー (ID: ${playlistId}): ` + e.toString());
      return;
    }

    if (newItems.length === 0) {
      Logger.log(`[${targetSheetName}] 新着動画はありませんでした。`);
      return;
    }

    // 古い順（追加された順）に並び替え
    newItems.reverse();

    // 3. 動画の詳細データ（タイトル、説明欄、再生数、いいね数）を50件ずつ一括取得
    const newVideoIds = newItems.map(item => item.videoId);
    const videoDetailsMap = new Map();

    for (let i = 0; i < newVideoIds.length; i += 50) {
      const chunkIds = newVideoIds.slice(i, i + 50).join(',');
      const videoResponse = YouTube.Videos.list('snippet,statistics', {
        id: chunkIds
      });

      (videoResponse.items || []).forEach(v => {
        const stats = v.statistics || {};
        const snippet = v.snippet || {};

        videoDetailsMap.set(v.id, {
          title: snippet.title || '',
          artist: snippet.videoOwnerChannelTitle || snippet.channelTitle || '不明',
          description: snippet.description || '',
          viewCount: Number(stats.viewCount || 0),
          likeCount: Number(stats.likeCount || 0),
          publishedAt: snippet.publishedAt ? snippet.publishedAt.substring(0, 10) : ''
        });
      });
    }

    // 4. 最新のシート配置（16列 P列まで）に合わせて一括追加用配列を構築
    const newRowsToAppend = [];
    newItems.forEach(item => {
      const detail = videoDetailsMap.get(item.videoId) || {
        title: '', artist: '不明', description: '', viewCount: 0, likeCount: 0, publishedAt: ''
      };

      // A(0) 〜 P(15) の 16列配列
      const rowArray = createCommonSongRow_();
      rowArray[0] = true;                                       // A列: 有効フラグ (TRUE)
      rowArray[3] = item.videoId;
      rowArray[4] = item.videoId;                               // E列: 動画IDリスト
      rowArray[6] = detail.title;                               // G列: 動画タイトル ★
      rowArray[1] = detail.artist;                              // B列: アーティスト名 ★
      rowArray[7] = detail.description;                         // H列: 説明欄 ★
      rowArray[8] = detail.viewCount;                           // I列: 再生数 ★
      rowArray[10] = detail.publishedAt || item.addedAt;        // K列: 投稿日 ★
      rowArray[9] = detail.likeCount;                          // J列: いいね数 ★
      rowArray[15] = config.genreTag || "プレイリスト自動追加";  // P列: タグ ★

      newRowsToAppend.push(rowArray);
    });

    // 5. シートの末尾へ一括書き込み (1回のAPI通信)
    if (newRowsToAppend.length > 0) {
      const lastRow = targetSheet.getLastRow();
      const startRow = (lastRow < 2) ? 2 : lastRow + 1;
      targetSheet.getRange(startRow, 1, newRowsToAppend.length, 16).setValues(newRowsToAppend);
      Logger.log(`[${targetSheetName}] に ${newRowsToAppend.length} 件の新しい動画を一括追記しました。`);
    }
  });
}

// ==========================================
// ユーティリティ関数
// ==========================================

function formatIsoDuration(isoDuration) {
  if (!isoDuration) return '';
  
  const match = isoDuration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return isoDuration;

  const hours = parseInt(match[1] || 0, 10);
  const minutes = parseInt(match[2] || 0, 10);
  const seconds = parseInt(match[3] || 0, 10);

  const formattedMinutes = hours > 0 ? String(minutes).padStart(2, '0') : String(minutes);
  const formattedSeconds = String(seconds).padStart(2, '0');

  if (hours > 0) {
    return `${hours}:${formattedMinutes}:${formattedSeconds}`;
  } else {
    return `${formattedMinutes}:${formattedSeconds}`;
  }
}

function importMultipleCoverPlaylists() {
  const PLAYLIST_INPUTS = [
    "PL_0A0t0-Y0AMGFPCuKDZ3o8PMVpKcNvOQ",
    "PLF54qiUa9cTsy9SnCxc5oa5GOZnp0RNKm",
    "PLpt61bADOMwXIZpLr09sCNeocN7CXjcoS",
    "PLMF8NocLqJtubiccXUm9SCZ79dDPSLq2B",
    "PLAo9RlHR2tDakgakxbOxT9dAZD0rpwvXQ",
    "PLYt9gyZFZDj_e7p0saJXMghPsjmBMwyiO",
    "PLQoA24ikdy_kFurrcDRRJufTwHo_Kv6ew",
    "PL7LdRPp7xCkOwyOEKkE5YNC1_pGCTU2Id",
    "PL7LdRPp7xCkP9JJYvAqrQJLScbSol7S4-",
    "PL6sZ3uYmeG1vGZVLNlWdJkm3W0R86y_Qa",
    "PL7EFb7T-o_1O4fE4frYhWI4toIh8q5LWH",
    "PLbk25bH74T5ZgCqaphR7q_dQ8Ja65Z4r_",
    "PLOLvMgy3Jf3FVLGFJ4HooifOvd3UIQ1eJ",
    "PLUp1t9SPBl6r4nqsQBg7HH7fPX2l9AGQk",
    "PLzYaEBSRCCt5kgnBdQkLM7yDDZt2_NIZG",
    "PLMUtkfkYDjAFJlZeDBNLAlrKjfue2DRRu",
    "PL3y-e0n6k6gDoG3mViEOESGV1WhY8aCrR",
    "PLI3I4eCwIlMiUDPebeFM5efCZQ26nKMEE",
    "PL9jli3ZqRhJD13ORC2CUTEYnjeZt9KMPb",
    "PLBoFNGFmekcMIhlW5Z-umsS3VPnQx7tI8",
    "PLKo9UD3uKyyFUSwUDM5hCTB4fEvZqYs0P",
    "PLZ34fLWik_iAP2AdGLOHthUhAJTrEXqGb",
    "PLIHyIgRAWkUz3MAUPbTg9XcuP_rzDJ1bk",
    "PL6gUpTCMieF7MBNPM5RROJj97i9EO-Yl7",
    "PL6gUpTCMieF7Dl87iAuDEWdB4WEwO9M9Q",
    "PLNvxSJodVoAu5ABQjfuPGnzqkRMLtN0y2",
    "PL7Mw3OdFU891DfLqncG9cbkwH80lHFZEl",
    "PLet0wKVboCzEMtX-hZQn1obu40HWVKn7q",
    "PLrQwEkvx5phi8umsHZd-VKAKZ1XZ_7OCy",
    "PLWGY2acU-ZeTnH4EvPHYEcqCUx145dDUP",
    "PL-ksV7rAwJHWfgsi38uZEsWpnY128WwBC",
    "PLD992QYJ2953ekTB3YuoHcF505PWPXlwb",
    "PLCxKJRIVDrYHVk6OfBsuGwIujL63ooOQV",
    "PLUfQ3xz0-Jen7EoEDeZAkl5yVuBdtw-Wl",
    "PLUfQ3xz0-JemBHsJIHPvt0ZTb7YpuwzYm",
    "PLi6TWx3pTf1dHW-DvMpPEW2NLDICBkW8z",
    "PLHeIKIUgnbJBzW9cvaqSG6P9_NvXldgCH",
    "PLdcktSYecSKfeQ7p5CrVO4avrmLzSL8mt",
    "PL0OWyC-pOgQnQ1POfg1fOiXffUPE-b3sR",
    "PL6DI34hJswwQ8um3dpgrZY1elCO-tNfrc",
    "PL6DI34hJswwRLDUbazrru0CHk-KDS9gHB",
    "PLuRWkjIMjanEVABxsKoHeDrRDXsMf1dbE",
    "PLrALGrrF-6IVnnurSv7Nxdxpo5zJmmjnC",
    "PL_36PeBLiASmMHr4O3Px0pQeyW4i3vDxa",
    "PLH6DjTF_aCGjhddtGEEr7E1BApvg0EpIM",
    "PLwtTNsTyL23S4MNl0-gdozd769CytE_9X",
    "PLB8Nt5W7hnKA_pG2qljWbgVmJPobrLTm4",
    "PLsAbCDORf2yRA9TvFeQkVXnSHdghrfNuH",
    "PLQZecHYc3j2plhVMu0squPpD4XRoQmX-y",
    "PLpBqtLy3mHw2Wlox1cPU2-WGM5VmMYR8m",
    "PLh_mnrb4W5H6ir9_rwZomjr5SKhPqDf7y",
    "PLcPAszg2ItaGAkxKIUiTTtgekfmgMmmEQ",
    "PLQNcTNavhAOg78zZNEYFEQYXA02QOmlxM",
    "PLpy43bHw-UIVGeA9pDHRyZSHGlktU90ro",
    "PLSJG6WXf5JvFn2FWv_B5HP7DklUtdeZiM",
    "PLf4O_VcbYo24EjtizbTZt_tTgwXaaK8N-",
    "PLhOafnpk0ZP_C70vgr0yhVMXP7bgPGnKN",
    "PL3Z_j65GhJVzEX-Ek6PGd9B9Q4dL-JC9y",
    "PLjrOpeHGfCDp_5-Zgwyfz9Oua4A_hCUT8",
    "PL9781Kn53ls9iSV-Kl-Gxk7X5VKeeNzoX",
    "PLeLzFij5dykeofPj21nmXDrCJUxmvAOKW",
    "PL6IyOjey3U8WiUwXw-1AIDv99qMB9otjp"
  ];
  
  const TARGET_SHEET_NAME = "hololive_cover_song";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!targetSheet) {
    Browser.msgBox(`エラー: シート '${TARGET_SHEET_NAME}' が見つかりません。`);
    return;
  }

  const existingVideoIds = new Set();
  TARGET_SHEETS.forEach(sName => {
    const s = ss.getSheetByName(sName);
    if (!s) return;
    const lastRow = s.getLastRow();
    if (lastRow >= 2) {
      const values = s.getRange(2, 4, lastRow - 1, 1).getValues();
      values.forEach(r => {
        const vid = String(r[0]).trim();
        if (vid) existingVideoIds.add(vid);
      });
    }
  });

  let totalAddedCount = 0;

  PLAYLIST_INPUTS.forEach((input, index) => {
    const playlistId = extractPlaylistId(input);
    if (!playlistId) {
      Logger.log(`[スキップ] 無効な入力です (インデックス: ${index})`);
      return;
    }

    const playlistItems = [];
    let nextPageToken = "";

    try {
      do {
        const res = YouTube.PlaylistItems.list('snippet,contentDetails', {
          playlistId: playlistId,
          maxResults: 50,
          pageToken: nextPageToken
        });

        (res.items || []).forEach(item => {
          const vid = item.contentDetails.videoId;
          if (vid && !existingVideoIds.has(vid)) {
            playlistItems.push({
              videoId: vid,
              addedAt: ""
            });
            existingVideoIds.add(vid);
          }
        });

        nextPageToken = res.nextPageToken;
      } while (nextPageToken);

    } catch (e) {
      Logger.log(`[エラー] プレイリスト取得失敗 (ID: ${playlistId}): ` + e.toString());
      return;
    }

    if (playlistItems.length === 0) {
      Logger.log(`[完了] プレイリスト (${playlistId}): 新規動画はありませんでした。`);
      return;
    }

    playlistItems.reverse();

    const newVideoIds = playlistItems.map(item => item.videoId);
    const videoDetailsMap = new Map();

    for (let i = 0; i < newVideoIds.length; i += 50) {
      const chunkIds = newVideoIds.slice(i, i + 50).join(',');
      const videoRes = YouTube.Videos.list('snippet,statistics,contentDetails', {
        id: chunkIds
      });

      (videoRes.items || []).forEach(v => {
        const stats = v.statistics || {};
        const snippet = v.snippet || {};
        const contentDetails = v.contentDetails || {};

        videoDetailsMap.set(v.id, {
          title: snippet.title || '',
          description: snippet.description || '',
          viewCount: Number(stats.viewCount || 0),
          likeCount: Number(stats.likeCount || 0),
          publishedAt: snippet.publishedAt ? snippet.publishedAt.split("T")[0] : '',
          duration: formatIsoDuration(contentDetails.duration),
          channelId: snippet.channelId || '',
          channelTitle: snippet.channelTitle || ''
        });
      });
    }

    playlistItems.forEach(item => {
      const detail = videoDetailsMap.get(item.videoId) || {
        title: '', description: '', viewCount: 0, likeCount: 0, publishedAt: '', duration: '', channelId: '', channelTitle: ''
      };

      targetSheet.appendRow([
        true,                       // A: 有効
        detail.channelTitle,        // B: 歌唱者
        "",               // C: 音源公開日
        item.videoId,               // D: 代表動画ID
        item.videoId,               // E: 動画IDリスト
        "手動プレイリスト一括追加",   // F: メモ
        detail.title,               // G: 動画タイトル
        detail.description,         // H: 説明欄
        detail.viewCount,           // I: 再生数
        detail.likeCount,           // J: いいね数
        detail.publishedAt,         // K: 投稿日時
        detail.duration,            // L: 動画の長さ
        detail.channelId,           // M: チャンネルID
        detail.channelTitle,        // N: チャンネル名
        null,                       // O: 指定ラントロ秒数
        "歌ってみた"                // P: タグ
      ]);
      totalAddedCount++;
    });

    Logger.log(`[成功] プレイリスト (${playlistId}) から ${playlistItems.length} 件追加しました。`);
  });

  Browser.msgBox(`処理完了: 合計 ${totalAddedCount} 件の新しい動画を '${TARGET_SHEET_NAME}' シートに追加しました。`);
}

function extractPlaylistId(input) {
  if (!input) return null;
  const match = String(input).match(/[&?]list=([^&]+)/);
  if (match) return match[1];
  
  const trimmed = String(input).trim();
  if (trimmed.length > 10) return trimmed;
  
  return null;
}

function fetchTop10VocaloidVideoIdsAuto() {
  const apiKey = getRequiredScriptProperty_('SCRAPINGBEE_API_KEY');
  const targetUrl = 'https://lvchart.com/weekly/2022-12-1/';
  const sheetName = 'LocalVoid_weekly_best10';
  const targetPeriod = '2022-12-1';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`シート 「${sheetName}」 が見つかりません。`);
    return;
  }

  // ScrapingBeeのAPIエンドポイントを経由してHTMLを取得
  const requestUrl = `https://app.scrapingbee.com/api/v1/?api_key=${apiKey}&url=${encodeURIComponent(targetUrl)}&render_js=false`;

  try {
    const response = UrlFetchApp.fetch(requestUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log(`取得失敗: ステータスコード ${response.getResponseCode()}`);
      return;
    }
    
    const html = response.getContentText();

    // 動画IDの抽出
    const regex = /i\.ytimg\.com\/vi\/([a-zA-Z0-9_-]{11})\//g;
    const videoIds = [];
    let match;

    while ((match = regex.exec(html)) !== null) {
      const id = match[1];
      if (!videoIds.includes(id)) {
        videoIds.push(id);
      }
      if (videoIds.length === 10) break;
    }

    if (videoIds.length === 0) {
      Logger.log('動画IDが見つかりませんでした。');
      return;
    }

    // 書き込みデータの作成
    const dValues = [];
    const pValues = [];
    for (let i = 0; i < videoIds.length; i++) {
      dValues.push([videoIds[i]]);
      pValues.push([`${targetPeriod},#${i + 1}`]);
    }

    // スプレッドシートへセット
    sheet.getRange(2, 4, videoIds.length, 1).setValues(dValues);  // D列
    sheet.getRange(2, 16, videoIds.length, 1).setValues(pValues); // P列

    Logger.log(`成功: ${videoIds.length} 件取得完了。1位: ${videoIds[0]}`);

  } catch (e) {
    Logger.log(`エラー: ${e.message}`);
  }
}

//LocalVoid情報一括取得
function fetchAllVocaloidRankings() {
  const apiKey = getRequiredScriptProperty_('SCRAPINGBEE_API_KEY'); // ScrapingBeeのAPIキーを入力
  const sheetName = 'LocalVoid_weekly_best10';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`シート 「${sheetName}」 が見つかりませんでした。`);
    return;
  }

  // --- 1. 既存データの読み込みと取得済み週の把握 ---
  const existingIdRowMap = new Map();
  const processedPeriods = new Set(); // 取得済みの週（例: "2022-12-1"）を保持
  const lastRow = sheet.getLastRow();
  
  if (lastRow >= 2) {
    const dValues = sheet.getRange(2, 4, lastRow - 1, 2).getValues();  // D列 (動画ID)
    const pValues = sheet.getRange(2, 16, lastRow - 1, 1).getValues(); // P列 (ランク履歴)

    for (let i = 0; i < dValues.length; i++) {
      const id = String(dValues[i][0]).trim();
      const pText = String(pValues[i][0]).trim();
      const currentRow = i + 2;

      if (id !== '') {
        existingIdRowMap.set(id, currentRow);
      }

      // P列から処理済みの時期（"YYYY-M-W"）を抽出してSetに追加
      if (pText !== '') {
        const entries = pText.split(',');
        for (const entry of entries) {
          const period = entry.split(',#')[0].split('#')[0]; // "2022-12-1" 部分を抽出
          if (period) {
            processedPeriods.add(period.trim());
          }
        }
      }
    }
  }

  // --- 2. 2022年12月1週〜2026年9月1週までのループ処理 ---
  const startYear = 2026;
  const endYear = 2026;

  for (let year = startYear; year <= endYear; year++) {
    const startMonth = (year === 2026) ? 2 : 1;
    const endMonth = (year === 2026) ? 9 : 12;

    for (let month = startMonth; month <= endMonth; month++) {
      const maxWeek = (year === 2026 && month === 9) ? 1 : 5;

      for (let week = 1; week <= maxWeek; week++) {
        const periodStr = `${year}-${month}-${week}`;

        // ★ 取得済みの週であればURLアクセスをスキップ
        if (processedPeriods.has(periodStr)) {
          Logger.log(`[スキップ (取得済み)] ${periodStr}`);
          continue;
        }

        const targetUrl = `https://lvchart.com/weekly/${periodStr}/`;
        Logger.log(`[取得開始] ${periodStr}`);

        // サーバー負荷軽減のため1秒待機
        Utilities.sleep(1000);

        // APIリクエスト
        const requestUrl = `https://app.scrapingbee.com/api/v1/?api_key=${apiKey}&url=${encodeURIComponent(targetUrl)}&render_js=false`;

        try {
          const response = UrlFetchApp.fetch(requestUrl, { muteHttpExceptions: true });
          const statusCode = response.getResponseCode();

          // ページが存在しない場合（404等）はスキップ
          if (statusCode !== 200) {
            Logger.log(`[スキップ (ページなし)] ${periodStr} (Status: ${statusCode})`);
            continue;
          }

          const html = response.getContentText();

          // サムネイルURLから動画IDを抽出
          const regex = /i\.ytimg\.com\/vi\/([a-zA-Z0-9_-]{11})\//g;
          const videoIds = [];
          let match;

          while ((match = regex.exec(html)) !== null) {
            const id = match[1];
            if (!videoIds.includes(id)) {
              videoIds.push(id);
            }
            if (videoIds.length === 10) break;
          }

          if (videoIds.length === 0) {
            Logger.log(`[注意] ${periodStr} から動画IDが抽出できませんでした。`);
            continue;
          }

          // --- 3. 抽出した上位10曲をシートへ書き込み/追記 ---
          for (let i = 0; i < videoIds.length; i++) {
            const id = videoIds[i];
            const rank = i + 1;
            const entryText = `${periodStr},#${rank}`;

            if (existingIdRowMap.has(id)) {
              // 既存ID: P列にカンマ区切りで追記
              const targetRow = existingIdRowMap.get(id);
              const pCell = sheet.getRange(targetRow, 16);
              const currentPValue = String(pCell.getValue()).trim();

              if (currentPValue === '') {
                pCell.setValue(entryText);
              } else {
                pCell.setValue(`${currentPValue},${entryText}`);
              }
            } else {
              // 新規ID: 末尾行に追加
              const newRow = sheet.getLastRow() < 1 ? 2 : sheet.getLastRow() + 1;
              sheet.getRange(newRow, 4).setValue(id);        // D列
              sheet.getRange(newRow, 16).setValue(entryText); // P列

              existingIdRowMap.set(id, newRow);
            }
          }

          // 処理済みセットに追加
          processedPeriods.add(periodStr);
          Logger.log(`[完了] ${periodStr} (${videoIds.length}曲)`);

        } catch (e) {
          Logger.log(`[エラー] ${periodStr}: ${e.message}`);
        }
      }
    }
  }

  Logger.log('==== 全期間の処理が終了しました ====');
  SpreadsheetApp.getUi().alert('全期間の更新処理が完了しました。');
}

function fetchNewVocaloidRankingsDaily() {
  const apiKey = getRequiredScriptProperty_('SCRAPINGBEE_API_KEY'); // ScrapingBeeのAPIキーを入力
  const sheetName = 'LocalVoid_weekly_best10';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`シート 「${sheetName}」 が見つかりませんでした。`);
    return;
  }

  // --- 1. 既存データの読み込みと「最後に取得した最新の週」を特定 ---
  const existingIdRowMap = new Map();
  const processedPeriods = new Set();
  const lastRow = sheet.getLastRow();

  let latestYear = 2022;
  let latestMonth = 12;
  let latestWeek = 1;

  if (lastRow >= 2) {
    const dValues = sheet.getRange(2, 4, lastRow - 1, 2).getValues();  // D列
    const pValues = sheet.getRange(2, 16, lastRow - 1, 1).getValues(); // P列

    for (let i = 0; i < dValues.length; i++) {
      const id = String(dValues[i][0]).trim();
      const pText = String(pValues[i][0]).trim();
      const currentRow = i + 2;

      if (id !== '') {
        existingIdRowMap.set(id, currentRow);
      }

      if (pText !== '') {
        const entries = pText.split(',');
        for (const entry of entries) {
          const period = entry.split(',#')[0].split('#')[0].trim();
          if (period) {
            processedPeriods.add(period);

            // "YYYY-M-W" から最新の週を比較・特定
            const parts = period.split('-').map(Number);
            if (parts.length === 3) {
              const [y, m, w] = parts;
              if (
                y > latestYear ||
                (y === latestYear && m > latestMonth) ||
                (y === latestYear && m === latestMonth && w > latestWeek)
              ) {
                latestYear = y;
                latestMonth = m;
                latestWeek = w;
              }
            }
          }
        }
      }
    }
  }

  Logger.log(`[確認] 取得済みの最終週: ${latestYear}-${latestMonth}-${latestWeek}`);

  // --- 2. 今日までの日付情報から探索範囲を算出 ---
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;

  // --- 3. 最新週の翌週から本日までの範囲をループ探索 ---
  for (let year = latestYear; year <= currentYear; year++) {
    const startMonth = (year === latestYear) ? latestMonth : 1;
    const endMonth = (year === currentYear) ? currentMonth : 12;

    for (let month = startMonth; month <= endMonth; month++) {
      const startWeek = (year === latestYear && month === latestMonth) ? latestWeek + 1 : 1;

      for (let week = startWeek; week <= 5; week++) {
        const periodStr = `${year}-${month}-${week}`;

        if (processedPeriods.has(periodStr)) {
          continue;
        }

        const targetUrl = `https://lvchart.com/weekly/${periodStr}/`;
        Logger.log(`[新着チェック] ${periodStr} を確認中...`);

        Utilities.sleep(1000);

        const requestUrl = `https://app.scrapingbee.com/api/v1/?api_key=${apiKey}&url=${encodeURIComponent(targetUrl)}&render_js=false`;

        try {
          const response = UrlFetchApp.fetch(requestUrl, { muteHttpExceptions: true });
          const statusCode = response.getResponseCode();

          if (statusCode !== 200) {
            Logger.log(`[未公開/存在なし] ${periodStr} はまだありません (Status: ${statusCode})`);
            continue;
          }

          const html = response.getContentText();
          const regex = /i\.ytimg\.com\/vi\/([a-zA-Z0-9_-]{11})\//g;
          const videoIds = [];
          let match;

          while ((match = regex.exec(html)) !== null) {
            const id = match[1];
            if (!videoIds.includes(id)) {
              videoIds.push(id);
            }
            if (videoIds.length === 10) break;
          }

          if (videoIds.length === 0) {
            Logger.log(`[注意] ${periodStr} から動画IDが抽出できませんでした。`);
            continue;
          }

          // --- 4. シートへの反映・更新 ---
          for (let i = 0; i < videoIds.length; i++) {
            const id = videoIds[i];
            const rank = i + 1;
            const entryText = `${periodStr},#${rank}`;

            if (existingIdRowMap.has(id)) {
              // 既存ID: P列に追記（A列やD列はそのまま）
              const targetRow = existingIdRowMap.get(id);
              const pCell = sheet.getRange(targetRow, 16);
              const currentPValue = String(pCell.getValue()).trim();

              if (currentPValue === '') {
                pCell.setValue(entryText);
              } else {
                pCell.setValue(`${currentPValue},${entryText}`);
              }
            } else {
              // 新規ID: 末尾行に追加
              const newRow = sheet.getLastRow() < 1 ? 2 : sheet.getLastRow() + 1;
              
              sheet.getRange(newRow, 1).setValue(true);      // ★ A列 (1列目) に TRUE を入力
              sheet.getRange(newRow, 4).setValue(id);        // D列 (4列目) に動画ID
              sheet.getRange(newRow, 16).setValue(entryText); // P列 (16列目) にランク履歴

              existingIdRowMap.set(id, newRow);
            }
          }

          processedPeriods.add(periodStr);
          Logger.log(`[成功] 新規ランキング ${periodStr} を追加しました (${videoIds.length}曲)`);

        } catch (e) {
          Logger.log(`[エラー] ${periodStr}: ${e.message}`);
        }
      }
    }
  }

  Logger.log('==== 日次チェック完了 ====');
}

// YouTube Data API を使用して再生リスト内の動画を取得する関数
function getPlaylistItemsDirect(playlistId) {
  const items = [];
  let nextPageToken = "";

  try {
    do {
      // YouTube Data API (YouTube v3) を呼び出し
      const res = YouTube.PlaylistItems.list("snippet,contentDetails", {
        playlistId: playlistId,
        maxResults: 50,
        pageToken: nextPageToken
      });

      if (res && res.items) {
        res.items.forEach(item => {
          const snippet = item.snippet;
          const vid = snippet.resourceId ? snippet.resourceId.videoId : "";
          if (vid && snippet.title !== "Private video" && snippet.title !== "Deleted video") {
            items.push({
              youtubeId: vid,
              videoTitle: snippet.title,
              artist: snippet.videoOwnerChannelTitle || snippet.channelTitle || "不明",
              channelName: snippet.channelTitle || "",
              postDate: item.contentDetails && item.contentDetails.videoPublishedAt ? item.contentDetails.videoPublishedAt.substring(0, 10) : "",
              views: 0,
              likes: 0,
              description: snippet.description || "",
              customStartSec: null,
              sheetName: "PlaylistDirect"
            });
          }
        });
      }
      nextPageToken = res.nextPageToken;
    } while (nextPageToken);
  } catch (err) {
    Logger.log("getPlaylistItemsDirect Error: " + err.message);
  }

  return items;
}

function syncPlaylistToGeneralSheet(playlistId, genreTags) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("General");
  if (!sheet) return [];

  // 1. YouTube API から再生リスト内の全動画IDを取得
  const rawVideoList = [];
  const playlistVideoIds = [];
  let nextPageToken = "";

  do {
    const res = YouTube.PlaylistItems.list("snippet,contentDetails", {
      playlistId: playlistId,
      maxResults: 50,
      pageToken: nextPageToken
    });

    if (res && res.items) {
      res.items.forEach(item => {
        const snippet = item.snippet;
        const vid = snippet.resourceId ? snippet.resourceId.videoId : "";
        if (vid && snippet.title !== "Private video" && snippet.title !== "Deleted video") {
          playlistVideoIds.push(vid);
          rawVideoList.push({
            youtubeId: vid,
            videoTitle: snippet.title,
            artist: snippet.videoOwnerChannelTitle || snippet.channelTitle || "不明",
            channelName: snippet.channelTitle || "",
            postDate: item.contentDetails && item.contentDetails.videoPublishedAt ? item.contentDetails.videoPublishedAt.substring(0, 10) : ""
          });
        }
      });
    }
    nextPageToken = res.nextPageToken;
  } while (nextPageToken);

  if (playlistVideoIds.length === 0) return [];

  // 2. 動画IDをもとに YouTube.Videos.list から「説明欄」「再生数」「いいね数」を一括取得 (50件ずつ処理)
  const videoDetailsMap = new Map();
  for (let i = 0; i < playlistVideoIds.length; i += 50) {
    const chunkIds = playlistVideoIds.slice(i, i + 50).join(",");
    const videoRes = YouTube.Videos.list("snippet,statistics", {
      id: chunkIds
    });

    if (videoRes && videoRes.items) {
      videoRes.items.forEach(vItem => {
        const stats = vItem.statistics || {};
        const snip = vItem.snippet || {};
        videoDetailsMap.set(vItem.id, {
          description: snip.description || "",
          views: Number(stats.viewCount) || 0,
          likes: Number(stats.likeCount) || 0
        });
      });
    }
  }

  // 取得した詳細情報を結合
  const playlistVideos = rawVideoList.map(v => {
    const details = videoDetailsMap.get(v.youtubeId) || { description: "", views: 0, likes: 0 };
    return {
      ...v,
      description: details.description,
      views: details.views,
      likes: details.likes
    };
  });

  // 3. シートのデータを全件取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 16列 (P列) 以上まで確実に範囲を取得
  const maxCols = Math.max(sheet.getLastColumn(), 16);
  const dataRange = sheet.getRange(2, 1, lastRow - 1, maxCols);
  const values = dataRange.getValues();

  // D列 (インデックス3) が代表YouTubeID
  const existingIdRowMap = new Map();
  for (let i = 0; i < values.length; i++) { const row = values[i]; const rowNumber = i + 2; [row[3], row[4]].forEach(cell => String(cell || "").split(/[,\s、，;；|]+/).forEach(id => { id = id.trim(); if (id && !existingIdRowMap.has(id)) existingIdRowMap.set(id, rowNumber); })); }

  // 4. A列を一旦すべて FALSE に設定
  const aColumnValues = values.map(() => [false]);
  
  // 5. 再生リストに含まれる動画のA列を TRUE に更新＆新規追加リスト作成
  const newRowsToAppend = [];
  const targetQuizData = [];

  const seenPlaylistIds = new Set(); playlistVideos.forEach(v => { const videoId = String(v.youtubeId || "").trim(); if (!videoId || seenPlaylistIds.has(videoId)) return; seenPlaylistIds.add(videoId); v.youtubeId = videoId;
    if (existingIdRowMap.has(v.youtubeId)) {
      // 既存曲：A列を TRUE に変更
      const rowIndex = existingIdRowMap.get(v.youtubeId) - 2;
      aColumnValues[rowIndex][0] = true;

      const rowData = values[rowIndex];
      targetQuizData.push({
        sheetName: "General",
        rowIndex: existingIdRowMap.get(v.youtubeId),
        youtubeId: v.youtubeId,
        videoTitle: rowData[6] || v.videoTitle,   // G列 (6): タイトル
        artist: rowData[1] || v.artist,           // B列 (1): アーティスト
        description: rowData[7] || v.description, // H列 (7): 説明欄
        views: rowData[8] || v.views,             // I列 (8): 再生数
        postDate: rowData[10] || v.postDate,      // K列 (10): 投稿日
        likes: rowData[9] || v.likes,            // J列 (9): いいね数
        tags: rowData[15] || genreTags,            // P列 (15): タグ
        customStartSec: rowData[14] !== "" && rowData[14] !== undefined ? Number(rowData[14]) : null,
        comment: rowData[5] || ""
      });
    } else {
      // 未登録曲：新規追加データを作成 (16列分 P列まで生成)
      // A(0), B(1), C(2), D(3), E(4), F(5), G(6), H(7), I(8), J(9), K(10), L(11), M(12), N(13), O(14), P(15)
      const rowArray = createCommonSongRow_();
      rowArray[0] = true;           // A列: フラグ (TRUE)
      rowArray[3] = v.youtubeId;
      rowArray[4] = v.youtubeId;    // E列: 動画IDリスト
      rowArray[6] = v.videoTitle;   // G列: タイトル ★
      rowArray[1] = v.artist;       // B列: アーティスト ★
      rowArray[7] = v.description;  // H列: 説明欄
      rowArray[8] = v.views;        // I列: 再生数
      rowArray[10] = v.postDate;    // K列: 投稿日 ★
      rowArray[9] = v.likes;       // J列: いいね数
      rowArray[15] = genreTags;     // P列: タグ ★

      newRowsToAppend.push(rowArray);
    }
  });

  // 6. A列の更新を一括書き込み
  sheet.getRange(2, 1, values.length, 1).setValues(aColumnValues);

  // 7. 未登録曲をシート末尾に追加
  if (newRowsToAppend.length > 0) {
    const startAppendRow = sheet.getLastRow() + 1;
    sheet.getRange(startAppendRow, 1, newRowsToAppend.length, 16).setValues(newRowsToAppend);

    // 新規追加分もクイズデータに含める
    newRowsToAppend.forEach((row, idx) => {
      targetQuizData.push({
        sheetName: "General",
        rowIndex: startAppendRow + idx,
        youtubeId: row[3],
        videoTitle: row[6],
        artist: row[1],
        description: row[7],
        views: Number(row[8] || 0),
        postDate: row[10],
        likes: Number(row[9] || 0),
        tags: row[15],
        customStartSec: null,
        comment: ""
      });
    });
  }

  return targetQuizData;
}

// --- 未登録曲をDB(Generalシート)へ保存する関数 ---
function importPlaylistToGeneralSheet(playlistId, genreTags) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("General");
  if (!sheet) return { addedCount: 0, skippedCount: 0 };

  // 1. YouTube API から再生リスト内の動画ID一覧を取得
  const rawVideoList = [];
  const playlistVideoIds = [];
  let nextPageToken = "";

  do {
    const res = YouTube.PlaylistItems.list("snippet,contentDetails", {
      playlistId: playlistId,
      maxResults: 50,
      pageToken: nextPageToken
    });

    if (res && res.items) {
      res.items.forEach(item => {
        const snippet = item.snippet;
        const vid = snippet.resourceId ? snippet.resourceId.videoId : "";
        if (vid && snippet.title !== "Private video" && snippet.title !== "Deleted video") {
          playlistVideoIds.push(vid);
          rawVideoList.push({
            youtubeId: vid,
            videoTitle: snippet.title,
            artist: snippet.videoOwnerChannelTitle || snippet.channelTitle || "不明",
            postDate: item.contentDetails && item.contentDetails.videoPublishedAt ? item.contentDetails.videoPublishedAt.substring(0, 10) : ""
          });
        }
      });
    }
    nextPageToken = res.nextPageToken;
  } while (nextPageToken);

  if (playlistVideoIds.length === 0) return { addedCount: 0, skippedCount: 0 };

  // 2. YouTube.Videos.list で「説明欄」「再生数」「いいね数」を一括取得
  const videoDetailsMap = new Map();
  for (let i = 0; i < playlistVideoIds.length; i += 50) {
    const chunkIds = playlistVideoIds.slice(i, i + 50).join(",");
    const videoRes = YouTube.Videos.list("snippet,statistics", {
      id: chunkIds
    });

    if (videoRes && videoRes.items) {
      videoRes.items.forEach(vItem => {
        const stats = vItem.statistics || {};
        const snip = vItem.snippet || {};
        videoDetailsMap.set(vItem.id, {
          description: snip.description || "",
          views: Number(stats.viewCount) || 0,
          likes: Number(stats.likeCount) || 0
        });
      });
    }
  }

  // 3. 既存のD列 (代表YouTubeID) を全件チェックして重複判定マップを作成
  const lastRow = sheet.getLastRow();
  const existingIds = new Set();
  
  if (lastRow >= 2) {
    const dValues = sheet.getRange(2, 4, lastRow - 1, 2).getValues(); // D列 (4列目)
    dValues.forEach(row => {
      const vid = String(row[0]).trim();
      if (vid) existingIds.add(vid); String(row[1] || "").split(/[,\s、，;；|]+/).forEach(id => { id = id.trim(); if (id) existingIds.add(id); });
    });
  }

  // 4. 未登録曲のみを抽出して新規定型フォーマット（16列 P列まで）を生成
  const newRowsToAppend = [];
  let skippedCount = 0;

  rawVideoList.forEach(v => {
    if (existingIds.has(v.youtubeId)) {
      skippedCount++;
    } else {
      const details = videoDetailsMap.get(v.youtubeId) || { description: "", views: 0, likes: 0 };
      
      // A(0) 〜 P(15) の配列を作成
      const rowArray = createCommonSongRow_();
      rowArray[0] = true;                 // A列: 有効フラグ (TRUE)
      rowArray[3] = v.youtubeId;
      rowArray[4] = v.youtubeId;          // E列: 動画IDリスト
      rowArray[6] = v.videoTitle;         // G列: タイトル
      rowArray[1] = v.artist;             // B列: アーティスト
      rowArray[7] = details.description;  // H列: 説明欄
      rowArray[8] = details.views;        // I列: 再生数
      rowArray[10] = v.postDate;          // K列: 投稿日
      rowArray[9] = details.likes;       // J列: いいね数
      rowArray[15] = genreTags;           // P列: タグ

      newRowsToAppend.push(rowArray);
      existingIds.add(v.youtubeId); // 同一再生リスト内の重複も防止
    }
  });

  // 5. シート末尾に未登録曲のみを一括追加
  if (newRowsToAppend.length > 0) {
    const startAppendRow = (lastRow < 2) ? 2 : lastRow + 1;
    sheet.getRange(startAppendRow, 1, newRowsToAppend.length, 16).setValues(newRowsToAppend);
  }

  return {
    addedCount: newRowsToAppend.length,
    skippedCount: skippedCount
  };
}

/*******************************************************
 * anison.online → Google Sheets
 *
 * A  有効
 * B  歌唱者・アーティスト
 * C  音源公開日
 * D  代表動画ID（自動生成）
 * E  動画IDリスト
 * F  メモ・コメント（変更しない）
 * G:N 既存処理（変更しない）
 * O  指定ラントロ秒数（変更しない）
 * P  タグ（変更しない）
 * Q  作品名
 * R  区分
 * S  順位
 * T  検索ステータス
 * U  照合スコア
 * V  照合理由
 * W  エラー
 *
 * 曲名などの取得用情報は _ANISON_CACHE に保存
 *******************************************************/


const CONFIG = {
  // ==============================
  // anison.online
  // ==============================

  SHEET_NAME: 'anime_song',

  BASE_URL:
    'https://anison.online/season?view=ranking&year={YEAR}',

  START_YEAR: 2008,
  get END_YEARS() {
    return Number(new Date().getFullYear()); // 確実に数値にする
  },

  TEST_YEARS: [2013, 2021],
  TEST_LIMIT: 3,

  // ==============================
  // Apps Script
  // ==============================

  // 1回の実行をこの時間以内に終了
  MAX_RUNTIME: 5 * 60 * 1000,

  // ==============================
  // Cache
  // ==============================

  CACHE_SHEET_NAME: '_ANISON_CACHE',

  // YouTube情報更新履歴
  YOUTUBE_CACHE_SHEET_NAME: '_YT_CACHE',

  // ==============================
  // YouTube API
  // ==============================

  // videos.list 1回につき最大50件
  YOUTUBE_BATCH_SIZE: 50,

  // 初期取り込み時に1回の実行で使う
  // videos.list の最大回数
  INITIAL_YOUTUBE_API_CALL_LIMIT: 10,

  // 定期更新時のvideos.list最大回数
  UPDATE_YOUTUBE_API_CALL_LIMIT: 5,

  // 既存動画情報を何日ごとに更新するか
  YOUTUBE_REFRESH_DAYS: 7,

  // ==============================
  // 定期監視
  // ==============================

  MONITOR_START_YEAR: 2008,
  get MONITOR_END_YEAR() {
    return Number(new Date().getFullYear());
  }
};


function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('アニソン取得')
    .addItem('② 初回取得を開始',
             'startInitialImport')
    .addItem('③ 初回取得を続行',
             'processInitialImport')
    .addItem('④ YouTube検索を実行',
             'processYouTubeQueue')
    .addItem('⑤ 週間更新',
             'weeklyUpdate')
    .addSeparator()
    .addItem('キャッシュシートを表示',
             'showAnisonCache')
    .addItem('キャッシュシートを非表示',
             'hideAnisonCache')
    .addToUi();
}


/*******************************************************
 * シート取得
 *******************************************************/

function getMainSheet_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (CONFIG.SHEET_NAME) {
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(
        '指定されたシートがありません: ' +
        CONFIG.SHEET_NAME
      );
    }

    return sheet;
  }

  throw new Error('CONFIG.SHEET_NAME must be set explicitly.');
}


/*******************************************************
 * キャッシュシート
 *******************************************************/

function getCacheSheet_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(CONFIG.CACHE_SHEET_NAME);

  if (!sheet) {

    sheet =
      ss.insertSheet(CONFIG.CACHE_SHEET_NAME);

    sheet.getRange(1, 1, 1, 12).setValues([[
      'KEY',
      'YEAR',
      'RANK',
      'DETAIL_URL',
      'SONG_TITLE',
      'WORK_TITLE',
      'ARTIST',
      'TYPE',
      'RELEASE_DATE',
      'EMBEDDED_IDS',
      'LAST_SEEN',
      'STATUS'
    ]]);

    sheet.hideSheet();
  }

  return sheet;
}


function showAnisonCache() {

  const sheet = getCacheSheet_();

  sheet.showSheet();

  SpreadsheetApp.setActiveSheet(sheet);
}


function hideAnisonCache() {

  const sheet = getCacheSheet_();

  sheet.hideSheet();
}


/*******************************************************
 * HTML取得
 *******************************************************/

function fetchHtml_(url) {

  const response =
    UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent':
          'Mozilla/5.0 (Google Apps Script)'
      }
    });

  const code =
    response.getResponseCode();

  const html =
    response.getContentText('UTF-8');

  Logger.log(
    '[HTTP] ' + url +
    ' / ' + code +
    ' / ' + html.length + ' bytes'
  );

  if (code !== 200) {

    throw new Error(
      'HTTP ' + code + ': ' + url
    );
  }

  return html;
}


/*******************************************************
 * HTMLデコード
 *******************************************************/

function decodeHtml_(text) {

  if (!text) return '';

  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#(\d+);/g, function(_, n) {
      return String.fromCharCode(Number(n));
    })
    .replace(/&#x([0-9a-f]+);/gi, function(_, n) {
      return String.fromCharCode(
        parseInt(n, 16)
      );
    });
}


/*******************************************************
 * タグ除去
 *******************************************************/

function stripTags_(html) {

  if (!html) return '';

  return decodeHtml_(html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]+>/g, ' ')
  )
    .replace(/\s+/g, ' ')
    .trim();
}


/*******************************************************
 * URL正規化
 *******************************************************/

function absoluteUrl_(url) {

  if (!url) return '';

  url = decodeHtml_(url.trim());

  if (/^https?:\/\//i.test(url)) {
    return url;
  }

  if (url.indexOf('//') === 0) {
    return 'https:' + url;
  }

  if (url.charAt(0) === '/') {
    return 'https://anison.online' + url;
  }

  return 'https://anison.online/' + url;
}


/*******************************************************
 * YouTube ID抽出
 *******************************************************/

function extractYouTubeIdsLegacy_(html) {

  const ids = [];

  if (!html) return ids;

  const patterns = [

    /youtube(?:-nocookie)?\.com\/embed\/([A-Za-z0-9_-]{11})/gi,

    /youtube(?:-nocookie)?\.com\/watch\?v=([A-Za-z0-9_-]{11})/gi,

    /youtu\.be\/([A-Za-z0-9_-]{11})/gi,

    /[?&]v=([A-Za-z0-9_-]{11})/gi

  ];

  patterns.forEach(function(re) {

    let m;

    while ((m = re.exec(html)) !== null) {

      const id = m[1];

      if (id && ids.indexOf(id) === -1) {
        ids.push(id);
      }
    }
  });

  return ids;
}


/*******************************************************
 * 日付抽出
 *******************************************************/

function extractReleaseDate_(text) {

  if (!text) return '';

  let m =
    text.match(
      /(?:公開日|発売日|音源公開日)\s*[:：]?\s*(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/
    );

  if (m) {
    return (
      m[1] + '-' +
      ('0' + m[2]).slice(-2) + '-' +
      ('0' + m[3]).slice(-2)
    );
  }

  m =
    text.match(
      /(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/
    );

  if (m) {
    return (
      m[1] + '-' +
      ('0' + m[2]).slice(-2) + '-' +
      ('0' + m[3]).slice(-2)
    );
  }

  return '';
}


/*******************************************************
 * ランキング解析
 *
 * anison.online の年間ランキングでは、
 *
 *   /anime/3499#song7749
 *       ↓
 *   曲名
 *       ↓
 *   アーティスト
 *       ↓
 *   シーズン
 *       ↓
 *   作品名
 *       ↓
 *   OP / ED / 主題歌 / 挿入歌
 *
 * という構造になっている。
 *
 * #songXXXX を「1曲」として扱う。
 *******************************************************/

function parseRanking_(html, year, limit) {

  const results = [];

  /*
   * script / styleを除去
   */
  const cleanHtml =
    html
      .replace(
        /<script[\s\S]*?<\/script>/gi,
        ''
      )
      .replace(
        /<style[\s\S]*?<\/style>/gi,
        ''
      );


  /*
   * 年間ランキング開始位置
   */
  const rankingStart =
    cleanHtml.search(
      /Youtube再生数降順/i
    );


  if (rankingStart < 0) {

    Logger.log(
      '[ランキング開始位置] 見つかりません'
    );

    return [];
  }


  /*
   * ランキング部分だけ
   */
  const rankingHtml =
    cleanHtml.substring(
      rankingStart
    );


  /*
   * #songXXXX を持つリンクを取得
   *
   * これが「1曲 = 1件」になる。
   */
  const songLinkRe =
    /<a\b[^>]*href=["']([^"']*\/anime\/\d+#song[^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;


  const songLinks = [];

  let m;

  while (
    (m = songLinkRe.exec(rankingHtml)) !== null
  ) {

    const href =
      absoluteUrl_(m[1]);

    const text =
      stripTags_(m[2]);

    songLinks.push({

      url: href,

      text: text,

      index: m.index,

      end: songLinkRe.lastIndex

    });
  }


  Logger.log(
    '[曲リンク数] ' +
    songLinks.length
  );


  /*
   * 1曲ずつ処理
   */
  for (
    let i = 0;
    i < songLinks.length;
    i++
  ) {

    if (
      results.length >= limit
    ) {
      break;
    }


    const songLink =
      songLinks[i];


    /*
     * 曲リンクの直後から、
     *
     * アーティスト
     * シーズン
     * 作品
     * 区分
     *
     * を取得する。
     */
    const after =
      rankingHtml.substring(
        songLink.end,
        Math.min(
          rankingHtml.length,
          songLink.end + 5000
        )
      );


    /*
     * アーティスト
     */
    const artist =
      extractNextArtist_(
        after
      );


    /*
     * 作品
     */
    const workInfo =
      extractNextAnimeWork_(
        after
      );


    /*
     * 区分
     */
    const type =
      extractTypeFromRankingArea_(
        after
      );


    /*
     * YouTube埋め込みID
     *
     * ランキングページ内にある場合のみ取得。
     * 実際の埋め込みは作品ページにもあるので、
     * 最終的には作品ページ側でも取得する。
     */
    const nearby =
      rankingHtml.substring(
        songLink.index,
        Math.min(
          rankingHtml.length,
          songLink.end + 4000
        )
      );


    const embeddedIds =
      extractYouTubeIds_(
        nearby
      );


    /*
     * 曲名
     *
     * リンクテキストが空の場合がある。
     *
     * その場合は後で作品ページの
     * #songXXXX から取得する。
     */
    let songTitle =
      songLink.text;


    /*
     * #songXXXX
     */
    const anchorMatch =
      songLink.url.match(
        /#(song\d+)$/i
      );


    const songAnchor =
      anchorMatch
        ? anchorMatch[1]
        : '';


    /*
     * ログ
     */
    Logger.log(
      '[RANK] #' +
      (i + 1) +
      ' / 曲=' +
      songTitle +
      ' / artist=' +
      artist +
      ' / 作品=' +
      workInfo.title +
      ' / 区分=' +
      type +
      ' / URL=' +
      songLink.url
    );


    /*
     * キャッシュキー
     */
    const key =
      makeCacheKey_(
        year,
        i + 1,
        songLink.url,
        songTitle,
        artist
      );


    results.push({

      key: key,

      year: year,

      rank: i + 1,

      /*
       * #songXXXX付きURLを保持
       */
      detailUrl:
        songLink.url,

      /*
       * 曲名
       */
      songTitle:
        songTitle,

      /*
       * 作品名
       */
      workTitle:
        workInfo.title,

      /*
       * アーティスト
       */
      artist:
        artist,

      /*
       * 区分
       */
      type:
        type,

      /*
       * 公開日は作品ページから取得
       */
      releaseDate:
        '',

      /*
       * YouTube ID
       */
      embeddedIds:
        embeddedIds,

      /*
       * song7749
       */
      songAnchor:
        songAnchor,

      lastSeen:
        new Date(),

      status:
        'RANKING'

    });
  }


  return results;
}


/*******************************************************
 * 曲リンク直後のアーティスト取得
 *******************************************************/

function extractNextArtist_(html) {

  /*
   * /artist/ リンクを探す
   */
  const re =
    /<a\b[^>]*href=["']([^"']*\/artist\/[^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi;


  const m =
    re.exec(html);


  if (m) {

    return stripTags_(
      m[2]
    );
  }


  /*
   * /artist/ URLがない場合、
   * 最初のリンクを候補にする
   */
  const linkRe =
    /<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;


  while (
    (m = linkRe.exec(html)) !== null
  ) {

    const url =
      m[1];

    const text =
      stripTags_(m[2]);


    if (!text) {
      continue;
    }


    /*
     * animeリンク等は除外
     */
    if (
      /\/anime\/\d+/i.test(url) ||
      /\/season/i.test(url) ||
      /facebook|twitter|x\.com/i.test(url)
    ) {
      continue;
    }


    return text;
  }


  return '';
}


/*******************************************************
 * 曲リンク直後の作品リンク取得
 *******************************************************/

function extractNextAnimeWork_(html) {

  /*
   * #songではない /anime/XXXX を探す。
   *
   * ランキングでは
   *
   * 曲リンク
   * ↓
   * アーティスト
   * ↓
   * 作品リンク
   *
   * の順なので、#songを除外する。
   */
  const re =
    /<a\b[^>]*href=["']([^"']*\/anime\/\d+[^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi;


  let m;

  while (
    (m = re.exec(html)) !== null
  ) {

    const url =
      m[1];

    const title =
      stripTags_(m[2]);


    /*
     * #songは別の曲なので除外
     */
    if (
      /#song\d+/i.test(url)
    ) {
      continue;
    }


    /*
     * 空タイトルも除外
     */
    if (!title) {
      continue;
    }


    return {

      url:
        absoluteUrl_(url),

      title:
        title
    };
  }


  return {

    url: '',

    title: ''
  };
}


/*******************************************************
 * 区分取得
 *******************************************************/

function extractTypeFromRankingArea_(
  html
) {

  /*
   * 作品リンク以降に登場する
   * 最初の区分を取得。
   */
  const text =
    stripTags_(
      html.substring(
        0,
        Math.min(
          html.length,
          2500
        )
      )
    );


  const types = [

    '主題歌',

    '挿入歌',

    'OP',

    'ED',

    'Opening Theme',

    'Ending Theme',

    'Opening',

    'Ending'

  ];


  /*
   * 作品名などに含まれる文字との誤判定を避ける
   */
  for (
    let i = 0;
    i < types.length;
    i++
  ) {

    if (
      text.indexOf(
        types[i]
      ) !== -1
    ) {

      return types[i];
    }
  }


  return '';
}

/*******************************************************
 * animeリンクの直前からアーティストを探す
 *******************************************************/

function extractArtistBeforeAnime_(
  html,
  animeIndex,
  rankingStart
) {

  /*
   * animeリンクより前のHTML
   */
  const before =
    html.substring(
      rankingStart,
      animeIndex
    );


  /*
   * 直前の /artist/ リンク
   */
  const re =
    /<a\b[^>]*href=["']([^"']*\/artist\/[^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi;

  let m;
  let last = null;

  while ((m = re.exec(before)) !== null) {

    const title =
      stripTags_(m[2]);

    if (title) {
      last = title;
    }
  }


  if (last) {
    return last;
  }


  /*
   * artist URLがない場合は、
   * 直前のリンクを候補にする。
   */
  const linkRe =
    /<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;

  let candidate = null;

  while ((m = linkRe.exec(before)) !== null) {

    const title =
      stripTags_(m[2]);

    if (!title) continue;

    if (
      /\/anime\/\d+/i.test(m[1]) ||
      /\/season/i.test(m[1]) ||
      /facebook|twitter|x\.com/i.test(m[1])
    ) {
      continue;
    }

    candidate = title;
  }


  return candidate || '';
}


/*******************************************************
 * animeリンク直後から区分を探す
 *******************************************************/

function extractTypeAfterAnime_(
  html,
  animeEnd
) {

  /*
   * animeリンクの直後3000文字
   */
  const after =
    html.substring(
      animeEnd,
      Math.min(
        html.length,
        animeEnd + 3000
      )
    );


  /*
   * まずHTMLをテキスト化
   */
  const text =
    stripTags_(after);


  /*
   * サイトで実際に使用される表記
   */
  const types = [
    '主題歌',
    '挿入歌',
    'Opening',
    'Ending',
    'Opening Theme',
    'Ending Theme',
    'OP',
    'ED'
  ];


  for (
    let i = 0;
    i < types.length;
    i++
  ) {

    if (
      text.indexOf(types[i]) !== -1
    ) {

      return types[i];
    }
  }


  return '';
}


/*******************************************************
 * YouTube ID抽出
 *******************************************************/

function extractYouTubeIds_(html) {

  const ids = [];

  if (!html) {
    return ids;
  }


  const patterns = [

    /youtube(?:-nocookie)?\.com\/embed\/([A-Za-z0-9_-]{11})/gi,

    /youtube(?:-nocookie)?\.com\/watch\?v=([A-Za-z0-9_-]{11})/gi,

    /youtu\.be\/([A-Za-z0-9_-]{11})/gi,

    /[?&]v=([A-Za-z0-9_-]{11})/gi

  ];


  patterns.forEach(function(re) {

    let m;

    while (
      (m = re.exec(html)) !== null
    ) {

      const id =
        m[1];

      if (
        id &&
        ids.indexOf(id) === -1
      ) {

        ids.push(id);
      }
    }
  });


  return ids;
}

/*******************************************************
 * 正規表現エスケープ
 *******************************************************/

function escapeRegExp_(text) {

  return text.replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&'
  );
}


/*******************************************************
 * キャッシュキー
 *******************************************************/

function makeCacheKey_(
  year,
  rank,
  detailUrl,
  songTitle,
  artist
) {

  return [
    year,
    rank,
    detailUrl,
    songTitle,
    artist
  ].join('|');
}


/*******************************************************
 * キャッシュ保存
 *******************************************************/

function saveCache_(items) {

  if (!items || !items.length) {
    return;
  }

  const sheet = getCacheSheet_();

  const values = sheet.getDataRange().getValues();

  const map = {};

  for (let i = 1; i < values.length; i++) {

    const key = values[i][0];

    if (key) {
      map[key] = i + 1;
    }
  }

  items.forEach(function(item) {

    const row = map[item.key];

    const rowValues = [[

      item.key,                         // A KEY
      item.year,                        // B YEAR
      item.rank,                        // C RANK
      item.detailUrl,                   // D DETAIL_URL
      item.songTitle,                   // E SONG_TITLE
      item.workTitle,                  // F WORK_TITLE
      item.artist,                     // G ARTIST
      item.type,                       // H TYPE
      item.releaseDate,                // I RELEASE_DATE
      unique_(item.embeddedIds || []).join(','), // J EMBEDDED_IDS
      item.lastSeen,                   // K LAST_SEEN
      item.status                      // L STATUS

    ]];

    if (row) {

      sheet
        .getRange(row, 1, 1, 12)
        .setValues(rowValues);

    } else {

      sheet
        .getRange(
          sheet.getLastRow() + 1,
          1,
          1,
          12
        )
        .setValues(rowValues);
    }

  });
}

/*******************************************************
 * 重複削除
 *******************************************************/

function unique_(array) {
  const seen = {};
  const result = [];

  (array || []).forEach(function(value) {
    const key = String(value);

    if (!key) return;

    if (seen[key]) return;

    seen[key] = true;
    result.push(value);
  });

  return result;
}

/*******************************************************
 * キャッシュ検索
 *******************************************************/

function findCacheByRowKey_(
  year,
  rank,
  workTitle,
  artist,
  embeddedId
) {

  const sheet =
    getCacheSheet_();

  const values =
    sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    const row = values[i];

    if (
      Number(row[1]) === Number(year) &&
      Number(row[2]) === Number(rank)
    ) {

      return {

        row: i + 1,

        key: row[0],

        year: row[1],

        rank: row[2],

        detailUrl: row[3],

        songTitle: row[4],

        workTitle: row[5],

        artist: row[6],

        type: row[7],

        releaseDate: row[8],

        embeddedIds:
          row[9]
            ? String(row[9]).split(',')
            : []

      };
    }
  }

  return null;
}


/*******************************************************
 * メインシートへ反映
 *******************************************************/

function writeRankingToSheet_(items) {

  const sheet = getMainSheet_();

  const lastRow =
    Math.max(sheet.getLastRow(), 1);

  const existing =
    lastRow > 1
      ? sheet.getRange(
          2,
          1,
          lastRow - 1,
          23
        ).getValues()
      : [];

  /*
   * 今回の処理ですでに使用した行
   *
   * 同じ作品・アーティストの複数曲が
   * 同じ行に入ることを防ぐ。
   */
  const usedRows = {};


  items.forEach(function(item) {

    const itemIds =
      unique_(item.embeddedIds || []);

    let targetRow = 0;


    /***************************************************
     * 1. E列の埋め込みIDで完全一致を探す
     ***************************************************/

    if (itemIds.length) {

      for (
        let i = 0;
        i < existing.length;
        i++
      ) {

        const actualRow =
          i + 2;

        if (usedRows[actualRow]) {
          continue;
        }

        const row =
          existing[i];

        const rowIds =
          parseVideoIds_(row[4]);

        let matched = false;

        for (
          let j = 0;
          j < itemIds.length;
          j++
        ) {

          if (
            rowIds.indexOf(itemIds[j]) !== -1
          ) {

            matched = true;
            break;
          }
        }

        if (!matched) {
          continue;
        }


        /*
         * 作品名も確認
         */
        const rowWork =
          String(row[16] || '');

        if (
          rowWork &&
          rowWork !== item.workTitle
        ) {
          continue;
        }


        targetRow =
          actualRow;

        break;
      }
    }


    /***************************************************
     * 2. E列で見つからなければ
     *    作品＋アーティスト＋順位
     ***************************************************/

    if (!targetRow) {

      for (
        let i = 0;
        i < existing.length;
        i++
      ) {

        const actualRow =
          i + 2;

        if (usedRows[actualRow]) {
          continue;
        }

        const row =
          existing[i];

        const rowArtist =
          String(row[1] || '');

        const rowWork =
          String(row[16] || '');

        const rowRank =
          Number(row[18] || 0);

        if (
          rowRank === Number(item.rank) &&
          rowWork === item.workTitle &&
          (
            !rowArtist ||
            rowArtist === item.artist
          )
        ) {

          targetRow =
            actualRow;

          break;
        }
      }
    }


    /***************************************************
     * 3. 新規行
     ***************************************************/

    if (!targetRow) {

      targetRow =
        sheet.getLastRow() + 1;

      sheet
        .getRange(
          targetRow,
          1,
          1,
          23
        )
        .setValues([[
          false,                         // A
          item.artist,                   // B
          item.releaseDate,              // C
          '',                            // D
          itemIds.join(','),             // E
          '',                            // F
          '',                            // G
          '',                            // H
          '',                            // I
          '',                            // J
          '',                            // K
          '',                            // L
          '',                            // M
          '',                            // N
          '',                            // O
          '',                            // P
          item.workTitle,                // Q
          item.type,                     // R
          item.rank,                     // S
          '未検索',                      // T
          '',                            // U
          '',                            // V
          ''                             // W
        ]]);

    } else {

      usedRows[targetRow] = true;


      /*************************************************
       * 既存行
       *
       * F / G:N / O / P は触らない
       *************************************************/

      sheet
        .getRange(targetRow, 2)
        .setValue(item.artist);

      sheet
        .getRange(targetRow, 3)
        .setValue(item.releaseDate);


      /*
       * E列は既存ID＋今回取得ID
       */
      sheet
        .getRange(targetRow, 5)
        .setValue(
          mergeVideoIds_(
            sheet
              .getRange(targetRow, 5)
              .getValue(),
            itemIds
          )
        );


      sheet
        .getRange(targetRow, 17)
        .setValue(item.workTitle);

      sheet
        .getRange(targetRow, 18)
        .setValue(item.type);

      sheet
        .getRange(targetRow, 19)
        .setValue(item.rank);
    }

  });
}

/*******************************************************
 * 動画ID結合
 *******************************************************/

function mergeVideoIdsLegacy_(
  current,
  newIds
) {

  const ids = [];

  String(current || '')
    .split(',')
    .forEach(function(id) {

      id = id.trim();

      if (
        id &&
        ids.indexOf(id) === -1
      ) {
        ids.push(id);
      }
    });


  (newIds || [])
    .forEach(function(id) {

      id = String(id || '').trim();

      if (
        id &&
        ids.indexOf(id) === -1
      ) {
        ids.push(id);
      }
    });


  return ids.join(',');
}



function startInitialImport() {

  const props =
    PropertiesService
      .getScriptProperties();

  props.setProperty(
    'ANISON_INITIAL_YEAR',
    String(CONFIG.START_YEAR)
  );

  props.setProperty(
    'ANISON_INITIAL_DONE',
    'false'
  );

  processInitialImport();
}


/*******************************************************
 * 初回取得
 *
 * 2008 → 2026
 *
 * ランキング取得後、
 * 各曲の作品ページから
 * ・曲名
 * ・アーティスト
 * ・区分
 * ・音源公開日
 * ・埋め込みYouTube ID
 * を確定してから保存する。
 *******************************************************/
function processInitialImport() {

  const start =
    Date.now();

  const props =
    PropertiesService
      .getScriptProperties();

  let year =
    Number(
      props.getProperty(
        'ANISON_INITIAL_YEAR'
      ) ||
      CONFIG.START_YEAR
    );


  while (
    year <= CONFIG.END_YEARS
  ) {

    /*
     * 実行時間制限
     */
    if (
      Date.now() - start >
      CONFIG.MAX_RUNTIME
    ) {

      props.setProperty(
        'ANISON_INITIAL_YEAR',
        String(year)
      );

      Logger.log(
        '[時間上限] 次回は ' +
        year +
        '年から再開'
      );

      return;
    }


    Logger.log(
      '================================'
    );

    Logger.log(
      '[初回取得] ' +
      year
    );


    try {

      const url =
        CONFIG.BASE_URL.replace(
          '{YEAR}',
          year
        );


      /*
       * 年間ランキング
       */
      const html =
        fetchHtml_(url);


      let items =
        parseRanking_(
          html,
          year,
          10000
        );
      
      Logger.log(
        '[ランキング取得件数] ' +
        year +
        ': ' +
        items.length  
      );
      
      
      /*
       * 曲名・公開日・埋め込みYouTube IDなどを補完
       */
      
      items =
        enrichRankingItems_(
          items
        );


      Logger.log(
        '[詳細情報補完完了] ' +
        year
      );


      saveCache_(items);

      writeRankingToSheet_(items);


      /*
       * 作品ページ解析
       *
       * parseRanking_() の結果を
       * enrichFromAnimePage_() で確定する。
       */
      for (
        let i = 0;
        i < items.length;
        i++
      ) {

        /*
         * 実行時間確認
         */
        if (
          Date.now() - start >
          CONFIG.MAX_RUNTIME
        ) {

          props.setProperty(
            'ANISON_INITIAL_YEAR',
            String(year)
          );

          Logger.log(
            '[時間上限] ' +
            year +
            '年の途中で停止'
          );

          return;
        }


        const item =
          items[i];


        Logger.log(
          '[作品解析] ' +
          year +
          ' / #' +
          item.rank +
          ' / ' +
          item.detailUrl
        );


        items[i] =
          enrichFromAnimePage_(
            item
          );
      }


      /*
       * 年完了
       */
      year++;

      props.setProperty(
        'ANISON_INITIAL_YEAR',
        String(year)
      );


      Logger.log(
        '[初回取得完了] ' +
        (year - 1)
      );

    } catch (e) {

      Logger.log(
        '[ERROR] ' +
        year +
        ': ' +
        e.stack
      );

      /*
       * この年は次回再試行
       */
      props.setProperty(
        'ANISON_INITIAL_YEAR',
        String(year)
      );

      return;
    }
  }


  props.setProperty(
    'ANISON_INITIAL_DONE',
    'true'
  );


  Logger.log(
    '[初回取得] ' +
    CONFIG.START_YEAR +
    '～' +
    CONFIG.END_YEARS +
    ' 完了'
  );
}

/**
 * D列の代表動画IDを確定し、
 * その動画のYouTube情報を取得してG:Nへ書き込む。
 *
 * 代表動画IDはYouTube検索では決めない。
 * anison.onlineの埋め込み動画IDをそのまま使用する。
 */
function processYouTubeQueue() {
  const sheet = getMainSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('[YouTube動画情報] 対象データなし');
    return;
  }

  const cacheRecords = getCacheRecords_();

  const values = sheet
    .getRange(2, 1, lastRow - 1, 23)
    .getValues();

  const targets = [];

  Logger.log('================================');
  Logger.log('[YouTube動画情報取得] 開始');
  Logger.log('[対象行] ' + values.length);
  Logger.log('================================');

  /*
   * まず各行についてD列を確定する。
   *
   * D = anison.online埋め込み動画ID
   * E = 既存ID + 埋め込みIDを統合
   */
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowNumber = i + 2;

    const cache = findCacheForSheetRow_(row, cacheRecords);

    if (!cache) {
      Logger.log('[キャッシュなし] 行=' + rowNumber);

      // 既存Dがある場合は、その動画情報だけ取得対象にする
      const existingVideoId = String(row[3] || '').trim();

      if (existingVideoId) {
        targets.push({
          rowNumber: rowNumber,
          videoId: existingVideoId,
          source: '既存D列'
        });
      }

      continue;
    }

    const embeddedIds = unique_(
      cache.embeddedIds || []
    ).filter(function(id) {
      return String(id || '').trim() !== '';
    });

    if (!embeddedIds.length) {
      Logger.log(
        '[埋め込み動画なし] 行=' +
        rowNumber +
        ' / 曲=' +
        String(cache.songTitle || '')
      );

      // Dは空にする
      sheet.getRange(rowNumber, 4).clearContent();

      // Aは動画なしとしてFALSE
      sheet.getRange(rowNumber, 1).setValue(false);

      sheet.getRange(rowNumber, 20, 1, 4).setValues([[
        '埋め込み動画なし',
        '',
        '',
        ''
      ]]);

      continue;
    }

    /*
     * 埋め込みIDの先頭を代表動画とする。
     */
    const representativeId = embeddedIds[0];

    /*
     * D列 = 代表動画ID
     */
    sheet
      .getRange(rowNumber, 4)
      .setValue(representativeId);

    /*
     * E列 = 既存ID + 埋め込みID
     */
    const existingIds = parseVideoIds_(row[4]);

    const mergedIds = unique_(
      existingIds.concat(embeddedIds)
    );

    sheet
      .getRange(rowNumber, 5)
      .setValue(mergedIds.join(','));

    /*
     * 有効
     */
    sheet
      .getRange(rowNumber, 1)
      .setValue(true);

    /*
     * YouTube API取得対象
     */
    targets.push({
      rowNumber: rowNumber,
      videoId: representativeId,
      source: 'anison.online埋め込み'
    });

    Logger.log(
      '[代表動画決定] 行=' +
      rowNumber +
      ' / ID=' +
      representativeId
    );
  }

  if (!targets.length) {
    Logger.log('[YouTube動画情報] 取得対象なし');
    return;
  }

  /*
   * 同じ動画IDが複数行に存在する可能性があるため、
   * API取得はID単位で重複排除する。
   */
  const videoIdMap = {};

  targets.forEach(function(target) {
    const id = String(target.videoId || '').trim();

    if (!id) return;

    if (!videoIdMap[id]) {
      videoIdMap[id] = [];
    }

    videoIdMap[id].push(target.rowNumber);
  });

  const videoIds = Object.keys(videoIdMap);

  Logger.log(
    '[YouTube API取得対象] ' +
    videoIds.length +
    '動画'
  );

  /*
   * videos.list は最大50 IDをまとめて取得。
   */
  for (
    let start = 0;
    start < videoIds.length;
    start += CONFIG.YOUTUBE_BATCH_SIZE
  ) {
    const batchIds = videoIds.slice(
      start,
      start + CONFIG.YOUTUBE_BATCH_SIZE
    );

    Logger.log(
      '[YouTube API] ' +
      (start + 1) +
      '～' +
      Math.min(
        start + CONFIG.YOUTUBE_BATCH_SIZE,
        videoIds.length
      ) +
      ' / ' +
      videoIds.length
    );

    let result;

    try {
      result = getYouTubeVideosByIds_(batchIds);
    } catch (e) {
      Logger.log(
        '[YouTube APIエラー] ' +
        e.message
      );

      /*
       * バッチ内の行をエラーにする。
       */
      batchIds.forEach(function(videoId) {
        const rows = videoIdMap[videoId] || [];

        rows.forEach(function(rowNumber) {
          sheet
            .getRange(rowNumber, 20, 1, 4)
            .setValues([[
              'エラー',
              '',
              'YouTube API取得失敗',
              e.message
            ]]);
        });
      });

      continue;
    }

    const foundMap = {};

    result.forEach(function(video) {
      const videoId = String(video.id || '');

      if (!videoId) return;

      foundMap[videoId] = video;

      const rows = videoIdMap[videoId] || [];

      rows.forEach(function(rowNumber) {
        writeYouTubeVideoInfo_(
          sheet,
          rowNumber,
          video
        );
      });
    });

    /*
     * APIから返ってこなかったIDを処理。
     *
     * 削除済み・非公開などの場合は
     * itemsに含まれない可能性がある。
     */
    batchIds.forEach(function(videoId) {
      if (foundMap[videoId]) return;

      const rows = videoIdMap[videoId] || [];

      rows.forEach(function(rowNumber) {
        sheet
          .getRange(rowNumber, 20, 1, 4)
          .setValues([[
            '動画取得失敗',
            '',
            'YouTube動画が取得できません',
            '動画が削除済み、非公開、または存在しない可能性があります'
          ]]);
      });

      Logger.log(
        '[動画取得失敗] ID=' +
        videoId
      );
    });
  }

  Logger.log('================================');
  Logger.log(
    '[YouTube動画情報取得] 完了 / ' +
    videoIds.length +
    '動画'
  );
  Logger.log('================================');
}

/*******************************************************
 * YouTube検索
 *******************************************************/

function searchYouTubeForTrack_(
  songTitle,
  artist,
  embeddedIds
) {

  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'YOUTUBE_API_KEY'
      );


  if (!apiKey) {

    throw new Error(
      'YOUTUBE_API_KEY がScript Propertiesにありません'
    );
  }


  /*
   * 検索語
   */
  const query =
    songTitle +
    ' ' +
    artist;


  const url =
    'https://www.googleapis.com/youtube/v3/search' +
    '?part=snippet' +
    '&type=video' +
    '&maxResults=' +
    CONFIG.YOUTUBE_MAX_RESULTS +
    '&q=' +
    encodeURIComponent(query) +
    '&key=' +
    encodeURIComponent(apiKey);


  const response =
    UrlFetchApp.fetch(
      url,
      {
        muteHttpExceptions: true
      }
    );


  const code =
    response.getResponseCode();

  const json =
    JSON.parse(
      response.getContentText()
    );


  if (code !== 200) {

    throw new Error(
      'YouTube API HTTP ' +
      code +
      ': ' +
      JSON.stringify(json)
    );
  }


  const items =
    json.items || [];


  if (!items.length) {

    return {
      found: false,
      score: 0,
      reason: '検索結果なし'
    };
  }


  let best = null;


  items.forEach(function(item) {

    const id =
      item.id &&
      item.id.videoId;


    if (!id) return;


    const title =
      item.snippet &&
      item.snippet.title
        ? item.snippet.title
        : '';


    const channel =
      item.snippet &&
      item.snippet.channelTitle
        ? item.snippet.channelTitle
        : '';


    const score =
      scoreYouTubeCandidate_(
        songTitle,
        artist,
        title,
        channel,
        embeddedIds,
        id
      );


    if (
      !best ||
      score.score > best.score
    ) {

      best = {

        videoId: id,

        score: score.score,

        reason: score.reason,

        title: title,

        channel: channel
      };
    }

    Logger.log(
      '[YouTube候補] ' +
      songTitle +
      ' / ' +
      artist +
      ' / ' +
      id +
      ' / ' +
      title +
      ' / ' +
      channel +
      ' / score=' +
      score.score +
      ' / ' +
      score.reason
    );

  });


  if (!best) {

    return {
      found: false,
      score: 0,
      reason: '候補なし'
    };
  }


  /*
   * 自動生成Topic/Music候補として
   * 十分高いものだけ採用
   */
  if (best.score < 65) {

    return {
      found: false,

      score: best.score,

      reason:
        'スコア不足: ' +
        best.reason
    };
  }


  return {

    found: true,

    videoId: best.videoId,

    score: best.score,

    reason: best.reason
  };
}


/*******************************************************
 * YouTube候補スコア
 *******************************************************/
function scoreYouTubeCandidate_(
  songTitle,
  artist,
  title,
  channel,
  embeddedIds,
  candidateId
) {

  let score = 0;

  const reasons = [];

  const s =
    normalizeText_(songTitle);

  const a =
    normalizeText_(artist);

  const t =
    normalizeText_(title);

  const c =
    normalizeText_(channel);


  /*
   * ---------------------------------------------------
   * 曲名
   * ---------------------------------------------------
   */

  if (
    s &&
    t.indexOf(s) !== -1
  ) {

    score += 40;

    reasons.push(
      '曲名一致'
    );
  }


  /*
   * 曲名の完全一致に近い場合は追加点
   *
   * 例:
   * 逆夢 - King Gnu
   */
  if (
    s &&
    (
      t === s ||
      t.indexOf(s + '-') === 0 ||
      t.indexOf(s + '–') === 0 ||
      t.indexOf(s + '—') === 0
    )
  ) {

    score += 10;

    reasons.push(
      '曲名強一致'
    );
  }


  /*
   * ---------------------------------------------------
   * アーティスト
   * ---------------------------------------------------
   */

  if (
    a &&
    t.indexOf(a) !== -1
  ) {

    score += 20;

    reasons.push(
      'タイトルにアーティスト'
    );
  }


  if (
    a &&
    c.indexOf(a) !== -1
  ) {

    score += 15;

    reasons.push(
      'チャンネル名にアーティスト'
    );
  }


  /*
   * ---------------------------------------------------
   * Topic
   * ---------------------------------------------------
   */

  if (
    /-topic$/i.test(
      String(channel || '').trim()
    )
  ) {

    score += 30;

    reasons.push(
      'Topicチャンネル'
    );
  }


  /*
   * ---------------------------------------------------
   * Provided to YouTube
   * ---------------------------------------------------
   */

  if (
    /provided\s+to\s+youtube/i.test(
      title
    )
  ) {

    score += 25;

    reasons.push(
      'Provided to YouTube'
    );
  }


  /*
   * ---------------------------------------------------
   * 自動生成トラックではない可能性が高い候補
   * ---------------------------------------------------
   */

  if (
    /official\s+(music\s+)?video/i.test(
      title
    )
  ) {

    score -= 20;

    reasons.push(
      'Official Video'
    );
  }


  if (
    /official\s+mv/i.test(
      title
    )
  ) {

    score -= 20;

    reasons.push(
      'Official MV'
    );
  }


  if (
    /歌ってみた|cover|カバー/i.test(
      title
    )
  ) {

    score -= 30;

    reasons.push(
      '歌ってみた/Cover'
    );
  }


  if (
    /remix|リミックス/i.test(
      title
    )
  ) {

    score -= 20;

    reasons.push(
      'Remix'
    );
  }


  /*
   * ---------------------------------------------------
   * anison.online埋め込み動画
   *
   * これは完全な除外条件ではないので、
   * 弱い減点にする。
   * ---------------------------------------------------
   */

  if (
    embeddedIds &&
    embeddedIds.indexOf(candidateId) !== -1
  ) {

    score -= 5;

    reasons.push(
      'サイト埋込動画と同一'
    );
  }


  return {

    score:
      Math.max(
        0,
        score
      ),

    reason:
      reasons.join(' / ') ||
      '一致条件なし'
  };
}

/*******************************************************
 * 文字列正規化
 *******************************************************/

function normalizeText_(text) {

  return String(text || '')
    .toLowerCase()
    .replace(
      /[\s　・「」『』【】（）()［］\[\]「」]/g,
      ''
    )
    .replace(
      /feat\.?|featuring|with/gi,
      ''
    );
}


/*******************************************************
 * 週間更新
 *******************************************************/

function weeklyUpdate() {

  const start =
    Date.now();

  Logger.log(
    '================================'
  );

  Logger.log(
    '[週間更新] 開始'
  );


  for (
    let year = CONFIG.START_YEAR;
    year <= CONFIG.END_YEARS;
    year++
  ) {

    if (
      Date.now() - start >
      CONFIG.MAX_RUNTIME
    ) {

      Logger.log(
        '[時間上限] 更新を終了'
      );

      return;
    }


    try {

      const url =
        CONFIG.BASE_URL.replace(
          '{YEAR}',
          year
        );


      const html =
        fetchHtml_(url);


      const items =
        parseRanking_(
          html,
          year,
          10000
        );


      saveCache_(items);

      updateExistingRanking_(
        items
      );


    } catch (e) {

      Logger.log(
        '[週間更新 ERROR] ' +
        year +
        ': ' +
        e.message
      );
    }
  }


  Logger.log(
    '[週間更新] 完了'
  );
}


/*******************************************************
 * 週間ランキング更新
 *******************************************************/

function updateExistingRanking_(items) {

  const sheet =
    getMainSheet_();

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    writeRankingToSheet_(items);

    return;
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        23
      )
      .getValues();


  /*
   * 同じ行を2曲に割り当てない
   */
  const usedRows = {};


  items.forEach(function(item) {

    const itemIds =
      unique_(
        item.embeddedIds || []
      );

    let foundRow = 0;


    /***************************************************
     * 1. 埋め込みYouTube IDで検索
     ***************************************************/

    if (itemIds.length) {

      for (
        let i = 0;
        i < values.length;
        i++
      ) {

        const actualRow =
          i + 2;

        if (usedRows[actualRow]) {
          continue;
        }

        const row =
          values[i];

        const rowIds =
          parseVideoIds_(row[4]);

        let matched = false;

        for (
          let j = 0;
          j < itemIds.length;
          j++
        ) {

          if (
            rowIds.indexOf(itemIds[j]) !== -1
          ) {

            matched = true;
            break;
          }
        }

        if (!matched) {
          continue;
        }


        /*
         * 作品名も確認
         */
        const work =
          String(row[16] || '');

        if (
          work &&
          work !== item.workTitle
        ) {
          continue;
        }


        foundRow =
          actualRow;

        break;
      }
    }


    /***************************************************
     * 2. IDで見つからなければ
     *    作品＋アーティスト＋順位
     ***************************************************/

    if (!foundRow) {

      for (
        let i = 0;
        i < values.length;
        i++
      ) {

        const actualRow =
          i + 2;

        if (usedRows[actualRow]) {
          continue;
        }

        const row =
          values[i];

        const work =
          String(row[16] || '');

        const artist =
          String(row[1] || '');

        const rank =
          Number(row[18] || 0);

        if (
          work === item.workTitle &&
          artist === item.artist &&
          rank === Number(item.rank)
        ) {

          foundRow =
            actualRow;

          break;
        }
      }
    }


    /***************************************************
     * 3. 見つからない → 新曲
     ***************************************************/

    if (!foundRow) {

      writeRankingToSheet_([
        item
      ]);

      return;
    }


    usedRows[foundRow] = true;


    /***************************************************
     * 既存行を更新
     ***************************************************/

    sheet
      .getRange(foundRow, 2)
      .setValue(item.artist);

    sheet
      .getRange(foundRow, 3)
      .setValue(item.releaseDate);


    /*
     * E列はIDを追加するだけ
     */
    sheet
      .getRange(foundRow, 5)
      .setValue(
        mergeVideoIds_(
          sheet
            .getRange(foundRow, 5)
            .getValue(),
          itemIds
        )
      );


    /*
     * Q～S
     */
    sheet
      .getRange(foundRow, 17)
      .setValue(item.workTitle);

    sheet
      .getRange(foundRow, 18)
      .setValue(item.type);

    sheet
      .getRange(foundRow, 19)
      .setValue(item.rank);

  });
}

/*******************************************************
 * 作品ページから楽曲情報を確定
 *******************************************************/
function enrichFromAnimePage_(
  item,
  htmlOverride
) {

  if (!item.detailUrl) {
    return item;
  }

  try {

    const match =
      item.detailUrl.match(/#(song\d+)$/i);

    const songAnchor =
      match ? match[1] : '';

    const pageUrl =
    normalizeDetailUrl_(
      item.detailUrl
    );

    const html =
    htmlOverride ||
    fetchHtml_(pageUrl);
    
    /*
     * 音源公開日は作品ページ全体から取得
     */
    const text =
      stripTags_(html);

    const releaseDate =
      extractReleaseDate_(text);

    if (releaseDate) {
      item.releaseDate = releaseDate;
    }

    /*
     * 楽曲単位で情報を取得
     */
    if (songAnchor) {

      const songInfo =
        extractSongByAnchor_(
          html,
          songAnchor
        );

      if (songInfo) {

        /*
         * ランキングページで取得済みの値を優先。
         * 詳細ページは空欄の補完だけにする。
         */
        if (
          (!item.songTitle ||
          isProbablyUiText_(item.songTitle)) &&
          songInfo.songTitle &&
          !isProbablyUiText_(songInfo.songTitle)
        ) {
          item.songTitle =
          songInfo.songTitle;
        }

        if (!item.artist && songInfo.artist) {
          item.artist =
            songInfo.artist;
        }

        if (!item.type && songInfo.type) {
          item.type =
            songInfo.type;
        }

        /*
         * YouTube IDも対象曲だけ追加
         */
        if (
          songInfo.youtubeIds &&
          songInfo.youtubeIds.length
        ) {

          item.embeddedIds =
            unique_(
              (item.embeddedIds || [])
                .concat(songInfo.youtubeIds)
            );
        }
      }
    }

  } catch (e) {

    Logger.log(
      '[作品ページ解析失敗] ' +
      item.detailUrl +
      ' / ' +
      e.message
    );
  }

  return item;
}

/*******************************************************
 * 作品ページから特定 songXXXX の情報を取得
 *******************************************************/
function extractSongByAnchor_(html, songAnchor) {

  if (!html || !songAnchor) {
    return null;
  }

  const result = {
    songTitle: '',
    artist: '',
    type: '',
    youtubeIds: []
  };

  try {

    /*
     * songXXXX を含むリンクを探す
     *
     * 例:
     * href="/anime/1875#song4177"
     */
    const linkRe =
      new RegExp(
        '<a\\b[^>]*href=["\\\'][^"\\\']*#' +
        escapeRegExp_(songAnchor) +
        '["\\\'][^>]*>[\\s\\S]*?<\\/a>',
        'i'
      );

    const linkMatch =
      linkRe.exec(html);

    if (!linkMatch) {

      Logger.log(
        '[曲リンクなし] ' +
        songAnchor
      );

      return result;
    }

    const linkHtml =
      linkMatch[0];

    const linkPos =
      linkMatch.index;

    /*
     * リンクそのものの表示文字列
     */
    let linkText =
      cleanSongText_(
        linkHtml
          .replace(/<img\b[^>]*>/gi, '')
          .replace(/<svg\b[\s\S]*?<\/svg>/gi, '')
      );

    /*
     * href直後～少し後ろを取得。
     * artistリンクや区分がこの近辺に存在するケースに対応。
     */
    const nearbyStart =
      Math.max(
        0,
        linkPos - 1500
      );

    const nearbyEnd =
      Math.min(
        html.length,
        linkPos + 4000
      );

    const nearby =
      html.substring(
        nearbyStart,
        nearbyEnd
      );

    /***************************************************
     * アーティスト
     ***************************************************/
    const artistRe =
      /<a\b[^>]*href=["'][^"']*\/artist\/[^"']*["'][^>]*>([\s\S]*?)<\/a>/gi;

    let artistMatch;

    /*
     * songリンクより後ろを優先
     */
    const afterLink =
      html.substring(
        linkPos,
        Math.min(
          html.length,
          linkPos + 2500
        )
      );

    artistMatch =
      artistRe.exec(afterLink);

    if (artistMatch) {

      result.artist =
        cleanSongText_(
          artistMatch[1]
        );
    }

    /*
     * 後ろで取れなければ周辺HTMLから探す
     */
    if (!result.artist) {

      artistRe.lastIndex = 0;

      artistMatch =
        artistRe.exec(nearby);

      if (artistMatch) {

        result.artist =
          cleanSongText_(
            artistMatch[1]
          );
      }
    }

    /***************************************************
     * 曲タイトル
     ***************************************************/
    /*
     * まずリンク自身のテキスト
     */
    if (linkText) {

      /*
       * 明らかにUI文字列の場合は除外
       */
      if (
        !isProbablyUiText_(linkText)
      ) {

        result.songTitle =
          linkText;
      }
    }

    /*
     * リンク文字列が空の場合、
     * 対象songの詳細ブロックを探す。
     */
    if (!result.songTitle) {

      result.songTitle =
        extractSongTitleFromBlock_(
          html,
          songAnchor
        );
    }

    /***************************************************
     * 区分
     ***************************************************/
    /*
     * 「最初に出てきたOP」を拾う方式はやめる。
     *
     * songリンクの直前にある短いテキストを調べる。
     */
    result.type =
      extractSongTypeNearLink_(
        html,
        linkPos
      );

    /***************************************************
     * YouTube ID
     ***************************************************/
    /*
     * 作品ページ全体ではなく、
     * 対象曲の詳細ブロックだけから取得する。
     */
    result.youtubeIds =
      extractSongYoutubeIds_(
        html,
        songAnchor
      );

    Logger.log(
      '[曲情報] ' +
      songAnchor +
      ' / ' +
      result.songTitle +
      ' / ' +
      result.artist +
      ' / ' +
      result.type +
      ' / YouTube=' +
      JSON.stringify(result.youtubeIds)
    );

    return result;

  } catch (e) {

    Logger.log(
      '[extractSongByAnchor_失敗] ' +
      songAnchor +
      ' / ' +
      e.message
    );

    return result;
  }
}

/*******************************************************
 * 作品ページのアニソン情報
 *******************************************************/

function extractAnimeSongInfo_(html) {

  /*
   * 「アニソン一覧」の位置
   */
  const pos =
    html.search(
      /アニソン一覧|Anime Song List/i
    );


  if (pos < 0) {
    return null;
  }


  /*
   * アニソン一覧以降だけ見る
   */
  const area =
    html.substring(
      pos,
      Math.min(
        html.length,
        pos + 15000
      )
    );


  /*
   * 区分
   */
  let type = '';

  const typeMatch =
    area.match(
      /(?:主題歌|挿入歌|OP|ED|Opening|Ending)/
    );

  if (typeMatch) {
    type =
      typeMatch[0];
  }


  /*
   * 曲名
   *
   * h2タグを優先
   */
  let songTitle = '';

  let m =
    area.match(
      /<h2[^>]*>\s*([^<]+?)\s*<\/h2>/i
    );

  if (m) {

    songTitle =
      stripTags_(m[1]);
  }


  /*
   * h2がない場合、曲名リンクを探す
   */
  if (!songTitle) {

    m =
      area.match(
        /<a\b[^>]*href=["'][^"']*\/song[^"']*["'][^>]*>([\s\S]*?)<\/a>/i
      );

    if (m) {

      songTitle =
        stripTags_(m[1]);
    }
  }


  /*
   * アーティスト
   */
  let artist = '';

  /*
   * songTitleの後ろにあるartistリンクを探す
   */
  if (songTitle) {

    const songPos =
      area.indexOf(
        songTitle
      );

    if (songPos >= 0) {

      const afterSong =
        area.substring(
          songPos,
          Math.min(
            area.length,
            songPos + 5000
          )
        );


      m =
        afterSong.match(
          /<a\b[^>]*href=["'][^"']*\/artist\/[^"']*["'][^>]*>([\s\S]*?)<\/a>/i
        );


      if (m) {

        artist =
          stripTags_(m[1]);
      }
    }
  }


  /*
   * artistリンクがない場合
   */
  if (!artist) {

    /*
     * animeページではアーティスト名が
     * songTitleの直後に存在するケースが多い
     */
    const text =
      stripTags_(area);


    if (songTitle) {

      const p =
        text.indexOf(
          songTitle
        );

      if (p >= 0) {

        const after =
          text.substring(
            p + songTitle.length
          )
          .trim();

        const parts =
          after
            .split(/\s+/)
            .filter(Boolean);

        if (parts.length) {

          artist =
            parts[0];
        }
      }
    }
  }


  return {

    songTitle:
      songTitle,

    artist:
      artist,

    type:
      type
  };
}

/*******************************************************
 * 楽曲タイトル文字列を整理
 *******************************************************/
function cleanSongText_(value) {

  if (!value) {
    return '';
  }

  return stripTags_(value)
    .replace(/\s+/g, ' ')
    .replace(/^[\s　]+|[\s　]+$/g, '')
    .trim();
}


function isProbablyUiText_(text) {

  if (!text) {
    return true;
  }

  const t =
    text
      .replace(/\s+/g, ' ')
      .trim();

  const bad = [
    'OP',
    'ED',
    '主題歌',
    '挿入歌',
    'Opening',
    'Ending',

    'アニソン一覧',
    'アーティスト一覧',
    '作品一覧',
    '一覧',
    '詳細',
    '歌詞',
    '動画',
    '動画一覧',
    '関連動画',
    'YouTube',
    '公式サイト',

    '検索',
    'ログイン',
    'メニュー',
    'ホーム'
  ];

  return bad.indexOf(t) !== -1;
}

/*******************************************************
 * songXXXX の詳細ブロックからタイトルを取得
 *******************************************************/
function extractSongTitleFromBlock_(html, songAnchor) {

  if (!html || !songAnchor) {
    return '';
  }

  /*
   * songXXXX の数値
   */
  const songId =
    songAnchor.replace(/^song/i, '');

  /*
   * ---------------------------------------------------
   * 1. href="#songXXXX" のリンクを直接探す
   * ---------------------------------------------------
   *
   * このページでは、
   *
   *   href="...#song4177"
   *
   * のリンクが対象曲への入口になっている。
   *
   * まずこれを最優先する。
   */
  const hrefRe =
    new RegExp(
      '<a\\b[^>]*href=["\\\'][^"\\\']*#' +
      escapeRegExp_(songAnchor) +
      '["\\\'][^>]*>([\\s\\S]*?)<\\/a>',
      'gi'
    );

  let match;

  while (
    (match = hrefRe.exec(html))
  ) {

    const text =
      cleanSongText_(
        match[1]
      );

    if (
      text &&
      !isProbablyUiText_(text) &&
      text.length <= 200
    ) {

      Logger.log(
        '[曲タイトル直接取得] ' +
        songAnchor +
        ' → ' +
        text
      );

      return text;
    }
  }


  /*
   * ---------------------------------------------------
   * 2. id="songXXXX" / name="songXXXX"
   * ---------------------------------------------------
   */
  const anchorRe =
    new RegExp(
      '(?:id|name)\\s*=\\s*["\\\']' +
      escapeRegExp_(songAnchor) +
      '["\\\']',
      'i'
    );

  const anchorMatch =
    anchorRe.exec(html);

  if (!anchorMatch) {

    Logger.log(
      '[曲タイトルアンカーなし] ' +
      songAnchor
    );

    return '';
  }

  const start =
    anchorMatch.index;


  /*
   * 次の songXXXX までを対象範囲にする
   */
  const nextRe =
    /(?:id|name)\s*=\s*["']song\d+["']/gi;

  nextRe.lastIndex =
    start +
    anchorMatch[0].length;

  const nextMatch =
    nextRe.exec(html);

  const end =
    nextMatch
      ? nextMatch.index
      : Math.min(
          html.length,
          start + 12000
        );

  const block =
    html.substring(
      start,
      end
    );


  /*
   * ---------------------------------------------------
   * 3. 対象 block 内の href="#songXXXX"
   * ---------------------------------------------------
   */
  const blockLinkRe =
    new RegExp(
      '<a\\b[^>]*href=["\\\'][^"\\\']*#' +
      escapeRegExp_(songAnchor) +
      '["\\\'][^>]*>([\\s\\S]*?)<\\/a>',
      'i'
    );

  const blockLink =
    blockLinkRe.exec(block);

  if (blockLink) {

    const text =
      cleanSongText_(
        blockLink[1]
      );

    if (
      text &&
      !isProbablyUiText_(text) &&
      text.length <= 200
    ) {

      Logger.log(
        '[曲タイトルblock取得] ' +
        songAnchor +
        ' → ' +
        text
      );

      return text;
    }
  }


  /*
   * ---------------------------------------------------
   * 4. 見出しを探す
   * ---------------------------------------------------
   */
  const headingRe =
    /<(?:h1|h2|h3|h4|h5)\b[^>]*>([\s\S]*?)<\/(?:h1|h2|h3|h4|h5)>/gi;

  let heading;

  while (
    (heading = headingRe.exec(block))
  ) {

    const text =
      cleanSongText_(
        heading[1]
      );

    if (
      text &&
      !isProbablyUiText_(text) &&
      text.length <= 200
    ) {

      Logger.log(
        '[曲タイトル見出し取得] ' +
        songAnchor +
        ' → ' +
        text
      );

      return text;
    }
  }


  return '';
}

/*******************************************************
 * songリンク付近から区分を取得
 *******************************************************/
function extractSongTypeNearLink_(html, linkPos) {

  /*
   * リンクより前の範囲だけを見る。
   * 作品ページ全体の先頭にあるOPを拾わないため。
   */
  const start =
    Math.max(
      0,
      linkPos - 1200
    );

  const before =
    html.substring(
      start,
      linkPos
    );

  /*
   * 最後に出てきた区分を取得
   */
  const typeRe =
    /(?:主題歌|挿入歌|OP|ED|Opening|Ending)/gi;

  let match;
  let last = '';

  while (
    (match = typeRe.exec(before))
  ) {
    last = match[0];
  }

  if (!last) {
    return '';
  }

  /*
   * 表記を統一
   */
  switch (
    last.toLowerCase()
  ) {

    case 'opening':
    case 'op':
      return 'OP';

    case 'ending':
    case 'ed':
      return 'ED';

    default:
      return last;
  }
}


/*******************************************************
 * 特定 songXXXX に対応する YouTube IDを取得
 *******************************************************/
function extractSongYoutubeIds_(html, songAnchor) {

  if (!html || !songAnchor) {
    return [];
  }

  const ids = [];

  const songId =
    songAnchor.replace(
      /^song/i,
      ''
    );

  /***************************************************
   * 1. data-song-id="XXXX" の要素を探す
   ***************************************************/
  const dataSongRe =
    new RegExp(
      'data-song-id\\s*=\\s*["\\\']' +
      escapeRegExp_(songId) +
      '["\\\']',
      'i'
    );

  const dataMatch =
    dataSongRe.exec(html);

  if (dataMatch) {

    /*
     * data-song-id を持つ要素そのものを探す。
     *
     * iframe の場合は、
     * iframe開始タグ～終了 > まで。
     */
    const before =
      html.substring(
        Math.max(
          0,
          dataMatch.index - 1000
        ),
        dataMatch.index + 1000
      );

    /*
     * data-song-id の近辺にあるiframe
     */
    const iframeRe =
      /<iframe\b[^>]*data-song-id\s*=\s*["'][^"']+["'][^>]*>/gi;

    let iframeMatch;

    while (
      (iframeMatch = iframeRe.exec(before))
    ) {

      const iframeHtml =
        iframeMatch[0];

      const iframeIds =
        extractYouTubeIds_(
          iframeHtml
        );

      if (iframeIds.length) {
        ids.push(
          ...iframeIds
        );
      }
    }

    /*
     * data-song-idを持つタグ自身のsrcも確認
     */
    const tagStart =
      Math.max(
        0,
        dataMatch.index - 1000
      );

    const tagEnd =
      Math.min(
        html.length,
        dataMatch.index + 1000
      );

    const around =
      html.substring(
        tagStart,
        tagEnd
      );

    /*
     * data-song-idから最も近いiframeを探す
     */
    const iframeAroundRe =
      /<iframe\b[^>]*>/gi;

    let m;

    while (
      (m = iframeAroundRe.exec(around))
    ) {

      const pos =
        m.index;

      /*
       * data-song-idから近いiframeだけ採用
       */
      const distance =
        Math.abs(
          (tagStart + pos) -
          dataMatch.index
        );

      if (distance <= 500) {

        const found =
          extractYouTubeIds_(
            m[0]
          );

        if (found.length) {
          ids.push(
            ...found
          );
        }
      }
    }
  }

  /***************************************************
   * 2. id="song-XXXX" の要素を探す
   ***************************************************/
  if (!ids.length) {

    const songIdRe =
      new RegExp(
        'id\\s*=\\s*["\\\']song-' +
        escapeRegExp_(songId) +
        '["\\\']',
        'i'
      );

    const songIdMatch =
      songIdRe.exec(html);

    if (songIdMatch) {

      /*
       * 対象要素の周辺だけを調べる
       */
      const start =
        Math.max(
          0,
          songIdMatch.index - 1000
        );

      const end =
        Math.min(
          html.length,
          songIdMatch.index + 2500
        );

      const area =
        html.substring(
          start,
          end
        );

      const found =
        extractYouTubeIds_(
          area
        );

      ids.push(
        ...found
      );
    }
  }

  /***************************************************
   * 3. 最終フォールバック
   ***************************************************/
  if (!ids.length) {

    const anchorRe =
      new RegExp(
        '(?:id|name)\\s*=\\s*["\\\']' +
        escapeRegExp_(songAnchor) +
        '["\\\']',
        'i'
      );

    const anchorMatch =
      anchorRe.exec(html);

    if (anchorMatch) {

      const start =
        anchorMatch.index;

      const nextRe =
        /(?:id|name)\s*=\s*["']song\d+["']/gi;

      nextRe.lastIndex =
        start +
        anchorMatch[0].length;

      const nextMatch =
        nextRe.exec(html);

      const end =
        nextMatch
          ? nextMatch.index
          : Math.min(
              html.length,
              start + 5000
            );

      const block =
        html.substring(
          start,
          end
        );

      ids.push(
        ...extractYouTubeIds_(
          block
        )
      );
    }
  }

  return unique_(ids);
}

/*******************************************************
 * detailUrlから songXXXX を取得
 *******************************************************/
function getSongKeyFromUrl_(url) {

  if (!url) {
    return '';
  }

  const m =
    String(url).match(
      /#(song\d+)$/i
    );

  return m
    ? m[1].toLowerCase()
    : '';
}

/*******************************************************
 * キャッシュから songXXXX で検索
 *******************************************************/
function findCacheBySongKey_(
  songKey
) {

  if (!songKey) {
    return null;
  }


  const sheet =
    getCacheSheet_();


  if (!sheet) {
    return null;
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return null;
  }


  const values =
    sheet.getRange(
      2,
      1,
      lastRow - 1,
      12
    ).getValues();


  const target =
    String(
      songKey
    ).toLowerCase();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    /*
     * D = DETAIL_URL
     */
    const detailUrl =
      String(
        values[i][3] || ''
      );


    const key =
      getSongKeyFromUrl_(
        detailUrl
      );


    if (
      key &&
      key === target
    ) {

      return {

        key: key,

        year:
          values[i][1],

        rank:
          values[i][2],

        detailUrl:
          values[i][3],

        songTitle:
          values[i][4],

        workTitle:
          values[i][5],

        artist:
          values[i][6],

        type:
          values[i][7],

        releaseDate:
          values[i][8],

        embeddedIds:
          values[i][9]
            ? String(values[i][9])
                .split(',')
                .map(function(x) {
                  return x.trim();
                })
                .filter(Boolean)
            : [],

        lastSeen:
          values[i][10],

        status:
          values[i][11]
      };
    }
  }


  return null;
}

/*******************************************************
 * キャッシュ全件取得
 *******************************************************/

function getCacheRecords_() {

  const sheet = getCacheSheet_();

  const values = sheet.getDataRange().getValues();

  const result = [];

  for (let i = 1; i < values.length; i++) {

    const row = values[i];

    if (!row[0]) {
      continue;
    }

    result.push({

      row: i + 1,

      key: row[0],

      year: row[1],

      rank: row[2],

      detailUrl: row[3],

      songTitle: row[4],

      workTitle: row[5],

      artist: row[6],

      type: row[7],

      releaseDate: row[8],

      embeddedIds: parseVideoIds_(row[9])

    });

  }

  return result;
}


/*******************************************************
 * 動画ID文字列を配列化
 *******************************************************/

function parseVideoIds_(value) {
  if (!value) return [];

  if (Array.isArray(value)) {
    return unique_(value);
  }

  const text = String(value).trim();

  if (!text) return [];

  try {
    const json = JSON.parse(text);

    if (Array.isArray(json)) {
      return unique_(json);
    }
  } catch (e) {}

  return unique_(
    text
      .split(',')
      .map(function(id) {
        return String(id).trim();
      })
      .filter(function(id) {
        return id !== '';
      })
  );
}

/*******************************************************
 * URL正規化
 *******************************************************/

function normalizeDetailUrl_(url) {

  return String(url || '')
    .trim()
    .replace(/#.*$/, '')
    .replace(/\/+$/, '');
}


/*******************************************************
 * 詳細URLからキャッシュを検索
 *******************************************************/

function findCacheByDetailUrl_(
  detailUrl,
  cacheRecords
) {

  const target =
    normalizeDetailUrl_(detailUrl);

  if (!target) {
    return null;
  }

  for (let i = 0; i < cacheRecords.length; i++) {

    const cache =
      cacheRecords[i];

    if (
      normalizeDetailUrl_(cache.detailUrl) === target
    ) {
      return cache;
    }
  }

  return null;
}


/*******************************************************
 * メインシートの行からキャッシュを検索
 *
 * 優先順位：
 *
 * 1. E列の埋め込みYouTube ID
 * 2. 作品名＋アーティスト＋順位
 *
 * E列を使うことで、
 * 同一作品・同一アーティストの複数曲にも対応。
 *******************************************************/

function findCacheForSheetRow_(
  sheetRow,
  cacheRecords
) {

  const workTitle =
    String(sheetRow[16] || '');

  const artist =
    String(sheetRow[1] || '');

  const rank =
    Number(sheetRow[18] || 0);

  const sheetVideoIds =
    parseVideoIds_(sheetRow[4]);

  /*
   * まず E列のYouTube IDで探す
   */
  if (sheetVideoIds.length) {

    let best = null;
    let bestScore = 0;

    cacheRecords.forEach(function(cache) {

      if (
        workTitle &&
        cache.workTitle !== workTitle
      ) {
        return;
      }

      if (
        artist &&
        cache.artist &&
        cache.artist !== artist
      ) {
        return;
      }

      let score = 0;

      cache.embeddedIds.forEach(function(id) {

        if (
          sheetVideoIds.indexOf(id) !== -1
        ) {
          score++;
        }

      });

      if (score > bestScore) {

        bestScore = score;
        best = cache;

      }

    });

    if (best) {
      return best;
    }
  }


  /*
   * E列で見つからない場合、
   * 作品＋アーティスト＋順位で検索
   */
  for (let i = 0; i < cacheRecords.length; i++) {

    const cache =
      cacheRecords[i];

    if (
      cache.workTitle === workTitle &&
      cache.artist === artist &&
      Number(cache.rank) === rank
    ) {

      return cache;

    }
  }

  return null;
}

/*******************************************************
 * ランキング項目を作品ページ情報で補完
 *
 * 同じ作品ページは1回だけ取得する。
 *******************************************************/

function enrichRankingItems_(items, forceRefresh) {

  if (!items || !items.length) {
    return [];
  }

  const cacheRecords =
    getCacheRecords_();

  const cacheByUrl = {};

  cacheRecords.forEach(function(cache) {

    const url =
      normalizeDetailUrl_(cache.detailUrl);

    if (url) {
      cacheByUrl[url] = cache;
    }

  });


  /*
   * 同じ作品ページの二重取得を防ぐ
   */
  const pageCache = {};

  const result = [];


  items.forEach(function(item) {

    const detailKey =
      normalizeDetailUrl_(item.detailUrl);

    const oldCache =
      cacheByUrl[detailKey];


    /*
     * 既存キャッシュに十分な情報があれば
     * 作品ページを再取得しない
     */
    if (
      !forceRefresh &&
      oldCache &&
      oldCache.songTitle &&
      oldCache.embeddedIds &&
      oldCache.embeddedIds.length
    ) {

      if (!item.songTitle) {
        item.songTitle =
          oldCache.songTitle;
      }

      if (!item.artist) {
        item.artist =
          oldCache.artist;
      }

      if (!item.type) {
        item.type =
          oldCache.type;
      }

      if (!item.releaseDate) {
        item.releaseDate =
          oldCache.releaseDate;
      }

      item.embeddedIds =
        unique_(
          (item.embeddedIds || [])
            .concat(oldCache.embeddedIds || [])
        );

      result.push(item);

      return;
    }


    /*
     * detailUrlがなければ何もしない
     */
    if (!item.detailUrl) {

      result.push(item);

      return;
    }


    /*
     * URLから #songXXXX を除いた
     * 作品ページURL
     */
    const pageUrl =
      normalizeDetailUrl_(
        item.detailUrl
      );


    let html;


    /*
     * 同じ作品ページなら再利用
     */
    if (
      Object.prototype.hasOwnProperty.call(
        pageCache,
        pageUrl
      )
    ) {

      html =
        pageCache[pageUrl];

    } else {

      html =
        fetchHtml_(pageUrl);

      pageCache[pageUrl] =
        html;
    }


    /*
     * 作品ページから対象曲を抽出
     */
    item =
      enrichFromAnimePage_(
        item,
        html
      );


    result.push(item);

  });


  return result;
}

/**
 * YouTube動画1件の情報を
 * メインシートG:Nへ書き込む。
 *
 * G 動画タイトル
 * H 説明欄
 * I 再生数
 * J いいね数
 * K 投稿日時
 * L 動画の長さ
 * M チャンネルID
 * N チャンネル名
 */
/**
 * YouTube動画情報をG:Nへ書き込む
 */
function writeYouTubeVideoInfo_(
  sheet,
  rowNumber,
  video
) {
  const snippet = video.snippet || {};
  const statistics = video.statistics || {};
  const contentDetails = video.contentDetails || {};

  const title = String(
    snippet.title || ''
  );

  const description = String(
    snippet.description || ''
  );

  const viewCount =
    statistics.viewCount !== undefined
      ? Number(statistics.viewCount)
      : '';

  const likeCount =
    statistics.likeCount !== undefined
      ? Number(statistics.likeCount)
      : '';

  const publishedAt =
    snippet.publishedAt
      ? new Date(snippet.publishedAt)
      : '';

  const duration =
    formatYouTubeDuration_(
      contentDetails.duration
    );

  const channelId =
    String(snippet.channelId || '');

  const channelTitle =
    String(snippet.channelTitle || '');

  sheet
    .getRange(rowNumber, 7, 1, 8)
    .setValues([[
      title,
      description,
      viewCount,
      likeCount,
      publishedAt,
      duration,
      channelId,
      channelTitle
    ]]);

  sheet
    .getRange(rowNumber, 20, 1, 4)
    .setValues([[
      '動画情報取得成功',
      '',
      'anison.online埋め込み動画',
      ''
    ]]);

  Logger.log(
    '[動画情報取得成功] 行=' +
    rowNumber +
    ' / ' +
    title +
    ' / ' +
    channelTitle
  );
}

/**
 * YouTubeのISO 8601 durationを
 * H:MM:SS / M:SS に変換。
 *
 * 例:
 * PT4M32S    -> 4:32
 * PT1H23M45S -> 1:23:45
 */
function formatYouTubeDuration_(duration) {
  if (!duration) {
    return '';
  }

  const text = String(duration);

  const match = text.match(
    /^PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?$/
  );

  if (!match) {
    return text;
  }

  const hours = Number(match[1] || 0);
  const minutes = Number(match[2] || 0);
  const seconds = Number(match[3] || 0);

  if (hours > 0) {
    return (
      hours +
      ':' +
      String(minutes).padStart(2, '0') +
      ':' +
      String(seconds).padStart(2, '0')
    );
  }

  return (
    minutes +
    ':' +
    String(seconds).padStart(2, '0')
  );
}

/**
 * E列の既存動画IDと新しい動画IDを統合。
 */
function mergeVideoIds_(existingValue, newIds) {
  const oldIds = parseVideoIds_(existingValue);

  const addIds = Array.isArray(newIds)
    ? newIds
    : [newIds];

  return unique_(
    oldIds.concat(
      addIds
        .map(function(id) {
          return String(id || '').trim();
        })
        .filter(function(id) {
          return id !== '';
        })
    )
  ).join(',');
}

/**
 * YouTube Data API videos.list
 *
 * 指定された動画IDをまとめて取得する。
 *
 * 取得情報:
 * - snippet
 * - contentDetails
 * - statistics
 */
function getYouTubeVideosByIds_(videoIds) {
  if (!videoIds || !videoIds.length) {
    return [];
  }

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('YOUTUBE_API_KEY');

  if (!apiKey) {
    throw new Error(
      'Script Properties に YOUTUBE_API_KEY がありません。'
    );
  }

  const ids = unique_(
    videoIds
      .map(function(id) {
        return String(id || '').trim();
      })
      .filter(function(id) {
        return id !== '';
      })
  );

  if (!ids.length) {
    return [];
  }

  const url =
    'https://www.googleapis.com/youtube/v3/videos' +
    '?part=snippet%2CcontentDetails%2Cstatistics' +
    '&id=' +
    encodeURIComponent(ids.join(',')) +
    '&key=' +
    encodeURIComponent(apiKey);

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const httpCode = response.getResponseCode();
  const text = response.getContentText();

  Logger.log(
    '[YouTube HTTP] ' +
    httpCode +
    ' / ' +
    ids.length +
    ' IDs'
  );

  if (httpCode !== 200) {
    let message = text;

    try {
      const errorJson = JSON.parse(text);

      if (
        errorJson &&
        errorJson.error &&
        errorJson.error.message
      ) {
        message = errorJson.error.message;
      }
    } catch (e) {}

    throw new Error(
      'YouTube API HTTP ' +
      httpCode +
      ': ' +
      message
    );
  }

  const json = JSON.parse(text);

  return json.items || [];
}

/*******************************************************
 * NicoNicoクイズ用データ取得
 *
 * VOCALOID_niconico シートから
 * 有効な行だけを取得してJSON用オブジェクトにする。
 *
 * YouTube側の処理・YouTube APIは一切使用しない。
 *******************************************************/
function getNiconicoQuizData_() {

  const sheetName = NICONICO_SHEET_NAME;

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      'NicoNico用シートが見つかりません: ' +
      sheetName
    );
  }

  const lastRow =
    sheet.getLastRow();

  /*
   * ヘッダーしかない場合
   */
  if (lastRow < 2) {
    return [];
  }

  /*
   * A:Pをまとめて取得
   */
  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        16
      )
      .getValues();

  const result = [];

  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const row =
      values[i];

    /*
     * シート上の実際の行番号
     */
    const rowIndex =
      i + 2;

    /*
     * A列：有効
     */
    const enabled =
      row[0] === true ||
      String(row[0]).trim().toLowerCase() === 'true' ||
      String(row[0]).trim() === '1' ||
      String(row[0]).trim() === 'yes' ||
      String(row[0]).trim() === '有効';

    if (!enabled) {
      continue;
    }

    /*
     * D列：ニコニコ動画ID
     */
    const niconicoId =
      String(
        row[3] || ''
      ).trim();

    /*
     * IDがない行はクイズ対象外
     */
    if (!niconicoId) {
      Logger.log(
        '[NicoNico] IDなし / 行=' +
        rowIndex
      );

      continue;
    }

    /*
     * ニコニコ動画URL
     *
     * E列にURLがあればそれを優先。
     * 空の場合はD列から自動生成。
     */
    let niconicoUrl =
      String(
        row[4] || ''
      ).trim();

    if (!niconicoUrl) {

      niconicoUrl =
        'https://www.nicovideo.jp/watch/' +
        encodeURIComponent(niconicoId);

    }

    /*
     * 日付をJSONで扱いやすい文字列にする。
     */
    const releaseDate =
      normalizeNiconicoDate_(
        row[2]
      );

    const postDate =
      normalizeNiconicoDate_(
        row[10]
      );

    /*
     * O列：指定ラントロ秒数
     */
    let customStartSec = null;

    if (
      row[14] !== '' &&
      row[14] !== null &&
      row[14] !== undefined
    ) {

      const sec =
        Number(row[14]);

      if (
        Number.isFinite(sec) &&
        sec >= 0
      ) {
        customStartSec = sec;
      }
    }

    /*
     * P列：タグ
     */
    const tags =
      parseNiconicoTags_(
        row[15]
      );

    /*
     * I/J列
     */
    const views =
      normalizeNumber_(
        row[8]
      );

    const likes =
      normalizeNumber_(
        row[9]
      );

    /*
     * L列：動画の長さ
     *
     * 現在DBに入っている文字列を
     * そのまま返す。
     */
    const duration =
      row[11] === null ||
      row[11] === undefined
        ? ''
        : String(row[11]);

    result.push({

      /*
       * 共通情報
       */
      sheetName: sheetName,

      rowIndex: rowIndex,

      artist:
        String(
          row[1] || ''
        ),

      publishedAt:
        releaseDate,

      /*
       * NicoNico固有情報
       */
      niconicoId:
        niconicoId,

      niconicoUrl:
        niconicoUrl,

      /*
       * DB情報
       */
      comment:
        String(
          row[5] || ''
        ),

      videoTitle:
        String(
          row[6] || ''
        ),

      description:
        String(
          row[7] || ''
        ),

      views:
        views,

      likes:
        likes,

      postDate:
        postDate,

      duration:
        duration,

      uploaderId:
        String(
          row[12] || ''
        ),

      uploaderName:
        String(
          row[13] || ''
        ),

      customStartSec:
        customStartSec,

      tags:
        tags
    });
  }

  Logger.log(
    '[NicoNico] クイズデータ取得: ' +
    result.length +
    '件'
  );

  return result;
}


/*******************************************************
 * NicoNico用日付正規化
 *******************************************************/
function normalizeNiconicoDate_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return '';
  }

  /*
   * Sheetsの日付セル
   */
  if (
    Object.prototype.toString.call(value) ===
    '[object Date]'
  ) {

    if (
      isNaN(value.getTime())
    ) {
      return '';
    }

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd'T'HH:mm:ss"
    );
  }

  /*
   * 文字列
   */
  return String(value);
}


/*******************************************************
 * NicoNico用数値正規化
 *******************************************************/
function normalizeNumber_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return 0;
  }

  const number =
    Number(
      String(value)
        .replace(/,/g, '')
        .trim()
    );

  return Number.isFinite(number)
    ? number
    : 0;
}


/*******************************************************
 * NicoNico用タグ解析
 *******************************************************/
function parseNiconicoTags_(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return [];
  }

  /*
   * すでに配列の場合
   */
  if (Array.isArray(value)) {

    return value
      .map(function(tag) {
        return String(tag).trim();
      })
      .filter(function(tag) {
        return tag !== '';
      });
  }

  return String(value)
    .split(',')
    .map(function(tag) {
      return tag.trim();
    })
    .filter(function(tag) {
      return tag !== '';
    });
}

/** 共通列構成 A:P の空行を作成 */
function createCommonSongRow_() {
    return new Array(16).fill("");
}



