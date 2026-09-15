// シート上の表示形式の差（全角数字、余分な空白、`3.0` など）を吸収して照合する。
function normalizeLookupText(value) {
  return String(value === undefined || value === null ? '' : value)
    .normalize('NFKC')
    .replace(/[\s\u3000]/g, '')
    .toLowerCase();
}

function normalizeRaceId(value) {
  const id = normalizeLookupText(value);
  // スプレッドシートの数値セルが "3" と "3.0" のいずれでも同じレースとして扱う。
  if (/^\d+(?:\.0+)?$/.test(id)) return String(Number(id));
  return id;
}

function doGet(e) {
  const action = e.parameter ? e.parameter.action : '';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- A. ログイン・初回確認認証 (action=login) ---
  if (action === 'login') {
    const rawName = String((e.parameter && (e.parameter.name || e.parameter.id)) || '');
    const rawPin = String((e.parameter && e.parameter.pin) || '');
    const cleanInputName = rawName.replace(/[\s\u3000]/g, '').toLowerCase();
    const cleanInputPin = rawPin.trim();
    
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return createJsonResponse({ success: false, message: 'Usersシートが見つかりません' });
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createJsonResponse({ success: false, message: 'Usersシートにデータがありません' });
    
    const headers = data[0].map(function(h) { return String(h).trim(); });
    let nameIdx = headers.findIndex(function(h) { return h.includes('名') || h.includes('ユーザー') || h.includes('ID'); });
    let pinIdx = headers.findIndex(function(h) { return h.toUpperCase().includes('PIN') || h.includes('暗証番号') || h.includes('パスワード'); });
    let ptIdx = headers.findIndex(function(h) { return h.includes('ポイント') || h.includes('pt') || h.includes('PT'); });
    let statusIdx = headers.findIndex(function(h) { return h.includes('ステータス') || h.includes('設定'); });

    if (nameIdx === -1) nameIdx = 0;
    if (pinIdx === -1) pinIdx = 2;
    if (ptIdx === -1) ptIdx = 3;
    
    let foundUserRow = null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const col0 = String(row[0] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      const col1 = String(row[1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      const colName = String(row[nameIdx] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      
      if ((col0 && col0 === cleanInputName) || (col1 && col1 === cleanInputName) || (colName && colName === cleanInputName)) {
        foundUserRow = {
          rowIndex: i + 1,
          name: row[nameIdx] || row[1] || row[0],
          pin: String(row[pinIdx] !== undefined ? row[pinIdx] : '').trim(),
          points: row[ptIdx] !== undefined ? row[ptIdx] : 0,
          status: statusIdx !== -1 ? String(row[statusIdx]).trim() : ''
        };
        break;
      }
    }
    
    if (!foundUserRow) {
      return createJsonResponse({ success: false, message: '「' + rawName + '」というお名前が見つかりません' });
    }

    // ステータスが「設定済み」でない、またはPINが空の場合は初回パスワード設定へ
    const isConfigured = (foundUserRow.status === '設定済み' && foundUserRow.pin !== '');
    
    if (!isConfigured) {
      return createJsonResponse({ success: true, isFirstTime: true, name: String(foundUserRow.name) });
    }

    // 設定済みの人のみ、PINの照合を行う（平文比較）
    const dbPin = String(foundUserRow.pin).trim();
    
    if (dbPin === cleanInputPin) {
      return createJsonResponse({ success: true, isFirstTime: false, name: String(foundUserRow.name), points: Number(foundUserRow.points) });
    } else {
      return createJsonResponse({ success: false, message: '暗証番号（PIN）が間違っています' });
    }
  }

  // --- B. 自分のベット履歴取得 (action=myBets) ---
  if (action === 'myBets') {
    const rawName = String((e.parameter && e.parameter.name) || '');
    const cleanInputName = rawName.replace(/[\s\u3000]/g, '').toLowerCase();

    // ★ キャッシュから取得を試みる（キー: myBets_ユーザー名）
    try {
      const cache = CacheService.getScriptCache();
      const cacheKey = 'myBets_' + cleanInputName;
      const cached = cache.get(cacheKey);
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
    } catch(cacheReadErr) { /* キャッシュ読み込み失敗は無視 */ }

    const betsSheet = ss.getSheetByName('Bets');
    if (!betsSheet) return createJsonResponse([]);
    const data = betsSheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return createJsonResponse([]);
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userNameInSheet = String(row[1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      if (userNameInSheet === cleanInputName) {
        result.push({
          rowIndex: i + 1,
          timestamp: row[0],
          userName: row[1],
          raceId: row[2],
          pick1: row[3],
          pick2: row[4],
          pick3: row[5],
          betPt: row[6],
          status: row[7] || '',
          hitCount: row[8] || '',
          payoutPt: row[9] || ''
        });
      }
    }
    const resultJson = JSON.stringify(result);
    try {
      const cache = CacheService.getScriptCache();
      cache.put('myBets_' + cleanInputName, resultJson, 60); // 60秒キャッシュ
    } catch(cacheWriteErr) { /* キャッシュ保存失敗は無視 */ }
    return ContentService.createTextOutput(resultJson).setMimeType(ContentService.MimeType.JSON);
  }

  // --- C. ランキング取得 (action=ranking) ---
  if (action === 'ranking') {
    // ★ キャッシュから取得を試みる
    try {
      const cache = CacheService.getScriptCache();
      const cacheKey = 'ranking_all';
      const cached = cache.get(cacheKey);
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
    } catch(cacheReadErr) { /* キャッシュ読み込み失敗は無視してスプレッドシートから取得 */ }

    const sheet = ss.getSheetByName('Users');
    if (!sheet) return createJsonResponse([]);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return String(h); });
    const result = [];
    for (let i = 1; i < data.length; i++) {
      let obj = {};
      headers.forEach(function(h, idx) {
        // Date型など非JSON型を安全に文字列化
        const val = data[i][idx];
        obj[h] = (val instanceof Date) ? val.toISOString() : val;
      });
      result.push(obj);
    }
    const resultJson = JSON.stringify(result);
    try {
      const cache = CacheService.getScriptCache();
      cache.put('ranking_all', resultJson, 30); // 30秒キャッシュ
    } catch(cacheWriteErr) { /* キャッシュ保存失敗は無視 */ }
    return ContentService.createTextOutput(resultJson).setMimeType(ContentService.MimeType.JSON);
  }

  // --- D. レース一覧取得 ---
  // ★ キャッシュから取得を試みる
  try {
    const raceCache = CacheService.getScriptCache();
    const raceCached = raceCache.get('races_all');
    if (raceCached) {
      return ContentService.createTextOutput(raceCached).setMimeType(ContentService.MimeType.JSON);
    }
  } catch(cacheReadErr) { /* キャッシュ読み込み失敗は無視 */ }

  const racesSheet = ss.getSheetByName('Races');
  if (!racesSheet) return createJsonResponse([]);
  const racesData = racesSheet.getDataRange().getDisplayValues();
  const racesHeaders = racesData[0];
  const racesList = [];
  for (let i = 1; i < racesData.length; i++) {
    let obj = {};
    racesHeaders.forEach(function(h, idx) { obj[h] = racesData[i][idx]; });
    racesList.push(obj);
  }
  const racesJson = JSON.stringify(racesList);
  try {
    const raceCache = CacheService.getScriptCache();
    raceCache.put('races_all', racesJson, 120); // 120秒キャッシュ
  } catch(cacheWriteErr) { /* キャッシュ保存失敗は無視 */ }
  return ContentService.createTextOutput(racesJson).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return createJsonResponse({ success: false, message: '混雑しています。数秒後にもう一度ボタンを押してください。' });
  }

  try {
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action || 'bet';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- 初回パスワード設定 (action=setupPin) ---
    if (action === 'setupPin') {
      const rawName = String(contents.userName || '');
      const newPin = String(contents.newPin || '').trim();
      const cleanInputName = rawName.replace(/[\s\u3000]/g, '').toLowerCase();

      if (!newPin || newPin.length !== 4 || isNaN(newPin)) {
        return createJsonResponse({ success: false, message: '4桁の数字を入力してください' });
      }

      const usersSheet = ss.getSheetByName('Users');
      const data = usersSheet.getDataRange().getValues();
      const headers = data[0].map(function(h) { return String(h).trim(); });

      let pinIdx = headers.findIndex(function(h) { return h.toUpperCase().includes('PIN') || h.includes('暗証番号'); });
      let statusIdx = headers.findIndex(function(h) { return h.includes('ステータス') || h.includes('設定'); });
      let ptIdx = headers.findIndex(function(h) { return h.includes('ポイント') || h.includes('pt'); });

      if (pinIdx === -1) pinIdx = 2;
      if (ptIdx === -1) ptIdx = 3;

      if (statusIdx === -1) {
        statusIdx = headers.length;
        usersSheet.getRange(1, statusIdx + 1).setValue('ステータス');
      }

      let userRowIndex = -1;
      let currentPt = 1000;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const col0 = String(row[0] || '').replace(/[\s\u3000]/g, '').toLowerCase();
        const col1 = String(row[1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
        if (col0 === cleanInputName || col1 === cleanInputName) {
          userRowIndex = i + 1;
          currentPt = Number(row[ptIdx] !== undefined ? row[ptIdx] : 1000);
          break;
        }
      }

      if (userRowIndex === -1) {
        return createJsonResponse({ success: false, message: 'ユーザーが見つかりません' });
      }

      // PINとステータス「設定済み」を保存
      usersSheet.getRange(userRowIndex, pinIdx + 1).setValue(newPin);
      usersSheet.getRange(userRowIndex, statusIdx + 1).setValue('設定済み');

      return createJsonResponse({ success: true, newPt: currentPt });
    }

    // --- キャンセル処理 (action=cancelBet) ---
    if (action === 'cancelBet') {
      const rawName = String(contents.userName || '');
      const rowIndex = Number(contents.rowIndex);
      const cleanInputName = rawName.replace(/[\s\u3000]/g, '').toLowerCase();
      const betsSheet = ss.getSheetByName('Bets');
      const usersSheet = ss.getSheetByName('Users');

      if (!betsSheet || !usersSheet || !rowIndex) {
        return createJsonResponse({ success: false, message: 'キャンセル対象が見つかりません' });
      }

      const rowData = betsSheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
      // Betsシートの列: 0:日時, 1:カタカナ名, 2:レースID, 3:1着, 4:2着, 5:3着, 6:賭けたpt, 7:状態
      const sheetUser = String(rowData[1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      const status = String(rowData[7] || '').trim();
      const refundPt = Number(rowData[6] || 0);

      if (sheetUser !== cleanInputName) {
        return createJsonResponse({ success: false, message: '他のユーザーのベットは削除できません' });
      }

      if (status !== '受付済' && status !== '受付済(フォーム)') {
        return createJsonResponse({ success: false, message: '確定済み、または既にキャンセルされたベットは取り消せません' });
      }

      // ユーザーのPIN確認とポイント返還処理
      const usersData = usersSheet.getDataRange().getValues();
      const headers = usersData[0].map(function(h) { return String(h).trim(); });
      let ptIdx = headers.findIndex(function(h) { return h.includes('ポイント') || h.includes('pt') || h.includes('PT'); });
      let pinIdx = headers.findIndex(function(h) { return h.toUpperCase().includes('PIN') || h.includes('暗証番号') || h.includes('パスワード'); });
      if (ptIdx === -1) ptIdx = 3;
      if (pinIdx === -1) pinIdx = 2;

      let validUser = false;
      let newPt = 0;
      const cleanInputPin = String(contents.pin || '').trim();

      for (let i = 1; i < usersData.length; i++) {
        const uRow = usersData[i];
        const uCol0 = String(uRow[0] || '').replace(/[\s\u3000]/g, '').toLowerCase();
        const uCol1 = String(uRow[1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
        
        if (uCol0 === cleanInputName || uCol1 === cleanInputName) {
          const dbPin = String(uRow[pinIdx] !== undefined ? uRow[pinIdx] : (uRow[2] || '')).trim();
          
          if (dbPin === cleanInputPin) {
            validUser = true;
            const currentPt = Number(uRow[ptIdx] || 0);
            newPt = currentPt + refundPt;
            usersSheet.getRange(i + 1, ptIdx + 1).setValue(newPt);
          }
          break;
        }
      }

      if (!validUser) {
        return createJsonResponse({ success: false, message: '認証エラー：本人確認ができません' });
      }

      // 論理削除に変更
      betsSheet.getRange(rowIndex, 8).setValue('キャンセル済');

      // ★ キャンセル後はmyBetsとrankingのキャッシュを削除
      try {
        const cache = CacheService.getScriptCache();
        cache.remove('myBets_' + cleanInputName);
        cache.remove('ranking_all');
      } catch(cacheErr) { /* キャッシュ削除失敗は無視 */ }

      return createJsonResponse({ success: true, newPt: newPt });
    }

    // --- 通常のベット処理（3人予想方式） ---
    const rawName = String(contents.userName || contents.userId || '');
    const rawPin = String(contents.pin || '');
    const raceId = String(contents.raceId || '').trim();
    const pick1 = String(contents.pick1 || '').trim();
    const pick2 = String(contents.pick2 || '').trim();
    const pick3 = String(contents.pick3 || '').trim();
    const betPt = Number(contents.betPt || 0);

    const cleanInputName = rawName.replace(/[\s\u3000]/g, '').toLowerCase();
    const cleanInputPin = rawPin.trim();

    if (!cleanInputName || !cleanInputPin || !raceId || !pick1 || !pick2 || !pick3 || betPt <= 0) {
      return createJsonResponse({ success: false, message: '入力内容が不正です' });
    }

    // 同じ選手の重複チェック
    if (pick1 === pick2 || pick2 === pick3 || pick1 === pick3) {
      return createJsonResponse({ success: false, message: '同じ選手を複数回選択することはできません' });
    }

    const ss2 = SpreadsheetApp.getActiveSpreadsheet();

    // ★フールプルーフ6: サーバー側での締切時刻チェック
    const racesSheet = ss2.getSheetByName('Races');
    if (racesSheet) {
      const racesData = racesSheet.getDataRange().getDisplayValues();
      const racesHeaders = racesData[0];
      const rIdIdx = racesHeaders.findIndex(function(h) { return h.includes('レースID') || h === 'raceId'; });
      const timeIdx = racesHeaders.findIndex(function(h) { return h.includes('発走時刻') || h === 'time' || h === 'startTime'; });
      
      let targetTimeStr = null;
      for (let i = 1; i < racesData.length; i++) {
        const rowId = String(racesData[i][rIdIdx >= 0 ? rIdIdx : 0]).trim();
        if (rowId === raceId) {
          targetTimeStr = String(racesData[i][timeIdx >= 0 ? timeIdx : 4]).trim();
          break;
        }
      }
      
      if (targetTimeStr) {
        // "12:30" などの文字列から本日のDateを作成
        const timeMatch = targetTimeStr.match(/(\d{1,2}):(\d{2})/);
        if (timeMatch) {
          const hours = parseInt(timeMatch[1], 10);
          const mins = parseInt(timeMatch[2], 10);
          const now = new Date();
          const targetTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hours, mins, 0);
          
          if (now.getTime() >= targetTime.getTime()) {
             return createJsonResponse({ success: false, message: 'このレースはすでに発走時刻を過ぎているためベットできません' });
          }
        }
      }
    }
    const usersSheet = ss2.getSheetByName('Users');
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return String(h).trim(); });

    let nameIdx = headers.findIndex(function(h) { return h.includes('名') || h.includes('ユーザー') || h.includes('ID'); });
    let pinIdx = headers.findIndex(function(h) { return h.toUpperCase().includes('PIN') || h.includes('暗証番号') || h.includes('パスワード'); });
    let ptIdx = headers.findIndex(function(h) { return h.includes('ポイント') || h.includes('pt') || h.includes('PT'); });

    if (nameIdx === -1) nameIdx = 0;
    if (pinIdx === -1) pinIdx = 2;
    if (ptIdx === -1) ptIdx = 3;

    let userRowIndex = -1;
    let currentPt = 0;
    let canonicalName = rawName;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const col0 = String(row[0] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      const col1 = String(row[1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      const colName = String(row[nameIdx] || '').replace(/[\s\u3000]/g, '').toLowerCase();

      const isNameMatch = (col0 && col0 === cleanInputName) || (col1 && col1 === cleanInputName) || (colName && colName === cleanInputName);
      const dbPin = String(row[pinIdx] !== undefined ? row[pinIdx] : (row[2] || '')).trim();
      const isPinMatch = (dbPin === cleanInputPin);

      if (isNameMatch && isPinMatch) {
        userRowIndex = i + 1;
        currentPt = Number(row[ptIdx] !== undefined ? row[ptIdx] : (row[3] || 0));
        canonicalName = row[nameIdx] || row[1] || row[0] || rawName;
        break;
      }
    }

    if (userRowIndex === -1) {
      return createJsonResponse({ success: false, message: 'ユーザー認証に失敗しました' });
    }

    if (currentPt < betPt) {
      return createJsonResponse({ success: false, message: '所持ポイントが不足しています' });
    }

    let betsSheet = ss2.getSheetByName('Bets');
    if (!betsSheet) {
      betsSheet = ss2.insertSheet('Bets');
      betsSheet.appendRow(['日時', 'カタカナ名', 'レースID', '1着予想', '2着予想', '3着予想', '賭けたpt', '状態', '的中数', '払戻pt']);
    }

    // ★ サーバー側での重複ベットチェック
    const betsData = betsSheet.getDataRange().getValues();
    const normalizedCanonicalName = canonicalName.replace(/[\s\u3000]/g, '').toLowerCase();
    for (let i = 1; i < betsData.length; i++) {
      const bRow = betsData[i];
      const bUser = String(bRow[1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      const bRaceId = String(bRow[2] || '').trim();
      const bStatus = String(bRow[7] || '').trim();
      
      if (bUser === normalizedCanonicalName && bRaceId === raceId && bStatus !== 'キャンセル済') {
        return createJsonResponse({ success: false, message: 'このレースにはすでにベットしています（重複ベットエラー）' });
      }
    }

    const newPt = currentPt - betPt;
    usersSheet.getRange(userRowIndex, ptIdx + 1).setValue(newPt);

    betsSheet.appendRow([new Date(), canonicalName, raceId, pick1, pick2, pick3, betPt, '受付済', '', 0]);

    // ★ ベット後はこのユーザーのmyBetsキャッシュを削除（次回は最新データを取得）
    try {
      const cache = CacheService.getScriptCache();
      const normalizedForCache = canonicalName.replace(/[\s\u3000]/g, '').toLowerCase();
      cache.remove('myBets_' + normalizedForCache);
      cache.remove('ranking_all'); // ポイント変動があるのでランキングも削除
    } catch(cacheErr) { /* キャッシュ削除失敗は無視 */ }

    return createJsonResponse({ success: true, newPt: newPt });

  } catch (err) {
    return createJsonResponse({ success: false, message: 'エラーが発生しました: ' + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========================================
// カスタムメニュー（スプレッドシート上部）
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏇 ダービー管理')
    .addItem('結果確定＆払い戻し実行', 'processRaceResults')
    .addItem('不的中の再判定・払い戻し補正', 'repairMissedPayouts')
    .addToUi();
}

// ========================================
// 結果確定・自動払い戻し処理（順位一致方式）
// ========================================
function processRaceResults(options) {
  const repairMisses = options && options.repairMisses === true;
  // ★フールプルーフ5: 結果確定前の確認ダイアログ
  const ui = SpreadsheetApp.getUi();
  const dialogTitle = repairMisses ? '不的中の再判定・払い戻し補正' : '結果確定と払い戻し';
  const dialogMessage = repairMisses
    ? '確定済みレースの「不的中」ベットだけを再判定します。新たに的中となったベットにのみ払い戻しを加算します。\n本当によろしいですか？'
    : '結果を確定し、払い戻しを実行します。\nこの操作は取り消せません。本当によろしいですか？';
  const response = ui.alert(dialogTitle, dialogMessage, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultsSheet = ss.getSheetByName('Results');
  const betsSheet = ss.getSheetByName('Bets');
  const usersSheet = ss.getSheetByName('Users');
  const racesSheet = ss.getSheetByName('Races');

  if (!resultsSheet) {
    SpreadsheetApp.getUi().alert('Resultsシートが見つかりません。\n「レースID, 1着レーン, 2着レーン, 3着レーン, ステータス」のカラムで作成してください。');
    return;
  }
  if (!betsSheet) {
    SpreadsheetApp.getUi().alert('Betsシートが見つかりません。');
    return;
  }

  const resultsData = resultsSheet.getDataRange().getValues();
  // ヘッダー: レースID(0), 1着レーン(1), 2着レーン(2), 3着レーン(3), ステータス(4)

  if (!racesSheet) {
    SpreadsheetApp.getUi().alert('Racesシートが見つかりません。オッズ・選手情報が参照できないため処理を中止します。');
    return;
  }

  const racesData = racesSheet.getDataRange().getValues();
  const racesHeaders = racesData[0].map(function(h) { return String(h).trim(); });
  const raceIdIdx = racesHeaders.findIndex(function(h) { return h.includes('レースID') || h === 'raceId'; });
  const laneIdx   = racesHeaders.findIndex(function(h) { return h.includes('レーン') || h === 'lane'; });
  const nameIdx   = racesHeaders.findIndex(function(h) { return h.includes('選手名') || h === 'name'; });
  const oddsIdx   = racesHeaders.findIndex(function(h) { return h.includes('オッズ') || h === 'odds'; });

  // レーン番号または選手名 → 選手名。Resultsと旧形式のBetsの両方に使う。
  const resultEntryToAthleteMap = {};
  // レースID+選手名 → オッズ のマップ
  const oddsMap = {};

  for (let i = 1; i < racesData.length; i++) {
    const rId        = normalizeRaceId(racesData[i][raceIdIdx >= 0 ? raceIdIdx : 0]);
    const lane       = normalizeLookupText(racesData[i][laneIdx   >= 0 ? laneIdx   : 1]);
    const athlete    = String(racesData[i][nameIdx   >= 0 ? nameIdx   : 2] || '').trim();
    const odds       = parseFloat(racesData[i][oddsIdx >= 0 ? oddsIdx : 6] || 1);
    if (rId && lane && athlete) {
      resultEntryToAthleteMap[rId + '_' + lane] = athlete;
      resultEntryToAthleteMap[rId + '_' + normalizeLookupText(athlete)] = athlete;
      oddsMap[rId + '_' + normalizeLookupText(athlete)] = odds;
    }
  }

  // Betsシートのデータ
  const betsData = betsSheet.getDataRange().getValues();
  // ヘッダー: 日時(0), カタカナ名(1), レースID(2), 1着予想(3), 2着予想(4), 3着予想(5), 賭けたpt(6), 状態(7), 的中数(8), 払戻pt(9)

  // Usersシートのポイント列を特定
  const usersData = usersSheet.getDataRange().getValues();
  const usersHeaders = usersData[0].map(function(h) { return String(h).trim(); });
  let ptIdx = usersHeaders.findIndex(function(h) { return h.includes('ポイント') || h.includes('pt') || h.includes('PT'); });
  if (ptIdx === -1) ptIdx = 3;

  let processedCount = 0;
  let totalPayout = 0;

  // 未確定レースを処理
  for (let r = 1; r < resultsData.length; r++) {
    const resultRow    = resultsData[r];
    const targetRaceId = normalizeRaceId(resultRow[0]);
    const lane1st      = normalizeLookupText(resultRow[1]);
    const lane2nd      = normalizeLookupText(resultRow[2]);
    const lane3rd      = normalizeLookupText(resultRow[3]);
    const status       = String(resultRow[4] || '').trim();

    if ((!repairMisses && status === '確定済') || (repairMisses && status !== '確定済') || !targetRaceId || !lane1st) continue;

    // ★ レーン番号を選手名に変換 & 存在チェック
    const actual1st = resultEntryToAthleteMap[targetRaceId + '_' + lane1st];
    const actual2nd = resultEntryToAthleteMap[targetRaceId + '_' + lane2nd];
    const actual3rd = resultEntryToAthleteMap[targetRaceId + '_' + lane3rd];

    // ★フールプルーフ3: レーン重複チェック
    const lanes = [lane1st];
    if (lane2nd) lanes.push(lane2nd);
    if (lane3rd) lanes.push(lane3rd);
    const uniqueLanes = new Set(lanes);
    if (uniqueLanes.size !== lanes.length) {
      SpreadsheetApp.getUi().alert(
        '⚠️ エラー: レースID「' + targetRaceId + '」の結果に\n' +
        '同じレーン番号が重複して入力されています。\n\n' +
        'Resultsシートの着順レーンを確認し、修正してから再実行してください。\n処理を中止します。'
      );
      return;
    }

    if (!actual1st) {
      SpreadsheetApp.getUi().alert(
        '⚠️ エラー: レースID「' + targetRaceId + '」に\n' +
        '1着レーン「' + lane1st + '」は存在しません。\n\n' +
        'Racesシートのレーン番号を確認して再実行してください。\n処理を中止します。'
      );
      return;
    }
    if (lane2nd && !actual2nd) {
      SpreadsheetApp.getUi().alert(
        '⚠️ エラー: レースID「' + targetRaceId + '」に\n' +
        '2着レーン「' + lane2nd + '」は存在しません。\n\n' +
        'Racesシートのレーン番号を確認して再実行してください。\n処理を中止します。'
      );
      return;
    }
    if (lane3rd && !actual3rd) {
      SpreadsheetApp.getUi().alert(
        '⚠️ エラー: レースID「' + targetRaceId + '」に\n' +
        '3着レーン「' + lane3rd + '」は存在しません。\n\n' +
        'Racesシートのレーン番号を確認して再実行してください。\n処理を中止します。'
      );
      return;
    }

    // 対象レースのBetsを走査
    for (let b = 1; b < betsData.length; b++) {
      const betRow   = betsData[b];
      const betRaceId = normalizeRaceId(betRow[2]);
      const betStatus = String(betRow[7] || '').trim();

      if (betRaceId !== targetRaceId) continue;
      if (repairMisses) {
        // 補正では既存の不的中だけを扱うため、二重払い戻しは発生しない。
        if (betStatus !== '不的中') continue;
      } else if (betStatus === '確定' || betStatus === '的中' || betStatus === '不的中' || betStatus === 'キャンセル済') {
        continue;
      }

      const userPick1 = String(betRow[3] || '').trim();
      const userPick2 = String(betRow[4] || '').trim();
      const userPick3 = String(betRow[5] || '').trim();
      // 旧ベットには選手名ではなくレーン番号が保存されていることがある。
      const pick1Athlete = resultEntryToAthleteMap[targetRaceId + '_' + normalizeLookupText(userPick1)] || userPick1;
      const pick2Athlete = resultEntryToAthleteMap[targetRaceId + '_' + normalizeLookupText(userPick2)] || userPick2;
      const pick3Athlete = resultEntryToAthleteMap[targetRaceId + '_' + normalizeLookupText(userPick3)] || userPick3;
      const betPt     = Number(betRow[6] || 0);
      const betUserName = String(betRow[1] || '');

      // 順位一致判定（ユーザー予想＝選手名、actual＝選手名で比較）
      let hitCount = 0;
      let multiplier = 1;

      if (userPick1 && normalizeLookupText(pick1Athlete) === normalizeLookupText(actual1st)) {
        hitCount++;
        multiplier *= (oddsMap[targetRaceId + '_' + normalizeLookupText(pick1Athlete)] || 1);
      }
      if (userPick2 && actual2nd && normalizeLookupText(pick2Athlete) === normalizeLookupText(actual2nd)) {
        hitCount++;
        multiplier *= (oddsMap[targetRaceId + '_' + normalizeLookupText(pick2Athlete)] || 1);
      }
      if (userPick3 && actual3rd && normalizeLookupText(pick3Athlete) === normalizeLookupText(actual3rd)) {
        hitCount++;
        multiplier *= (oddsMap[targetRaceId + '_' + normalizeLookupText(pick3Athlete)] || 1);
      }

      let payoutPt = 0;
      let resultStatus = '不的中';
      if (hitCount > 0) {
        payoutPt = Math.floor(betPt * multiplier);
        resultStatus = '的中';
        totalPayout += payoutPt;

        // Usersシートのポイントを加算
        const cleanBetUser = betUserName.replace(/[\s\u3000]/g, '').toLowerCase();
        for (let u = 1; u < usersData.length; u++) {
          const uCol0 = String(usersData[u][0] || '').replace(/[\s\u3000]/g, '').toLowerCase();
          const uCol1 = String(usersData[u][1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
          if (uCol0 === cleanBetUser || uCol1 === cleanBetUser) {
            const currentPt = Number(usersData[u][ptIdx] || 0);
            const newPt = currentPt + payoutPt;
            usersSheet.getRange(u + 1, ptIdx + 1).setValue(newPt);
            usersData[u][ptIdx] = newPt; // メモリ上も更新
            break;
          }
        }
      }

      // Betsシートを更新
      betsSheet.getRange(b + 1, 8).setValue(resultStatus);   // 状態
      betsSheet.getRange(b + 1, 9).setValue(hitCount + '/3'); // 的中数
      betsSheet.getRange(b + 1, 10).setValue(payoutPt);       // 払戻pt
      processedCount++;
    }

    // Resultsシートのステータスを「確定済」に
    resultsSheet.getRange(r + 1, 5).setValue('確定済');
  }

  // ★ 結果確定後は全ユーザーのキャッシュを一括削除（全員に最新データを届けるため）
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('ranking_all');
    cache.remove('races_all');
    // myBetsは個別キーのため、ユーザー分を全削除
    const usersSheetForCache = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (usersSheetForCache) {
      const allUsers = usersSheetForCache.getDataRange().getValues();
      const keysToRemove = [];
      for (let u = 1; u < allUsers.length; u++) {
        const uName = String(allUsers[u][0] || allUsers[u][1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
        if (uName) keysToRemove.push('myBets_' + uName);
      }
      if (keysToRemove.length > 0) cache.removeAll(keysToRemove);
    }
  } catch(cacheErr) { /* キャッシュ削除失敗は無視 */ }

  SpreadsheetApp.getUi().alert(
    (repairMisses ? '🏇 再判定・払い戻し補正完了' : '🏇 結果確定完了') + '\n\n' +
    '処理したベット数: ' + processedCount + '\n' +
    '合計払い戻し: ' + totalPayout + ' pt'
  );
}

// 誤判定済みの不的中を、安全に一度だけ再判定する管理用メニュー。
function repairMissedPayouts() {
  processRaceResults({ repairMisses: true });
}

// ========================================
// Googleフォーム連携（バックアップ）
// ========================================
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return; // ロック取得失敗
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const responses = e.values; // [タイムスタンプ, お名前, 種目, 1着予想, 2着予想, 3着予想, 賭けるポイント]

    const formName = String(responses[1] || '').trim();
    const formSubject = String(responses[2] || '').trim();
    const formPick1 = String(responses[3] || '').trim();
    const formPick2 = String(responses[4] || '').trim();
    const formPick3 = String(responses[5] || '').trim();
    const formBetPt = Number(responses[6] || 0);

    if (!formName || !formSubject || !formPick1 || !formPick2 || !formPick3 || formBetPt <= 0) return;

    // 種目名からレースIDを検索
    const racesSheet = ss.getSheetByName('Races');
    let raceId = '';
    if (racesSheet) {
      const racesData = racesSheet.getDataRange().getValues();
      const racesHeaders = racesData[0].map(function(h) { return String(h).trim(); });
      const raceIdIdx = racesHeaders.findIndex(function(h) { return h.includes('レースID'); });
      const subjectIdx = racesHeaders.findIndex(function(h) { return h.includes('種目'); });

      for (let i = 1; i < racesData.length; i++) {
        const subjectInSheet = String(racesData[i][subjectIdx >= 0 ? subjectIdx : 4] || '').trim();
        if (subjectInSheet === formSubject) {
          raceId = String(racesData[i][raceIdIdx >= 0 ? raceIdIdx : 0] || '').trim();
          break;
        }
      }
    }

    if (!raceId) return; // 種目名が見つからない場合は無視

    // Usersシートからポイントを差し引き
    const usersSheet = ss.getSheetByName('Users');
    const usersData = usersSheet.getDataRange().getValues();
    const usersHeaders = usersData[0].map(function(h) { return String(h).trim(); });
    let ptIdx = usersHeaders.findIndex(function(h) { return h.includes('ポイント') || h.includes('pt'); });
    if (ptIdx === -1) ptIdx = 3;

    const cleanFormName = formName.replace(/[\s\u3000]/g, '').toLowerCase();
    let userFound = false;

    for (let i = 1; i < usersData.length; i++) {
      const col0 = String(usersData[i][0] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      const col1 = String(usersData[i][1] || '').replace(/[\s\u3000]/g, '').toLowerCase();
      if (col0 === cleanFormName || col1 === cleanFormName) {
        const currentPt = Number(usersData[i][ptIdx] || 0);
        if (currentPt >= formBetPt) {
          usersSheet.getRange(i + 1, ptIdx + 1).setValue(currentPt - formBetPt);
          userFound = true;
        }
        break;
      }
    }

    if (!userFound) return; // ユーザーが見つからないかポイント不足

    // Betsシートに記録
    let betsSheet = ss.getSheetByName('Bets');
    if (!betsSheet) {
      betsSheet = ss.insertSheet('Bets');
      betsSheet.appendRow(['日時', 'カタカナ名', 'レースID', '1着予想', '2着予想', '3着予想', '賭けたpt', '状態', '的中数', '払戻pt']);
    }
    betsSheet.appendRow([new Date(), formName, raceId, formPick1, formPick2, formPick3, formBetPt, '受付済(フォーム)', '', 0]);

  } catch (err) {
    // エラー時はログに記録
    console.error('onFormSubmit error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// ========================================
// サーバー側でのPINハッシュ化ユーティリティ
// 過去の平文PIN（4桁数字）との後方互換性を保つために使用
// ========================================
function computeSHA256(str) {
  const signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return signature.map(function(byte) {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}
