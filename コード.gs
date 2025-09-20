// ----------------------------------------
// グローバル設定
// ----------------------------------------
const SPREADSHEET_ID = '1xEg4Cth_rosTqgj0nMrwZ-cOBduCoX2mi6B-SyUVoUo'; // メイン用
// const SPREADSHEET_ID = '1qTSkENn-CxotdKvoetRF7CthWSzsOHbDMzdQexslGV4'; // コード修正用
const SHEET_STUDENT_MASTER = '生徒マスタ';
const SHEET_ATTENDANCE_LOG = '入退室記録';
const SHEET_SUMMARY = '学習時間サマリー';
const SHEET_MONTHLY_SUMMARY = '月別学習時間集計';
//const SHEET_MONTHLY_SUMMARY = '月別学習時間集計(テスト)';
const SHEET_CURRENT_STATUS = '学習状況';
const SHEET_GOAL = '目標管理';

// ----------------------------------------
// ★★★ 最終修正版：日付変換ヘルパー関数 ★★★
// ----------------------------------------
/**
 * どんな形式の入力値からでも、有効なDateオブジェクトを返すことを試みる関数
 * @param {*} value - 日時データ (Dateオブジェクト、文字列など)
 * @returns {Date|null} - 有効なDateオブジェクト、または無効な場合はnull
 */
function getValidDate(value) {
  // 1. 既に有効なDateオブジェクトの場合、そのまま返す
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  // 2. 文字列の場合、手動でパース（解析）する
  if (typeof value === 'string' && value.trim() !== '') {
    // "2025-07-29 19:19:35" を [2025, 7, 29, 19, 19, 35] のような数値の配列に変換
    const parts = value.split(/[\s:-]/).map(part => parseInt(part, 10));
    if (parts.length >= 6 && !parts.some(isNaN)) {
      // new Date(年, 月-1, 日, 時, 分, 秒) でオブジェクトを生成 (月は0始まりのため-1する)
      const dt = new Date(parts[0], parts[1] - 1, parts[2], parts[3], parts[4], parts[5]);
      if (!isNaN(dt.getTime())) {
        return dt; // 正常に変換できたら返す
      }
    }
  }
  // 3. 上記のいずれにも当てはまらない場合はnullを返す
  return null;
}

// ----------------------------------------
// Webアプリケーションのエントリーポイント (doGet)
// ----------------------------------------
function doGet(e) {
  Logger.log('doGet called with parameters: ' + JSON.stringify(e.parameter));
  const page = e.parameter.page;
  if (page === 'admin') {
    return showAdminPage(e);
  } else if (page === 'main') {
    return showMainPage(e);
  } else if (page === 'goal') {
    return showGoalPage(e);
  } else {
    return showLoginPage(e);
  }
}

// ----------------------------------------
// ページ表示関数
// ----------------------------------------
function showLoginPage(e) {
  Logger.log('Rendering login page.');
  const template = HtmlService.createTemplateFromFile('login');
  template.webAppUrl = ScriptApp.getService().getUrl();
  return template.evaluate().setTitle('自習室管理システム - ログイン');
}

function showAdminPage(e) {
  Logger.log('Rendering admin page.');
  const template = HtmlService.createTemplateFromFile('admin');
  template.webAppUrl = ScriptApp.getService().getUrl();
  return template.evaluate().setTitle('管理者用ダッシュボード');
}

function showMainPage(e) {
  if (!e.parameter.userId || !e.parameter.studentName) {
    return showLoginPage(e);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    Logger.log('Rendering main page for user: ' + e.parameter.userId);
    const template = HtmlService.createTemplateFromFile('index');
    template.userId = e.parameter.userId;
    template.studentName = decodeURIComponent(e.parameter.studentName);
    template.webAppUrl = ScriptApp.getService().getUrl();

    const today = new Date();
    const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
    template.todayAsString = Utilities.formatString("%d年%02d月%02d日（%s）", today.getFullYear(), today.getMonth() + 1, today.getDate(), weekdays[today.getDay()]);

    let monthlyStudyTime = 0;
    const yearlyStudyData = [['月', '勉強時間(実績)', 'トレンドライン']];
    const monthNames = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];
    const yearlyActualTotals = Array(12).fill(0);
    try {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const summarySheet = spreadsheet.getSheetByName(SHEET_SUMMARY);
      const monthlyAggregationSheet = spreadsheet.getSheetByName(SHEET_MONTHLY_SUMMARY);
      if (summarySheet) {
        const summaryData = summarySheet.getDataRange().getValues();
        for (let i = 1; i < summaryData.length; i++) {
          if (summaryData[i][0] && summaryData[i][0].toString().trim() === e.parameter.userId.trim()) {
            monthlyStudyTime = parseFloat(summaryData[i][6]) || 0;
            break;
          }
        }
      }

      if (monthlyAggregationSheet) {
        const monthlyData = monthlyAggregationSheet.getDataRange().getValues();
        const currentYearStr = today.getFullYear().toString();
        for (let i = 1; i < monthlyData.length; i++) {
          const rowUserId = monthlyData[i][0] ? monthlyData[i][0].toString().trim() : '';
          let processedYearMonth = '';
          if (monthlyData[i][2]) {
            const dt = getValidDate(monthlyData[i][2]);
            if (dt) {
              processedYearMonth = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM');
            } else {
              processedYearMonth = monthlyData[i][2].toString().substring(0, 7).trim();
            }
          }
          const rowMonthlyMinutes = parseFloat(monthlyData[i][3]) || 0;
          if (rowUserId === e.parameter.userId.trim() && processedYearMonth.startsWith(currentYearStr + "-")) {
            const monthPart = parseInt(processedYearMonth.split('-')[1], 10);
            if (monthPart >= 1 && monthPart <= 12) {
              yearlyActualTotals[monthPart - 1] = rowMonthlyMinutes;
            }
          }
        }
      }

      yearlyActualTotals[today.getMonth()] = monthlyStudyTime;
      for (let i = 0; i < 12; i++) {
        yearlyStudyData.push([monthNames[i], yearlyActualTotals[i], yearlyActualTotals[i]]);
      }
      template.yearlyStudyDataJson = JSON.stringify(yearlyStudyData);
      template.monthlyStudyTime = monthlyStudyTime;
    } catch (err) {
      Logger.log('Error fetching study data for main page: ' + err.toString());
      template.monthlyStudyTime = 0;
      const emptyYearlyData = [['月', '勉強時間(実績)', 'トレンドライン']];
      for(let i=0; i<12; i++) { emptyYearlyData.push([monthNames[i], 0, 0]);}
      template.yearlyStudyDataJson = JSON.stringify(emptyYearlyData);
    }
    
    return template.evaluate().setTitle('自習室管理システム');
  } finally {
    lock.releaseLock();
  }
}

function showGoalPage(e) {
  if (!e.parameter.userId || !e.parameter.studentName) {
    return showLoginPage(e);
  }
  Logger.log('Rendering goal setting page for user: ' + e.parameter.userId);
  const template = HtmlService.createTemplateFromFile('goal');
  template.userId = e.parameter.userId;
  template.studentName = decodeURIComponent(e.parameter.studentName);
  template.webAppUrl = ScriptApp.getService().getUrl();
  template.lastMonthGoal = e.parameter.lastMonthGoal ? decodeURIComponent(e.parameter.lastMonthGoal) : '';
  template.currentYear = e.parameter.currentYear;
  template.currentMonth = e.parameter.currentMonth;
  template.lastMonthYear = e.parameter.lastMonthYear;
  template.lastMonth = e.parameter.lastMonth;

  return template.evaluate().setTitle('目標設定');
}

function authenticateUser(userId, password) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[認証開始] ユーザーID: ${userId}`);

    if (userId === 'admin' && password === 'admin_password') {
      return { success: true, isAdmin: true };
    }
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_STUDENT_MASTER);
    if (!sheet) throw new Error('生徒マスタが見つかりません。');
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === userId.trim() && data[i][2].toString() === password) {
        const studentName = data[i][1].toString();
        Logger.log(`[認証成功] 生徒名: ${studentName}`);
        
        const goalStatus = checkGoalSettingStatus(userId);
        // ▼▼▼ ログ追加 ▼▼▼
        Logger.log(`[目標チェック結果] goalStatus: ${JSON.stringify(goalStatus)}`);
        
        return { 
          success: true, 
          isAdmin: false, 
          studentName: studentName,
          goalStatus: goalStatus
        };
      }
    }
    Logger.log('[認証失敗] 生徒IDまたはパスワードが不正');
    return { success: false, message: '生徒IDまたはパスワードが正しくありません。' };
  } catch (error) {
    Logger.log(`[認証エラー] ${error.toString()}`);
    return { success: false, message: `認証エラー: ${error.message}` };
  } finally {
    lock.releaseLock();
  }
}

function checkGoalSettingStatus(userId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[目標チェック開始] ユーザーID: ${userId}`);

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) {
      Logger.log('[目標チェック] 目標管理シートが見つかりません。スキップします。');
      return { required: false };
    }
    
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[目標チェック] 現在の年月: ${currentYear}年${currentMonth}月`);
    
    const data = sheet.getDataRange().getValues();
    
    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log('[目標チェック] シートの各行と現在の年月を比較します...');
    
    const currentMonthGoalExists = data.slice(1).some(row => {
      const rowUserId = row[0].toString().trim();
      const rowYear = parseInt(row[2], 10);
      const rowMonth = parseInt(row[3], 10);
      const isMatch = (rowUserId === userId && rowYear === currentYear && rowMonth === currentMonth);
      
      // ▼▼▼ ログ追加 ▼▼▼
      // 比較対象のユーザーIDが一致する場合のみログを出力し、ログが煩雑になるのを防ぐ
      if (rowUserId === userId) {
        Logger.log(`  - シートの値: {id: ${rowUserId}, year: ${row[2]}, month: ${row[3]}} | 比較結果: ${isMatch}`);
      }
      
      return isMatch;
    });

    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[目標チェック] 最終判定: 今月の目標は存在しますか？ -> ${currentMonthGoalExists}`);

    if (!currentMonthGoalExists) {
      Logger.log('[目標チェック] 結果: 目標設定が必要です。');
      const lastMonthDate = new Date(today.getFullYear(), today.getMonth(), 0);
      const lastMonthYear = lastMonthDate.getFullYear();
      const lastMonth = lastMonthDate.getMonth() + 1;
      let lastMonthData = { goal: '', reflection: '', comment: '' };
      const lastMonthRow = data.slice(1).find(row =>
        row[0].toString().trim() === userId && 
        parseInt(row[2], 10) === lastMonthYear && 
        parseInt(row[3], 10) === lastMonth
      );
      if (lastMonthRow) {
        lastMonthData.goal = lastMonthRow[4] || '';
      }
      
      return { 
        required: true, 
        lastMonthData: lastMonthData,
        currentYear: currentYear,
        currentMonth: currentMonth,
        lastMonthYear: lastMonthYear,
        lastMonth: lastMonth
      };
    }

    Logger.log('[目標チェック] 結果: 目標設定は不要です。');
    return { required: false };
  } catch (e) {
    Logger.log(`[目標チェックエラー] ${e.toString()}`);
    return { required: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveGoalAndReflection(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[目標保存開始] 受信データ: ${JSON.stringify(data)}`);

    const { userId, currentYear, currentMonth, newGoal, lastMonthYear, lastMonth, reflection } = data;
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) throw new Error("目標管理シートが見つかりません。");
    const studentName = getStudentNameById(userId);
    if (!studentName) throw new Error("生徒が見つかりません。");
    const allData = sheet.getDataRange().getValues();

    if (reflection && lastMonthYear && lastMonth) {
      let rowIndex = -1;
      for(let i = 1; i < allData.length; i++) {
        if (allData[i][0].toString().trim() === userId && parseInt(allData[i][2], 10) === lastMonthYear && parseInt(allData[i][3], 10) === lastMonth) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 6).setValue(reflection);
        Logger.log(`${studentName}の${lastMonthYear}年${lastMonth}月の振り返りを更新しました。`);
      }
    }
    
    sheet.appendRow([userId, studentName, currentYear.toString(), currentMonth.toString(), newGoal, '', '']);
    Logger.log(`${studentName}の${currentYear}年${currentMonth}月の目標を追加しました。`);
    
    Logger.log(`[目標保存完了]`);
    return { success: true };
  } catch (e) {
    Logger.log(`[目標保存エラー] ${e.toString()}`);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveCoachComment(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { userId, year, month, comment } = data;
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) throw new Error("目標管理シートが見つかりません。");
    const allData = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for(let i = 1; i < allData.length; i++) {
      if (allData[i][0].toString().trim() === userId && parseInt(allData[i][2], 10) === year && parseInt(allData[i][3], 10) === month) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 7).setValue(comment);
      Logger.log(`Comment saved for ${userId} for ${year}-${month}.`);
      return { success: true };
    } else {
      throw new Error("対象の目標データが見つかりません。");
    }
  } catch (e) {
    Logger.log(`Error in saveCoachComment: ${e}`);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getGoalData(userId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[目標データ取得開始] ユーザーID: ${userId}`);

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) return { success: true, data: null };
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    const lastMonthDate = new Date(today.getFullYear(), today.getMonth(), 0);
    const lastMonthYear = lastMonthDate.getFullYear();
    const lastMonth = lastMonthDate.getMonth() + 1;

    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[目標データ取得] 検索対象: ${currentYear}年${currentMonth}月 と ${lastMonthYear}年${lastMonth}月`);

    const allData = sheet.getDataRange().getValues();
    let currentGoalData = null;
    let lastMonthGoalData = null;
    for (let i = allData.length - 1; i >= 1; i--) {
      const rowUserId = allData[i][0].toString().trim();
      if (rowUserId === userId) {
        const year = parseInt(allData[i][2], 10);
        const month = parseInt(allData[i][3], 10);
        if (year === currentYear && month === currentMonth && !currentGoalData) {
          currentGoalData = {
            year: year, month: month,
            goal: allData[i][4] || '', reflection: allData[i][5] || '', comment: allData[i][6] || ''
          };
        } else if (year === lastMonthYear && month === lastMonth && !lastMonthGoalData) {
          lastMonthGoalData = {
            year: year, month: month,
            goal: allData[i][4] || '', reflection: allData[i][5] || '', comment: allData[i][6] || ''
          };
        }
      }
      if (currentGoalData && lastMonthGoalData) break;
    }
    
    // ▼▼▼ ログ追加 ▼▼▼
    Logger.log(`[目標データ取得] 検索結果: {今月データ: ${currentGoalData ? 'あり' : 'なし'}, 先月データ: ${lastMonthGoalData ? 'あり' : 'なし'}}`);
    
    return { success: true, data: { current: currentGoalData, last: lastMonthGoalData } };
  } catch (e) {
    Logger.log(`[目標データ取得エラー] ${e.toString()}`);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getStudentsNeedingComment() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) return { success: false, message: '目標管理シートが見つかりません。' };
    const today = new Date();
    const lastMonthDate = new Date(today.getFullYear(), today.getMonth(), 0);
    const targetYear = lastMonthDate.getFullYear();
    const targetMonth = lastMonthDate.getMonth() + 1;
    const allData = sheet.getDataRange().getValues();
    const studentsNeedingComment = [];
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const userId = row[0] ? row[0].toString().trim() : '';
      const studentName = row[1] ? row[1].toString().trim() : '';
      const year = parseInt(row[2], 10);
      const month = parseInt(row[3], 10);
      const reflection = row[5] ? row[5].toString().trim() : '';
      const comment = row[6] ? row[6].toString().trim() : '';
      if (year === targetYear && month === targetMonth && reflection !== '' && comment === '') {
        studentsNeedingComment.push({
          userId: userId, studentName: studentName, year: year, month: month, reflection: reflection
        });
      }
    }
    return { success: true, data: studentsNeedingComment, targetYear: targetYear, targetMonth: targetMonth };
  } catch (e) {
    Logger.log(`Error in getStudentsNeedingComment: ${e.toString()}`);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------
// 学習状況・アクション記録関連
// ----------------------------------------
function getStudentCurrentStatus(userId, statusSheet) {
  if (!statusSheet) return null;
  const data = statusSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === userId.trim()) {
      return {
        rowIndex: i,
        studentName: data[i][1] ? data[i][1].toString() : '',
        isLearning: data[i][2] === true, // ★より厳密にbooleanのtrueのみをチェック
        startTime: getValidDate(data[i][3]) // ★getValidDateを使用
      };
    }
  }
  return null;
}

function setStudentLearningStatus(userId, studentName, startTime, statusSheet) {
  if (!statusSheet) {
    return;
  }
  const status = getStudentCurrentStatus(userId, statusSheet);
  // ★★★ 先頭にシングルクォートを追加して、強制的に文字列として書き込む ★★★
  const startTimeStr = "'" + Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  if (status) {
    // C列とD列を一度に更新
    statusSheet.getRange(status.rowIndex + 1, 3, 1, 2).setValues([[true, startTimeStr]]);
  } else {
    statusSheet.appendRow([userId, studentName, true, startTimeStr]);
  }
  Logger.log(`学習状況を更新(開始): UserID=${userId}, Name=${studentName}, StartTime=${startTimeStr}`);
}

function clearStudentLearningStatus(userId, statusSheet) {
  if (!statusSheet) return;
  const status = getStudentCurrentStatus(userId, statusSheet);
  if (status) {
    // C列をfalseに、D列を空にする
    statusSheet.getRange(status.rowIndex + 1, 3, 1, 2).setValues([[false, '']]);
    Logger.log(`学習状況を更新(終了): UserID=${userId}`);
  }
}

function recordAction(userId, studentName, action, options = {}) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const endTime = options.endTime || new Date();
    // ★★★ 先頭にシングルクォートを追加 ★★★
    const timestamp = "'" + Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const dateStr = Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = spreadsheet.getSheetByName(SHEET_ATTENDANCE_LOG);
    const summarySheet = spreadsheet.getSheetByName(SHEET_SUMMARY);
    const statusSheet = spreadsheet.getSheetByName(SHEET_CURRENT_STATUS);
    if (!logSheet || !summarySheet || !statusSheet) {
        return { success: false, message: '必要なシートが見つかりません。' };
    }

    const currentStatus = getStudentCurrentStatus(userId, statusSheet);
      
    if (action === '開始') {
        if (currentStatus && currentStatus.isLearning) {
          const alreadyStartedTime = currentStatus.startTime ? Utilities.formatDate(currentStatus.startTime, Session.getScriptTimeZone(), 'HH:mm') : "不明";
          return { success: false, message: `既に ${alreadyStartedTime} から学習を開始しています。` };
        }
        setStudentLearningStatus(userId, studentName, new Date(), statusSheet);
      // ★★★ こちらの書き込みにもシングルクォートを追加 ★★★
        logSheet.appendRow(["'" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), userId, studentName, action, dateStr]);
      return { success: true, message: `${studentName}さんの学習を開始しました。(${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm')})` };
    } else if (action === '終了') {
        if (!currentStatus || !currentStatus.isLearning || !currentStatus.startTime) {
          return { success: false, message: 'まだ学習を開始していません。' };
        }
        const startTime = currentStatus.startTime;
        let sessionMinutes = 0;
        if (startTime) { // startTimeがnullでないことを確認
        sessionMinutes = Math.round((endTime.getTime() - startTime.getTime()) / (1000 * 60));
        }
        
        clearStudentLearningStatus(userId, statusSheet);
        logSheet.appendRow([timestamp, userId, studentName, action, dateStr]);
        updateStudentSummaryAfterSession(userId, studentName, dateStr, sessionMinutes, summarySheet, logSheet);
      return { success: true, message: `${studentName}さんの学習を終了しました。今回の学習時間: ${sessionMinutes}分` };
    }
    return { success: false, message: '不明な操作です。' };
  } finally {
    lock.releaseLock();
  }
}

function updateStudentSummaryAfterSession(userId, studentName, currentActionDateStr, sessionMinutes, summarySheet, attendanceLogSheet) {
  const summaryData = summarySheet.getDataRange().getValues();
  let summaryRowIndex = -1;
  for (let i = 1; i < summaryData.length; i++) {
    if (summaryData[i][0] && summaryData[i][0].toString().trim() === userId.trim()) {
      summaryRowIndex = i;
      break;
    }
  }

  const logData = attendanceLogSheet.getDataRange().getValues();
  const userRecordsForToday = [];
  for (let i = 1; i < logData.length; i++) {
    let logActionDateStr = '';
    const dateVal = getValidDate(logData[i][4]);
    if (dateVal) logActionDateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (logData[i][1] && logData[i][1].toString().trim() === userId.trim() && logActionDateStr === currentActionDateStr.trim()) {
      const ts = getValidDate(logData[i][0]);
      if (ts) userRecordsForToday.push({ timestamp: ts, action: logData[i][3] });
    }
  }
  userRecordsForToday.sort((a, b) => a.timestamp - b.timestamp);
  let todayFirstStartTimeStr = '';
  let todayLastEndTimeStr = '';
  const firstStartRec = userRecordsForToday.find(r => r.action === '開始');
  if (firstStartRec) todayFirstStartTimeStr = Utilities.formatDate(firstStartRec.timestamp, Session.getScriptTimeZone(), 'HH:mm:ss');
  
  const endRecs = userRecordsForToday.filter(r => r.action === '終了');
  if (endRecs.length > 0) todayLastEndTimeStr = Utilities.formatDate(endRecs[endRecs.length - 1].timestamp, Session.getScriptTimeZone(), 'HH:mm:ss');
  let currentDaily = 0, currentMonthly = 0, currentOverall = 0, prevDateStr = "";
  if (summaryRowIndex !== -1) {
    prevDateStr = summaryData[summaryRowIndex][2] ? summaryData[summaryRowIndex][2].toString().trim() : "";
    currentDaily = parseFloat(summaryData[summaryRowIndex][5]) || 0;
    currentMonthly = parseFloat(summaryData[summaryRowIndex][6]) || 0;
    currentOverall = parseFloat(summaryData[summaryRowIndex][7]) || 0;
  }

  currentDaily = (prevDateStr === currentActionDateStr.trim()) ? currentDaily + sessionMinutes : sessionMinutes;
  
  const currentMonthStr = Utilities.formatDate(new Date(currentActionDateStr), Session.getScriptTimeZone(), 'yyyy-MM');
  const prevMonthStr = prevDateStr ? Utilities.formatDate(new Date(prevDateStr), Session.getScriptTimeZone(), 'yyyy-MM') : "";
  currentMonthly = (prevMonthStr === currentMonthStr) ? currentMonthly + sessionMinutes : currentDaily;

  currentOverall += sessionMinutes;
  if (summaryRowIndex !== -1) {
    summarySheet.getRange(summaryRowIndex + 1, 3, 1, 6).setValues([[currentActionDateStr, todayFirstStartTimeStr, todayLastEndTimeStr, currentDaily, currentMonthly, currentOverall]]);
  } else {
    summarySheet.appendRow([userId, studentName, currentActionDateStr, todayFirstStartTimeStr, todayLastEndTimeStr, currentDaily, currentMonthly, currentOverall]);
  }
  return { success: true };
}

// ----------------------------------------
// 管理者向け・補助関数
// ----------------------------------------
function getRealTimeStatus() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const statusSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CURRENT_STATUS);
    if (!statusSheet) throw new Error("「学習状況」シートが見つかりません。");
    
    const data = statusSheet.getDataRange().getValues();
    const learningStudents = [];
    const now = new Date();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[2] === true) { // C列がtrueの行のみを対象
        const startTime = getValidDate(row[3]);
        if (startTime) {
          const duration = Math.round((now.getTime() - startTime.getTime()) / (1000 * 60));
          learningStudents.push({
            userId: row[0] || 'ID不明',
            studentName: row[1] || '名前不明',
            startTime: Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm'),
            duration: duration > 0 ? duration : 0
          });
        }
      }
    }
    return { success: true, data: learningStudents };
  } catch (error) {
    Logger.log('Error in getRealTimeStatus: ' + error.toString());
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function forceEndStudy(userId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const studentName = getStudentNameById(userId);
    if (!studentName) throw new Error(`生徒IDが見つかりません: ${userId}`);
    return recordAction(userId, studentName, '終了');
  } catch (error) {
    Logger.log(`Error in forceEndStudy for ${userId}: ` + error.toString());
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function autoEndOverdueStudiesAt2230() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const statusSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CURRENT_STATUS);
    if (!statusSheet) return;
    const data = statusSheet.getDataRange().getValues();
    const now = new Date();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][2] === true) {
        const endTime = new Date();
        endTime.setHours(22, 30, 0, 0);
        if (now >= endTime) {
          recordAction(data[i][0], data[i][1], '終了', { endTime: endTime });
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function getAllStudents() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const masterSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_STUDENT_MASTER);
    if (!masterSheet) throw new Error("「生徒マスタ」シートが見つかりません。");
    const data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 2).getValues();
    return { success: true, data: data.map(row => ({ id: row[0], name: row[1] })) };
  } catch (error) {
    Logger.log('Error in getAllStudents: ' + error.toString());
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getStudentData(userId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summarySheet = spreadsheet.getSheetByName(SHEET_SUMMARY);
    const monthlySheet = spreadsheet.getSheetByName(SHEET_MONTHLY_SUMMARY);
    const logSheet = spreadsheet.getSheetByName(SHEET_ATTENDANCE_LOG);
    const goalSheet = spreadsheet.getSheetByName(SHEET_GOAL);

    // 1. サマリー取得
    let summary = { isDataFound: false };
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = 1; i < summaryData.length; i++) {
      if (summaryData[i][0] && summaryData[i][0].toString().trim() === userId) {
        const lastActivityDate = getValidDate(summaryData[i][2]);
        summary = {
          isDataFound: true,
          lastActivityDate: lastActivityDate ? Utilities.formatDate(lastActivityDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '記録なし',
          monthlyTotal: summaryData[i][6] || 0,
          overallTotal: summaryData[i][7] || 0
        };
        break;
      }
    }

    // 2. 年間グラフデータ取得
    const yearlyStudyData = [['月', '勉強時間(実績)', 'トレンドライン']];
    const monthNames = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"];
    if (summary.isDataFound) {
      const yearlyTotals = Array(12).fill(0);
      const currentYearStr = new Date().getFullYear().toString();

      const monthlyData = monthlySheet.getDataRange().getValues();
      for (let i = 1; i < monthlyData.length; i++) {
        const rowUserId = monthlyData[i][0] ? monthlyData[i][0].toString().trim() : '';
        const yearMonth = monthlyData[i][2] ? monthlyData[i][2].toString() : '';
        if (rowUserId === userId && yearMonth.startsWith(currentYearStr)) {
          const monthIndex = parseInt(yearMonth.split('-')[1], 10) - 1;
          if(monthIndex >= 0 && monthIndex < 12){
              yearlyTotals[monthIndex] = parseFloat(monthlyData[i][3]) || 0;
          }
        }
      }
      yearlyTotals[new Date().getMonth()] = summary.monthlyTotal;
      for (let i = 0; i < 12; i++) {
          yearlyStudyData.push([monthNames[i], yearlyTotals[i], yearlyTotals[i]]);
      }
    } else {
      // ★★★ ここが修正点 ★★★
      // データが見つからない場合（新規生徒など）は、全て0のグラフデータを作成する
      for (let i = 0; i < 12; i++) {
        yearlyStudyData.push([monthNames[i], 0, 0]);
      }
    }

    // 3. 入退室ログ取得
    const logs = [];
    if(summary.isDataFound) {
        const logData = logSheet.getDataRange().getValues();
        for (let i = logData.length - 1; i >= 1; i--) {
          if (logData[i][1] && logData[i][1].toString().trim() === userId) {
            const ts = getValidDate(logData[i][0]);
            if (ts) {
              logs.push({
                timestamp: Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
                action: logData[i][3]
              });
            }
            if (logs.length >= 50) break;
          }
        }
    }
    
    // 4. 目標履歴取得 (過去1年分)
    const goalHistory = [];
    if (goalSheet) {
      const goalData = goalSheet.getDataRange().getValues();
      const oneYearAgo = new Date();
      oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
      oneYearAgo.setDate(1);
      for (let i = goalData.length - 1; i >= 1; i--) {
        const rowUserId = goalData[i][0].toString().trim();
        if (rowUserId === userId) {
          const year = parseInt(goalData[i][2], 10);
          const month = parseInt(goalData[i][3], 10) -1;
          const goalDate = new Date(year, month, 1);

          if (goalDate >= oneYearAgo) {
            goalHistory.push({
              year: goalData[i][2].toString(),
              month: goalData[i][3].toString(),
              goal: goalData[i][4] || '',
              reflection: goalData[i][5] || '',
              comment: goalData[i][6] || ''
            });
          }
        }
      }
    }

    return { 
      success: true, 
      data: { summary: summary, yearlyData: yearlyStudyData, logs: logs, goalHistory: goalHistory }
    };
  } catch (error) {
    Logger.log(`Error in getStudentData for ${userId}: ` + error.toString() + " Stack: " + error.stack);
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getStudentNameById(userId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_STUDENT_MASTER);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === userId.trim()) {
        return data[i][1].toString();
      }
    }
    return null;
  } catch(e) {
    Logger.log(`Error in getStudentNameById: ${e}`);
    return null;
  } finally {
    lock.releaseLock();
  }
}

function archiveMonthlySummaryAndReset() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const today = new Date();
    const firstDayOfCurrentMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    // 1日前の日付を取得することで、先月の最終日を得る
    const lastDayOfPreviousMonth = new Date(firstDayOfCurrentMonth.getTime() - 1);
    // フォーマットを 'YYYY-MM' 形式にする
    const targetYearMonth = Utilities.formatDate(lastDayOfPreviousMonth, Session.getScriptTimeZone(), 'yyyy-MM');

    Logger.log(`月次集計とリセット処理を開始します。対象年月: ${targetYearMonth}`);

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summarySheet = spreadsheet.getSheetByName(SHEET_SUMMARY);
    const monthlySummarySheet = spreadsheet.getSheetByName(SHEET_MONTHLY_SUMMARY);

    if (!summarySheet || !monthlySummarySheet) {
      Logger.log('必要なシートが見つかりませんでした。処理を中断します。');
      return;
    }

    // ヘッダー行を除いてデータを取得 (2行目から最終行まで)
    const summaryData = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 7).getValues();
    const dataToArchive = [];
    const rowsToReset = [];

    summaryData.forEach((row, index) => {
      const studentId = row[0]; // A列: 生徒ID
      const studentName = row[1]; // B列: 生徒名
      const monthlyMinutes = parseFloat(row[6]) || 0; // G列: 今月の累計(分)

      // 月間学習時間が0分より大きい場合のみ処理
      if (studentId && monthlyMinutes > 0) {
        // 1. 月別集計シートにアーカイブするデータを作成
        dataToArchive.push([
          studentId,
          studentName,
          targetYearMonth,
          
          monthlyMinutes
        ]);
        Logger.log(`アーカイブ対象: ${studentName}さん (${studentId}) の ${targetYearMonth} の学習時間 ${monthlyMinutes} 分`);

        // 2. サマリーシートでリセットする行のインデックスを記録
        // getRangeは1から始まるため、indexに2を足す
        rowsToReset.push(index + 2); 
      }
    });
    // アーカイブデータがあれば、月別集計シートに一括で追記
    if (dataToArchive.length > 0) {
      monthlySummarySheet.getRange(monthlySummarySheet.getLastRow() + 1, 1, dataToArchive.length, 4).setValues(dataToArchive);
      Logger.log(`${dataToArchive.length} 件のデータを月別集計シートに転記しました。`);
    } else {
      Logger.log('アーカイブ対象のデータはありませんでした。');
    }

    // カウンターをリセット
    rowsToReset.forEach(rowIndex => {
      // C列(最終活動日)からG列(今月の累計)までをリセット
      // 最終活動日、初回・最終時刻はクリアし、日次・月次累計は0にする
      summarySheet.getRange(rowIndex, 3, 1, 5).setValues([['', '', '', 0, 0]]);
    });
    if (rowsToReset.length > 0) {
      Logger.log(`${rowsToReset.length} 人の生徒の学習時間カウンターをリセットしました。`);
    }

    Logger.log('月次集計とリセット処理が完了しました。');
  } finally {
    lock.releaseLock();
  }
}