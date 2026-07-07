// ----------------------------------------
// グローバル設定
// ----------------------------------------
const SPREADSHEET_ID = '1qTSkENn-CxotdKvoetRF7CthWSzsOHbDMzdQexslGV4'; // コード修正用
const SHEET_STUDENT_MASTER = '生徒マスタ';
const SHEET_ATTENDANCE_LOG = '入退室記録';
const SHEET_SUMMARY = '学習時間サマリー';
const SHEET_MONTHLY_SUMMARY = '月別学習時間集計(テスト)';
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
    const normalized = value.trim().replace(/^'/, '');
    // "2025-07-29 19:19:35" を [2025, 7, 29, 19, 19, 35] のような数値の配列に変換
    const parts = normalized.split(/[\s:-]/).map(part => parseInt(part, 10));
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

function normalizeUserId(value) {
  const raw = String(value == null ? '' : value)
    .trim()
    .replace(/^['"]+|['"]+$/g, '');
  if (!raw) return '';
  if (/^\d+$/.test(raw)) {
    return String(parseInt(raw, 10));
  }
  return raw;
}

function normalizeStudentName(value) {
  return String(value == null ? '' : value)
    .trim()
    .replace(/^['"]+|['"]+$/g, '');
}

const SESSION_DURATION_MS = 10 * 60 * 60 * 1000; // 10時間

function createSession(userId, studentName, isAdmin) {
  const normalizedUserId = normalizeUserId(userId);
  const normalizedStudentName = normalizeStudentName(studentName);
  const token = Utilities.getUuid();
  const expiresAt = Date.now() + SESSION_DURATION_MS;

  const sessionData = {
    userId: normalizedUserId,
    studentName: normalizedStudentName,
    isAdmin: !!isAdmin,
    expiresAt: expiresAt
  };

  PropertiesService.getScriptProperties().setProperty(
    'session_' + token,
    JSON.stringify(sessionData)
  );

  return {
    token: token,
    expiresAt: expiresAt
  };
}

function getSession(token) {
  if (!token) return null;

  const raw = PropertiesService.getScriptProperties().getProperty('session_' + token);
  if (!raw) return null;

  try {
    const session = JSON.parse(raw);

    if (!session.expiresAt || Date.now() > session.expiresAt) {
      PropertiesService.getScriptProperties().deleteProperty('session_' + token);
      return null;
    }

    return session;
  } catch (e) {
    PropertiesService.getScriptProperties().deleteProperty('session_' + token);
    return null;
  }
}

function deleteSession(token) {
  if (!token) return;
  PropertiesService.getScriptProperties().deleteProperty('session_' + token);
}

function requireValidSession(token) {
  const session = getSession(token);
  if (!session) {
    throw new Error('セッションの有効期限が切れています。もう一度ログインしてください。');
  }
  return session;
}

// ----------------------------------------
// Webアプリケーションのエントリーポイント (doGet)
// ----------------------------------------
function doGet(e) {
  Logger.log('doGet called with parameters: ' + JSON.stringify(e.parameter));
  const page = e.parameter.page;
  const token = e.parameter.token || '';

  if (page === 'admin') {
    const session = getSession(token);
    if (!session || !session.isAdmin) return showLoginPage(e);
    return showAdminPage(e, session);
  } else if (page === 'main') {
    const session = getSession(token);
    if (!session) return showLoginPage(e);
    return showMainPage(e, session);
  } else if (page === 'goal') {
    const session = getSession(token);
    if (!session) return showLoginPage(e);
    return showGoalPage(e, session);
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

function showAdminPage(e, session) {
  Logger.log('Rendering admin page.');
  const template = HtmlService.createTemplateFromFile('admin');
  template.webAppUrl = ScriptApp.getService().getUrl();
  template.sessionToken = e.parameter.token || '';
  template.sessionExpiresAt = session.expiresAt;
  return template.evaluate().setTitle('管理者用ダッシュボード');
}

function showMainPage(e, session) {
  const currentUserId = normalizeUserId(session.userId);
  const currentStudentName = normalizeStudentName(session.studentName);
  Logger.log('Rendering main page for user: ' + currentUserId);
  const template = HtmlService.createTemplateFromFile('index');
  template.userId = currentUserId;
  template.studentName = currentStudentName;
  template.sessionToken = e.parameter.token || '';
  template.sessionExpiresAt = session.expiresAt;
  template.webAppUrl = ScriptApp.getService().getUrl();

  const today = new Date();
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
  template.todayAsString = Utilities.formatString("%d年%02d月%02d日（%s）", today.getFullYear(), today.getMonth() + 1, today.getDate(), weekdays[today.getDay()]);

  let monthlyStudyTime = 0;
  let dailyChartBundle = {
    currentMonth: '',
    months: [],
    dailyByMonth: {}
  };

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summarySheet = spreadsheet.getSheetByName(SHEET_SUMMARY);
    const attendanceLogSheet = spreadsheet.getSheetByName(SHEET_ATTENDANCE_LOG);

    if (summarySheet) {
      const summaryData = summarySheet.getDataRange().getValues();
      const currentYearMonth = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM');

      for (let i = summaryData.length - 1; i >= 1; i--) {
        if (summaryData[i][0] && summaryData[i][0].toString().trim() === currentUserId) {
          let summaryYearMonth = '';
          const summaryDate = getValidDate(summaryData[i][2]);

          if (summaryDate) {
            summaryYearMonth = Utilities.formatDate(summaryDate, Session.getScriptTimeZone(), 'yyyy-MM');
          } else if (summaryData[i][2]) {
            summaryYearMonth = summaryData[i][2].toString().substring(0, 7).trim();
          }

          monthlyStudyTime = (summaryYearMonth === currentYearMonth)
            ? (parseFloat(summaryData[i][6]) || 0)
            : 0;
          if (summaryYearMonth === currentYearMonth) {
            break;
          }
        }
      }
    }

    if (attendanceLogSheet) {
      dailyChartBundle = buildDailyChartBundle(
        currentUserId,
        attendanceLogSheet.getDataRange().getValues(),
        today
      );
    }

    template.dailyChartDataJson = JSON.stringify(dailyChartBundle);
    template.monthlyStudyTime = monthlyStudyTime;

      // --- 追加: 今月の学習時間の順位を計算してテンプレートに渡す ---
      try {
        // 指示に従い、`学習時間サマリー` シートの A列=userId, G列=月間合計 を用いてランキングを算出
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const summary = ss.getSheetByName(SHEET_SUMMARY);
        let studentRank = '';
        let studentRankTotal = 0;

      if (summary) {
        const vals = summary.getDataRange().getValues();
        const masterSheet = ss.getSheetByName(SHEET_STUDENT_MASTER);
        const masterData = masterSheet.getDataRange().getValues();
        const currentYearMonth = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM');

        const list = [];

        // 生徒マスタの全員をランキング対象にする
        for (let i = 1; i < masterData.length; i++) {
          const uid = masterData[i][0] ? String(masterData[i][0]).trim() : '';
          if (!uid) continue;

          let minutes = 0;

          // 学習時間サマリーから、その生徒の今月の学習時間を探す
          for (let j = 1; j < vals.length; j++) {
            const rowUid = vals[j][0] ? String(vals[j][0]).trim() : '';
            if (rowUid !== uid) continue;

            let rowYearMonth = '';
            const rowDate = getValidDate(vals[j][2]);

            if (rowDate) {
              rowYearMonth = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM');
            } else if (vals[j][2]) {
              rowYearMonth = String(vals[j][2]).substring(0, 7).trim();
            }

            if (rowYearMonth === currentYearMonth) {
              const g = vals[j][6];
              if (typeof g === 'number') {
                minutes = g;
              } else if (g) {
                const digits = String(g).replace(/[^0-9-]/g, '');
                minutes = digits ? parseInt(digits, 10) || 0 : 0;
              }
              break;
            }
          }

          list.push({ userId: uid, monthly: minutes });
        }

        // 分母は生徒マスタの人数
        studentRankTotal = list.length;

        // 勉強時間の多い順に並べる
        list.sort(function(a, b) { return b.monthly - a.monthly; });

        // 順位を計算
        let rank = 0;
        let prev = null;
        const currentId = currentUserId;

        for (let i = 0; i < list.length; i++) {
          if (prev === null || list[i].monthly !== prev) {
            rank++;
            prev = list[i].monthly;
          }
          if (String(list[i].userId) === currentId) {
            studentRank = rank;
            break;
          }
        }
      }

        template.studentRank = studentRank;
        template.studentRankTotal = studentRankTotal;
      } catch (rankErr) {
        template.studentRank = '';
        template.studentRankTotal = '';
        Logger.log('Error computing student rank from summary G column: ' + rankErr);
      }
  } catch (err) {
    Logger.log('Error fetching study data for main page: ' + err.toString());
    template.monthlyStudyTime = 0;
    template.dailyChartDataJson = JSON.stringify({
      currentMonth: '',
      months: [],
      dailyByMonth: {}
    });
  }
  
  return template.evaluate().setTitle('自習室管理システム');
}

function showGoalPage(e, session) {
  Logger.log('Rendering goal setting page for user: ' + session.userId);
  const template = HtmlService.createTemplateFromFile('goal');
  template.userId = normalizeUserId(session.userId);
  template.studentName = normalizeStudentName(session.studentName);
  template.sessionToken = e.parameter.token || '';
  template.sessionExpiresAt = session.expiresAt;
  template.webAppUrl = ScriptApp.getService().getUrl();
  template.lastMonthGoal = e.parameter.lastMonthGoal ? decodeURIComponent(e.parameter.lastMonthGoal) : '';
  template.currentYear = e.parameter.currentYear;
  template.currentMonth = e.parameter.currentMonth;
  template.lastMonthYear = e.parameter.lastMonthYear;
  template.lastMonth = e.parameter.lastMonth;

  return template.evaluate().setTitle('目標設定');
}

function authenticateUser(userId, password) {
  try {
    const normalizedUserId = normalizeUserId(userId);
    if (userId === 'admin' && password === 'admin_password') {
      const sessionInfo = createSession(userId, 'admin', true);
      return {
        success: true,
        isAdmin: true,
        sessionToken: sessionInfo.token,
        sessionExpiresAt: sessionInfo.expiresAt
      };
    }
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_STUDENT_MASTER);
    if (!sheet) throw new Error('生徒マスタが見つかりません。');
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (normalizeUserId(data[i][0]) === normalizedUserId && data[i][2].toString() === password) {
        const studentName = normalizeStudentName(data[i][1]);
        Logger.log(`Authentication successful for ${studentName}`);
        
        const goalStatus = checkGoalSettingStatus(normalizedUserId);
        const sessionInfo = createSession(normalizedUserId, studentName, false);
        return { 
          success: true, 
          isAdmin: false, 
          studentName: studentName,
          goalStatus: goalStatus,
          sessionToken: sessionInfo.token,
          sessionExpiresAt: sessionInfo.expiresAt
        };
      }
    }
    Logger.log('Authentication failed for user: ' + userId);
    return { success: false, message: '生徒IDまたはパスワードが正しくありません。' };
  } catch (error) {
    Logger.log('Authentication error: ' + error.toString());
    return { success: false, message: `認証エラー: ${error.message}` };
  }
}

function checkGoalSettingStatus(userId) {
  try {
    const normalizedUserId = normalizeUserId(userId);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) {
      Logger.log('目標管理シートが見つからないため、目標設定をスキップします。');
      return { required: false };
    }
    
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    
    const data = sheet.getDataRange().getValues();
    const currentMonthGoalExists = data.slice(1).some(row => 
      normalizeUserId(row[0]) === normalizedUserId && 
      parseInt(row[2], 10) === currentYear && 
      parseInt(row[3], 10) === currentMonth
    );
    
    if (!currentMonthGoalExists) {
      const lastMonthDate = new Date(today.getFullYear(), today.getMonth(), 0);
      const lastMonthYear = lastMonthDate.getFullYear();
      const lastMonth = lastMonthDate.getMonth() + 1;
      let lastMonthData = { goal: '', reflection: '', comment: '' };
      
      const lastMonthRow = data.slice(1).find(row =>
        normalizeUserId(row[0]) === normalizedUserId && 
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
    return { required: false };
  } catch (e) {
    Logger.log(`Error in checkGoalSettingStatus for ${userId}: ${e}`);
    return { required: false, error: e.toString() };
  }
}

function saveGoalAndReflection(data) {
  try {
    const { userId, studentName, sessionToken, currentYear, currentMonth, newGoal, lastMonthYear, lastMonth, reflection } = data;
    const session = sessionToken ? getSession(sessionToken) : null;
    const effectiveUserId = session && session.userId ? normalizeUserId(session.userId) : normalizeUserId(userId);
    const fallbackStudentName = normalizeStudentName(studentName);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) throw new Error("目標管理シートが見つかりません。");
    const effectiveStudentName =
      (session && session.studentName ? normalizeStudentName(session.studentName) : '') ||
      fallbackStudentName ||
      getStudentNameById(effectiveUserId);
    if (!effectiveUserId) throw new Error("生徒IDが取得できません。");
    if (!effectiveStudentName) throw new Error("生徒が見つかりません。");
    const allData = sheet.getDataRange().getValues();

    if (reflection && lastMonthYear && lastMonth) {
      let rowIndex = -1;
      for(let i = 1; i < allData.length; i++) {
        if (normalizeUserId(allData[i][0]) === effectiveUserId && parseInt(allData[i][2], 10) === lastMonthYear && parseInt(allData[i][3], 10) === lastMonth) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 6).setValue(reflection);
        Logger.log(`${effectiveStudentName}の${lastMonthYear}年${lastMonth}月の振り返りを更新しました。`);
      }
    }
    
    sheet.appendRow([effectiveUserId, effectiveStudentName, currentYear.toString(), currentMonth.toString(), newGoal, '', '']);
    Logger.log(`${effectiveStudentName}の${currentYear}年${currentMonth}月の目標を追加しました。`);

    const webAppUrl = ScriptApp.getService().getUrl();
    let redirectToken = sessionToken;
    if (!session) {
      const sessionInfo = createSession(effectiveUserId, effectiveStudentName, false);
      redirectToken = sessionInfo.token;
    }

    const redirectUrl = `${webAppUrl}?page=main&token=${encodeURIComponent(redirectToken)}`;
    
    return { success: true, redirectUrl: redirectUrl };
  } catch (e) {
    Logger.log(`Error in saveGoalAndReflection: ${e}`);
    return { success: false, message: e.toString() };
  }
}

function saveCoachComment(data) {
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
  }
}

function getGoalData(userId) {
  try {
    const normalizedUserId = normalizeUserId(userId);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_GOAL);
    if (!sheet) return { success: true, data: null };
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    const lastMonthDate = new Date(today.getFullYear(), today.getMonth(), 0);
    const lastMonthYear = lastMonthDate.getFullYear();
    const lastMonth = lastMonthDate.getMonth() + 1;
    const allData = sheet.getDataRange().getValues();
    let currentGoalData = null;
    let lastMonthGoalData = null;
    for (let i = allData.length - 1; i >= 1; i--) {
      const rowUserId = normalizeUserId(allData[i][0]);
      if (rowUserId === normalizedUserId) {
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
    return { success: true, data: { current: currentGoalData, last: lastMonthGoalData } };
  } catch (e) {
    Logger.log(`Error in getGoalData for ${userId}: ${e}`);
    return { success: false, message: e.toString() };
  }
}

function buildDailyChartDataForMonth(userId, logData, targetYear, targetMonthZeroBased, lastDay) {
  const normalizedUserId = normalizeUserId(userId);
  const dailyTotals = Array(lastDay).fill(0);
  const records = [];

  for (let i = 1; i < logData.length; i++) {
    const rowUserId = normalizeUserId(logData[i][1]);
    if (rowUserId !== normalizedUserId) continue;

    const ts = getValidDate(logData[i][0]);
    if (!ts) continue;
    if (ts.getFullYear() !== targetYear || ts.getMonth() !== targetMonthZeroBased) continue;

    records.push({
      timestamp: ts,
      action: logData[i][3] ? String(logData[i][3]).trim() : ''
    });
  }

  records.sort(function(a, b) {
    return a.timestamp - b.timestamp;
  });

  let currentStart = null;
  for (let i = 0; i < records.length; i++) {
    const record = records[i];

    if (record.action === '開始') {
      currentStart = record.timestamp;
      continue;
    }

    if (record.action === '終了' && currentStart) {
      const sameDay =
        currentStart.getFullYear() === record.timestamp.getFullYear() &&
        currentStart.getMonth() === record.timestamp.getMonth() &&
        currentStart.getDate() === record.timestamp.getDate();

      if (sameDay) {
        const diffMinutes = Math.max(
          0,
          Math.round((record.timestamp.getTime() - currentStart.getTime()) / (1000 * 60))
        );
        dailyTotals[currentStart.getDate() - 1] += diffMinutes;
      }
      currentStart = null;
    }
  }

  const chartData = [['日', '勉強時間(分)']];
  for (let day = 1; day <= lastDay; day++) {
    chartData.push([String(day), dailyTotals[day - 1]]);
  }
  return chartData;
}

function buildDailyChartBundle(userId, logData, referenceDate) {
  const normalizedUserId = normalizeUserId(userId);
  const targetDate = referenceDate || new Date();
  const monthMap = {};

  for (let i = 1; i < logData.length; i++) {
    const rowUserId = normalizeUserId(logData[i][1]);
    if (rowUserId !== normalizedUserId) continue;

    const ts = getValidDate(logData[i][0]);
    if (!ts) continue;

    const key = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM');
    monthMap[key] = {
      year: ts.getFullYear(),
      monthZeroBased: ts.getMonth()
    };
  }

  const currentKey = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM');
  if (!monthMap[currentKey]) {
    monthMap[currentKey] = {
      year: targetDate.getFullYear(),
      monthZeroBased: targetDate.getMonth()
    };
  }

  const monthKeys = Object.keys(monthMap).sort().reverse().slice(0, 12);
  const months = [];
  const dailyByMonth = {};

  for (let i = 0; i < monthKeys.length; i++) {
    const key = monthKeys[i];
    const item = monthMap[key];
    const lastDay =
      key === currentKey
        ? targetDate.getDate()
        : new Date(item.year, item.monthZeroBased + 1, 0).getDate();

    months.push({
      value: key,
      label: `${item.year}年${item.monthZeroBased + 1}月`
    });
    dailyByMonth[key] = buildDailyChartDataForMonth(
      userId,
      logData,
      item.year,
      item.monthZeroBased,
      lastDay
    );
  }

  return {
    currentMonth: currentKey,
    months: months,
    dailyByMonth: dailyByMonth
  };
}

function getStudentsNeedingComment() {
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
  }
}

// ----------------------------------------
// 学習状況・アクション記録関連
// ----------------------------------------
function getStudentCurrentStatus(userId, statusSheet) {
  if (!statusSheet) return null;
  const normalizedUserId = normalizeUserId(userId);
  const data = statusSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeUserId(data[i][0]) === normalizedUserId) {
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
  const normalizedUserId = normalizeUserId(userId);
  const normalizedStudentName = normalizeStudentName(studentName);
  const status = getStudentCurrentStatus(userId, statusSheet);
  // ★★★ 先頭にシングルクォートを追加して、強制的に文字列として書き込む ★★★
  const startTimeStr = "'" + Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  if (status) {
    // C列とD列を一度に更新
    statusSheet.getRange(status.rowIndex + 1, 3, 1, 2).setValues([[true, startTimeStr]]);
  } else {
    statusSheet.appendRow([normalizedUserId, normalizedStudentName, true, startTimeStr]);
  }
  Logger.log(`学習状況を更新(開始): UserID=${normalizedUserId}, Name=${normalizedStudentName}, StartTime=${startTimeStr}`);
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
  const endTime = options.endTime || new Date();
  const normalizedUserId = normalizeUserId(userId);
  const normalizedStudentName = normalizeStudentName(studentName);
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
    const startTimestamp = new Date();
    setStudentLearningStatus(normalizedUserId, normalizedStudentName, startTimestamp, statusSheet);
    // ★★★ こちらの書き込みにもシングルクォートを追加 ★★★
    logSheet.appendRow(["'" + Utilities.formatDate(startTimestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), normalizedUserId, normalizedStudentName, action, dateStr]);
    return {
      success: true,
      message: `${normalizedStudentName}さんの学習を開始しました。(${Utilities.formatDate(startTimestamp, Session.getScriptTimeZone(), 'HH:mm')})`,
      startTime: Utilities.formatDate(startTimestamp, Session.getScriptTimeZone(), 'HH:mm')
    };

  } else if (action === '終了') {
    if (!currentStatus || !currentStatus.isLearning || !currentStatus.startTime) {
      return { success: false, message: 'まだ学習を開始していません。' };
    }
    const startTime = currentStatus.startTime;
    let sessionMinutes = 0;
    if (startTime) { // startTimeがnullでないことを確認
      sessionMinutes = Math.round((endTime.getTime() - startTime.getTime()) / (1000 * 60));
    }
    
    clearStudentLearningStatus(normalizedUserId, statusSheet);
    logSheet.appendRow([timestamp, normalizedUserId, normalizedStudentName, action, dateStr]);
    const summaryResult = updateStudentSummaryAfterSession(normalizedUserId, normalizedStudentName, dateStr, sessionMinutes, summarySheet, logSheet);
    
    return {
      success: true,
      message: `${normalizedStudentName}さんの学習を終了しました。今回の学習時間: ${sessionMinutes}分`,
      sessionMinutes: sessionMinutes,
      monthlyTotal: summaryResult.monthlyTotal
    };
  }
  return { success: false, message: '不明な操作です。' };
}

function updateStudentSummaryAfterSession(userId, studentName, currentActionDateStr, sessionMinutes, summarySheet, attendanceLogSheet) {
  const normalizedUserId = normalizeUserId(userId);
  const summaryData = summarySheet.getDataRange().getValues();
  let summaryRowIndex = -1;
  for (let i = 1; i < summaryData.length; i++) {
    if (normalizeUserId(summaryData[i][0]) === normalizedUserId) {
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
    
    if (normalizeUserId(logData[i][1]) === normalizedUserId && logActionDateStr === currentActionDateStr.trim()) {
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
    summarySheet.appendRow([normalizedUserId, studentName, currentActionDateStr, todayFirstStartTimeStr, todayLastEndTimeStr, currentDaily, currentMonthly, currentOverall]);
  }
  return { success: true, monthlyTotal: currentMonthly, overallTotal: currentOverall, dailyTotal: currentDaily };
}


function getDailyLogs(userId) {
  const normalizedUserId = normalizeUserId(userId);
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("入退室記録");
  const data = sheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < data.length; i++) {
    if (normalizeUserId(data[i][1]) === normalizedUserId) {
      const timestamp = getValidDate(data[i][0]);
      if (!timestamp) continue;
      records.push({
        timestamp: timestamp,
        action: data[i][3] ? data[i][3].toString().trim() : ""
      });
    }
  }

  records.sort((a, b) => a.timestamp - b.timestamp);

  const results = [];
  let currentStart = null;

  for (let i = 0; i < records.length; i++) {
    const record = records[i];

    if (record.action === "開始") {
      currentStart = record.timestamp;
      continue;
    }

    if (record.action === "終了") {
      if (currentStart) {
        const startDate = Utilities.formatDate(currentStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const endDate = Utilities.formatDate(record.timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const sameDay = startDate === endDate;
        const duration = sameDay
          ? Math.max(0, Math.round((record.timestamp - currentStart) / (1000 * 60)))
          : 0;

        results.push({
          date: startDate,
          start: Utilities.formatDate(currentStart, Session.getScriptTimeZone(), "HH:mm"),
          end: Utilities.formatDate(record.timestamp, Session.getScriptTimeZone(), "HH:mm"),
          duration: duration,
          sortTimestamp: record.timestamp.getTime()
        });
        currentStart = null;
      } else {
        results.push({
          date: Utilities.formatDate(record.timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          start: "",
          end: Utilities.formatDate(record.timestamp, Session.getScriptTimeZone(), "HH:mm"),
          duration: 0,
          sortTimestamp: record.timestamp.getTime()
        });
      }
    }
  }

  if (currentStart) {
    results.push({
      date: Utilities.formatDate(currentStart, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      start: Utilities.formatDate(currentStart, Session.getScriptTimeZone(), "HH:mm"),
      end: "",
      duration: 0,
      sortTimestamp: currentStart.getTime()
    });
  }

  results.sort((a, b) => b.sortTimestamp - a.sortTimestamp);

  return {
    success: true,
    data: results.map(({ sortTimestamp, ...rest }) => rest)
  };
}


function updateStudyLog(userId, date, newStart, newEnd) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName("入退室記録");
  const summarySheet = ss.getSheetByName("学習時間サマリー");
  if (!logSheet || !summarySheet) {
    return { success: false, message: "必要なシートが見つかりません。" };
  }

  const data = logSheet.getDataRange().getValues();
  let updated = false;

  // ---- ① 入退室記録シートの更新 ----
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] && row[1].toString().trim() === userId) {
      const logDate = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const action = row[3];

      if (logDate === date && action === "開始" && newStart) {
        const d = new Date(`${date}T${newStart}:00`);
        logSheet.getRange(i + 1, 1).setValue(d);
        updated = true;
      }
      if (logDate === date && action === "終了" && newEnd) {
        const d = new Date(`${date}T${newEnd}:00`);
        logSheet.getRange(i + 1, 1).setValue(d);
        updated = true;
      }
    }
  }

  // ---- ② 対応する日の再集計 ----
  const logData = logSheet.getDataRange().getValues();
  const dayRecords = [];
  for (let i = 1; i < logData.length; i++) {
    if (logData[i][1] && logData[i][1].toString().trim() === userId) {
      const d = Utilities.formatDate(new Date(logData[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (d === date) {
        const time = new Date(logData[i][0]);
        dayRecords.push({ action: logData[i][3], time });
      }
    }
  }

  if (dayRecords.length === 0) {
    return { success: false, message: "該当日のデータが見つかりません。" };
  }

  const startRec = dayRecords.find(r => r.action === "開始");
  const endRec = dayRecords.reverse().find(r => r.action === "終了");

  let newStartTime = startRec ? startRec.time : null;
  let newEndTime = endRec ? endRec.time : null;
  let durationMin = 0;
  if (newStartTime && newEndTime) {
    durationMin = Math.round((newEndTime - newStartTime) / (1000 * 60));
  }

  // ---- ③ 学習時間サマリーシートの更新 ----
  const summaryData = summarySheet.getDataRange().getValues();
  let summaryUpdated = false;
  for (let i = 1; i < summaryData.length; i++) {
    const row = summaryData[i];
    const rowUser = row[0] ? row[0].toString().trim() : "";
    const rowDate = row[2] ? row[2].toString().trim() : "";
    if (rowUser === userId && rowDate === date) {
      summarySheet.getRange(i + 1, 4).setValue(newStartTime ? Utilities.formatDate(newStartTime, Session.getScriptTimeZone(), "HH:mm:ss") : "");
      summarySheet.getRange(i + 1, 5).setValue(newEndTime ? Utilities.formatDate(newEndTime, Session.getScriptTimeZone(), "HH:mm:ss") : "");
      summarySheet.getRange(i + 1, 6).setValue(durationMin);
      summaryUpdated = true;
      break;
    }
  }

  if (!summaryUpdated) {
    // サマリーに行がない場合は追加
    const studentName = getStudentNameById(userId);
    summarySheet.appendRow([
      userId,
      studentName,
      date,
      newStartTime ? Utilities.formatDate(newStartTime, Session.getScriptTimeZone(), "HH:mm:ss") : "",
      newEndTime ? Utilities.formatDate(newEndTime, Session.getScriptTimeZone(), "HH:mm:ss") : "",
      durationMin,
      durationMin,
      durationMin
    ]);
  }

  return {
    success: true,
    message: "修正を保存し、学習時間サマリーを更新しました。"
  };
}



// ----------------------------------------
// 管理者向け・補助関数
// ----------------------------------------
function getRealTimeStatus() {
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
  }
}

function forceEndStudy(userId) {
  try {
    const studentName = getStudentNameById(userId);
    if (!studentName) throw new Error(`生徒IDが見つかりません: ${userId}`);
    return recordAction(userId, studentName, '終了');
  } catch (error) {
    Logger.log(`Error in forceEndStudy for ${userId}: ` + error.toString());
    return { success: false, message: error.toString() };
  }
}

function autoEndOverdueStudiesAt2230() {
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
}

function getAllStudents() {
  try {
    const masterSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_STUDENT_MASTER);
    if (!masterSheet) throw new Error("「生徒マスタ」シートが見つかりません。");
    const data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 2).getValues();
    return { success: true, data: data.map(row => ({ id: row[0], name: row[1] })) };
  } catch (error) {
    Logger.log('Error in getAllStudents: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function getAllStudentsMonthlyOverview() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const masterSheet = ss.getSheetByName(SHEET_STUDENT_MASTER);
    const summarySheet = ss.getSheetByName(SHEET_SUMMARY);
    const goalSheet = ss.getSheetByName(SHEET_GOAL);

    if (!masterSheet || !summarySheet || !goalSheet) {
      throw new Error('必要なシートが見つかりません。');
    }

    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;

    const lastMonthDate = new Date(today.getFullYear(), today.getMonth(), 0);
    const lastMonthYear = lastMonthDate.getFullYear();
    const lastMonth = lastMonthDate.getMonth() + 1;

    const masterData = masterSheet.getDataRange().getValues();
    const summaryData = summarySheet.getDataRange().getValues();
    const goalData = goalSheet.getDataRange().getValues();

    const students = [];

    for (let i = 1; i < masterData.length; i++) {
      const userId = masterData[i][0] ? String(masterData[i][0]).trim() : '';
      const studentName = masterData[i][1] ? String(masterData[i][1]).trim() : '';
      if (!userId) continue;

      let currentMonthStudyTime = 0;
      let lastMonthStudyTime = 0;
      let currentGoal = '';
      let lastReflection = '';
      let lastCoachComment = '';

      // 学習時間サマリーから今月・先月を取得
      for (let j = 1; j < summaryData.length; j++) {
        const rowUserId = summaryData[j][0] ? String(summaryData[j][0]).trim() : '';
        if (rowUserId !== userId) continue;

        const rowDate = getValidDate(summaryData[j][2]);
        if (!rowDate) continue;

        const rowYear = rowDate.getFullYear();
        const rowMonth = rowDate.getMonth() + 1;
        const monthlyTotal = parseFloat(summaryData[j][6]) || 0;

        if (rowYear === currentYear && rowMonth === currentMonth) {
          currentMonthStudyTime = monthlyTotal;
        } else if (rowYear === lastMonthYear && rowMonth === lastMonth) {
          lastMonthStudyTime = monthlyTotal;
        }
      }

      // 目標管理から今月の目標・先月の振り返りを取得
      for (let j = goalData.length - 1; j >= 1; j--) {
        const rowUserId = goalData[j][0] ? String(goalData[j][0]).trim() : '';
        if (rowUserId !== userId) continue;

        const year = parseInt(goalData[j][2], 10);
        const month = parseInt(goalData[j][3], 10);

        if (year === currentYear && month === currentMonth && !currentGoal) {
          currentGoal = goalData[j][4] ? String(goalData[j][4]) : '';
        }

        if (year === lastMonthYear && month === lastMonth) {
          if (!lastReflection) {
            lastReflection = goalData[j][5] ? String(goalData[j][5]) : '';
          }
          if (!lastCoachComment) {
            lastCoachComment = goalData[j][6] ? String(goalData[j][6]) : '';
          }
        }

        if (currentGoal !== '' && lastReflection !== '') {
          break;
        }
      }

      students.push({
        userId: userId,
        studentName: studentName,
        currentMonthStudyTime: currentMonthStudyTime,
        lastMonthStudyTime: lastMonthStudyTime,
        currentGoal: currentGoal,
        lastReflection: lastReflection,
        coachCommentStatus: lastCoachComment ? '済' : '未'
      });
    }

    students.sort(function(a, b) {
      return b.currentMonthStudyTime - a.currentMonthStudyTime;
    });

    return {
      success: true,
      data: students,
      currentYear: currentYear,
      currentMonth: currentMonth,
      lastMonthYear: lastMonthYear,
      lastMonth: lastMonth
    };
  } catch (e) {
    Logger.log('Error in getAllStudentsMonthlyOverview: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getStudentData(userId) {
  try {
    const normalizedUserId = normalizeUserId(userId);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summarySheet = spreadsheet.getSheetByName(SHEET_SUMMARY);
    const monthlySheet = spreadsheet.getSheetByName(SHEET_MONTHLY_SUMMARY);
    const logSheet = spreadsheet.getSheetByName(SHEET_ATTENDANCE_LOG);
    const goalSheet = spreadsheet.getSheetByName(SHEET_GOAL);

    // -----------------------------
    // 基本情報（サマリー）
    // -----------------------------
    let summary = { isDataFound: false, monthlyTotal: 0 };
    let currentMonthTotal = 0;
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = summaryData.length - 1; i >= 1; i--) {
      const rowUserId = summaryData[i][0] ? summaryData[i][0].toString().trim() : '';
      if (normalizeUserId(rowUserId) === normalizedUserId) {
        const lastActivityDate = getValidDate(summaryData[i][2]);
        const currentYearMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

        let summaryYearMonth = '';
        if (lastActivityDate) {
          summaryYearMonth = Utilities.formatDate(lastActivityDate, Session.getScriptTimeZone(), 'yyyy-MM');
        } else if (summaryData[i][2]) {
          summaryYearMonth = summaryData[i][2].toString().substring(0, 7).trim();
        }

        const monthlyTotal = (summaryYearMonth === currentYearMonth)
          ? (parseFloat(summaryData[i][6]) || 0)
          : 0;

        summary = {
          isDataFound: true,
          lastActivityDate: lastActivityDate
            ? Utilities.formatDate(lastActivityDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : '記録なし',
          monthlyTotal: monthlyTotal,
          overallTotal: summaryData[i][7] || 0
        };
        currentMonthTotal = monthlyTotal;
        if (summaryYearMonth === currentYearMonth) {
          break;
        }
      }
    }

    // -----------------------------
    // 今月の日別グラフデータ
    // -----------------------------
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;
    const dailyChartBundle = buildDailyChartBundle(userId, logSheet.getDataRange().getValues(), now);

    // -----------------------------
    // 入退室ログ
    // -----------------------------
    const logs = [];
    const logData = logSheet.getDataRange().getValues();
    for (let i = logData.length - 1; i >= 1; i--) {
      if (normalizeUserId(logData[i][1]) === normalizedUserId) {
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

    // -----------------------------
    // 目標履歴（過去1年）
    // -----------------------------
    const goalHistory = [];
    if (goalSheet) {
      const allGoalData = goalSheet.getDataRange().getValues();
      const oneYearAgo = currentYear - 1;
      for (let i = 1; i < allGoalData.length; i++) {
        const row = allGoalData[i];
        const uid = row[0] ? row[0].toString().trim() : '';
        const year = parseInt(row[2], 10);
        const month = parseInt(row[3], 10);
        if (normalizeUserId(uid) === normalizedUserId && year >= oneYearAgo) {
          goalHistory.push({
            year,
            month,
            goal: row[4] || '',
            reflection: row[5] || '',
            comment: row[6] || ''
          });
        }
      }
      goalHistory.sort((a, b) => (a.year === b.year ? b.month - a.month : b.year - a.year));
    }

    // -----------------------------
    // 最終返却
    // -----------------------------
    return {
      success: true,
      data: {
        summary,
        dailyChartBundle: dailyChartBundle,
        logs,
        goalHistory
      }
    };
  } catch (error) {
    Logger.log(`Error in getStudentData for ${userId}: ${error.toString()}`);
    return { success: false, message: error.toString() };
  }
}

function getStudentNameById(userId) {
  try {
    const normalizedUserId = normalizeUserId(userId);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_STUDENT_MASTER);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (normalizeUserId(data[i][0]) === normalizedUserId) {
        return data[i][1].toString();
      }
    }
    return null;
  } catch(e) {
    Logger.log(`Error in getStudentNameById: ${e}`);
    return null;
  }
}

// ヘルパー: 概要シートの値配列から指定ユーザーの今月順位を計算する（単独でテスト可能）
function computeStudentRankFromSummaryData(values, targetUserId) {
  if (!values || values.length <= 1) return { rank: null, total: 0 };

  const currentYearMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const list = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const uid = row[0] ? String(row[0]).trim() : '';
    if (!uid) continue;

    let rowYearMonth = '';
    const rowDate = getValidDate(row[2]);
    if (rowDate) {
      rowYearMonth = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM');
    } else if (row[2]) {
      rowYearMonth = String(row[2]).substring(0, 7).trim();
    }

    let monthlyNum = 0;
    if (rowYearMonth === currentYearMonth) {
      const monthlyVal = row[6];
      if (typeof monthlyVal === 'number') {
        monthlyNum = monthlyVal;
      } else if (monthlyVal) {
        const digits = String(monthlyVal).replace(/[^0-9-]/g, '');
        monthlyNum = digits ? parseInt(digits, 10) || 0 : 0;
      }
    }

    list.push({ userId: uid, monthly: monthlyNum });
  }

  list.sort(function(a, b) { return b.monthly - a.monthly; });

  let rank = null;
  let dense = 0;
  let prev = null;
  for (let i = 0; i < list.length; i++) {
    if (prev === null || list[i].monthly !== prev) {
      dense++;
      prev = list[i].monthly;
    }
    if (String(list[i].userId) === String(targetUserId)) {
      rank = dense;
    }
  }

  return { rank: rank, total: list.length };
}

//自習開始ボタン関数
function getCurrentLearningStatus(userId) {
  try {
    const statusSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CURRENT_STATUS);
    if (!statusSheet) {
      return { success: false, message: '学習状況シートが見つかりません。' };
    }

    const status = getStudentCurrentStatus(userId, statusSheet);

    if (status && status.isLearning && status.startTime) {
      return {
        success: true,
        isLearning: true,
        startTime: Utilities.formatDate(
          status.startTime,
          Session.getScriptTimeZone(),
          'HH:mm'
        )
      };
    }

    return {
      success: true,
      isLearning: false,
      startTime: ''
    };
  } catch (e) {
    Logger.log('Error in getCurrentLearningStatus: ' + e);
    return { success: false, message: e.toString() };
  }
}

// テスト用関数（Apps Script エディタから実行できます）
function testComputeStudentRank() {
  const sample = [
    ['userId','a','b','c','d','e', 'monthly'],
    ['u1', '', '', '', '', '', 120],
    ['u2', '', '', '', '', '', 200],
    ['u3', '', '', '', '', '', 120],
    ['u4', '', '', '', '', '', '50']
  ];
  Logger.log('Sample ranking for u1: ' + JSON.stringify(computeStudentRankFromSummaryData(sample, 'u1')));
  Logger.log('Sample ranking for u2: ' + JSON.stringify(computeStudentRankFromSummaryData(sample, 'u2')));
  Logger.log('Sample ranking for u4: ' + JSON.stringify(computeStudentRankFromSummaryData(sample, 'u4')));
}
