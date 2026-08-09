const PARTICIPANT_SHEET_NAME = '参加者リスト'; // 参加者シート名

// スプレッドシートの列インデックス（参加者リストシート）
const ATTENDANCE_COL_INDEX = 2; // C列：チェックイン状態
const TIME_COL_INDEX = 3;       // D列：チェックイン時刻
const SUBSIDY_COL_INDEX = 4;    // E列：交通費補助の有無

function doGet(e) {
  Logger.log('doGet: Webアプリが起動しました');
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * QRリーダー端末：参加者をチェックインし、時刻を記録する
 * @param {string} registeredId - 読み取られた登録番号
 * @returns {object} 処理結果
 */
function checkInAndRecord(registeredId) {
  Logger.log(`checkInAndRecord: 処理開始 ID: ${registeredId}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantSheet = ss.getSheetByName(PARTICIPANT_SHEET_NAME);

  if (!participantSheet) {
    Logger.log(`エラー: シート「${PARTICIPANT_SHEET_NAME}」が見つかりません`);
    return {
      success: false,
      message: 'エラー: 参加者リストが見つかりません',
      name: '---',
      status: 'error'
    };
  }

  const data = participantSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(registeredId).trim()) {
      const name = data[i][1];
      const currentStatus = String(data[i][ATTENDANCE_COL_INDEX]).trim();
      const subsidyStatus = String(data[i][SUBSIDY_COL_INDEX]).trim();
      const hasSubsidy = subsidyStatus === 'あり';

      // 既にチェックイン済みの場合は再記録しない
      if (currentStatus === '出席済み') {
        Logger.log(`checkInAndRecord: 既にチェックイン済み ${name}`);
        return {
          success: true,
          message: `${name} 様は既にチェックイン済みです`,
          name: name,
          status: 'already_checked',
          hasSubsidy: hasSubsidy
        };
      }

      // C列に状態、D列に時刻のみを記録
      const ssTimeZone = ss.getSpreadsheetTimeZone();
      const timeString = Utilities.formatDate(new Date(), ssTimeZone, 'HH:mm:ss');

      participantSheet
        .getRange(i + 1, ATTENDANCE_COL_INDEX + 1)
        .setValue('出席済み');

      participantSheet
        .getRange(i + 1, TIME_COL_INDEX + 1)
        .setNumberFormat('@')
        .setValue(timeString);

      Logger.log(`checkInAndRecord: チェックイン完了 ${name} ${timeString}`);

      return {
        success: true,
        message: `${name} 様のチェックインが完了しました`,
        name: name,
        status: 'checked_in',
        hasSubsidy: hasSubsidy
      };
    }
  }

  Logger.log(`checkInAndRecord: 登録番号が見つかりません ID:${registeredId}`);
  return {
    success: false,
    message: 'エラー: 登録番号が見つかりません',
    name: '---',
    status: 'error'
  };
}

/**
 * チェックイン状況確認端末：参加者をチェックイン済み／未チェックインに分類する
 * @returns {object} チェックイン状況
 */
function getCheckInOverview() {
  Logger.log('getCheckInOverview: チェックイン状況取得開始');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantSheet = ss.getSheetByName(PARTICIPANT_SHEET_NAME);

  if (!participantSheet) {
    Logger.log(`エラー: シート「${PARTICIPANT_SHEET_NAME}」が見つかりません`);
    return { checkedIn: [], unchecked: [] };
  }

  const data = participantSheet.getDataRange().getValues();
  const checkedIn = [];
  const unchecked = [];
  const ssTimeZone = ss.getSpreadsheetTimeZone();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[0]).trim();
    const name = String(row[1]).trim();

    // 完全な空行は一覧に出さない
    if (!id && !name) continue;

    const attendanceStatus = String(row[ATTENDANCE_COL_INDEX]).trim();
    const rawCheckInTime = row[TIME_COL_INDEX];
    const subsidyStatus = String(row[SUBSIDY_COL_INDEX]).trim();
    const hasSubsidy = subsidyStatus === 'あり';

    // D列に過去のDate型データが残っていても、画面には時刻だけ出す
    const checkInTime = rawCheckInTime instanceof Date
      ? Utilities.formatDate(rawCheckInTime, ssTimeZone, 'HH:mm:ss')
      : String(rawCheckInTime || '').trim();

    const participant = {
      id: id,
      name: name,
      hasSubsidy: hasSubsidy
    };

    if (attendanceStatus === '出席済み') {
      participant.checkInTime = checkInTime;
      checkedIn.push(participant);
    } else {
      unchecked.push(participant);
    }
  }

  Logger.log(`getCheckInOverview: チェックイン済み=${checkedIn.length}, 未チェックイン=${unchecked.length}`);
  return {
    checkedIn: checkedIn,
    unchecked: unchecked
  };
}
