const CONFIG = {
  STATUS_SHEET: 'ステータス',
  ROSTER_SHEET: '名簿',
  DEFAULT_ADMIN_PASSWORD: '0000',
  EQUIPMENT: [
    ['A1mini-1', 'A1 mini 1号機'],
    ['A1mini-2', 'A1 mini 2号機'],
    ['A1mini-3', 'A1 mini 3号機'],
    ['A1mini-4', 'A1 mini 4号機'],
    ['P2S-1', 'P2S 1号機'],
    ['LASER-1', 'レーザーカッター 1号機'],
    ['LASER-2', 'レーザーカッター 2号機']
  ]
};

// 画面表示
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ファブラボ機材管理');
}

// 初期設定
function initializeDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('スプレッドシートからApps Scriptを開いてください。');
  }

  let status = ss.getSheetByName(CONFIG.STATUS_SHEET);
  if (!status) status = ss.insertSheet(CONFIG.STATUS_SHEET);

  status.getRange(1, 1, 1, 9).setValues([[
    '機材ID',
    '機材名',
    '使用中フラグ',
    '使用者名',
    '使用者学籍番号',
    '利用開始日時',
    '利用区分',
    '利用予定時間（分）',
    '予定終了日時'
  ]]);

  const existingIds = status.getLastRow() >= 2
    ? status.getRange(2, 1, status.getLastRow() - 1, 1)
        .getValues().flat().map(String)
    : [];

  CONFIG.EQUIPMENT.forEach(item => {
    if (!existingIds.includes(item[0])) {
      status.appendRow([item[0], item[1], false, '', '', '', '', '', '']);
    }
  });

  let roster = ss.getSheetByName(CONFIG.ROSTER_SHEET);
  if (!roster) roster = ss.insertSheet(CONFIG.ROSTER_SHEET);

  if (roster.getLastRow() === 0) {
    roster.getRange(1, 1, 1, 2).setValues([['学籍番号', '氏名']]);
  }

  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty('ADMIN_PASSWORD')) {
    props.setProperty('ADMIN_PASSWORD', CONFIG.DEFAULT_ADMIN_PASSWORD);
  }

  SpreadsheetApp.flush();
  return '初期設定が完了しました。';
}

// 管理者パスワード確認
function verifyAdmin(password) {
  const saved = PropertiesService.getScriptProperties()
    .getProperty('ADMIN_PASSWORD') || CONFIG.DEFAULT_ADMIN_PASSWORD;

  return String(password) === saved;
}

// 管理者パスワード変更
function setAdminPassword(newPassword, currentPassword) {
  if (!verifyAdmin(currentPassword)) {
    throw new Error('現在のパスワードが違います。');
  }

  if (!newPassword || String(newPassword).length < 4) {
    throw new Error('パスワードは4文字以上にしてください。');
  }

  PropertiesService.getScriptProperties()
    .setProperty('ADMIN_PASSWORD', String(newPassword));

  return 'パスワードを変更しました。';
}

// 学籍番号確認
function validateStudentId(studentId) {
  return /^B\d{7}$/.test(String(studentId).trim().toUpperCase());
}

// 氏名取得
function getStudentName(studentId) {
  studentId = String(studentId).trim().toUpperCase();

  if (!validateStudentId(studentId)) {
    throw new Error('学籍番号が正しくありません。');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.ROSTER_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return '未登録ユーザー(' + studentId + ')';
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

  for (const row of data) {
    if (String(row[0]).trim().toUpperCase() === studentId) {
      return String(row[1]);
    }
  }

  return '未登録ユーザー(' + studentId + ')';
}

// 機材状況取得
function getEquipmentStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.STATUS_SHEET);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  return data.filter(row => row[0]).map(row => ({
    id: String(row[0]),
    name: String(row[1]),
    inUse: row[2] === true || String(row[2]).toUpperCase() === 'TRUE',
    userName: String(row[3] || ''),
    studentId: String(row[4] || ''),
    startTime: row[5] instanceof Date ? row[5].getTime() :
      (row[5] ? new Date(row[5]).getTime() : null),
    group: String(row[6] || ''),
    plannedMinutes: Number(row[7]) || 0,
    plannedEndTime: row[8] instanceof Date ? row[8].getTime() :
      (row[8] ? new Date(row[8]).getTime() : null)
  }));
}

// 機材の利用開始・本人による返却
function updateEquipment(equipmentId, studentId, mode, group, plannedHours) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    equipmentId = String(equipmentId).trim();
    studentId = String(studentId).trim().toUpperCase();

    if (!validateStudentId(studentId)) {
      throw new Error('学生証の番号が正しくありません。');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.STATUS_SHEET);

    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('機材データがありません。');
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    const index = data.findIndex(row => String(row[0]) === equipmentId);

    if (index === -1) throw new Error('機材が見つかりません。');

    const rowNumber = index + 2;
    const row = data[index];
    const inUse = row[2] === true || String(row[2]).toUpperCase() === 'TRUE';

    if (mode === 'checkout') {
      if (inUse) throw new Error('この機材はすでに使用中です。');

      group = String(group || '').trim();
      plannedHours = Number(plannedHours);

      if (!group) {
        throw new Error('利用区分を選択してください。');
      }

      if (!Number.isFinite(plannedHours) ||
          plannedHours < 0.25 || plannedHours > 24) {
        throw new Error('利用予定時間は0.25～24時間で入力してください。');
      }

      const studentName = getStudentName(studentId);
      const now = new Date();
      const plannedMinutes = Math.round(plannedHours * 60);
      const endTime = new Date(now.getTime() + plannedMinutes * 60000);

      sheet.getRange(rowNumber, 3, 1, 7).setValues([[
        true,
        studentName,
        studentId,
        now,
        group,
        plannedMinutes,
        endTime
      ]]);

      return studentName + ' さんの利用を開始しました。';

    } else if (mode === 'return') {
      if (!inUse) throw new Error('この機材は現在使用されていません。');

      if (String(row[4]).trim().toUpperCase() !== studentId) {
        throw new Error('この機材を借りた本人の学生証ではありません。');
      }

      sheet.getRange(rowNumber, 3, 1, 7).setValues([[
        false, '', '', '', '', '', ''
      ]]);

      return '返却が完了しました。';
    }

    throw new Error('操作が正しくありません。');

  } finally {
    lock.releaseLock();
  }
}

// ボタン1つで返却する機能
function quickReturnEquipment(equipmentId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    equipmentId = String(equipmentId).trim();

    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.STATUS_SHEET);

    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('機材データがありません。');
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    const index = data.findIndex(row => String(row[0]) === equipmentId);

    if (index === -1) throw new Error('機材が見つかりません。');

    const rowNumber = index + 2;
    const row = data[index];
    const inUse = row[2] === true || String(row[2]).toUpperCase() === 'TRUE';

    if (!inUse) {
      throw new Error('この機材はすでに返却されています。');
    }

    const equipmentName = String(row[1]);

    sheet.getRange(rowNumber, 3, 1, 7).setValues([[
      false, '', '', '', '', '', ''
    ]]);

    return equipmentName + ' の返却が完了しました。';

  } finally {
    lock.releaseLock();
  }
}

// 管理者による強制返却
function adminRelease(equipmentId, password) {
  if (!verifyAdmin(password)) {
    throw new Error('管理者パスワードが違います。');
  }

  return quickReturnEquipment(equipmentId);
}

// 名簿メンバー追加（1件ずつ）
function addRosterMember(studentId, studentName, password) {
  if (!verifyAdmin(password)) {
    throw new Error('管理者パスワードが違います。');
  }

  studentId = String(studentId).trim().toUpperCase();
  studentName = String(studentName).trim();

  if (!validateStudentId(studentId)) {
    throw new Error('学籍番号はB＋数字7桁で入力してください。');
  }

  if (!studentName) throw new Error('氏名を入力してください。');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.ROSTER_SHEET);

    if (!sheet) throw new Error('名簿シートがありません。');

    const data = sheet.getLastRow() >= 2
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
      : [];

    if (data.some(row => String(row[0]).trim().toUpperCase() === studentId)) {
      throw new Error('この学籍番号はすでに登録されています。');
    }

    sheet.appendRow([studentId, studentName]);
    return '名簿に追加しました。';

  } finally {
    lock.releaseLock();
  }
}

// 名簿メンバー一括追加（CSV用）
function bulkAddRosterMembers(members, password) {
  if (!verifyAdmin(password)) {
    throw new Error('管理者パスワードが違います。');
  }

  if (!Array.isArray(members) || members.length === 0) {
    throw new Error('登録するデータが見つかりません。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.ROSTER_SHEET);

    if (!sheet) throw new Error('名簿シートがありません。');

    const lastRow = sheet.getLastRow();
    const existingData = lastRow >= 2
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      : [];

    const existingIds = new Set(
      existingData.map(row => String(row[0]).trim().toUpperCase()).filter(Boolean)
    );

    const newRows = [];
    let skippedDuplicate = 0;
    let skippedInvalid = 0;

    members.forEach(item => {
      const id = String(item.id || '').trim().toUpperCase();
      const name = String(item.name || '').trim();

      if (!validateStudentId(id) || !name) {
        skippedInvalid++;
        return;
      }

      if (existingIds.has(id)) {
        skippedDuplicate++;
        return;
      }

      existingIds.add(id);
      newRows.push([id, name]);
    });

    if (newRows.length > 0) {
      sheet.getRange(lastRow + 1, 1, newRows.length, 2).setValues(newRows);
    }

    let message = newRows.length + ' 件を名簿に追加しました。';
    if (skippedDuplicate > 0 || skippedInvalid > 0) {
      message += '\n（登録済み重複：' + skippedDuplicate + '件 / 形式エラー：' + skippedInvalid + '件をスキップ）';
    }

    return message;

  } finally {
    lock.releaseLock();
  }
}

// 機材追加
function addEquipment(equipmentId, equipmentName, password) {
  if (!verifyAdmin(password)) {
    throw new Error('管理者パスワードが違います。');
  }

  equipmentId = String(equipmentId).trim();
  equipmentName = String(equipmentName).trim();

  if (!equipmentId || !equipmentName) {
    throw new Error('機材IDと機材名を入力してください。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.STATUS_SHEET);

    if (!sheet) throw new Error('ステータスシートがありません。');

    const data = sheet.getLastRow() >= 2
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      : [];

    if (data.some(row => String(row[0]).trim() === equipmentId)) {
      throw new Error('この機材IDはすでに登録されています。');
    }

    sheet.appendRow([equipmentId, equipmentName, false, '', '', '', '', '', '']);
    return '機材を追加しました。';

  } finally {
    lock.releaseLock();
  }
}

// 機材名変更
function renameEquipment(equipmentId, newName, password) {
  if (!verifyAdmin(password)) {
    throw new Error('管理者パスワードが違います。');
  }

  equipmentId = String(equipmentId).trim();
  newName = String(newName).trim();

  if (!newName) throw new Error('新しい機材名を入力してください。');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.STATUS_SHEET);

    if (!sheet) throw new Error('ステータスシートがありません。');

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === equipmentId) {
        sheet.getRange(i + 1, 2).setValue(newName);
        return '機材名を変更しました。';
      }
    }

    throw new Error('機材が見つかりません。');

  } finally {
    lock.releaseLock();
  }
}
