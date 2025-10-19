/**
 * Shared Drive重複ファイル検出・削除スクリプト v1
 *
 * 機能：
 * 1. 指定したShared Drive内の全ファイルをスキャン
 * 2. 重複ファイルを検出（ファイル名 or MD5ハッシュ）
 * 3. 重複リストをスプレッドシートに出力
 * 4. 重複ファイルを削除（オプション）
 */

// ============================================================
// 設定
// ============================================================
const CONFIG = {
  // Shared DriveのID（URLから取得: https://drive.google.com/drive/folders/{SHARED_DRIVE_ID}）
  SHARED_DRIVE_ID: '【ここにShared DriveのIDを入力】',

  // 重複判定基準
  // 'name': ファイル名で判定
  // 'md5': MD5ハッシュで判定（同じ内容のファイル）
  // 'both': 両方で判定
  DUPLICATE_CRITERIA: 'name',

  // 出力先スプレッドシート名
  OUTPUT_SHEET_NAME: '重複ファイルリスト',

  // 削除時に残すファイルの優先順位
  // 'oldest': 最も古いファイルを残す
  // 'newest': 最も新しいファイルを残す
  KEEP_PRIORITY: 'newest'
};

// ============================================================
// メイン関数：メニューを追加
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('重複ファイル管理')
    .addItem('1. 重複ファイルをスキャン', 'scanDuplicates')
    .addItem('2. 重複ファイルを削除（要確認）', 'deleteDuplicates')
    .addItem('3. 初期設定シート作成', 'createConfigSheet')
    .addToUi();
}

// ============================================================
// 初期設定シートの作成
// ============================================================
function createConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('設定');

  if (!configSheet) {
    configSheet = ss.insertSheet('設定');
  }
  configSheet.clear();

  // ヘッダー
  configSheet.getRange('A1:B1').setValues([['設定項目', '値']]);
  configSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

  // 設定値
  const configs = [
    ['Shared Drive ID', CONFIG.SHARED_DRIVE_ID],
    ['重複判定基準（name/md5/both）', CONFIG.DUPLICATE_CRITERIA],
    ['残すファイルの優先順位（oldest/newest）', CONFIG.KEEP_PRIORITY],
    ['', ''],
    ['【使い方】', ''],
    ['1. Shared Drive IDを入力', ''],
    ['2. メニュー「重複ファイル管理」→「重複ファイルをスキャン」を実行', ''],
    ['3. 結果を確認後、削除する場合は「重複ファイルを削除」を実行', '']
  ];

  configSheet.getRange(2, 1, configs.length, 2).setValues(configs);
  configSheet.setColumnWidth(1, 350);
  configSheet.setColumnWidth(2, 300);

  SpreadsheetApp.getUi().alert('設定シートを作成しました。\nShared Drive IDを入力してください。');
}

// ============================================================
// 重複ファイルのスキャン
// ============================================================
function scanDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('設定');

  // 設定の読み込み
  let driveId = CONFIG.SHARED_DRIVE_ID;
  let criteria = CONFIG.DUPLICATE_CRITERIA;

  if (configSheet) {
    driveId = configSheet.getRange('B2').getValue() || driveId;
    criteria = configSheet.getRange('B3').getValue() || criteria;
  }

  if (driveId === '【ここにShared DriveのIDを入力】' || !driveId) {
    SpreadsheetApp.getUi().alert('エラー', 'Shared Drive IDを設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getUi().alert('スキャンを開始します。\n大量のファイルがある場合、時間がかかる場合があります。');

  const startTime = new Date();
  Logger.log('スキャン開始: ' + startTime);

  try {
    // Shared Driveのルートフォルダを取得
    const rootFolder = DriveApp.getFolderById(driveId);

    // 全ファイルを収集
    const allFiles = [];
    collectFiles(rootFolder, allFiles, '');

    Logger.log('総ファイル数: ' + allFiles.length);

    // 重複を検出
    const duplicates = findDuplicates(allFiles, criteria);

    Logger.log('重複グループ数: ' + duplicates.length);

    // 結果をスプレッドシートに出力
    outputDuplicates(duplicates);

    const endTime = new Date();
    const elapsedSeconds = Math.round((endTime - startTime) / 1000);

    SpreadsheetApp.getUi().alert(
      'スキャン完了',
      `総ファイル数: ${allFiles.length}\n` +
      `重複グループ数: ${duplicates.length}\n` +
      `処理時間: ${elapsedSeconds}秒\n\n` +
      `結果を「${CONFIG.OUTPUT_SHEET_NAME}」シートに出力しました。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラー', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================================
// ファイルを再帰的に収集
// ============================================================
function collectFiles(folder, fileList, path) {
  const currentPath = path + '/' + folder.getName();

  // フォルダ内のファイルを収集
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    fileList.push({
      id: file.getId(),
      name: file.getName(),
      md5: file.getMd5Checksum(),
      size: file.getSize(),
      mimeType: file.getMimeType(),
      createdDate: file.getDateCreated(),
      modifiedDate: file.getLastUpdated(),
      owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
      path: currentPath,
      url: file.getUrl()
    });
  }

  // サブフォルダを再帰的に処理
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    collectFiles(subFolder, fileList, currentPath);
  }
}

// ============================================================
// 重複ファイルを検出
// ============================================================
function findDuplicates(files, criteria) {
  const duplicateGroups = [];
  const fileMap = new Map();

  // ファイルをキーでグループ化
  for (const file of files) {
    let key;

    if (criteria === 'name') {
      key = file.name;
    } else if (criteria === 'md5') {
      key = file.md5 || 'no-md5-' + file.id; // MD5がない場合（Google Docs等）
    } else if (criteria === 'both') {
      key = file.name + '::' + (file.md5 || 'no-md5');
    }

    if (!fileMap.has(key)) {
      fileMap.set(key, []);
    }
    fileMap.get(key).push(file);
  }

  // 重複があるグループのみを抽出
  for (const [key, group] of fileMap) {
    if (group.length > 1) {
      duplicateGroups.push({
        key: key,
        files: group
      });
    }
  }

  return duplicateGroups;
}

// ============================================================
// 重複リストをスプレッドシートに出力
// ============================================================
function outputDuplicates(duplicates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);

  if (!outputSheet) {
    outputSheet = ss.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
  }
  outputSheet.clear();

  // ヘッダー
  const headers = [
    '重複グループ',
    'ファイル名',
    'MD5',
    'サイズ(bytes)',
    '作成日',
    '更新日',
    '所有者',
    'パス',
    'URL',
    'ファイルID',
    '削除対象'
  ];

  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  outputSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');

  let currentRow = 2;

  // 設定の読み込み
  const configSheet = ss.getSheetByName('設定');
  let keepPriority = CONFIG.KEEP_PRIORITY;
  if (configSheet) {
    keepPriority = configSheet.getRange('B4').getValue() || keepPriority;
  }

  // 重複グループごとに出力
  for (let i = 0; i < duplicates.length; i++) {
    const group = duplicates[i];
    const groupName = `グループ${i + 1}`;

    // ソート：残すファイルを決定
    const sortedFiles = [...group.files];
    sortedFiles.sort((a, b) => {
      if (keepPriority === 'newest') {
        return b.modifiedDate.getTime() - a.modifiedDate.getTime();
      } else {
        return a.modifiedDate.getTime() - b.modifiedDate.getTime();
      }
    });

    // グループ内の各ファイルを出力
    for (let j = 0; j < sortedFiles.length; j++) {
      const file = sortedFiles[j];
      const isKeep = j === 0; // 最初のファイルを残す

      const row = [
        groupName,
        file.name,
        file.md5 || 'N/A',
        file.size,
        file.createdDate,
        file.modifiedDate,
        file.owner,
        file.path,
        file.url,
        file.id,
        isKeep ? '残す' : '削除'
      ];

      outputSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);

      // 色分け
      if (isKeep) {
        outputSheet.getRange(currentRow, 1, 1, row.length).setBackground('#d9ead3'); // 緑
      } else {
        outputSheet.getRange(currentRow, 1, 1, row.length).setBackground('#f4cccc'); // 赤
      }

      currentRow++;
    }

    // グループ間の空白行
    currentRow++;
  }

  // 列幅調整
  outputSheet.setColumnWidth(1, 120);
  outputSheet.setColumnWidth(2, 250);
  outputSheet.setColumnWidth(3, 120);
  outputSheet.setColumnWidth(4, 100);
  outputSheet.setColumnWidth(5, 150);
  outputSheet.setColumnWidth(6, 150);
  outputSheet.setColumnWidth(7, 200);
  outputSheet.setColumnWidth(8, 400);
  outputSheet.setColumnWidth(9, 300);
  outputSheet.setColumnWidth(10, 250);
  outputSheet.setColumnWidth(11, 80);

  // フィルター設定
  outputSheet.getRange(1, 1, currentRow - 1, headers.length).createFilter();
}

// ============================================================
// 重複ファイルを削除
// ============================================================
function deleteDuplicates() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    '「削除」と表示されているファイルをゴミ箱に移動します。\n' +
    'この操作は元に戻せません（ゴミ箱から復元は可能）。\n\n' +
    '本当に削除しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);

  if (!outputSheet) {
    ui.alert('エラー', '重複リストが見つかりません。先にスキャンを実行してください。', ui.ButtonSet.OK);
    return;
  }

  const data = outputSheet.getDataRange().getValues();
  let deletedCount = 0;
  let errorCount = 0;

  // ヘッダーをスキップして処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const deleteFlag = row[10]; // K列：削除対象
    const fileId = row[9]; // J列：ファイルID

    if (deleteFlag === '削除' && fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        file.setTrashed(true);
        deletedCount++;
        Logger.log(`削除: ${row[1]} (ID: ${fileId})`);
      } catch (error) {
        errorCount++;
        Logger.log(`削除失敗: ${row[1]} (ID: ${fileId}) - ${error}`);
      }
    }
  }

  ui.alert(
    '削除完了',
    `削除成功: ${deletedCount}ファイル\n` +
    `削除失敗: ${errorCount}ファイル\n\n` +
    `詳細はログを確認してください。`,
    ui.ButtonSet.OK
  );
}
