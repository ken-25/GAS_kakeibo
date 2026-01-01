function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('家計簿ダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 指定された年月のデータを取得・集計するAPI
 * @param {number} year 
 * @param {number} month (1-12)
 * @returns {Object} 集計結果と明細データ
 */
function getDataByMonth(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB');
  
  if (!dbSheet) {
    // シートがない場合のダミーデータまたはエラーハンドリング
    // ここではエラーを返さず空データを返してアプリが落ちないようにする
    return {
      error: 'シート「DB」が見つかりません。',
      summary: { income: 0, expense: 0, balance: 0 },
      byCategory: {},
      details: []
    };
  }
  
  // A列～G列: [タイムスタンプ, 日付, 金額, 収支, 項目, メモ, タグ]
  const lastRow = dbSheet.getLastRow();
  const lastCol = dbSheet.getLastColumn();
  
  if (lastRow <= 1) {
    return {
      year: year,
      month: month,
      summary: { income: 0, expense: 0, balance: 0 },
      byCategory: {},
      details: []
    };
  }
  
  const data = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  let incomeTotal = 0;
  let expenseTotal = 0;
  
  const categorySummary = {};
  const details = [];
  
  data.forEach(row => {
    // 日付 (B列: index 1) を優先、なければタイムスタンプ (A列: index 0)
    let dateVal = row[1] || row[0]; 
    if (!dateVal) return;
    
    let dateObj = new Date(dateVal);
    if (isNaN(dateObj.getTime())) return; // 無効な日付
    
    // 年月チェック
    // Javascriptのmonthは0-11, 引数は1-12想定
    if (dateObj.getFullYear() === year && dateObj.getMonth() === month - 1) {
      const amount = Number(row[2]) || 0;
      // D列(index 3)の「収支」より、C列(index 2)の「金額(正負)」を信頼して判定
      // ユーザー仕様: 収入は+, 支出は-
      let type = row[3]; 
      const category = row[4] || '未分類';
      const memo = row[5] || '';
      const tag = row[6] || '';
      
      // Amountの正負でタイプを補完（もし空なら）
      if (!type) {
        if (amount >= 0) type = '収入';
        else type = '支出';
      }

      // 集計
      if (amount > 0) {
         // 収入
         incomeTotal += amount;
      } else if (amount < 0) {
         // 支出（マイナス値）
         // 合計やグラフには絶対値を使用
         const absAmount = Math.abs(amount);
         expenseTotal += absAmount;
         
         // カテゴリ集計
         if (!categorySummary[category]) {
           categorySummary[category] = 0;
         }
         categorySummary[category] += absAmount;
      }
      
      details.push({
        date: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
        amount: amount, // 明細には元の値（正負）を渡す
        type: type,
        category: category,
        memo: memo,
        tag: tag
      });
    }
  });

  // 日付順にソート (新しい順)
  details.sort((a, b) => new Date(b.date) - new Date(a.date));

  return {
    year: year,
    month: month,
    summary: {
      income: incomeTotal,
      expense: expenseTotal,
      balance: incomeTotal - expenseTotal // 収入(正) - 支出(正の合計) = 差引
    },
    byCategory: categorySummary,
    details: details
  };
}

/**
 * 初期表示用：現在の年月のデータを返す
 */
function getInitialData() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  return getDataByMonth(year, month);
}

/**
 * 全期間のタグ別集計データを取得する
 */
function getTagSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB');
  
  if (!dbSheet) return { tags: [] };
  
  const lastRow = dbSheet.getLastRow();
  const lastCol = dbSheet.getLastColumn();
  
  if (lastRow <= 1) return { tags: [] };
  
  const data = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  // タグごとの集計オブジェクト
  // Key: tagName, Value: { income: 0, expense: 0, count: 0 }
  const tagStats = {};
  
  data.forEach(row => {
    // 日付またはTSがあれば有効データとみなす
    if (!row[0] && !row[1]) return;
    
    // 金額
    const amount = Number(row[2]) || 0;
    // タグ (G列, index 6)
    const tagName = row[6] ? String(row[6]) : '';
    
    // タグなしは集計対象外（ユーザー要望）
    if (!tagName) return;

    if (!tagStats[tagName]) {
      tagStats[tagName] = { income: 0, expense: 0, count: 0 };
    }
    
    const stats = tagStats[tagName];
    stats.count++;
    
    if (amount > 0) {
      stats.income += amount;
    } else if (amount < 0) {
      stats.expense += Math.abs(amount);
    }
  });
  
  // 配列に変換してソート
  // デフォルトは「支出額が多い順」
  const result = Object.keys(tagStats).map(tag => {
    const s = tagStats[tag];
    return {
      tag: tag,
      income: s.income,
      expense: s.expense,
      balance: s.income - s.expense,
      count: s.count
    };
  });
  
  result.sort((a, b) => b.expense - a.expense);
  
  return { tags: result };
}

/**
 * 指定タグの全期間明細を取得する
 */
function getTagDetails(tagName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB');
  if (!dbSheet) return { error: 'シートが見つかりません' };
  
  const lastRow = dbSheet.getLastRow();
  const lastCol = dbSheet.getLastColumn();
  if (lastRow <= 1) return { details: [] };
  
  const data = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const details = [];
  
  data.forEach(row => {
    let dateVal = row[1] || row[0]; 
    if (!dateVal) return;
    let dateObj = new Date(dateVal);
    if (isNaN(dateObj.getTime())) return;
    
    const rowTag = row[6] ? String(row[6]) : '';
    
    if (rowTag === tagName) {
       const amount = Number(row[2]) || 0;
       let type = row[3];
       if (!type) {
         type = amount >= 0 ? '収入' : '支出';
       }
       
       details.push({
        date: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
        amount: amount,
        type: type,
        category: row[4] || '未分類',
        memo: row[5] || '',
        tag: rowTag
      });
    }
  });
  
  // 新しい順
  details.sort((a, b) => new Date(b.date) - new Date(a.date));
  
  return { details: details };
}

/**
 * 指定年の年間データを取得（月別推移・カテゴリ別集計）
 */
function getAnnualData(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB');
  
  // 初期値: 1月~12月のデータ配列
  const monthlyTrend = Array.from({ length: 12 }, (_, i) => ({
    month: i + 1,
    income: 0,
    expense: 0,
    balance: 0
  }));
  
  const yearCategory = {};
  
  if (!dbSheet) return { monthlyTrend, yearCategory };
  
  const lastRow = dbSheet.getLastRow();
  const lastCol = dbSheet.getLastColumn();
  
  if (lastRow <= 1) return { monthlyTrend, yearCategory };
  
  const data = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  data.forEach(row => {
    let dateVal = row[1] || row[0];
    if (!dateVal) return;
    let dateObj = new Date(dateVal);
    if (isNaN(dateObj.getTime())) return;
    
    // 指定年のみ対象
    if (dateObj.getFullYear() !== year) return;
    
    const monthIndex = dateObj.getMonth(); // 0-11
    const amount = Number(row[2]) || 0;
    const category = row[4] || '未分類';
    
    // 月別集計
    if (amount > 0) {
      monthlyTrend[monthIndex].income += amount;
    } else {
      const absAmount = Math.abs(amount);
      monthlyTrend[monthIndex].expense += absAmount;
      
      // カテゴリ別集計 (支出のみ)
      if (!yearCategory[category]) yearCategory[category] = 0;
      yearCategory[category] += absAmount;
    }
    
    monthlyTrend[monthIndex].balance += amount;
  });
  
  return { monthlyTrend, yearCategory };
}

/**
 * 全期間の年別推移を取得
 */
function getMultiYearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('DB');
  if (!dbSheet) return { yearStats: [] };

  const lastRow = dbSheet.getLastRow();
  const lastCol = dbSheet.getLastColumn();
  if (lastRow <= 1) return { yearStats: [] };

  const data = dbSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const statsMap = {};

  data.forEach(row => {
    let dateVal = row[1] || row[0];
    if (!dateVal) return;
    let dateObj = new Date(dateVal);
    if (isNaN(dateObj.getTime())) return;

    const y = dateObj.getFullYear();
    const amount = Number(row[2]) || 0;

    if (!statsMap[y]) {
      statsMap[y] = { year: y, income: 0, expense: 0, count: 0 };
    }

    statsMap[y].count++;
    if (amount > 0) {
      statsMap[y].income += amount;
    } else {
      statsMap[y].expense += Math.abs(amount);
    }
  });

  // 配列化して年昇順ソート
  const result = Object.values(statsMap).sort((a, b) => a.year - b.year);
  return { yearStats: result };
}