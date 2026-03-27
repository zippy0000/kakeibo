function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('家計簿入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// フォームから送信されたデータを受け取り、スプレッドシートに書き込む機能
function processForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheets()[0]; // 1番目のシート（明細用）
  const budgetSheet = ss.getSheetByName('予算'); // 「予算」という名前のシート
  
  // 1. データを明細シートの下に追加
  txSheet.appendRow([
    data.date,      
    data.category,  
    data.amount,    
    data.shop,      
    data.details,
    data.wallet === '家計の財布' ? '' : data.wallet,
    '', // G列: 精算フラグ
    new Date()      
  ]);
  
  // 予算シートが無ければ、単純なOKだけ返す
  if (!budgetSheet) {
    return "✅ 登録完了しました！\n（残高を表示するには「予算」シートを作成してください）";
  }

  try {
    const inputDate = new Date(data.date);
    const inputYearMonth = inputDate.getFullYear() + '-' + (inputDate.getMonth() + 1);
    
    // 2. 予算データを取得
    const budgetData = budgetSheet.getDataRange().getValues();
    let budgetAmount = 0;
    let budgetType = ''; // '月' or '半期'
    
    // 予算シートの1行目は見出しと想定して、2行目(index: 1)から探す
    for (let i = 1; i < budgetData.length; i++) { 
      if (budgetData[i][0] == data.category) {
        budgetAmount = Number(budgetData[i][1]) || 0; // B列: 予算額
        budgetType = budgetData[i][2] || '月'; // C列: 月 or 半期
        break;
      }
    }
    
    // 予算額が設定されていない費目の場合はそのまま返す
    if (budgetAmount === 0) {
      return `✅ 登録完了！ (${data.category})`;
    }

    // 3. 使用金額の合計を計算
    const txData = txSheet.getDataRange().getValues();
    let usedAmount = 0;
    
    for (let i = 1; i < txData.length; i++) { // 1行目は見出し
      const rowCategory = txData[i][1];
      const rowAmount = Number(txData[i][2]) || 0;
      
      // 日付を「テキスト」から「日付形式」に変換して年・月を取得
      const rowDateCell = txData[i][0];
      const rowDate = (rowDateCell instanceof Date) ? rowDateCell : new Date(rowDateCell);
      
      if (rowCategory === data.category) {
        if (budgetType === '月') {
          // 月予算なら、今回入力された月と同じ明細だけ合算する
          const rowYearMonth = rowDate.getFullYear() + '-' + (rowDate.getMonth() + 1);
          if (rowYearMonth === inputYearMonth) {
            usedAmount += rowAmount;
          }
        } else {
          // 半期予算なら、その明細シート内の該当費目をすべて合算する
          usedAmount += rowAmount;
        }
      }
    }
    
    // 4. 残金計算 (Math.maxを外してマイナスも許容)
    const remaining = budgetAmount - usedAmount;
    
    // 5. 返信メッセージの作成
    const prefix = (budgetType === '月') ? '今月の' : '半期の';
    const numFormat = new Intl.NumberFormat('ja-JP').format(remaining);
    
    // マイナスの場合は目立たせる
    const balanceMessage = remaining < 0 ? `⚠️ ${numFormat}円 (予算超過！)` : `${numFormat}円`;
    
    return `✅ 登録完了！\n${prefix}「${data.category}」の残高は ${balanceMessage} です✨`;
    
  } catch(e) {
    return "✅ 登録完了！（エラー：" + e.message + "）";
  }
}

// 全予算の現在の残高一覧を取得する機能
function getBalances() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheets()[0]; 
    const budgetSheet = ss.getSheetByName('予算'); 
    
    if (!budgetSheet) return "⚠️ 「予算」シートが見つかりません。残高を表示するには作成してください。";

    const budgetData = budgetSheet.getDataRange().getValues();
    const txData = txSheet.getDataRange().getValues();
    
    // 今月の判定用
    const today = new Date();
    const currentYearMonth = today.getFullYear() + '-' + (today.getMonth() + 1);
    
    let monthlyBalances = [];
    let halfYearBalances = [];
    
    for (let i = 1; i < budgetData.length; i++) {
      const category = budgetData[i][0];
      const budgetAmount = Number(budgetData[i][1]) || 0;
      const budgetType = budgetData[i][2] || '月';
      
      if (!category || budgetAmount === 0) continue;
      
      let usedAmount = 0;
      for (let j = 1; j < txData.length; j++) {
        const rowCategory = txData[j][1];
        const rowAmount = Number(txData[j][2]) || 0;
        const rowDateCell = txData[j][0];
        const rowDate = (rowDateCell instanceof Date) ? rowDateCell : new Date(rowDateCell);
        
        if (rowCategory === category) {
          if (budgetType === '月') {
            const rowYearMonth = rowDate.getFullYear() + '-' + (rowDate.getMonth() + 1);
            if (rowYearMonth === currentYearMonth) {
              usedAmount += rowAmount;
            }
          } else {
            usedAmount += rowAmount;
          }
        }
      }
      
      const remaining = budgetAmount - usedAmount;
      const numFormat = new Intl.NumberFormat('ja-JP').format(remaining);
      const budgetFormat = new Intl.NumberFormat('ja-JP').format(budgetAmount);
      const displayStr = remaining < 0 ? `⚠️ ${numFormat}円 / ${budgetFormat}円` : `${numFormat}円 / ${budgetFormat}円`;
      
      if (budgetType === '月') {
        monthlyBalances.push(`・${category}: ${displayStr}`);
      } else {
        halfYearBalances.push(`・${category}: ${displayStr}`);
      }
    }
    
    if (monthlyBalances.length === 0 && halfYearBalances.length === 0) {
      return "予算データがありません。";
    }
    
    let res = "";
    if (monthlyBalances.length > 0) {
      res += "📅 【今月の予算残高】\n" + monthlyBalances.join("\n") + "\n\n";
    }
    if (halfYearBalances.length > 0) {
      res += "🗓️ 【半期の予算残高】\n" + halfYearBalances.join("\n") + "\n\n";
    }
    
    // 立て替え金額の計算
    let husbandMap = {};
    let wifeMap = {};
    let tatekaeHusband = 0;
    let tatekaeWife = 0;
    
    // 見出し行以降
    for (let j = 1; j < txData.length; j++) {
      const category = txData[j][1]; // B列: 費目
      const amount = Number(txData[j][2]) || 0;
      const wallet = txData[j][5]; // F列: 支払元
      const status = txData[j][6]; // G列: 精算フラグ
      
      if (wallet === '夫の財布' && status !== '済') {
        tatekaeHusband += amount;
        husbandMap[category] = (husbandMap[category] || 0) + amount;
      }
      if (wallet === '妻の財布' && status !== '済') {
        tatekaeWife += amount;
        wifeMap[category] = (wifeMap[category] || 0) + amount;
      }
    }
    
    if (tatekaeHusband > 0 || tatekaeWife > 0) {
      res += "💴 【未清算の立て替え額】\n";
      
      if (tatekaeHusband > 0) {
        res += `👨 夫へ：計 ${new Intl.NumberFormat('ja-JP').format(tatekaeHusband)}円\n`;
        for (let cat in husbandMap) {
          res += `　└ ${cat}: ${new Intl.NumberFormat('ja-JP').format(husbandMap[cat])}円\n`;
        }
      }
      if (tatekaeWife > 0) {
        res += `👩 妻へ：計 ${new Intl.NumberFormat('ja-JP').format(tatekaeWife)}円\n`;
        for (let cat in wifeMap) {
          res += `　└ ${cat}: ${new Intl.NumberFormat('ja-JP').format(wifeMap[cat])}円\n`;
        }
      }
    }
    
    return res.trim();
  } catch(e) {
    return "エラーが発生しました: " + e.message;
  }
}

// 立て替えを「精算済」にする機能
function settleReimbursement(person) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheets()[0];
    const txData = txSheet.getDataRange().getValues();
    
    if (txData.length < 2) {
      return `⚠️ 「${person}」の未清算データはありませんでした。`;
    }
    
    // G列のデータを一括で取得して一括で書き換える（超高速化）
    const statusRange = txSheet.getRange(1, 7, txData.length, 1);
    const statusValues = statusRange.getValues();
    
    let settledCount = 0;
    let settledAmount = 0;
    
    for (let j = 1; j < txData.length; j++) {
      const wallet = txData[j][5]; // F列
      const status = txData[j][6]; // G列
      const amount = Number(txData[j][2]) || 0;
      
      if (wallet === person && status !== '済') {
        statusValues[j][0] = '済'; // メモリ上で書き換え
        settledAmount += amount;
        settledCount++;
      }
    }
    
    if (settledCount === 0) {
      return `⚠️ 「${person}」の未清算データはありませんでした。`;
    }
    
    // 書き換え処理を一括で実行
    statusRange.setValues(statusValues);
    
    const numFormat = new Intl.NumberFormat('ja-JP').format(settledAmount);
    return `✅ 【精算完了】\n「${person}」の立て替え ${settledCount}件（合計 ${numFormat}円）を精算済にしました！`;
  } catch(e) {
    return "エラーが発生しました: " + e.message;
  }
}
