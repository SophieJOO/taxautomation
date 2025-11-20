/**
 * 세금계산서 매칭 시스템
 * 
 * 주요 기능:
 * 1. 세금계산서 CSV 업로드 및 파싱
 * 2. 전처리 (이름 정규화)
 * 3. 매칭 알고리즘 (정확, 퍼지, 부분합)
 */

// ========================================
// 1. 세금계산서 업로드 처리
// ========================================

function processTaxInvoiceCSV(csvData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('세금계산서매칭');
    
    if (!sheet) {
      throw new Error('[세금계산서매칭] 시트를 찾을 수 없습니다.');
    }
    
    // 기존 데이터 삭제 (헤더 제외)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    let imported = 0;
    const dataToImport = [];
    
    // 헤더 행 건너뛰기 (csvData[0]은 헤더)
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      
      // 빈 행 스킵
      if (!row[0] || row[0] === '') continue;
      
      // 홈택스 엑셀/CSV 형식 추정
      // 보통: 작성일자, 공급받는자등록번호, 공급받는자상호, ..., 합계금액, ...
      // 여기서는 사용자가 업로드한 CSV의 열 순서가 중요함.
      // 일단 [일자, 공급받는자, 공급가액, 세액, 합계금액, 비고] 순서로 가정하거나
      // CSVUploader에서 매핑을 해야 하는데, 
      // 현재는 간단히 0:일자, 1:상호, 2:공급가액, 3:세액, 4:합계금액, 5:비고 라고 가정
      // (실제 홈택스 파일은 컬럼이 매우 많으므로, 사용자가 필요한 컬럼만 남겨서 업로드한다고 가정하거나
      //  추후 컬럼 매핑 기능이 필요할 수 있음. 여기서는 단순화)
      
      const date = normalizeDate(row[0]);
      const vendor = row[1] || '';
      const supplyAmount = parseFloat(row[2]) || 0;
      const taxAmount = parseFloat(row[3]) || 0;
      const totalAmount = parseFloat(row[4]) || 0;
      const remark = row[5] || '';
      
      if (!date) continue;
      
      dataToImport.push([
        date,
        vendor,
        supplyAmount,
        taxAmount,
        totalAmount,
        remark,
        '미매칭', // 매칭상태
        '',       // 매칭점수
        '',       // 거래일자
        '',       // 거래처
        '',       // 입금액
        '',       // 출금액
        ''        // 메모
      ]);
      
      imported++;
    }
    
    if (dataToImport.length > 0) {
      sheet.getRange(2, 1, dataToImport.length, dataToImport[0].length).setValues(dataToImport);
    }
    
    return {
      imported: imported,
      categorized: 0,
      uncategorized: 0,
      type: 'tax'
    };
    
  } catch (error) {
    Logger.log('processTaxInvoiceCSV 오류: ' + error.toString());
    throw error;
  }
}

// ========================================
// 2. 매칭 알고리즘 실행
// ========================================

function runTaxInvoiceMatching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taxSheet = ss.getSheetByName('세금계산서매칭');
  const txnSheet = ss.getSheetByName('거래내역통합');
  const ui = SpreadsheetApp.getUi();
  
  if (!taxSheet || !txnSheet) {
    ui.alert('필요한 시트가 없습니다.');
    return;
  }
  
  const taxData = taxSheet.getRange(2, 1, taxSheet.getLastRow()-1, 13).getValues();
  const txnData = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  
  // 입출금 내역 분리 및 전처리
  const transactions = txnData.map((row, index) => ({
    id: index + 2,
    date: new Date(row[0]),
    vendor: normalizeName(row[2]),
    originalVendor: row[2],
    amount: (parseFloat(row[3]) || 0) + (parseFloat(row[4]) || 0), // 입금+출금 (절대값)
    type: parseFloat(row[4]) > 0 ? '입금' : '출금',
    memo: row[9],
    matched: false
  })).filter(t => t.amount > 0); // 0원 거래 제외
  
  let matchCount = 0;
  
  // 각 세금계산서에 대해 매칭 시도
  for (let i = 0; i < taxData.length; i++) {
    // 이미 매칭된 경우 스킵 (수동 매칭 등)
    if (taxData[i][6] === '매칭완료') continue;
    
    const taxDate = new Date(taxData[i][0]);
    const taxVendor = normalizeName(taxData[i][1]);
    const taxAmount = parseFloat(taxData[i][4]); // 합계금액
    
    // 1단계: 정확한 매칭 (날짜 ±3일, 금액 일치, 이름 일치)
    let bestMatch = null;
    
    // 후보군 필터링 (날짜 범위)
    const candidates = transactions.filter(t => {
      if (t.matched) return false;
      const dayDiff = Math.abs((t.date - taxDate) / (1000 * 60 * 60 * 24));
      return dayDiff <= 3;
    });
    
    // 1-1. 완전 일치 (금액 & 이름)
    const exactMatch = candidates.find(t => 
      Math.abs(t.amount - taxAmount) < 1 && 
      t.vendor === taxVendor
    );
    
    if (exactMatch) {
      updateMatchResult(taxSheet, i + 2, exactMatch, '정확', 100);
      exactMatch.matched = true;
      matchCount++;
      continue;
    }
    
    // 1-2. 금액 일치 & 이름 유사 (Fuzzy)
    const fuzzyMatch = candidates.find(t => 
      Math.abs(t.amount - taxAmount) < 1 && 
      calculateSimilarity(t.vendor, taxVendor) > 0.7
    );
    
    if (fuzzyMatch) {
      const score = Math.round(calculateSimilarity(fuzzyMatch.vendor, taxVendor) * 100);
      updateMatchResult(taxSheet, i + 2, fuzzyMatch, '유사', score);
      fuzzyMatch.matched = true;
      matchCount++;
      continue;
    }
    
    // 2단계: 복합 매칭 (Subset Sum)
    // 1:N (하나의 세금계산서 = 여러 입금내역 합계)
    // 예: 100만원 세금계산서 = 30만원 + 70만원 입금
    
    // 아직 매칭되지 않은 후보군 다시 필터링
    const remainingCandidates = candidates.filter(t => !t.matched);
    
    if (remainingCandidates.length > 1) {
      // 부분합 문제 해결 (최대 5개까지만 조합 시도하여 성능 저하 방지)
      const subsetMatch = findSubsetSum(remainingCandidates, taxAmount, 5);
      
      if (subsetMatch) {
        // 매칭 성공
        const matchedTxns = subsetMatch.txns;
        const totalMatchedAmount = subsetMatch.sum;
        
        // 결과 기록 (첫 번째 행에 메인 정보, 나머지는 비고에 표시)
        const mainTxn = matchedTxns[0];
        const otherTxns = matchedTxns.slice(1);
        
        let memo = `[복합매칭] ${mainTxn.date.toISOString().slice(0,10)} ${mainTxn.amount.toLocaleString()}`;
        otherTxns.forEach(t => {
          memo += `, ${t.date.toISOString().slice(0,10)} ${t.amount.toLocaleString()}`;
          t.matched = true; // 다른 내역도 매칭 처리
        });
        
        updateMatchResult(taxSheet, i + 2, mainTxn, '복합', 90);
        taxSheet.getRange(i + 2, 13).setValue(memo); // 메모 덮어쓰기
        
        mainTxn.matched = true;
        matchCount++;
      }
    }
  }
  
  ui.alert(`매칭 완료! 총 ${matchCount}건이 매칭되었습니다.`);
}

/**
 * 부분합 문제 해결 (Subset Sum Problem)
 * 주어진 거래 내역 중 합계가 targetAmount와 일치하는 조합을 찾음
 */
function findSubsetSum(transactions, targetAmount, maxDepth) {
  // 재귀 함수로 조합 탐색
  function search(index, currentSum, currentPath) {
    // 목표값과 정확히 일치 (오차 범위 1원)
    if (Math.abs(currentSum - targetAmount) < 1) {
      return { sum: currentSum, txns: currentPath };
    }
    
    // 목표값 초과하거나 탐색 깊이 초과 시 중단
    if (currentSum > targetAmount + 1 || currentPath.length >= maxDepth || index >= transactions.length) {
      return null;
    }
    
    // 현재 항목 포함
    const res1 = search(index + 1, currentSum + transactions[index].amount, [...currentPath, transactions[index]]);
    if (res1) return res1;
    
    // 현재 항목 미포함
    const res2 = search(index + 1, currentSum, currentPath);
    if (res2) return res2;
    
    return null;
  }
  
  return search(0, 0, []);
}
/**
 * 이름 정규화: (주), 공백, 특수문자 제거
 */
function normalizeName(name) {
  if (!name) return '';
  return name.toString()
    .replace(/\(주\)/g, '')
    .replace(/주식회사/g, '')
    .replace(/\s+/g, '') // 공백 제거
    .replace(/[^\w가-힣]/g, '') // 특수문자 제거
    .toLowerCase();
}

/**
 * 문자열 유사도 계산 (Levenshtein Distance 기반)
 */
function calculateSimilarity(s1, s2) {
  if (s1 === s2) return 1.0;
  if (s1.length === 0 || s2.length === 0) return 0.0;
  
  const longer = s1.length > s2.length ? s1 : s2;
  const shorter = s1.length > s2.length ? s2 : s1;
  
  if (longer.length === 0) return 1.0;
  
  const editDistance = getEditDistance(longer, shorter);
  return (longer.length - editDistance) / parseFloat(longer.length);
}

function getEditDistance(s1, s2) {
  const costs = new Array();
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i == 0)
        costs[j] = j;
      else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (s1.charAt(i - 1) != s2.charAt(j - 1))
            newValue = Math.min(Math.min(newValue, lastValue),
              costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0)
      costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

/**
 * 월별 매칭 통계 보고서 생성
 */
function generateTaxInvoiceReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('세금계산서매칭');
  const ui = SpreadsheetApp.getUi();

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('통계 없음', '데이터가 없습니다. 먼저 세금계산서를 업로드하고 매칭을 실행하세요.', ui.ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues(); // A:작성일자 ~ H:매칭점수

  // 월별 통계
  const monthlyStats = {};

  data.forEach(row => {
    const dateValue = row[0]; // A: 작성일자
    const amount = parseFloat(row[4]) || 0; // E: 합계금액
    const status = row[6]; // G: 매칭상태

    if (!dateValue) return;

    // 날짜를 문자열로 변환
    const dateStr = normalizeDate(dateValue);
    if (!dateStr) return;
    
    const month = dateStr.substring(0, 7); // YYYY-MM

    if (!monthlyStats[month]) {
      monthlyStats[month] = {
        total: 0,
        matched: 0,
        unmatched: 0,
        totalAmount: 0,
        unmatchedAmount: 0
      };
    }

    monthlyStats[month].total++;
    monthlyStats[month].totalAmount += amount;

    if (status === '정확' || status === '유사' || status === '복합') {
      monthlyStats[month].matched++;
    } else {
      monthlyStats[month].unmatched++;
      monthlyStats[month].unmatchedAmount += amount;
    }
  });

  // 메시지 생성
  let message = '📊 월별 세금계산서 매칭 통계\n\n';

  const months = Object.keys(monthlyStats).sort().reverse();
  months.forEach(month => {
    const stats = monthlyStats[month];
    const matchRate = stats.total > 0 ? ((stats.matched / stats.total) * 100).toFixed(1) : 0;

    message += `${month}\n`;
    message += `  총 ${stats.total}건 (${stats.totalAmount.toLocaleString()}원)\n`;
    message += `  ✅ 매칭: ${stats.matched}건 (${matchRate}%)\n`;
    message += `  ⚠️ 미매칭: ${stats.unmatched}건 (${stats.unmatchedAmount.toLocaleString()}원)\n\n`;
  });

  ui.alert('월별 통계', message, ui.ButtonSet.OK);
}
