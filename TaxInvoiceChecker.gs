/**
 * 세금계산서 대조 시스템 v1.0
 * 메인 로직 파일
 */

/**
 * 메뉴 추가
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🔧 시스템 설정')
    .addItem('⚡ 초기 설정 실행', 'setupTaxInvoiceChecker')
    .addToUi();

  ui.createMenu('🔍 세금계산서 대조')
    .addItem('▶️ 전체 대조 실행', 'runFullComparison')
    .addSeparator()
    .addItem('📊 통계 보기', 'showStatistics')
    .addItem('🗑️ 결과 초기화', 'clearResults')
    .addToUi();

  SpreadsheetApp.getActive().toast('세금계산서 대조 시스템 v1.0 준비 완료!', '알림', 3);
}

/**
 * 전체 대조 실행
 */
function runFullComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ibkSheet = ss.getSheetByName('기업은행거래내역');
  const invoiceSheet = ss.getSheetByName('세금계산서발행내역');
  let resultSheet = ss.getSheetByName('대조결과');
  const ui = SpreadsheetApp.getUi();

  // 시트 확인
  if (!ibkSheet) {
    ui.alert('오류', '[기업은행거래내역] 시트를 찾을 수 없습니다!\n\n먼저 [시스템 설정] > [초기 설정 실행]을 클릭하세요.', ui.ButtonSet.OK);
    return;
  }

  if (!invoiceSheet) {
    ui.alert('오류', '[세금계산서발행내역] 시트를 찾을 수 없습니다!\n\n먼저 [시스템 설정] > [초기 설정 실행]을 클릭하세요.', ui.ButtonSet.OK);
    return;
  }

  // 진행 상황 표시
  SpreadsheetApp.getActive().toast('대조 시작...', '진행중', -1);

  // 결과 시트 초기화
  if (!resultSheet) {
    resultSheet = ss.insertSheet('대조결과');
  } else {
    resultSheet.clear();
  }

  // 헤더 작성
  const headers = [['일자', '거래처', '금액', '입금/출금', '매칭상태', '비고']];
  resultSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  resultSheet.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 1. 기업은행 거래내역 로드
  const ibkLastRow = ibkSheet.getLastRow();
  if (ibkLastRow < 2) {
    ui.alert('오류', '[기업은행거래내역] 시트에 데이터가 없습니다!\n\n거래 데이터를 입력한 후 다시 시도하세요.', ui.ButtonSet.OK);
    SpreadsheetApp.getActive().toast('', '', 1);
    return;
  }

  const ibkData = ibkSheet.getRange(2, 1, ibkLastRow - 1, 5).getValues();
  const ibkTransactions = [];

  ibkData.forEach((row, index) => {
    const date = formatDate(row[0]);  // A열: 일자
    const merchant = (row[1] || '').toString().trim();  // B열: 거래처
    const amount = parseFloat(row[2]) || 0;  // C열: 금액
    const transactionType = (row[3] || '').toString().trim();  // D열: 입금/출금
    const memo = (row[4] || '').toString().trim();  // E열: 메모

    if (date && merchant && amount > 0) {
      ibkTransactions.push({
        rowNum: index + 2,
        date: date,
        merchant: merchant,
        amount: amount,
        transactionType: transactionType,
        memo: memo
      });
    }
  });

  if (ibkTransactions.length === 0) {
    ui.alert('오류', '유효한 기업은행 거래 데이터가 없습니다!\n\n일자, 거래처, 금액이 모두 입력되어야 합니다.', ui.ButtonSet.OK);
    SpreadsheetApp.getActive().toast('', '', 1);
    return;
  }

  SpreadsheetApp.getActive().toast(`기업은행 거래 ${ibkTransactions.length}건 로드 완료`, '진행중', 2);

  // 2. 세금계산서 발행내역 로드
  const invoiceLastRow = invoiceSheet.getLastRow();
  if (invoiceLastRow < 2) {
    ui.alert('오류', '[세금계산서발행내역] 시트에 데이터가 없습니다!\n\n홈택스 데이터를 입력한 후 다시 시도하세요.', ui.ButtonSet.OK);
    SpreadsheetApp.getActive().toast('', '', 1);
    return;
  }

  const invoiceData = invoiceSheet.getRange(2, 1, invoiceLastRow - 1, 7).getValues();
  const issuedInvoices = [];

  invoiceData.forEach(row => {
    const date = formatDate(row[0]);  // A열: 발행일자
    const merchant = (row[1] || '').toString().trim();  // B열: 거래처명
    const supplyAmount = parseFloat(row[2]) || 0;  // C열: 공급가액
    const taxAmount = parseFloat(row[3]) || 0;  // D열: 세액
    const totalAmount = parseFloat(row[4]) || 0;  // E열: 합계금액
    const approvalNum = (row[5] || '').toString().trim();  // F열: 승인번호

    // 합계금액이 없으면 공급가액+세액으로 계산
    const amount = totalAmount > 0 ? totalAmount : (supplyAmount + taxAmount);

    if (amount > 0 && merchant) {
      issuedInvoices.push({
        date: date,
        merchant: merchant,
        amount: amount,
        approvalNum: approvalNum
      });
    }
  });

  if (issuedInvoices.length === 0) {
    ui.alert('오류', '유효한 세금계산서 데이터가 없습니다!\n\n발행일자, 거래처명, 금액이 모두 입력되어야 합니다.', ui.ButtonSet.OK);
    SpreadsheetApp.getActive().toast('', '', 1);
    return;
  }

  SpreadsheetApp.getActive().toast(`세금계산서 ${issuedInvoices.length}건 로드 완료`, '진행중', 2);

  // 3. 대조 작업
  SpreadsheetApp.getActive().toast('대조 중...', '진행중', -1);

  const unmatchedTransactions = [];
  const matchedTransactions = [];

  ibkTransactions.forEach((transaction, txIndex) => {
    let matched = false;
    let matchInfo = '';

    // 트랜잭션 데이터 검증
    if (!transaction || typeof transaction !== 'object') {
      Logger.log(`경고: Transaction ${txIndex}가 유효하지 않습니다: ${JSON.stringify(transaction)}`);
      return;
    }

    // 거래처명과 금액으로 매칭
    for (const invoice of issuedInvoices) {
      const transactionMerchantNorm = normalizeMerchantName(transaction.merchant);
      const invoiceMerchantNorm = normalizeMerchantName(invoice.merchant);

      // 거래처명 매칭: 정확히 일치하거나 부분 일치
      const exactMatch = transactionMerchantNorm === invoiceMerchantNorm;
      const partialMatch = transactionMerchantNorm.includes(invoiceMerchantNorm) ||
                          invoiceMerchantNorm.includes(transactionMerchantNorm);
      const merchantMatch = exactMatch || (partialMatch && Math.min(transactionMerchantNorm.length, invoiceMerchantNorm.length) >= 2);

      // 금액 매칭 (±1% 또는 ±1,000원 허용)
      const amountTolerance = Math.max(transaction.amount * 0.01, 1000);
      const amountMatch = Math.abs(transaction.amount - invoice.amount) <= amountTolerance;

      if (merchantMatch && amountMatch) {
        matched = true;
        matchInfo = `매칭됨 (발행일: ${invoice.date}, 금액: ${invoice.amount.toLocaleString()}원)`;

        // 배열 생성 전 검증
        const matchedRow = [
          String(transaction.date || ''),
          String(transaction.merchant || ''),
          Number(transaction.amount || 0),
          String(transaction.transactionType || ''),
          '✅ 발행확인',
          String(matchInfo)
        ];

        Logger.log(`매칭 성공 ${txIndex}: ${JSON.stringify(matchedRow)}`);
        matchedTransactions.push(matchedRow);
        break;
      }
    }

    // 매칭되지 않은 경우
    if (!matched) {
      const unmatchedRow = [
        String(transaction.date || ''),
        String(transaction.merchant || ''),
        Number(transaction.amount || 0),
        String(transaction.transactionType || ''),
        '⚠️ 미발행 의심',
        '홈택스 발행내역에서 찾을 수 없음'
      ];

      Logger.log(`미매칭 ${txIndex}: ${JSON.stringify(unmatchedRow)}`);
      unmatchedTransactions.push(unmatchedRow);
    }
  });

  // 4. 결과 작성 (미발행 의심 건을 먼저, 그 다음 발행확인 건)
  const resultData = [...unmatchedTransactions, ...matchedTransactions];

  if (resultData.length === 0) {
    ui.alert('대조할 데이터가 없습니다!', ui.ButtonSet.OK);
    SpreadsheetApp.getActive().toast('', '', 1);
    return;
  }

  // 데이터 유효성 검사 및 정규화
  const validatedData = resultData.map((row, index) => {
    if (!Array.isArray(row)) {
      Logger.log(`경고: Row ${index}가 배열이 아닙니다: ${JSON.stringify(row)}`);
      return ['', '', 0, '', '오류', '데이터 형식 오류'];
    }
    if (row.length !== 6) {
      Logger.log(`경고: Row ${index}의 열 개수가 ${row.length}개입니다 (예상: 6개): ${JSON.stringify(row)}`);
      // 6개로 맞추기
      while (row.length < 6) row.push('');
      row = row.slice(0, 6);
    }
    // 각 셀이 유효한지 확인
    return [
      row[0] || '',  // 일자
      row[1] || '',  // 거래처
      row[2] || 0,   // 금액
      row[3] || '',  // 입금/출금
      row[4] || '',  // 매칭상태
      row[5] || ''   // 비고
    ];
  });

  resultSheet.getRange(2, 1, validatedData.length, headers[0].length).setValues(validatedData);

  // 숫자 포맷
  resultSheet.getRange(2, 3, validatedData.length, 1).setNumberFormat('#,##0');

  // 조건부 서식
  const statusRange = resultSheet.getRange(2, 5, validatedData.length, 1);

  const unmatchedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('미발행 의심')
    .setBackground('#fee2e2')
    .setFontColor('#991b1b')
    .setRanges([statusRange])
    .build();

  const matchedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('발행확인')
    .setBackground('#d1fae5')
    .setFontColor('#065f46')
    .setRanges([statusRange])
    .build();

  resultSheet.setConditionalFormatRules([unmatchedRule, matchedRule]);

  // 열 너비 자동 조정
  resultSheet.autoResizeColumns(1, headers[0].length);
  resultSheet.setFrozenRows(1);

  // 통계
  const totalAmount = ibkTransactions.reduce((sum, t) => sum + t.amount, 0);
  const unmatchedAmount = unmatchedTransactions.reduce((sum, row) => sum + row[2], 0);
  const depositCount = ibkTransactions.filter(t => t.transactionType === '입금').length;
  const debitCount = ibkTransactions.filter(t => t.transactionType === '출금').length;

  SpreadsheetApp.getActive().toast('', '', 1);

  ui.alert(
    '✅ 대조 완료!',
    `[대조결과] 시트에 결과가 생성되었습니다.\n\n` +
    `📊 대조 결과:\n` +
    `• 기업은행 거래 총 ${ibkTransactions.length}건\n` +
    `  - 입금: ${depositCount}건\n` +
    `  - 출금: ${debitCount}건\n` +
    `• 총 금액: ${totalAmount.toLocaleString()}원\n\n` +
    `• ✅ 세금계산서 발행확인: ${matchedTransactions.length}건\n` +
    `• ⚠️ 미발행 의심: ${unmatchedTransactions.length}건 (${unmatchedAmount.toLocaleString()}원)\n\n` +
    `💡 빨간색으로 표시된 항목을 확인하세요!\n` +
    `💡 거래처명 부분 일치도 지원합니다 (예: "한메디"와 "한메디로")`,
    ui.ButtonSet.OK
  );
}

/**
 * 통계 보기
 */
function showStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('대조결과');
  const ui = SpreadsheetApp.getUi();

  if (!resultSheet || resultSheet.getLastRow() < 2) {
    ui.alert('통계 없음', '먼저 [전체 대조 실행]을 클릭하여 대조를 실행하세요.', ui.ButtonSet.OK);
    return;
  }

  const lastRow = resultSheet.getLastRow();
  const data = resultSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // 월별 통계
  const monthlyStats = {};

  data.forEach(row => {
    const dateValue = row[0];
    const amount = row[2];
    const status = row[4];

    if (!dateValue) return;

    // 날짜를 문자열로 변환 (Date 객체일 수 있음)
    const dateStr = typeof dateValue === 'string' ? dateValue : formatDate(dateValue);
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

    if (status.includes('발행확인')) {
      monthlyStats[month].matched++;
    } else {
      monthlyStats[month].unmatched++;
      monthlyStats[month].unmatchedAmount += amount;
    }
  });

  // 메시지 생성
  let message = '📊 월별 통계\n\n';

  const months = Object.keys(monthlyStats).sort().reverse();
  months.forEach(month => {
    const stats = monthlyStats[month];
    const matchRate = stats.total > 0 ? ((stats.matched / stats.total) * 100).toFixed(1) : 0;

    message += `${month}\n`;
    message += `  총 ${stats.total}건 (${stats.totalAmount.toLocaleString()}원)\n`;
    message += `  ✅ 발행: ${stats.matched}건 (${matchRate}%)\n`;
    message += `  ⚠️ 미발행: ${stats.unmatched}건 (${stats.unmatchedAmount.toLocaleString()}원)\n\n`;
  });

  ui.alert('월별 통계', message, ui.ButtonSet.OK);
}

/**
 * 결과 초기화
 */
function clearResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('대조결과');
  const ui = SpreadsheetApp.getUi();

  if (!resultSheet) {
    ui.alert('결과가 없습니다.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '결과 초기화',
    '[대조결과] 시트의 모든 데이터를 삭제하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    resultSheet.clear();

    // 헤더만 다시 작성
    const headers = [['일자', '거래처', '금액', '입금/출금', '매칭상태', '비고']];
    resultSheet.getRange(1, 1, 1, headers[0].length).setValues([headers]);

    resultSheet.getRange(1, 1, 1, headers[0].length)
      .setFontWeight('bold')
      .setBackground('#ea4335')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    ui.alert('결과가 초기화되었습니다.', ui.ButtonSet.OK);
  }
}

/**
 * 날짜 포맷 변환
 */
function formatDate(date) {
  if (!date) return '';

  try {
    let d;

    if (date instanceof Date) {
      d = date;
    } else if (typeof date === 'number') {
      // Excel 날짜 시리얼 번호
      d = new Date((date - 25569) * 86400 * 1000);
    } else {
      d = new Date(date);
    }

    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;
  } catch (e) {
    return date.toString();
  }
}

/**
 * 거래처명 정규화 (대조를 위한 문자열 정리)
 */
function normalizeMerchantName(name) {
  if (!name) return '';

  return name
    .toString()
    .trim()
    .replace(/\s+/g, '')  // 모든 공백 제거
    .replace(/\(.*?\)/g, '')  // 괄호 안 내용 제거
    .replace(/주식회사|유한회사|㈜|㈜/g, '')  // 회사 형태 제거
    .toLowerCase();  // 소문자 변환
}
