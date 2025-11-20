/**
 * 아현재한의원 회계 자동화 시스템 v3.3
 * 완전 자동화 버전 - 사람 개입 최소화
 * v3.3 업데이트 (최신):
 * - 세금계산서 관리 기능 추가
 * - 입금내역과 세금계산서 발행여부 대조
 * - 미발행 내역 자동 검사
 * - 월별 대조 보고서 생성
 * v3.2 업데이트:
 * - HTML 기반 CSV 파일 업로더 추가 (드래그 앤 드롭 지원!)
 * - 파일 업로드 후 자동 파싱 및 분류
 * - 사용자 친화적 UI 제공
 * v3.1 업데이트:
 * - 결제내역 파싱 오류 수정 (날짜 정규화)
 * - 자동분류 로직 개선 (수식 자동 복구)
 * - 기존 데이터 복구 기능 추가
 * - 중복 체크 개선 (부동소수점 오차 처리)
 */

// ========================================
// 1. 초기 설정 및 메뉴
// ========================================

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('💰 한의원 회계')
      .addItem('🚀 원클릭 자동처리', 'oneClickAutomation')
      .addSeparator()
      .addItem('📤 CSV 파일 업로드 (신규!)', 'showCSVUploader')
      .addItem('📥 CSV 데이터 가져오기', 'importCSVData')
      .addItem('🔄 자동분류 실행', 'runAutoCategory')
      .addItem('📊 월간 보고서 생성', 'generateMonthlyReport')
      .addSeparator()
      .addSubMenu(ui.createMenu('🧾 세금계산서 관리')
        .addItem('① 세금계산서 업로드', 'showCSVUploader')
        .addItem('② 매칭 실행', 'runTaxInvoiceMatching')
        .addItem('📊 월별 대조 보고서', 'generateTaxInvoiceReport'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📁 세무사 전달용')
        .addItem('① 거래상세내역 (전체)', 'exportDetailedTransactions')
        .addItem('② 계정과목별 집계', 'exportCategorySummary')
        .addItem('③ 사업지출만 (간단)', 'exportForAccountant'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🔧 시스템 설정')
        .addItem('⚡ 초기 설정 실행', 'setupAhyunClinicSheets')
        .addItem('🔄 시트 재생성', 'recreateAllSheets')
        .addItem('📖 설정 가이드', 'showSetupGuide'))
      .addToUi();
  } catch (error) {
    Logger.log('메뉴 생성 오류: ' + error.toString());
  }
}

// ========================================
// 2. CSV 파일 업로더 (신규!)
// ========================================

/**
 * CSV 파일 업로더 다이얼로그 표시
 */
function showCSVUploader() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('CSVUploader')
      .setWidth(650)
      .setHeight(700)
      .setTitle('CSV 파일 업로드');

    SpreadsheetApp.getUi().showModalDialog(html, 'CSV 파일 업로드');
  } catch (error) {
    Logger.log('CSV 업로더 표시 오류: ' + error.toString());
    SpreadsheetApp.getUi().alert('오류', 'CSV 업로더를 열 수 없습니다: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 업로드된 CSV 데이터 처리 (HTML에서 호출)
 */
function processUploadedCSV(csvData, uploadType = 'bank') {
  try {
    // 세금계산서 업로드인 경우
    if (uploadType === 'tax') {
      return processTaxInvoiceCSV(csvData);
    }
    
    // 세금계산서용 은행내역 업로드인 경우 (신규)
    if (uploadType === 'tax_bank') {
      return processTaxBankCSV(csvData);
    }

    // 기존 은행/카드 거래내역 처리 (월간 회계용)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txnSheet = ss.getSheetByName('거래내역통합');

    if (!txnSheet) {
      throw new Error('[거래내역통합] 시트를 찾을 수 없습니다. Setup.gs를 먼저 실행하세요.');
    }

    let imported = 0;
    const lastRow = txnSheet.getLastRow();

    // 헤더 행 건너뛰기 (csvData[0]은 헤더)
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];

      // 빈 행 스킵
      if (!row[0] || row[0] === '') continue;

      // 날짜 정규화
      const normalizedDate = normalizeDate(row[0]);
      if (!normalizedDate) continue;

      // 중복 체크
      const isDuplicate = checkDuplicate(txnSheet, normalizedDate, row[2], row[3]);
      if (isDuplicate) continue;

      // [거래내역통합]에 추가
      const newRow = lastRow + imported + 1;
      txnSheet.getRange(newRow, 1).setValue(normalizedDate);  // A: 일자
      txnSheet.getRange(newRow, 2).setValue(row[1] || '');  // B: 카드/계좌
      txnSheet.getRange(newRow, 3).setValue(row[2] || '');  // C: 거래처
      txnSheet.getRange(newRow, 4).setValue(parseFloat(row[3]) || 0);  // D: 출금액
      txnSheet.getRange(newRow, 5).setValue(parseFloat(row[4]) || 0);  // E: 입금액
      txnSheet.getRange(newRow, 8).setFormula('=IF(G' + newRow + '<>"",G' + newRow + ',F' + newRow + ')');  // H: 최종분류
      txnSheet.getRange(newRow, 10).setValue(row[5] || '');  // J: 메모

      imported++;
    }

    // 자동분류 실행
    const categorized = runAutoCategory(true);

    // 미분류 개수 확인
    const uncategorized = countUncategorized();

    // 결과 반환
    return {
      imported: imported,
      categorized: categorized,
      uncategorized: uncategorized,
      type: 'bank'
    };

  } catch (error) {
    Logger.log('processUploadedCSV 오류: ' + error.toString());
    throw new Error('CSV 처리 중 오류가 발생했습니다: ' + error.toString());
  }
}

// ========================================
// 3. 원클릭 자동처리 (핵심 기능!)
// ========================================

function oneClickAutomation() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🚀 원클릭 자동처리',
    '다음 작업을 자동으로 수행합니다:\n\n' +
    '1. CSV 데이터 가져오기\n' +
    '2. 자동분류 실행\n' +
    '3. 월간 보고서 생성\n' +
    '4. 미분류 항목 알림\n\n' +
    '[CSV임시] 시트에 데이터를 붙여넣고 확인을 누르세요.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response != ui.Button.OK) return;
  
  try {
    SpreadsheetApp.getActive().toast('1/4: CSV 데이터 가져오는 중...', '처리중', -1);
    const imported = importCSVData(true);  // silent mode
    
    if (imported === 0) {
      ui.alert('❌ [CSV임시] 시트에 데이터가 없습니다!');
      return;
    }
    
    SpreadsheetApp.getActive().toast('2/4: 자동분류 실행 중...', '처리중', -1);
    const categorized = runAutoCategory(true);  // silent mode
    
    SpreadsheetApp.getActive().toast('3/4: 월간 보고서 생성 중...', '처리중', -1);
    generateMonthlyReport(true);  // silent mode
    
    SpreadsheetApp.getActive().toast('4/4: 최종 확인 중...', '처리중', -1);
    const uncategorized = countUncategorized();
    
    // 완료 메시지
    let message = `✅ 자동처리 완료!\n\n`;
    message += `📥 가져온 거래: ${imported}건\n`;
    message += `✅ 자동분류: ${categorized}건\n`;
    message += `❓ 미분류: ${uncategorized}건\n\n`;
    
    if (uncategorized > 0) {
      message += `⚠️ 미분류 항목이 있습니다.\n`;
      message += `[미분류 항목 보기]를 눌러 확인하세요.`;
    } else {
      message += `🎉 모든 거래가 분류되었습니다!`;
    }
    
    SpreadsheetApp.getActive().toast('완료!', '자동처리', 1);
    ui.alert('원클릭 자동처리', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('오류 발생: ' + error.toString());
  }
}

// ========================================
// 4. CSV 데이터 가져오기 (개선 버전)
// ========================================

function importCSVData(silentMode = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error('스프레드시트를 찾을 수 없습니다.');
    }

    const tempSheet = ss.getSheetByName('CSV임시');
    const txnSheet = ss.getSheetByName('거래내역통합');
    const ui = SpreadsheetApp.getUi();

    if (!tempSheet) {
      const message = '[CSV임시] 시트를 찾을 수 없습니다!\n\nSetup.gs를 먼저 실행하세요:\n1. [확장 프로그램] > [Apps Script]\n2. Setup.gs 열기\n3. setupAhyunClinicSheets 실행';
      if (!silentMode) ui.alert('오류', message, ui.ButtonSet.OK);
      throw new Error('[CSV임시] 시트가 없습니다.');
    }

    if (!txnSheet) {
      const message = '[거래내역통합] 시트를 찾을 수 없습니다!\n\nSetup.gs를 먼저 실행하세요.';
      if (!silentMode) ui.alert('오류', message, ui.ButtonSet.OK);
      throw new Error('[거래내역통합] 시트가 없습니다.');
    }

    const data = tempSheet.getDataRange().getValues();

    if (data.length < 2) {
      if (!silentMode) ui.alert('[CSV임시] 시트가 비어있습니다!');
      return 0;
    }

    let imported = 0;
    const lastRow = txnSheet.getLastRow();

    // 헤더 행 건너뛰기 (1행)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 빈 행 스킵
      if (!row[0] || row[0] === '') continue;

      // 날짜 정규화
      const normalizedDate = normalizeDate(row[0]);
      if (!normalizedDate) continue;

      // 중복 체크 (같은 날짜, 거래처, 금액)
      const isDuplicate = checkDuplicate(txnSheet, normalizedDate, row[2], row[3]);
      if (isDuplicate) continue;

      // [거래내역통합]에 추가
      const newRow = lastRow + imported + 1;
      txnSheet.getRange(newRow, 1).setValue(normalizedDate);  // A: 일자
      txnSheet.getRange(newRow, 2).setValue(row[1] || '');  // B: 카드/계좌
      txnSheet.getRange(newRow, 3).setValue(row[2] || '');  // C: 거래처
      txnSheet.getRange(newRow, 4).setValue(parseFloat(row[3]) || 0);  // D: 출금액
      txnSheet.getRange(newRow, 5).setValue(parseFloat(row[4]) || 0);  // E: 입금액
      // F: 자동분류 (비워둠)
      // G: 수동분류 (비워둠)
      txnSheet.getRange(newRow, 8).setFormula('=IF(G' + newRow + '<>"",G' + newRow + ',F' + newRow + ')');  // H: 최종분류
      // I: 사업/개인 (비워둠)
      txnSheet.getRange(newRow, 10).setValue(row[5] || '');  // J: 메모

      imported++;
    }

    // CSV임시 시트 비우기
    tempSheet.clear();

    // 헤더 다시 추가
    const headers = [['일자', '카드/계좌', '거래처', '출금액', '입금액', '메모']];
    tempSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    tempSheet.getRange(1, 1, 1, headers[0].length)
      .setFontWeight('bold')
      .setBackground('#9e9e9e')
      .setFontColor('#ffffff');

    if (!silentMode) {
      ui.alert(
        '가져오기 완료!',
        `${imported}건의 거래를 가져왔습니다.`,
        ui.ButtonSet.OK
      );
    }

    return imported;
  } catch (error) {
    Logger.log('importCSVData 오류: ' + error.toString());
    throw error;
  }
}

/**
 * 날짜 정규화 함수
 */
function normalizeDate(date) {
  if (!date) return null;

  try {
    let d;

    // 이미 Date 객체인 경우
    if (date instanceof Date) {
      d = date;
    }
    // 문자열인 경우
    else if (typeof date === 'string') {
      // YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD 형식 지원
      d = new Date(date.replace(/\./g, '-').replace(/\//g, '-'));
    }
    // 숫자인 경우 (엑셀 시리얼 날짜)
    else if (typeof date === 'number') {
      d = new Date((date - 25569) * 86400 * 1000);
    }
    else {
      return null;
    }

    // 유효한 날짜인지 확인
    if (isNaN(d.getTime())) return null;

    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    Logger.log('날짜 정규화 오류: ' + e.toString() + ', 입력값: ' + date);
    return null;
  }
}

/**
 * 중복 거래 체크 (개선 버전)
 */
function checkDuplicate(sheet, date, merchant, amount) {
  if (!date || !merchant) return false;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;  // 헤더만 있으면 중복 없음

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  // 날짜 정규화
  const normalizedDate = normalizeDate(date);
  const normalizedAmount = parseFloat(amount) || 0;

  for (let i = 0; i < data.length; i++) {
    const rowDate = normalizeDate(data[i][0]);
    const rowMerchant = data[i][2];
    const rowAmount = parseFloat(data[i][3]) || 0;

    if (rowDate === normalizedDate &&
        rowMerchant === merchant &&
        Math.abs(rowAmount - normalizedAmount) < 0.01) {  // 부동소수점 오차 고려
      return true;
    }
  }

  return false;
}

// ========================================
// 4. 자동분류 실행 (개선 버전)
// ========================================

function runAutoCategory(silentMode = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error('스프레드시트를 찾을 수 없습니다.');
    }

    const txnSheet = ss.getSheetByName('거래내역통합');
    const rulesSheet = ss.getSheetByName('분류규칙');
    const ui = SpreadsheetApp.getUi();

    if (!txnSheet) {
      const message = '[거래내역통합] 시트를 찾을 수 없습니다!\n\nSetup.gs를 먼저 실행하세요.';
      if (!silentMode) ui.alert('오류', message, ui.ButtonSet.OK);
      throw new Error('[거래내역통합] 시트가 없습니다.');
    }

    if (!rulesSheet) {
      const message = '[분류규칙] 시트를 찾을 수 없습니다!\n\nSetup.gs를 먼저 실행하세요.';
      if (!silentMode) ui.alert('오류', message, ui.ButtonSet.OK);
      throw new Error('[분류규칙] 시트가 없습니다.');
    }

    // 분류 규칙 로드 (개선된 에러 처리)
    const rulesLastRow = rulesSheet.getLastRow();
    if (rulesLastRow < 2) {
      if (!silentMode) ui.alert('[분류규칙] 시트에 규칙을 먼저 입력하세요!');
      return 0;
    }

    const rulesData = rulesSheet.getRange(2, 1, rulesLastRow - 1, 6).getValues();
    const rules = rulesData
      .filter(r => r[0] !== '' && r[0] !== null && r[3] !== '' && r[3] !== null)  // 키워드가 있는 것만
      .sort((a, b) => b[0] - a[0]);  // 우선순위 내림차순

    if (rules.length === 0) {
      if (!silentMode) ui.alert('[분류규칙] 시트에 유효한 규칙이 없습니다!\n\n키워드가 입력된 규칙을 추가하세요.');
      return 0;
    }

    // 거래 데이터 로드 (개선된 에러 처리)
    const txnLastRow = txnSheet.getLastRow();
    if (txnLastRow < 2) {
      if (!silentMode) ui.alert('[거래내역통합] 시트에 데이터가 없습니다!');
      return 0;
    }

    const txnData = txnSheet.getRange(2, 1, txnLastRow - 1, 10).getValues();

    if (txnData.length === 0) {
      if (!silentMode) ui.alert('[거래내역통합] 시트에 데이터가 없습니다!');
      return 0;
    }

    let categorized = 0;
    let skipped = 0;
    let formulaFixed = 0;

    if (!silentMode) {
      SpreadsheetApp.getActive().toast('자동분류 시작...', '진행중', -1);
    }

    // 각 거래 분류
    for (let i = 0; i < txnData.length; i++) {
      const rowNum = i + 2;
      const merchant = txnData[i][2];  // C열: 거래처
      const manualCategory = txnData[i][6];  // G열: 수동분류

      // H열에 수식이 없으면 추가 (기존 데이터 복구)
      const finalCategoryCell = txnSheet.getRange(rowNum, 8);
      const formula = finalCategoryCell.getFormula();
      if (!formula || formula === '') {
        finalCategoryCell.setFormula('=IF(G' + rowNum + '<>"",G' + rowNum + ',F' + rowNum + ')');
        formulaFixed++;
      }

      // 이미 수동 분류된 것은 스킵
      if (manualCategory && manualCategory !== '') {
        skipped++;
        continue;
      }

      if (!merchant || merchant === '') continue;

      // 규칙 매칭 (개선: 부분 일치 + 정규식)
      let matched = false;
      for (const rule of rules) {
        if (!rule[3]) continue;

        const keywords = rule[3].toString().toLowerCase().split('|');
        const merchantLower = merchant.toLowerCase().trim();

        for (const keyword of keywords) {
          const trimmedKeyword = keyword.trim();
          if (trimmedKeyword === '') continue;

          // 부분 일치 또는 정규식 매칭
          if (merchantLower.includes(trimmedKeyword) || matchRegex(merchantLower, trimmedKeyword)) {
            txnSheet.getRange(rowNum, 6).setValue(rule[2]);  // F: 자동분류 (중분류/계정과목)
            txnSheet.getRange(rowNum, 9).setValue(rule[4]);  // I: 사업/개인
            categorized++;
            matched = true;
            break;
          }
        }
        if (matched) break;
      }

      // 진행 상황 표시
      if (i % 50 === 0 && i > 0 && !silentMode) {
        SpreadsheetApp.getActive().toast(
          `${i}/${txnData.length}건 처리 중...`,
          '진행중', 2
        );
      }
    }

    if (!silentMode) {
      SpreadsheetApp.getActive().toast('완료!', '자동분류', 1);

      let message = `총 ${txnData.length}건 중\n\n` +
        `✅ 자동분류: ${categorized}건\n` +
        `⏭️ 수동분류 유지: ${skipped}건\n` +
        `❓ 미분류: ${txnData.length - categorized - skipped}건`;

      if (formulaFixed > 0) {
        message += `\n\n🔧 수식 복구: ${formulaFixed}건`;
      }

      ui.alert('자동분류 완료!', message, ui.ButtonSet.OK);
    }

    return categorized;
  } catch (error) {
    Logger.log('runAutoCategory 오류: ' + error.toString());
    throw error;
  }
}

/**
 * 정규식 매칭 (간단한 와일드카드 지원)
 */
function matchRegex(text, pattern) {
  try {
    const regex = new RegExp(pattern, 'i');
    return regex.test(text);
  } catch (e) {
    return false;
  }
}

// ========================================
// 5. 월간 보고서 자동 생성
// ========================================

function generateMonthlyReport(silentMode = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  let reportSheet = ss.getSheetByName('월간보고서');
  const ui = SpreadsheetApp.getUi();
  
  // 보고서 시트 생성 (없으면)
  if (!reportSheet) {
    reportSheet = ss.insertSheet('월간보고서');
  } else {
    reportSheet.clear();
  }
  
  // 헤더 작성
  const headers = ['월', '대분류', '계정과목', '사업지출', '개인지출', '합계', '거래건수'];
  reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  reportSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  
  // 데이터 집계
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  
  const monthlyData = {};
  
  data.forEach(row => {
    const date = new Date(row[0]);
    const month = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    const category = row[7] || '미분류';  // H열: 최종분류
    const businessType = row[8] || '확인필요';  // I열: 사업/개인
    const amount = parseFloat(row[3]) || 0;  // D열: 출금액
    
    const key = `${month}|${category}`;
    
    if (!monthlyData[key]) {
      monthlyData[key] = {
        month: month,
        category: category,
        business: 0,
        personal: 0,
        count: 0
      };
    }
    
    if (businessType === '사업') {
      monthlyData[key].business += amount;
    } else if (businessType === '개인') {
      monthlyData[key].personal += amount;
    }
    
    monthlyData[key].count++;
  });
  
  // 보고서 작성
  const reportData = [];
  Object.values(monthlyData).forEach(item => {
    reportData.push([
      item.month,
      '',  // 대분류 (추후 추가)
      item.category,
      item.business,
      item.personal,
      item.business + item.personal,
      item.count
    ]);
  });
  
  // 월별, 금액순 정렬
  reportData.sort((a, b) => {
    if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
    return b[5] - a[5];
  });
  
  if (reportData.length > 0) {
    reportSheet.getRange(2, 1, reportData.length, headers.length).setValues(reportData);
    
    // 숫자 포맷
    reportSheet.getRange(2, 4, reportData.length, 3).setNumberFormat('#,##0');
  }
  
  // 열 너비 자동 조정
  reportSheet.autoResizeColumns(1, headers.length);
  
  if (!silentMode) {
    ui.alert(
      '월간 보고서 생성 완료!',
      `[월간보고서] 시트에 ${reportData.length}개 항목이 생성되었습니다.`,
      ui.ButtonSet.OK
    );
  }
}

// ========================================
// 6. 미분류 항목 보기
// ========================================

function showUncategorized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  const ui = SpreadsheetApp.getUi();
  
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  
  const uncategorized = [];
  for (let i = 0; i < data.length; i++) {
    const finalCategory = data[i][7];  // H열: 최종분류
    if (!finalCategory || finalCategory === '' || finalCategory === '미분류') {
      uncategorized.push({
        row: i + 2,
        date: data[i][0],
        merchant: data[i][2],
        amount: data[i][3]
      });
    }
  }
  
  if (uncategorized.length === 0) {
    ui.alert('미분류 항목이 없습니다! 🎉');
    return;
  }
  
  let message = `미분류 항목 ${uncategorized.length}건:\n\n`;
  
  // 거래처별로 그룹화
  const merchantCounts = {};
  uncategorized.forEach(item => {
    if (!merchantCounts[item.merchant]) {
      merchantCounts[item.merchant] = { count: 0, total: 0 };
    }
    merchantCounts[item.merchant].count++;
    merchantCounts[item.merchant].total += item.amount;
  });
  
  // 빈도순 정렬
  const sorted = Object.entries(merchantCounts)
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, 15);
  
  sorted.forEach(([merchant, data], index) => {
    message += `${index+1}. ${merchant}\n`;
    message += `   ${data.count}건, ${data.total.toLocaleString()}원\n`;
  });
  
  message += `\n💡 [분류규칙] 시트에 키워드를 추가하세요!`;
  
  ui.alert('미분류 항목', message, ui.ButtonSet.OK);
}

/**
 * 미분류 항목 개수 반환
 */
function countUncategorized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  
  let count = 0;
  data.forEach(row => {
    const finalCategory = row[7];
    if (!finalCategory || finalCategory === '' || finalCategory === '미분류') {
      count++;
    }
  });
  
  return count;
}

// ========================================
// 7. 계정과목별 집계
// ========================================

function showCategoryTotals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  const ui = SpreadsheetApp.getUi();
  
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  
  const categoryTotals = {};
  const businessTotals = { '사업': 0, '개인': 0, '확인필요': 0 };
  
  data.forEach(row => {
    const category = row[7] || '미분류';
    const businessType = row[8] || '확인필요';
    const amount = parseFloat(row[3]) || 0;
    
    if (!categoryTotals[category]) {
      categoryTotals[category] = 0;
    }
    categoryTotals[category] += amount;
    businessTotals[businessType] += amount;
  });
  
  const sorted = Object.entries(categoryTotals)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);
  
  let message = '📊 계정과목별 지출 현황 (TOP 15)\n\n';
  
  sorted.forEach(([category, total], index) => {
    message += `${index+1}. ${category}\n`;
    message += `   ${total.toLocaleString()}원\n`;
  });
  
  message += `\n📈 구분별 합계:\n`;
  message += `💼 사업: ${businessTotals['사업'].toLocaleString()}원\n`;
  message += `🏠 개인: ${businessTotals['개인'].toLocaleString()}원\n`;
  message += `❓ 확인필요: ${businessTotals['확인필요'].toLocaleString()}원\n`;
  message += `\n💰 총합: ${(businessTotals['사업'] + businessTotals['개인'] + businessTotals['확인필요']).toLocaleString()}원`;
  
  ui.alert('계정과목별 집계', message, ui.ButtonSet.OK);
}

// ========================================
// 8. 세무사 전달용 파일 생성
// ========================================

function exportForAccountant() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  let exportSheet = ss.getSheetByName('세무사전달');
  const ui = SpreadsheetApp.getUi();
  
  // 시트 생성
  if (!exportSheet) {
    exportSheet = ss.insertSheet('세무사전달');
  } else {
    exportSheet.clear();
  }
  
  // 사업 지출만 필터링
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  const businessData = data.filter(row => row[8] === '사업');
  
  // 헤더
  const headers = ['일자', '계정과목', '거래처', '금액', '메모'];
  exportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  exportSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  
  // 데이터 작성
  const exportData = businessData.map(row => [
    row[0],  // 일자
    row[7],  // 최종분류
    row[2],  // 거래처
    row[3],  // 출금액
    row[9] || ''  // 메모
  ]);
  
  if (exportData.length > 0) {
    exportSheet.getRange(2, 1, exportData.length, headers.length).setValues(exportData);
    exportSheet.getRange(2, 4, exportData.length, 1).setNumberFormat('#,##0');
  }
  
  exportSheet.autoResizeColumns(1, headers.length);
  
  ui.alert(
    '세무사 전달용 파일 생성 완료!',
    `[세무사전달] 시트에 ${exportData.length}건의 사업 지출이 정리되었습니다.\n\n` +
    `이 시트를 복사하여 세무사님께 전달하세요.`,
    ui.ButtonSet.OK
  );
}

// ========================================
// 9. 분류규칙 자동 최적화
// ========================================

function optimizeRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  const rulesSheet = ss.getSheetByName('분류규칙');
  const ui = SpreadsheetApp.getUi();
  
  // 미분류 거래처 분석
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  const uncategorizedMerchants = {};
  
  data.forEach(row => {
    const finalCategory = row[7];
    const merchant = row[2];
    
    if (!finalCategory || finalCategory === '' || finalCategory === '미분류') {
      if (!uncategorizedMerchants[merchant]) {
        uncategorizedMerchants[merchant] = 0;
      }
      uncategorizedMerchants[merchant]++;
    }
  });
  
  const sorted = Object.entries(uncategorizedMerchants)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (sorted.length === 0) {
    ui.alert('최적화 완료!', '미분류 항목이 없습니다. 🎉', ui.ButtonSet.OK);
    return;
  }
  
  let message = '📊 자주 나오는 미분류 거래처 TOP 10:\n\n';
  sorted.forEach(([merchant, count], index) => {
    message += `${index+1}. ${merchant} (${count}건)\n`;
  });
  message += `\n💡 이 거래처들을 [분류규칙]에 추가하세요!`;
  
  ui.alert('분류규칙 최적화', message, ui.ButtonSet.OK);
}

// ========================================
// 10. 기존 데이터 복구 (신규 추가)
// ========================================

function fixExistingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  const ui = SpreadsheetApp.getUi();

  if (!txnSheet) {
    ui.alert('오류', '[거래내역통합] 시트를 찾을 수 없습니다!', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '🔧 기존 데이터 복구',
    '이 기능은 다음 작업을 수행합니다:\n\n' +
    '1. H열(최종분류)에 수식 추가/복구\n' +
    '2. 날짜 형식 정규화\n' +
    '3. 숫자 형식 정규화\n\n' +
    '⚠️ 주의: 기존 데이터가 수정될 수 있습니다.\n\n' +
    '계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (response != ui.Button.YES) return;

  try {
    SpreadsheetApp.getActive().toast('데이터 복구 시작...', '진행중', -1);

    const lastRow = txnSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('데이터가 없습니다!');
      return;
    }

    const data = txnSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    let formulaFixed = 0;
    let dateFixed = 0;
    let amountFixed = 0;

    for (let i = 0; i < data.length; i++) {
      const rowNum = i + 2;

      // 1. H열 수식 추가/복구
      const finalCategoryCell = txnSheet.getRange(rowNum, 8);
      const formula = finalCategoryCell.getFormula();
      if (!formula || formula === '') {
        finalCategoryCell.setFormula('=IF(G' + rowNum + '<>"",G' + rowNum + ',F' + rowNum + ')');
        formulaFixed++;
      }

      // 2. 날짜 정규화
      const dateCell = txnSheet.getRange(rowNum, 1);
      const currentDate = dateCell.getValue();
      if (currentDate) {
        const normalized = normalizeDate(currentDate);
        if (normalized && normalized !== currentDate) {
          dateCell.setValue(normalized);
          dateFixed++;
        }
      }

      // 3. 출금액/입금액 숫자 형식 확인
      const debitCell = txnSheet.getRange(rowNum, 4);
      const creditCell = txnSheet.getRange(rowNum, 5);

      const debitValue = debitCell.getValue();
      const creditValue = creditCell.getValue();

      if (debitValue !== '' && typeof debitValue !== 'number') {
        const parsed = parseFloat(debitValue);
        if (!isNaN(parsed)) {
          debitCell.setValue(parsed);
          amountFixed++;
        }
      }

      if (creditValue !== '' && typeof creditValue !== 'number') {
        const parsed = parseFloat(creditValue);
        if (!isNaN(parsed)) {
          creditCell.setValue(parsed);
          amountFixed++;
        }
      }

      // 진행 상황 표시
      if (i % 100 === 0 && i > 0) {
        SpreadsheetApp.getActive().toast(
          `${i}/${data.length}건 처리 중...`,
          '진행중', 2
        );
      }
    }

    SpreadsheetApp.getActive().toast('완료!', '데이터 복구', 1);

    ui.alert(
      '✅ 데이터 복구 완료!',
      `총 ${data.length}건 처리:\n\n` +
      `🔧 수식 복구: ${formulaFixed}건\n` +
      `📅 날짜 정규화: ${dateFixed}건\n` +
      `💰 금액 정규화: ${amountFixed}건\n\n` +
      `이제 [자동분류 실행]을 다시 실행해보세요!`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log('fixExistingData 오류: ' + error.toString());
    ui.alert('오류 발생', error.toString(), ui.ButtonSet.OK);
  }
}

// ========================================
// 11. 도움말
// ========================================

function showHelp() {
  const ui = SpreadsheetApp.getUi();

  const message = `🏥 아현재한의원 회계 자동화 시스템 v3.3\n\n` +
    `📖 사용 방법 (두 가지 방식):\n\n` +
    `✨ 방법 1: CSV 파일 업로드 (추천)\n` +
    `1️⃣ [CSV 파일 업로드] 메뉴 클릭\n` +
    `2️⃣ CSV 파일을 드래그하거나 선택\n` +
    `3️⃣ 자동으로 파싱 및 분류 완료!\n\n` +
    `📋 방법 2: 기존 방식\n` +
    `1️⃣ [CSV임시]에 데이터 붙여넣기\n` +
    `2️⃣ [원클릭 자동처리] 버튼 클릭\n\n` +
    `🧾 세금계산서 관리 (NEW!):\n` +
    `1️⃣ [입금내역 보기] - 입금건만 필터링\n` +
    `2️⃣ [미발행 내역 검사] - 미발행 항목 찾기\n` +
    `3️⃣ [월별 대조 보고서] - 월별 발행률 확인\n\n` +
    `💡 팁:\n` +
    `- CSV 파일 업로더가 가장 편리합니다!\n` +
    `- 세금계산서 열에 "발행" 또는 "미발행" 입력\n` +
    `- 자주 나오는 거래처는 [분류규칙]에 추가하세요\n` +
    `- 월간보고서는 자동 생성됩니다\n` +
    `- v3.3: 세금계산서 관리 기능 추가 (NEW!)\n` +
    `- v3.2: HTML 기반 파일 업로더 추가\n` +
    `- v3.1: 파싱/분류 오류 수정 및 데이터 복구\n\n` +
    `🆘 문제 발생시:\n` +
    `1. Setup.gs가 실행되었는지 확인\n` +
    `2. 모든 시트가 생성되었는지 확인\n` +
    `3. [기존 데이터 복구]를 실행해보세요\n` +
    `4. claude.ai에 질문하세요!`;

  ui.alert('도움말', message, ui.ButtonSet.OK);
}

// ========================================
// 12. 세무사 전달용 - 거래상세내역 (전체)
// ========================================

function exportDetailedTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  let exportSheet = ss.getSheetByName('거래상세내역');
  const ui = SpreadsheetApp.getUi();
  
  // 시트 생성
  if (!exportSheet) {
    exportSheet = ss.insertSheet('거래상세내역');
  } else {
    exportSheet.clear();
  }
  
  // 데이터 로드
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  
  // 헤더
  const headers = ['일자', '계좌/카드', '거래처', '계정과목', '출금액', '입금액', '사업/개인', '메모'];
  exportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 스타일
  exportSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // 데이터 작성
  const exportData = data.map(row => [
    formatDateForExport(row[0]),  // 일자
    row[1],  // 카드/계좌
    row[2],  // 거래처
    row[7] || '미분류',  // 최종분류
    row[3] || 0,  // 출금액
    row[4] || 0,  // 입금액
    row[8] || '확인필요',  // 사업/개인
    row[9] || ''  // 메모
  ]);
  
  if (exportData.length > 0) {
    exportSheet.getRange(2, 1, exportData.length, headers.length).setValues(exportData);
    
    // 숫자 포맷
    exportSheet.getRange(2, 5, exportData.length, 2).setNumberFormat('#,##0');
    
    // 조건부 서식 (사업/개인 구분)
    const businessRange = exportSheet.getRange(2, 7, exportData.length, 1);
    
    // 사업 = 파란색
    const businessRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('사업')
      .setBackground('#d0e0e3')
      .setRanges([businessRange])
      .build();
    
    // 개인 = 회색
    const personalRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('개인')
      .setBackground('#f4f4f4')
      .setRanges([businessRange])
      .build();
    
    // 확인필요 = 노란색
    const checkRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('확인필요')
      .setBackground('#fff2cc')
      .setRanges([businessRange])
      .build();
    
    exportSheet.setConditionalFormatRules([businessRule, personalRule, checkRule]);
  }
  
  exportSheet.autoResizeColumns(1, headers.length);
  exportSheet.setFrozenRows(1);
  
  ui.alert(
    '✅ 거래상세내역 생성 완료!',
    `[거래상세내역] 시트에 ${exportData.length}건의 거래가 정리되었습니다.\n\n` +
    `📋 포함 내용:\n` +
    `- 모든 거래 (사업용계좌 + 신용카드)\n` +
    `- 사업/개인 구분 (색상 표시)\n` +
    `- 출금/입금 분리\n\n` +
    `이 시트를 엑셀로 다운로드하여 세무사님께 전달하세요.`,
    ui.ButtonSet.OK
  );
}

// ========================================
// 13. 세무사 전달용 - 계정과목별 집계
// ========================================

function exportCategorySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  let exportSheet = ss.getSheetByName('계정과목별집계');
  const ui = SpreadsheetApp.getUi();
  
  // 시트 생성
  if (!exportSheet) {
    exportSheet = ss.insertSheet('계정과목별집계');
  } else {
    exportSheet.clear();
  }
  
  // 데이터 로드
  const data = txnSheet.getRange(2, 1, txnSheet.getLastRow()-1, 10).getValues();
  
  // 월별/계정과목별 집계
  const summary = {};
  
  data.forEach(row => {
    const date = new Date(row[0]);
    const month = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    const category = row[7] || '미분류';
    const businessType = row[8] || '확인필요';
    const amount = parseFloat(row[3]) || 0;
    
    // 사업 지출만 집계
    if (businessType !== '사업') return;
    
    const key = `${month}|${category}`;
    
    if (!summary[key]) {
      summary[key] = {
        month: month,
        category: category,
        amount: 0,
        count: 0
      };
    }
    
    summary[key].amount += amount;
    summary[key].count++;
  });
  
  // 헤더
  const headers = ['월', '계정과목', '금액', '거래건수', '평균금액'];
  exportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  exportSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // 데이터 작성
  const summaryData = Object.values(summary).map(item => [
    item.month,
    item.category,
    item.amount,
    item.count,
    Math.round(item.amount / item.count)
  ]);
  
  // 정렬 (월별, 금액순)
  summaryData.sort((a, b) => {
    if (a[0] !== b[0]) return b[0].localeCompare(a[0]);
    return b[2] - a[2];
  });
  
  if (summaryData.length > 0) {
    exportSheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
    
    // 숫자 포맷
    exportSheet.getRange(2, 3, summaryData.length, 3).setNumberFormat('#,##0');
    
    // 합계 행 추가
    const totalRow = summaryData.length + 2;
    exportSheet.getRange(totalRow, 1).setValue('총합');
    exportSheet.getRange(totalRow, 2).setValue('');
    exportSheet.getRange(totalRow, 3).setFormula(`=SUM(C2:C${totalRow-1})`);
    exportSheet.getRange(totalRow, 4).setFormula(`=SUM(D2:D${totalRow-1})`);
    exportSheet.getRange(totalRow, 5).setValue('');
    
    exportSheet.getRange(totalRow, 1, 1, 5)
      .setFontWeight('bold')
      .setBackground('#f4f4f4')
      .setNumberFormat('#,##0');
  }
  
  exportSheet.autoResizeColumns(1, headers.length);
  exportSheet.setFrozenRows(1);
  
  ui.alert(
    '✅ 계정과목별 집계 완료!',
    `[계정과목별집계] 시트에 ${summaryData.length}개 항목이 생성되었습니다.\n\n` +
    `📋 포함 내용:\n` +
    `- 사업 지출만 집계\n` +
    `- 월별/계정과목별 분류\n` +
    `- 거래건수 및 평균금액\n\n` +
    `이 시트를 엑셀로 다운로드하여 세무사님께 전달하세요.`,
    ui.ButtonSet.OK
  );
}

// ========================================
// 14. 날짜 포맷 변환 함수
// ========================================

function formatDateForExport(date) {
  if (!date) return '';

  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

// ========================================
// 15. 세금계산서 관리 기능
// ========================================

/**
 * 입금내역 보기
 */
function showIncomeTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  let incomeSheet = ss.getSheetByName('입금내역');
  const ui = SpreadsheetApp.getUi();

  if (!txnSheet) {
    ui.alert('오류', '[거래내역통합] 시트를 찾을 수 없습니다!', ui.ButtonSet.OK);
    return;
  }

  // 입금내역 시트 생성
  if (!incomeSheet) {
    incomeSheet = ss.insertSheet('입금내역');
  } else {
    incomeSheet.clear();
  }

  // 헤더
  const headers = ['일자', '거래처', '입금액', '세금계산서', '메모'];
  incomeSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  incomeSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 데이터 로드
  const lastRow = txnSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('데이터가 없습니다!');
    return;
  }

  const data = txnSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // 입금내역만 필터링 (E열: 입금액이 0보다 큰 것)
  const incomeData = [];
  data.forEach(row => {
    const creditAmount = parseFloat(row[4]) || 0;  // E열: 입금액
    if (creditAmount > 0) {
      incomeData.push([
        formatDateForExport(row[0]),  // 일자
        row[2] || '',  // 거래처
        creditAmount,  // 입금액
        row[10] || '',  // K열: 세금계산서
        row[9] || ''   // J열: 메모
      ]);
    }
  });

  if (incomeData.length === 0) {
    ui.alert('입금내역이 없습니다!');
    return;
  }

  // 데이터 작성 (날짜 최신순 정렬)
  incomeData.sort((a, b) => b[0].localeCompare(a[0]));
  incomeSheet.getRange(2, 1, incomeData.length, headers.length).setValues(incomeData);

  // 숫자 포맷
  incomeSheet.getRange(2, 3, incomeData.length, 1).setNumberFormat('#,##0');

  // 조건부 서식 (세금계산서 미발행 강조)
  const taxInvoiceRange = incomeSheet.getRange(2, 4, incomeData.length, 1);

  const issuedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('발행')
    .setBackground('#d1fae5')
    .setFontColor('#065f46')
    .setRanges([taxInvoiceRange])
    .build();

  const notIssuedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('미발행')
    .setBackground('#fee2e2')
    .setFontColor('#991b1b')
    .setRanges([taxInvoiceRange])
    .build();

  const emptyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([taxInvoiceRange])
    .build();

  incomeSheet.setConditionalFormatRules([issuedRule, notIssuedRule, emptyRule]);

  // 열 너비 자동 조정
  incomeSheet.autoResizeColumns(1, headers.length);
  incomeSheet.setFrozenRows(1);

  // 통계 계산
  const totalIncome = incomeData.reduce((sum, row) => sum + row[2], 0);
  const issuedCount = incomeData.filter(row => row[3] === '발행').length;
  const notIssuedCount = incomeData.filter(row => row[3] === '미발행' || row[3] === '').length;

  ui.alert(
    '✅ 입금내역 조회 완료!',
    `[입금내역] 시트에 ${incomeData.length}건이 생성되었습니다.\n\n` +
    `💰 총 입금액: ${totalIncome.toLocaleString()}원\n` +
    `✅ 세금계산서 발행: ${issuedCount}건\n` +
    `⚠️ 미발행/확인필요: ${notIssuedCount}건\n\n` +
    `세금계산서 열에 "발행" 또는 "미발행"을 직접 입력하세요.`,
    ui.ButtonSet.OK
  );
}

/**
 * 세금계산서 미발행 내역 검사
 */
function checkTaxInvoiceStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  const ui = SpreadsheetApp.getUi();

  if (!txnSheet) {
    ui.alert('오류', '[거래내역통합] 시트를 찾을 수 없습니다!', ui.ButtonSet.OK);
    return;
  }

  const lastRow = txnSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('데이터가 없습니다!');
    return;
  }

  const data = txnSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // 입금내역 중 세금계산서 미발행 항목 찾기
  const notIssued = [];
  data.forEach((row, index) => {
    const creditAmount = parseFloat(row[4]) || 0;  // E열: 입금액
    const taxInvoice = row[10] || '';  // K열: 세금계산서

    if (creditAmount > 0 && taxInvoice !== '발행') {
      notIssued.push({
        rowNum: index + 2,
        date: formatDateForExport(row[0]),
        merchant: row[2],
        amount: creditAmount,
        status: taxInvoice || '미입력'
      });
    }
  });

  if (notIssued.length === 0) {
    ui.alert('✅ 모든 입금내역에 세금계산서가 발행되었습니다! 🎉', ui.ButtonSet.OK);
    return;
  }

  // 거래처별 집계
  const merchantGroups = {};
  notIssued.forEach(item => {
    if (!merchantGroups[item.merchant]) {
      merchantGroups[item.merchant] = { count: 0, total: 0 };
    }
    merchantGroups[item.merchant].count++;
    merchantGroups[item.merchant].total += item.amount;
  });

  // 금액순 정렬
  const sorted = Object.entries(merchantGroups)
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 15);

  let message = `⚠️ 세금계산서 미발행 내역: ${notIssued.length}건\n\n`;
  message += `💰 총 미발행 금액: ${notIssued.reduce((sum, item) => sum + item.amount, 0).toLocaleString()}원\n\n`;
  message += `📋 거래처별 현황 (TOP 15):\n\n`;

  sorted.forEach(([merchant, data], index) => {
    message += `${index + 1}. ${merchant}\n`;
    message += `   ${data.count}건, ${data.total.toLocaleString()}원\n`;
  });

  message += `\n💡 [입금내역 보기]에서 세금계산서 상태를 확인하세요!`;

  ui.alert('세금계산서 미발행 검사', message, ui.ButtonSet.OK);
}

/**
 * 월별 세금계산서 대조 보고서 생성
 */
function generateTaxInvoiceReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName('거래내역통합');
  let reportSheet = ss.getSheetByName('세금계산서대조');
  const ui = SpreadsheetApp.getUi();

  if (!txnSheet) {
    ui.alert('오류', '[거래내역통합] 시트를 찾을 수 없습니다!', ui.ButtonSet.OK);
    return;
  }

  // 보고서 시트 생성
  if (!reportSheet) {
    reportSheet = ss.insertSheet('세금계산서대조');
  } else {
    reportSheet.clear();
  }

  // 헤더
  const headers = ['월', '총 입금액', '발행완료 금액', '미발행 금액', '발행건수', '미발행건수', '발행률'];
  reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  reportSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f59e0b')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 데이터 로드
  const lastRow = txnSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('데이터가 없습니다!');
    return;
  }

  const data = txnSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // 월별 집계
  const monthlyData = {};

  data.forEach(row => {
    const creditAmount = parseFloat(row[4]) || 0;  // E열: 입금액
    if (creditAmount <= 0) return;  // 입금이 아니면 스킵

    const date = new Date(row[0]);
    const month = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    const taxInvoice = row[10] || '';  // K열: 세금계산서

    if (!monthlyData[month]) {
      monthlyData[month] = {
        totalIncome: 0,
        issuedAmount: 0,
        notIssuedAmount: 0,
        issuedCount: 0,
        notIssuedCount: 0
      };
    }

    monthlyData[month].totalIncome += creditAmount;

    if (taxInvoice === '발행') {
      monthlyData[month].issuedAmount += creditAmount;
      monthlyData[month].issuedCount++;
    } else {
      monthlyData[month].notIssuedAmount += creditAmount;
      monthlyData[month].notIssuedCount++;
    }
  });

  // 보고서 데이터 생성
  const reportData = [];
  Object.entries(monthlyData).forEach(([month, data]) => {
    const totalCount = data.issuedCount + data.notIssuedCount;
    const issueRate = totalCount > 0 ? (data.issuedCount / totalCount * 100).toFixed(1) + '%' : '0%';

    reportData.push([
      month,
      data.totalIncome,
      data.issuedAmount,
      data.notIssuedAmount,
      data.issuedCount,
      data.notIssuedCount,
      issueRate
    ]);
  });

  // 월별 역순 정렬
  reportData.sort((a, b) => b[0].localeCompare(a[0]));

  if (reportData.length > 0) {
    reportSheet.getRange(2, 1, reportData.length, headers.length).setValues(reportData);

    // 숫자 포맷
    reportSheet.getRange(2, 2, reportData.length, 3).setNumberFormat('#,##0');

    // 조건부 서식 (미발행 금액이 0이 아니면 강조)
    const notIssuedRange = reportSheet.getRange(2, 4, reportData.length, 1);
    const warningRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#fee2e2')
      .setFontColor('#991b1b')
      .setRanges([notIssuedRange])
      .build();

    reportSheet.setConditionalFormatRules([warningRule]);

    // 합계 행 추가
    const totalRow = reportData.length + 2;
    reportSheet.getRange(totalRow, 1).setValue('총합');
    reportSheet.getRange(totalRow, 2).setFormula(`=SUM(B2:B${totalRow - 1})`);
    reportSheet.getRange(totalRow, 3).setFormula(`=SUM(C2:C${totalRow - 1})`);
    reportSheet.getRange(totalRow, 4).setFormula(`=SUM(D2:D${totalRow - 1})`);
    reportSheet.getRange(totalRow, 5).setFormula(`=SUM(E2:E${totalRow - 1})`);
    reportSheet.getRange(totalRow, 6).setFormula(`=SUM(F2:F${totalRow - 1})`);
    reportSheet.getRange(totalRow, 7).setFormula(`=IF(E${totalRow}+F${totalRow}>0,TEXT(E${totalRow}/(E${totalRow}+F${totalRow}),"0.0%"),"")`);

    reportSheet.getRange(totalRow, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f4f4f4')
      .setNumberFormat('#,##0');
  }

  // 열 너비 자동 조정
  reportSheet.autoResizeColumns(1, headers.length);
  reportSheet.setFrozenRows(1);

  ui.alert(
    '✅ 세금계산서 대조 보고서 생성 완료!',
    `[세금계산서대조] 시트에 ${reportData.length}개월 데이터가 생성되었습니다.\n\n` +
    `📋 포함 내용:\n` +
    `- 월별 총 입금액\n` +
    `- 세금계산서 발행/미발행 금액 및 건수\n` +
    `- 발행률\n\n` +
    `⚠️ 미발행 금액이 빨간색으로 강조됩니다.`,
    ui.ButtonSet.OK
  );
}

