/**
 * 세금계산서 대조 시스템 v1.0
 * 기업은행 거래내역과 홈택스 세금계산서 발행내역을 대조하는 독립 시스템
 */

/**
 * 메인 설정 함수
 */
function setupTaxInvoiceChecker() {
  try {
    const ss = SpreadsheetApp.getActive();

    if (!ss) {
      throw new Error('스프레드시트를 찾을 수 없습니다. Google Sheets에서 메뉴를 통해 실행해주세요.');
    }

    const ui = SpreadsheetApp.getUi();

    const response = ui.alert(
      '🏦 세금계산서 대조 시스템 설정',
      '다음 시트들이 자동으로 생성됩니다:\n\n' +
      '1. 기업은행거래내역 (1년치 거래 입력)\n' +
      '2. 세금계산서발행내역 (홈택스 데이터)\n' +
      '3. 대조결과 (자동 생성)\n\n' +
      '계속하시겠습니까?',
      ui.ButtonSet.YES_NO
    );

    if (response != ui.Button.YES) {
      ui.alert('취소되었습니다.');
      return;
    }

    // 진행 상황 표시
    SpreadsheetApp.getActive().toast('설정 시작...', '진행중', 3);

    // 1. 기업은행거래내역 시트
    SpreadsheetApp.getActive().toast('1/3: 기업은행거래내역 시트 생성 중...', '진행중', 2);
    createIBKTransactionSheet(ss);

    // 2. 세금계산서발행내역 시트
    SpreadsheetApp.getActive().toast('2/3: 세금계산서발행내역 시트 생성 중...', '진행중', 2);
    createTaxInvoiceSheet(ss);

    // 3. 대조결과 시트
    SpreadsheetApp.getActive().toast('3/3: 대조결과 시트 생성 중...', '진행중', 2);
    createComparisonResultSheet(ss);

    // 완료
    SpreadsheetApp.getActive().toast('설정 완료!', '완료', 3);

    ui.alert(
      '✅ 설정 완료!',
      '모든 시트가 생성되었습니다.\n\n' +
      '📋 생성된 시트:\n' +
      '- 기업은행거래내역\n' +
      '- 세금계산서발행내역\n' +
      '- 대조결과\n\n' +
      '다음 단계:\n' +
      '1. [기업은행거래내역] 시트에 1년치 거래 데이터 붙여넣기\n' +
      '2. [세금계산서발행내역] 시트에 홈택스 데이터 붙여넣기\n' +
      '3. [🔍 세금계산서 대조] > [▶️ 전체 대조 실행] 클릭\n\n' +
      '설정이 완료되었습니다! 🎉',
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log('설정 오류: ' + error.toString());
    Logger.log('스택 트레이스: ' + error.stack);

    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '오류 발생',
      '오류: ' + error.message + '\n\n' +
      '해결 방법:\n' +
      '1. Google Sheets에서 실행하고 있는지 확인\n' +
      '2. 메뉴를 통해 실행\n' +
      '3. Apps Script 에디터에서 직접 실행하지 마세요',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 기업은행거래내역 시트 생성
 */
function createIBKTransactionSheet(ss) {
  if (!ss) throw new Error('스프레드시트 객체가 없습니다.');

  let sheet = ss.getSheetByName('기업은행거래내역');

  if (!sheet) {
    sheet = ss.insertSheet('기업은행거래내역');
    Logger.log('기업은행거래내역 시트 생성됨');
  } else {
    // 기존 시트가 있으면 헤더만 재설정
    sheet.clear();
    Logger.log('기업은행거래내역 시트 클리어됨');
  }

  // 헤더
  const headers = [['일자', '거래처', '금액', '입금/출금', '메모']];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  // 스타일
  sheet.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 열 너비
  sheet.setColumnWidth(1, 100);  // 일자
  sheet.setColumnWidth(2, 250);  // 거래처
  sheet.setColumnWidth(3, 120);  // 금액
  sheet.setColumnWidth(4, 100);  // 입금/출금
  sheet.setColumnWidth(5, 300);  // 메모

  // 고정
  sheet.setFrozenRows(1);

  // 안내 메시지
  sheet.getRange(3, 1, 1, 5).merge();
  sheet.getRange(3, 1)
    .setValue('👆 기업은행 거래내역 1년치를 여기에 붙여넣으세요 (일자, 거래처, 금액, 입금/출금, 메모)')
    .setFontColor('#666666')
    .setFontStyle('italic');

  // 샘플 데이터
  sheet.getRange(5, 1, 3, 5).setValues([
    ['2024-01-15', '한메디', 71500, '출금', '의료용품'],
    ['2024-02-20', '행림원외탕전', 150000, '출금', '한약재'],
    ['2024-03-10', '서울병원', 200000, '입금', '진료수입']
  ]);

  Logger.log('기업은행거래내역 시트 설정 완료');
}

/**
 * 세금계산서발행내역 시트 생성
 */
function createTaxInvoiceSheet(ss) {
  if (!ss) throw new Error('스프레드시트 객체가 없습니다.');

  let sheet = ss.getSheetByName('세금계산서발행내역');

  if (!sheet) {
    sheet = ss.insertSheet('세금계산서발행내역');
    Logger.log('세금계산서발행내역 시트 생성됨');
  } else {
    sheet.clear();
    Logger.log('세금계산서발행내역 시트 클리어됨');
  }

  // 헤더
  const headers = [['발행일자', '거래처명', '공급가액', '세액', '합계금액', '승인번호', '메모']];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  // 스타일
  sheet.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#0f9d58')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 열 너비
  sheet.setColumnWidth(1, 100);  // 발행일자
  sheet.setColumnWidth(2, 250);  // 거래처명
  sheet.setColumnWidth(3, 120);  // 공급가액
  sheet.setColumnWidth(4, 100);  // 세액
  sheet.setColumnWidth(5, 120);  // 합계금액
  sheet.setColumnWidth(6, 150);  // 승인번호
  sheet.setColumnWidth(7, 200);  // 메모

  // 고정
  sheet.setFrozenRows(1);

  // 안내 메시지
  sheet.getRange(3, 1, 1, 7).merge();
  sheet.getRange(3, 1)
    .setValue('👆 홈택스에서 다운로드한 세금계산서 발행내역 1년치를 여기에 붙여넣으세요')
    .setFontColor('#666666')
    .setFontStyle('italic');

  // 샘플 데이터
  sheet.getRange(5, 1, 2, 7).setValues([
    ['2024-01-15', '한메디', 65000, 6500, 71500, 'INV-2024-001', ''],
    ['2024-02-20', '행림', 136364, 13636, 150000, 'INV-2024-002', '']
  ]);

  Logger.log('세금계산서발행내역 시트 설정 완료');
}

/**
 * 대조결과 시트 생성
 */
function createComparisonResultSheet(ss) {
  if (!ss) throw new Error('스프레드시트 객체가 없습니다.');

  let sheet = ss.getSheetByName('대조결과');

  if (!sheet) {
    sheet = ss.insertSheet('대조결과');
    Logger.log('대조결과 시트 생성됨');
  } else {
    sheet.clear();
    Logger.log('대조결과 시트 클리어됨');
  }

  // 헤더
  const headers = [['일자', '거래처', '금액', '입금/출금', '매칭상태', '비고']];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  // 스타일
  sheet.getRange(1, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 열 너비
  sheet.setColumnWidth(1, 100);  // 일자
  sheet.setColumnWidth(2, 250);  // 거래처
  sheet.setColumnWidth(3, 120);  // 금액
  sheet.setColumnWidth(4, 100);  // 입금/출금
  sheet.setColumnWidth(5, 150);  // 매칭상태
  sheet.setColumnWidth(6, 400);  // 비고

  // 고정
  sheet.setFrozenRows(1);

  // 안내 메시지
  sheet.getRange(3, 1, 1, 6).merge();
  sheet.getRange(3, 1)
    .setValue('👆 [🔍 세금계산서 대조] > [▶️ 전체 대조 실행]을 클릭하면 여기에 결과가 표시됩니다')
    .setFontColor('#666666')
    .setFontStyle('italic');

  Logger.log('대조결과 시트 설정 완료');
}
