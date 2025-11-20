/**
 * 세금계산서 매칭 로직 테스트 스크립트
 */

function testTaxMatchingLogic() {
  Logger.log('=== 테스트 시작 ===');
  
  // 1. Fuzzy Matching 테스트
  testFuzzyMatching();
  
  // 2. Subset Sum 테스트
  testSubsetSum();
  
  Logger.log('=== 테스트 완료 ===');
}

function testFuzzyMatching() {
  Logger.log('\n[Fuzzy Matching 테스트]');
  
  const cases = [
    { s1: '(주)아현재', s2: '아현재한의원', expected: true },
    { s1: '구글코리아 유한회사', s2: '구글코리아', expected: true },
    { s1: '스타벅스커피코리아', s2: '스타벅스', expected: true },
    { s1: '전혀다른이름', s2: '아현재', expected: false }
  ];
  
  cases.forEach(c => {
    const n1 = normalizeName(c.s1);
    const n2 = normalizeName(c.s2);
    const similarity = calculateSimilarity(n1, n2);
    const passed = (similarity > 0.4) === c.expected; // 임계값 0.4 가정
    
    Logger.log(`"${c.s1}" vs "${c.s2}" -> 유사도: ${similarity.toFixed(2)} [${passed ? 'PASS' : 'FAIL'}]`);
  });
}

function testSubsetSum() {
  Logger.log('\n[Subset Sum 테스트]');
  
  // Mock Transactions
  const transactions = [
    { id: 1, amount: 300000 },
    { id: 2, amount: 700000 },
    { id: 3, amount: 500000 },
    { id: 4, amount: 100000 }
  ];
  
  // Case 1: 100만원 찾기 (30 + 70)
  const target1 = 1000000;
  const result1 = findSubsetSum(transactions, target1, 5);
  
  if (result1 && result1.sum === target1) {
    Logger.log(`Target ${target1}: Found match! Items: ${result1.txns.map(t => t.amount).join('+')} [PASS]`);
  } else {
    Logger.log(`Target ${target1}: No match found [FAIL]`);
  }
  
  // Case 2: 60만원 찾기 (50 + 10)
  const target2 = 600000;
  const result2 = findSubsetSum(transactions, target2, 5);
  
  if (result2 && result2.sum === target2) {
    Logger.log(`Target ${target2}: Found match! Items: ${result2.txns.map(t => t.amount).join('+')} [PASS]`);
  } else {
    Logger.log(`Target ${target2}: No match found [FAIL]`);
  }
  
  // Case 3: 불가능한 금액 (99999)
  const target3 = 99999;
  const result3 = findSubsetSum(transactions, target3, 5);
  
  if (!result3) {
    Logger.log(`Target ${target3}: No match found (Expected) [PASS]`);
  } else {
    Logger.log(`Target ${target3}: Unexpected match found [FAIL]`);
  }
}

// --- 의존성 함수 복사 (테스트용) ---
function normalizeName(name) {
  if (!name) return '';
  return name.toString()
    .replace(/\(주\)/g, '')
    .replace(/주식회사/g, '')
    .replace(/\s+/g, '')
    .replace(/[^\w가-힣]/g, '')
    .toLowerCase();
}

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

function findSubsetSum(transactions, targetAmount, maxDepth) {
  function search(index, currentSum, currentPath) {
    if (Math.abs(currentSum - targetAmount) < 1) {
      return { sum: currentSum, txns: currentPath };
    }
    
    if (currentSum > targetAmount + 1 || currentPath.length >= maxDepth || index >= transactions.length) {
      return null;
    }
    
    const res1 = search(index + 1, currentSum + transactions[index].amount, [...currentPath, transactions[index]]);
    if (res1) return res1;
    
    const res2 = search(index + 1, currentSum, currentPath);
    if (res2) return res2;
    
    return null;
  }
  
  return search(0, 0, []);
}
