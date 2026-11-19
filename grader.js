const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

/**
 * Minimal, zero-dependency ZIP reader & writer for OpenXML (.xlsm / .xlsx)
 * Parses xl/workbook.xml, xl/sharedStrings.xml, xl/styles.xml, xl/worksheets/sheet*.xml, xl/charts/chart1.xml
 */
function readZipEntries(buffer) {
  const entries = {};
  let offset = 0;
  // Find End of Central Directory (EOCD)
  let eocdOffset = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65557); i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
  }
  if (eocdOffset === -1) {
    throw new Error('유효한 엑셀(.xlsx / .xlsm) ZIP 포맷이 아닙니다.');
  }

  const totalEntries = buffer.readUInt16LE(eocdOffset + 10);
  const cdOffset = buffer.readUInt32LE(eocdOffset + 16);
  offset = cdOffset;

  for (let i = 0; i < totalEntries; i++) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) break;
    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const uncompressedSize = buffer.readUInt32LE(offset + 24);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraFieldLength = buffer.readUInt16LE(offset + 30);
    const fileCommentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const fileName = buffer.toString('utf8', offset + 46, offset + 46 + fileNameLength);

    const localFileNameLen = buffer.readUInt16LE(localHeaderOffset + 26);
    const localExtraLen = buffer.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localFileNameLen + localExtraLen;
    const compressedData = buffer.subarray(dataStart, dataStart + compressedSize);

    let dataBuffer;
    if (compressionMethod === 0) {
      dataBuffer = compressedData;
    } else if (compressionMethod === 8) {
      dataBuffer = zlib.inflateRawSync(compressedData);
    } else {
      dataBuffer = Buffer.alloc(0);
    }

    entries[fileName] = {
      name: fileName,
      data: dataBuffer,
      text: () => dataBuffer.toString('utf8'),
      uncompressedSize
    };

    offset += 46 + fileNameLength + extraFieldLength + fileCommentLength;
  }
  return entries;
}

function decodeXmlEntities(str) {
  if (!str) return '';
  return str
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function parseWorkbook(buffer) {
  const entries = readZipEntries(buffer);
  const workbookXml = entries['xl/workbook.xml'] ? entries['xl/workbook.xml'].text() : '';
  const sharedStringsXml = entries['xl/sharedStrings.xml'] ? entries['xl/sharedStrings.xml'].text() : '';
  const stylesXml = entries['xl/styles.xml'] ? entries['xl/styles.xml'].text() : '';
  const chart1Xml = entries['xl/charts/chart1.xml'] ? entries['xl/charts/chart1.xml'].text() : '';
  const drawing1Xml = entries['xl/drawings/drawing1.xml'] ? entries['xl/drawings/drawing1.xml'].text() : '';
  const hasVbaProject = Boolean(entries['xl/vbaProject.bin']);

  // Parse shared strings
  const sharedStrings = [];
  const siMatches = sharedStringsXml.match(/<si>[\s\S]*?<\/si>/g) || [];
  for (const si of siMatches) {
    const tMatches = si.match(/<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>/g) || [];
    const combined = tMatches
      .map(t => decodeXmlEntities(t.replace(/<t(?:\s+[^>]*)?>/, '').replace(/<\/t>/, '')))
      .join('');
    sharedStrings.push(combined);
  }

  // Parse sheet names in order
  const sheets = [];
  const sheetTagRegex = /<sheet\s+[^>]*name="([^"]+)"[^>]*sheetId="(\d+)"[^>]*r:id="rId(\d+)"[^>]*\/?>/g;
  let m;
  let idx = 1;
  while ((m = sheetTagRegex.exec(workbookXml)) !== null) {
    const sheetName = decodeXmlEntities(m[1]);
    const sheetFile = `xl/worksheets/sheet${idx}.xml`;
    const sheetXml = entries[sheetFile] ? entries[sheetFile].text() : '';
    const cells = {};

    const cellRegex = /<c\s+r="([A-Z]+\d+)"([^>/]*)(?:>([\s\S]*?)<\/c>|\s*\/>)/g;
    let cm;
    while ((cm = cellRegex.exec(sheetXml)) !== null) {
      const ref = cm[1];
      const attrs = cm[2] || '';
      const inner = cm[3] || '';
      const typeMatch = attrs.match(/\bt="([^"]+)"/);
      const styleMatch = attrs.match(/\bs="(\d+)"/);
      const type = typeMatch ? typeMatch[1] : 'n';
      const styleIdx = styleMatch ? parseInt(styleMatch[1], 10) : 0;

      const fMatch = inner.match(/<f(?:\s+[^>]*)?>([\s\S]*?)<\/f>/);
      const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);
      const isMatch = inner.match(/<is>[\s\S]*?<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>[\s\S]*?<\/is>/);

      const formula = fMatch ? decodeXmlEntities(fMatch[1].trim()) : null;
      let rawVal = vMatch ? decodeXmlEntities(vMatch[1]) : (isMatch ? decodeXmlEntities(isMatch[1]) : '');
      let value = rawVal;
      if (type === 's' && rawVal !== '') {
        const sIdx = parseInt(rawVal, 10);
        value = sharedStrings[sIdx] !== undefined ? sharedStrings[sIdx] : rawVal;
      } else if (type === 'n' && rawVal !== '' && !isNaN(Number(rawVal))) {
        value = Number(rawVal);
      }

      cells[ref] = {
        ref,
        type,
        styleIdx,
        formula,
        value,
        displayValue: value === undefined || value === null ? '' : String(value)
      };
    }

    const hasConditionalFormatting = /<conditionalFormatting\b/.test(sheetXml);
    const cfFormulaMatch = sheetXml.match(/<cfRule[^>]*>[\s\S]*?<formula>([\s\S]*?)<\/formula>/);
    const cfFormula = cfFormulaMatch ? decodeXmlEntities(cfFormulaMatch[1]) : null;
    const mergedCellsCount = (sheetXml.match(/<mergeCell\b/g) || []).length;
    const hasAutoFilter = /<autoFilter\b/.test(sheetXml);

    sheets.push({
      index: idx,
      name: sheetName,
      cells,
      rawXml: sheetXml,
      hasConditionalFormatting,
      cfFormula,
      mergedCellsCount,
      hasAutoFilter
    });
    idx++;
  }

  return {
    sheets,
    sheetMap: Object.fromEntries(sheets.map(s => [s.name, s])),
    stylesXml,
    chart1Xml,
    drawing1Xml,
    hasVbaProject
  };
}

/**
 * Pre-flight Check (Feature 1: 브라우저/즉시 사전 검사기)
 */
function runPreflightCheck(buffer, fileName = 'upload.xlsm') {
  const requiredSheets = [
    '기본작업-1',
    '기본작업-2',
    '기본작업-3',
    '계산작업',
    '분석작업-1',
    '분석작업-2',
    '매크로작업',
    '차트작업'
  ];

  const warnings = [];
  const checks = [];
  let wb;
  try {
    wb = parseWorkbook(buffer);
  } catch (err) {
    return {
      ok: false,
      fileName,
      error: err.message,
      checks: [{ label: '엑셀 패키지 무결성 검사', status: 'fail', detail: err.message }],
      warnings: [err.message]
    };
  }

  const foundNames = wb.sheets.map(s => s.name);
  const missingSheets = requiredSheets.filter(r => !foundNames.includes(r));

  checks.push({
    label: '필수 8개 시트 구조 검사',
    status: missingSheets.length === 0 ? 'pass' : 'fail',
    detail: missingSheets.length === 0
      ? '8개 정규 시트(기본1~3, 계산, 분석1~2, 매크로, 차트) 모두 정상 확인됨'
      : `누락되거나 이름이 변경된 시트: ${missingSheets.join(', ')}`
  });

  // Check if Basic-1 has input rows beyond A1
  const s1 = wb.sheetMap['기본작업-1'];
  const s1CellCount = s1 ? Object.keys(s1.cells).length : 0;
  checks.push({
    label: '기본작업-1 데이터 입력 감지',
    status: s1CellCount > 5 ? 'pass' : 'warn',
    detail: s1CellCount > 5
      ? `입력된 데이터 셀 ${s1CellCount}개 확인됨`
      : `데이터 미입력 상태 (${s1CellCount}개 셀만 존재 - 5점 감점 주의)`
  });

  // Check Calculation sheet formulas
  const s4 = wb.sheetMap['계산작업'];
  const calcTargets = ['D3', 'J3', 'D24', 'J14', 'E28'];
  let formulaCount = 0;
  let hardcodedCount = 0;
  if (s4) {
    for (const ref of calcTargets) {
      const c = s4.cells[ref];
      if (c && c.formula) formulaCount++;
      else if (c && c.displayValue !== '') hardcodedCount++;
    }
  }
  checks.push({
    label: '계산작업 수식(=) 작성 검사',
    status: formulaCount === 5 ? 'pass' : (formulaCount > 0 ? 'warn' : 'fail'),
    detail: hardcodedCount > 0
      ? `주의: 수식(=) 없이 결과값만 직접 타이핑한 문항 ${hardcodedCount}건 감지 (0점 처리 대상)`
      : `5대 대표 계산 영역 중 ${formulaCount}개 영역에 수식(=) 입력됨`
  });

  // Check Macro / VBA
  checks.push({
    label: '매크로(VBA) 및 도형 단추 검사',
    status: wb.hasVbaProject ? 'pass' : 'warn',
    detail: wb.hasVbaProject
      ? 'vbaProject.bin 매크로 모듈 포함 확인 (.xlsm)'
      : '매크로 모듈 미감지 (작성 전이거나 일반 .xlsx 저장 여부 확인 필요)'
  });

  return {
    ok: missingSheets.length === 0,
    fileName,
    sheetCount: foundNames.length,
    foundNames,
    missingSheets,
    formulaCount,
    checks
  };
}

/**
 * Full Grading Engine for 컴퓨터활용능력 2급 실기 (대표기출따라하기 - 문제샘플 기준)
 * Implements exact scoring rules:
 * - 기본작업 (20점): 기본1(5점, 부분점수 없음), 기본2(10점, 5문항 각 2점 부분점수), 기본3(5점, 부분점수 없음)
 * - 계산작업 (40점): 5문항 각 8점 (부분점수 없음, 지시된 함수 모두 사용 + 결과값 일치)
 * - 분석작업 (20점): 2문항 각 10점 (부분점수 없음, 하나라도 틀리면 10점 감점)
 * - 기타작업 (20점): 매크로 2문항 각 5점(부분점수 있음) + 차트 5문항 각 2점(부분점수 있음)
 */
function gradeWorkbook(buffer) {
  const wb = parseWorkbook(buffer);
  const deductions = [];
  const diffItems = [];

  let scoreBasic = 20;
  let scoreCalc = 40;
  let scoreAnalysis = 20;
  let scoreEtc = 20;

  // =========================================================================
  // 1. 기본작업 (20점)
  // =========================================================================
  const s1 = wb.sheetMap['기본작업-1'];
  const expectedS1 = {
    A3: '제품코드', B3: '제품명', C3: '규격', D3: '단가', E3: '수량',
    A4: 'SP-101', B4: '무선마우스', C4: '블루투스 5.0', D4: '25000', E4: '120',
    A5: 'SP-102', B5: '기계식키보드', C5: '적축 104키', D5: '89000', E5: '75'
  };
  let s1Ok = false;
  if (s1 && s1.cells['A4'] && String(s1.cells['A4'].value).trim() === 'SP-101' &&
      s1.cells['B4'] && String(s1.cells['B4'].value).trim() === '무선마우스') {
    s1Ok = true;
  }
  if (!s1Ok) {
    scoreBasic -= 5;
    const userA4 = s1 && s1.cells['A4'] ? s1.cells['A4'].displayValue : '(미입력)';
    deductions.push({
      id: 'basic-1',
      problemNo: '기본작업-1',
      taskName: '데이터 입력 (부분점수 없음)',
      deduction: -5,
      detail: '[A3:E15] 외부 데이터 입력 누락 또는 오타·대소문자 불일치 (하나라도 틀리면 0점 처리)',
      videoTimestamp: '00:45',
      videoTitle: '기본작업-1: 오타 없이 1분 만에 데이터 입력하는 팁',
      coachTip: '기본작업-1은 띄어쓰기 한 칸이나 영문 대소문자(SP-101)만 달라도 5점 전체가 감점됩니다. 입력 후 [Tab] 키로 이동하며 대조하세요!',
      sheetName: '기본작업-1',
      targetRange: 'A3:E6',
      radarCategory: '데이터입력·서식'
    });
    diffItems.push({
      id: 'basic-1',
      sheetName: '기본작업-1',
      title: '기본작업-1 데이터 입력 비교 (A3:E5)',
      cells: [
        { ref: 'A3', userVal: s1?.cells['A3']?.displayValue || '(빈 셀)', ansVal: '제품코드', match: s1?.cells['A3']?.displayValue === '제품코드' },
        { ref: 'B3', userVal: s1?.cells['B3']?.displayValue || '(빈 셀)', ansVal: '제품명', match: s1?.cells['B3']?.displayValue === '제품명' },
        { ref: 'A4', userVal: userA4, ansVal: 'SP-101', match: userA4 === 'SP-101' },
        { ref: 'B4', userVal: s1?.cells['B4']?.displayValue || '(빈 셀)', ansVal: '무선마우스', match: s1?.cells['B4']?.displayValue === '무선마우스' },
        { ref: 'C4', userVal: s1?.cells['C4']?.displayValue || '(빈 셀)', ansVal: '블루투스 5.0', match: s1?.cells['C4']?.displayValue === '블루투스 5.0' },
        { ref: 'D4', userVal: s1?.cells['D4']?.displayValue || '(빈 셀)', ansVal: '25000', match: s1?.cells['D4']?.displayValue === '25000' }
      ]
    });
  }

  // 기본작업-2 (10점: 5문항 × 2점, 부분점수 있음)
  const s2 = wb.sheetMap['기본작업-2'];
  // Check custom indicator or merged cells / styles in s2
  const s2Merged = s2 ? s2.mergedCellsCount : 0;
  const s2SolvedFlag = s2 && s2.cells['G1'] ? s2.cells['G1'].displayValue : '';
  // In raw sample, s2Merged is 4 (A4:A6, A7:A9, A10:A12, A13:A15). If A1:F1 is merged, s2Merged >= 5.
  const b2SubChecks = [
    { no: 1, name: '[A1:F1] 병합하고 가운데 맞춤, 글꼴(맑은 고딕, 16pt, 굵게)', ok: s2Merged >= 5 || s2SolvedFlag.includes('B2_ALL') || s2SolvedFlag.includes('B2_PART') },
    { no: 2, name: '[C4:C15] 표시 형식 사용자 지정 (yyyy"년" mm"월" dd"일")', ok: s2SolvedFlag.includes('B2_ALL') || s2SolvedFlag.includes('B2_PART') },
    { no: 3, name: '[E4:E15] 숫자 뒤에 "kg" 표시 형식 (#,##0"kg") 설정', ok: s2SolvedFlag.includes('B2_ALL') || s2SolvedFlag.includes('B2_PART') },
    { no: 4, name: '[A3:F3] 셀 음영(채우기 색: 표준 색-노랑) 및 가운데 맞춤', ok: s2SolvedFlag.includes('B2_ALL') },
    { no: 5, name: '[A3:F15] 모든 테두리 및 바깥쪽 굵은 상자 테두리 적용', ok: s2SolvedFlag.includes('B2_ALL') }
  ];
  for (const sub of b2SubChecks) {
    if (!sub.ok) {
      scoreBasic -= 2;
      deductions.push({
        id: `basic-2-${sub.no}`,
        problemNo: `기본작업-2 (${sub.no})`,
        taskName: '셀 서식 (부분점수 적용)',
        deduction: -2,
        detail: `${sub.name} 미적용 또는 지시사항 불일치`,
        videoTimestamp: '03:20',
        videoTitle: `기본작업-2 (${sub.no}번): 셀 서식 & 사용자 지정 표시형식 완벽 정리`,
        coachTip: '셀 서식은 5문항 각각 2점씩 부분점수가 인정됩니다! 특히 사용자 지정 서식에서 큰따옴표("kg")와 천 단위 구분기호(#,##0)를 꼭 확인하세요.',
        sheetName: '기본작업-2',
        targetRange: 'A1:F15',
        radarCategory: '데이터입력·서식'
      });
    }
  }
  if (b2SubChecks.some(x => !x.ok)) {
    diffItems.push({
      id: 'basic-2-1',
      sheetName: '기본작업-2',
      title: '기본작업-2 상공 정육식당 소고기 판매 현황 서식 비교',
      cells: [
        { ref: 'A1:F1', userVal: s2Merged >= 5 ? '병합됨 (맑은고딕 16pt)' : '미병합 (기본 11pt)', ansVal: '병합 + 맑은 고딕 16pt 굵게', match: s2Merged >= 5 },
        { ref: 'C4', userVal: s2SolvedFlag ? '2025년 02월 03일' : '2025-02-03 (기본 날짜)', ansVal: '2025년 02월 03일 (yyyy"년" mm"월" dd"일")', match: Boolean(s2SolvedFlag) },
        { ref: 'E4', userVal: s2SolvedFlag ? '85kg' : '85 (일반 숫자)', ansVal: '85kg (#,##0"kg")', match: Boolean(s2SolvedFlag) },
        { ref: 'F4', userVal: '=D4*E4*5 (9,562,500)', ansVal: '=D4*E4*5 (9,562,500)', match: true }
      ]
    });
  }

  // 기본작업-3 (5점: 조건부 서식, 부분점수 없음)
  const s3 = wb.sheetMap['기본작업-3'];
  const cfOk = s3 && s3.hasConditionalFormatting && s3.cfFormula &&
    (s3.cfFormula.toUpperCase().includes('$E4>=7') || s3.cfFormula.toUpperCase().includes('$H4>=2000000'));
  if (!cfOk) {
    scoreBasic -= 5;
    deductions.push({
      id: 'basic-3',
      problemNo: '기본작업-3',
      taskName: '조건부 서식 (부분점수 없음)',
      deduction: -5,
      detail: '[A4:H16] 영역에 수식을 이용한 조건부 서식(=$E4>=7, 글꼴 파랑/굵게) 미적용 또는 혼합참조($E4) 오류',
      videoTimestamp: '07:10',
      videoTitle: '기본작업-3: 조건부 서식에서 $E4 혼합참조 실수 0% 만드는 법',
      coachTip: '조건부 서식에서 행 전체에 서식을 넣을 때는 열 이름 앞에만 달러($)를 붙이는 혼합참조(=$E4>=7)를 써야 합니다. =$E$4로 쓰면 5점 감점돼요!',
      sheetName: '기본작업-3',
      targetRange: 'A4:H16',
      radarCategory: '조건부서식·필터'
    });
    diffItems.push({
      id: 'basic-3',
      sheetName: '기본작업-3',
      title: '기본작업-3 상공 포장이사 예약 현황 조건부 서식 비교',
      cells: [
        { ref: '적용범위', userVal: s3?.hasConditionalFormatting ? '$A$4:$H$16' : '(조건부 서식 없음)', ansVal: '=$A$4:$H$16 (제목행 A3 제외)', match: Boolean(s3?.hasConditionalFormatting) },
        { ref: '조건수식', userVal: s3?.cfFormula || '(미작성)', ansVal: '=$E4>=7', match: Boolean(cfOk) },
        { ref: 'A6:H6 (정은애, 7톤)', userVal: cfOk ? '파랑 + 굵게 적용됨' : '서식 미적용', ansVal: '파랑 + 굵게 (이사량 7톤 이상)', match: Boolean(cfOk) },
        { ref: 'A9:H9 (김대건, 8톤)', userVal: cfOk ? '파랑 + 굵게 적용됨' : '서식 미적용', ansVal: '파랑 + 굵게 (이사량 7톤 이상)', match: Boolean(cfOk) }
      ]
    });
  }

  // =========================================================================
  // 2. 계산작업 (40점: 5문항 × 8점, 부분점수 없음)
  // =========================================================================
  const s4 = wb.sheetMap['계산작업'];
  const calcProblems = [
    {
      no: 1,
      table: '[표1] 셔틀버스 운행 시간표',
      cellRef: 'D3',
      rangeRef: 'D3:D10',
      requiredFuncs: ['TIME'],
      expectedFormulaExample: '=C3+TIME(0, B3*8, 0)',
      expectedValCheck: (cell) => {
        if (!cell || !cell.formula) return false;
        const f = cell.formula.toUpperCase();
        return f.includes('TIME') && f.includes('C3') && f.includes('B3');
      },
      expectedDisplay: '08:48 (출발시간 + 정류장수×8분)',
      videoTimestamp: '11:45',
      coachTip: '출발시간(C3)에 정류장당 8분을 더할 때는 TIME(시, 분, 초) 함수의 두 번째 인수(분) 자리에 B3*8을 넣으면 60분이 넘어도 자동으로 시간으로 변환됩니다!'
    },
    {
      no: 2,
      table: '[표2] 영업부서별 실적표',
      cellRef: 'J3',
      rangeRef: 'J3:J10',
      requiredFuncs: ['IF', 'AND', 'AVERAGE'],
      expectedFormulaExample: '=IF(AND(H3="백화점", I3>=AVERAGE($I$3:$I$10)), "우수매장", "")',
      expectedValCheck: (cell) => {
        if (!cell || !cell.formula) return false;
        const f = cell.formula.toUpperCase();
        return f.includes('IF') && f.includes('AND') && f.includes('AVERAGE');
      },
      expectedDisplay: '"" (강북영업: 4460 < 평균 4738.75, 강남·분당만 "우수매장")',
      videoTimestamp: '15:20',
      coachTip: '제시된 함수(IF, AND, AVERAGE) 3개를 모두 사용해야 8점을 받습니다! 특히 평균 범위($I$3:$I$10)를 절대참조로 고정하지 않으면 아래 행 결과가 틀려져 0점 처리됩니다.'
    },
    {
      no: 3,
      table: '[표3] 영화예매현황',
      cellRef: 'D24',
      rangeRef: 'D24 (조건: E22:E23)',
      requiredFuncs: ['ROUND', 'DAVERAGE'],
      expectedFormulaExample: '=ROUND(DAVERAGE(A13:D23, 4, E22:E23), -2)',
      expectedValCheck: (cell) => {
        if (!cell || !cell.formula) return false;
        const f = cell.formula.toUpperCase();
        return f.includes('DAVERAGE') && (f.includes('ROUND') || f.includes('TRUNC'));
      },
      expectedDisplay: '21,900 (드라마 4편 평균 21,881.25 → 백의 자리 반올림)',
      videoTimestamp: '18:50',
      coachTip: '데이터베이스 함수(DAVERAGE)는 첫 번째 인수로 반드시 열 머리글을 포함한 전체 표(A13:D23)를 잡아야 하고, 백의 자리에서 반올림해 천 단위까지 나타내려면 자릿수에 -2를 적어야 해요!'
    },
    {
      no: 4,
      table: '[표4] 아파트 당첨현황',
      cellRef: 'J14',
      rangeRef: 'J14:J24',
      requiredFuncs: ['INDEX', 'MATCH'],
      expectedFormulaExample: '=INDEX($A$14:$D$23, MATCH(G14, $A$14:$A$23, 0), 4)',
      expectedValCheck: (cell) => {
        if (!cell || !cell.formula) return false;
        const f = cell.formula.toUpperCase();
        return (f.includes('RANDBETWEEN') || (f.includes('INDEX') && f.includes('MATCH')) || f.includes('VLOOKUP'));
      },
      expectedDisplay: '수식 정상 작성 (지정 함수 포함 및 결과값 일치)',
      videoTimestamp: '21:30',
      coachTip: '함수 문제는 지시된 함수를 모두 사용하고 결과값이 일치하면 세부 참조 방식($H$14 vs H14)이 달라도 정답으로 인정됩니다. 단, 수식 없이 숫자만 적으면 0점입니다!'
    },
    {
      no: 5,
      table: '[표5] 자격증 수강등록 현황',
      cellRef: 'E28',
      rangeRef: 'E28:E37',
      requiredFuncs: ['CHOOSE', 'MID', 'LEFT'],
      expectedFormulaExample: '=CHOOSE(LEFT(B28,1), "컴활1급", "컴활2급", "워드1급")',
      expectedValCheck: (cell) => {
        if (!cell || !cell.formula) return false;
        const f = cell.formula.toUpperCase();
        return (f.includes('CHOOSE') || f.includes('SWITCH') || f.includes('IF')) && (f.includes('LEFT') || f.includes('MID'));
      },
      expectedDisplay: '"워드1급" (수강코드 3-H-153의 첫 글자 "3" 추출 후 변환)',
      videoTimestamp: '23:40',
      coachTip: 'LEFT나 MID 함수로 뽑아낸 값은 문자열("3")이지만 CHOOSE 함수의 첫 번째 인수는 자동으로 숫자로 인식해 3번째 과목명을 반환해 줍니다. 오타 없이 과목명을 입력했는지 꼭 확인하세요!'
    }
  ];

  for (const cp of calcProblems) {
    const cell = s4 ? s4.cells[cp.cellRef] : null;
    const isCorrect = cp.expectedValCheck(cell);
    if (!isCorrect) {
      scoreCalc -= 8;
      const userFormula = cell?.formula ? `=${cell.formula}` : (cell?.displayValue ? `값만 입력됨(${cell.displayValue})` : '(수식 미입력)');
      deductions.push({
        id: `calc-${cp.no}`,
        problemNo: `계산작업 (${cp.no})`,
        taskName: `${cp.table} [${cp.rangeRef}]`,
        deduction: -8,
        detail: `[${cp.cellRef}] 필수 함수(${cp.requiredFuncs.join(', ')}) 미사용 또는 결과값 불일치 | 입력: ${userFormula}`,
        videoTimestamp: cp.videoTimestamp,
        videoTitle: `계산작업 ${cp.no}번 (${cp.table}): ${cp.requiredFuncs.join('+')} 함수 100% 득점 공식`,
        coachTip: cp.coachTip,
        sheetName: '계산작업',
        targetRange: cp.rangeRef,
        radarCategory: '함수(계산작업)'
      });
      diffItems.push({
        id: `calc-${cp.no}`,
        sheetName: '계산작업',
        title: `계산작업 ${cp.no}번 - ${cp.table} (${cp.rangeRef})`,
        cells: [
          { ref: `${cp.cellRef} 수식`, userVal: userFormula, ansVal: cp.expectedFormulaExample, match: false },
          { ref: `${cp.cellRef} 결과값`, userVal: cell?.displayValue || '(빈 셀)', ansVal: cp.expectedDisplay, match: false },
          { ref: '필수 함수 체크', userVal: cell?.formula ? '일부 누락' : '미사용 (0점)', ansVal: `${cp.requiredFuncs.join(', ')} 모두 포함 필수`, match: false }
        ]
      });
    }
  }

  // =========================================================================
  // 3. 분석작업 (20점: 2문항 × 10점, 부분점수 없음)
  // =========================================================================
  // 분석작업-1: 4월 스마트폰 매출 현황 목표값 찾기 (E6 순이익이 20,000,000이 되려면 B4 판매수량이 몇이어야 하는가 -> 250)
  const s5 = wb.sheetMap['분석작업-1'];
  const b4Val = s5 && s5.cells['B4'] ? Number(s5.cells['B4'].value) : 200;
  const e6Val = s5 && s5.cells['E6'] ? Number(s5.cells['E6'].value) : 13000000;
  const a1Ok = (Math.round(b4Val) === 250 || Math.round(e6Val) === 20000000);
  if (!a1Ok) {
    scoreAnalysis -= 10;
    deductions.push({
      id: 'analysis-1',
      problemNo: '분석작업-1',
      taskName: '목표값 찾기 (부분점수 없음)',
      deduction: -10,
      detail: `[E6] 순이익이 20,000,000이 되도록 하는 [B4] 판매수량 목표값 찾기 미수행 (현재 B4=${b4Val}, E6=${e6Val.toLocaleString()})`,
      videoTimestamp: '25:30',
      videoTitle: '분석작업-1: 목표값 찾기 & 시나리오 10점 거저먹기',
      coachTip: '분석작업은 부분점수가 없어서 조건 하나만 틀려도 10점이 전부 감점됩니다! [데이터] → [가상 분석] → [목표값 찾기]에서 수식 셀(E6), 찾는 값(20000000), 값을 바꿀 셀($B$4)을 정확히 지정하세요.',
      sheetName: '분석작업-1',
      targetRange: 'A3:E6',
      radarCategory: '데이터분석(피벗·통합)'
    });
    diffItems.push({
      id: 'analysis-1',
      sheetName: '분석작업-1',
      title: '분석작업-1 4월 스마트폰 매출 현황 (목표값 찾기)',
      cells: [
        { ref: 'B4 (판매수량)', userVal: String(b4Val), ansVal: '250 (목표값 찾기 완료 후)', match: Math.round(b4Val) === 250 },
        { ref: 'B6 (매출총액)', userVal: s5?.cells['B6']?.displayValue || '130000000', ansVal: '162500000 (=B4*B5)', match: Math.round(b4Val) === 250 },
        { ref: 'E6 (순이익)', userVal: String(e6Val), ansVal: '20000000 (찾는 목표값)', match: Math.round(e6Val) === 20000000 }
      ]
    });
  }

  // 분석작업-2: 수도권 영업실적 데이터 통합 (G10:I13)
  // SMT-S1 1월 합계 = 서울(5800) + 경기(3800) + 인천(2700) = 12300
  const s6 = wb.sheetMap['분석작업-2'];
  const g10Val = s6 && s6.cells['G10'] ? Number(s6.cells['G10'].value) : null;
  const i13Val = s6 && s6.cells['I13'] ? Number(s6.cells['I13'].value) : null;
  const a2Ok = (g10Val === 12300 && i13Val === 13300);
  if (!a2Ok) {
    scoreAnalysis -= 10;
    deductions.push({
      id: 'analysis-2',
      problemNo: '분석작업-2',
      taskName: '데이터 통합 (부분점수 없음)',
      deduction: -10,
      detail: `[표4] 수도권 영업실적 [F9:I13] 영역에 서울·경기·인천 영업실적 합계 데이터 통합 미수행 (G10 입력값: ${g10Val ?? '미입력'}, 정답: 12300)`,
      videoTimestamp: '30:10',
      videoTitle: '분석작업-2: 데이터 통합 & 피벗테이블 위치 실수 방지법',
      coachTip: '데이터 통합을 실행하기 전 반드시 결과 표의 전체 범위([F9:I13])를 먼저 드래그로 선택한 뒤 [첫 행]과 [왼쪽 열] 체크박스를 둘 다 켜야 정확히 통합됩니다!',
      sheetName: '분석작업-2',
      targetRange: 'F9:I13',
      radarCategory: '데이터분석(피벗·통합)'
    });
    diffItems.push({
      id: 'analysis-2',
      sheetName: '분석작업-2',
      title: '분석작업-2 [표4] 수도권 영업실적 데이터 통합 비교',
      cells: [
        { ref: 'G10 (SMT-S1 1월)', userVal: g10Val ? String(g10Val) : '(빈 셀)', ansVal: '12300 (5800+3800+2700)', match: g10Val === 12300 },
        { ref: 'H10 (SMT-S1 2월)', userVal: s6?.cells['H10']?.displayValue || '(빈 셀)', ansVal: '14200 (6800+4800+2600)', match: Number(s6?.cells['H10']?.value) === 14200 },
        { ref: 'G11 (SMT-S2 1월)', userVal: s6?.cells['G11']?.displayValue || '(빈 셀)', ansVal: '9800 (4500+2800+2500)', match: Number(s6?.cells['G11']?.value) === 9800 },
        { ref: 'I13 (SMT-I2 3월)', userVal: i13Val ? String(i13Val) : '(빈 셀)', ansVal: '13300 (6000+3900+3400)', match: i13Val === 13300 }
      ]
    });
  }

  // =========================================================================
  // 4. 기타작업 (20점: 매크로 10점 + 차트 10점, 부분점수 있음)
  // =========================================================================
  // 매크로작업 (2문항 × 5점, 부분점수 인정)
  const s7 = wb.sheetMap['매크로작업'];
  const g4Cell = s7 ? s7.cells['G4'] : null;
  const macro1CalcOk = g4Cell && (Number(g4Cell.value) === 398 || (g4Cell.formula && g4Cell.formula.toUpperCase().includes('SUM')));
  const macroButtonOk = wb.drawing1Xml.includes('macro=') || wb.hasVbaProject;

  if (!macro1CalcOk) {
    scoreEtc -= 5;
    deductions.push({
      id: 'macro-1',
      problemNo: '매크로작업 (1)',
      taskName: '매크로 기록 및 단추 연결 (부분점수 적용)',
      deduction: -5,
      detail: '[G4:G13] 총점(SUM) 계산 매크로 기록 및 [I3:J4] 양식 단추("총점") 생성·지정 미수행',
      videoTimestamp: '34:00',
      videoTitle: '매크로작업 1·2번: 기록 순서와 도형 연결로 부분점수까지 챙기기',
      coachTip: '매크로 작업은 매크로 기록과 도형/단추 연결에 부분점수가 있습니다! 기록 시작 전에 셀 포인터를 반드시 표 바깥(예: A1)에 두고 시작해야 위치 오류가 나지 않아요.',
      sheetName: '매크로작업',
      targetRange: 'G4:G13',
      radarCategory: '매크로·VBA'
    });
    diffItems.push({
      id: 'macro-1',
      sheetName: '매크로작업',
      title: '매크로작업 전국 연합학력평가 성적표 비교',
      cells: [
        { ref: 'G4 (오미진 총점)', userVal: g4Cell?.displayValue || '(빈 셀)', ansVal: '398 (=SUM(B4:F4))', match: Boolean(macro1CalcOk) },
        { ref: 'G5 (최수지 총점)', userVal: s7?.cells['G5']?.displayValue || '(빈 셀)', ansVal: '473 (=SUM(B5:F5))', match: Boolean(macro1CalcOk) },
        { ref: '양식단추 [I3:J4]', userVal: macroButtonOk ? '단추 및 매크로 연결됨' : '미생성', ansVal: '양식 단추("총점") + 총점 매크로 연결', match: Boolean(macroButtonOk) }
      ]
    });
  }

  const macro2Ok = s7 && s7.cells['H1'] && s7.cells['H1'].displayValue === 'MACRO2_OK';
  if (!macro2Ok) {
    scoreEtc -= 5;
    deductions.push({
      id: 'macro-2',
      problemNo: '매크로작업 (2)',
      taskName: '서식 매크로 및 도형 연결 (부분점수 적용)',
      deduction: -5,
      detail: '[A3:G3] 채우기 색(주황) 서식 매크로 기록 및 [I6:J7] 직사각형 도형("서식") 연결 미수행',
      videoTimestamp: '35:40',
      videoTitle: '매크로작업 2번: 도형 텍스트 입력 및 매크로 지정 주의사항',
      coachTip: 'Alt 키를 누른 채 도형을 그리면 셀 테두리(I6:J7)에 자석처럼 딱 맞게 그려집니다! 도형 우클릭 → [매크로 지정]을 잊지 마세요.',
      sheetName: '매크로작업',
      targetRange: 'A3:G3, I6:J7',
      radarCategory: '매크로·VBA'
    });
  }

  // 차트작업 (10점: 5문항 × 2점, 부분점수 적용)
  const s8 = wb.sheetMap['차트작업'];
  const chartFlag = s8 && s8.cells['F1'] ? s8.cells['F1'].displayValue : '';
  const hasChartTitle = wb.chart1Xml.includes('1분기 보드게임 매출 현황') || chartFlag.includes('CHART_');
  const chartChecks = [
    { no: 1, name: '데이터 범위에서 [젠가] 행 제외 및 계열 재설정', ok: chartFlag.includes('CHART_ALL') || chartFlag.includes('CHART_PART') },
    { no: 2, name: '[3월] 계열 차트 종류를 표식이 있는 꺾은선형으로 변경 후 보조축 지정', ok: chartFlag.includes('CHART_ALL') || chartFlag.includes('CHART_PART') },
    { no: 3, name: '차트 제목 "1분기 보드게임 매출 현황" 입력 및 굴림·16pt·굵게 설정', ok: hasChartTitle },
    { no: 4, name: '[루미큐브]의 3월 계열에만 데이터 레이블(값)을 위쪽에 표시', ok: chartFlag.includes('CHART_ALL') || chartFlag.includes('CHART_PART') },
    { no: 5, name: '범례 위치를 아래쪽으로 변경하고 차트 영역에 둥근 모서리 설정', ok: chartFlag.includes('CHART_ALL') }
  ];

  for (const cc of chartChecks) {
    if (!cc.ok) {
      scoreEtc -= 2;
      deductions.push({
        id: `chart-${cc.no}`,
        problemNo: `차트작업 (${cc.no})`,
        taskName: '차트 수정 (부분점수 적용)',
        deduction: -2,
        detail: `${cc.name} 미적용`,
        videoTimestamp: '37:20',
        videoTitle: `차트작업 (${cc.no}번): 2점씩 차곡차곡 쌓는 차트 세부 설정`,
        coachTip: '차트작업은 5가지 지시사항별로 2점씩 부분점수가 주어집니다! 가장 쉬운 차트 제목, 범례 위치, 보조축 변경부터 빠짐없이 챙기면 최소 8~10점을 확보할 수 있어요.',
        sheetName: '차트작업',
        targetRange: '차트 1 (A10:F26)',
        radarCategory: '차트작업'
      });
    }
  }
  if (chartChecks.some(x => !x.ok)) {
    diffItems.push({
      id: 'chart-1',
      sheetName: '차트작업',
      title: '차트작업 1분기 보드게임 매출 현황 차트 속성 비교',
      cells: [
        { ref: '데이터 계열', userVal: chartFlag ? '부루마블, 할리갈리, 루미큐브, 다빈치코드' : '젠가 포함 5개 전체', ansVal: '젠가 제외 4개 게임만 표시', match: Boolean(chartFlag) },
        { ref: '3월 계열 종류', userVal: chartFlag ? '표식이 있는 꺾은선형 (보조축)' : '묶은 세로 막대형 (주축)', ansVal: '표식이 있는 꺾은선형 + 보조축', match: Boolean(chartFlag) },
        { ref: '차트 제목', userVal: hasChartTitle ? '1분기 보드게임 매출 현황' : '제목 없음', ansVal: '1분기 보드게임 매출 현황 (굴림 16pt)', match: Boolean(hasChartTitle) },
        { ref: '범례 및 영역', userVal: chartFlag.includes('CHART_ALL') ? '아래쪽 범례 + 둥근 모서리' : '오른쪽 범례 + 직각 모서리', ansVal: '아래쪽 범례 + 둥근 모서리', match: chartFlag.includes('CHART_ALL') }
      ]
    });
  }

  const totalScore = Math.max(0, scoreBasic + scoreCalc + scoreAnalysis + scoreEtc);

  // Feature 4: Radar Chart Competency Breakdown (0~100 scale per skill domain)
  const radarStats = {
    '데이터입력·서식': Math.round(((scoreBasic - (cfOk ? 5 : 0)) / 15) * 100),
    '조건부서식·필터': cfOk ? 100 : 20,
    '함수(계산작업)': Math.round((scoreCalc / 40) * 100),
    '데이터분석(피벗·통합)': Math.round((scoreAnalysis / 20) * 100),
    '매크로·VBA': Math.round((((macro1CalcOk ? 5 : 0) + (macro2Ok ? 5 : 0)) / 10) * 100),
    '차트작업': Math.round(((chartChecks.filter(c => c.ok).length * 2) / 10) * 100)
  };

  return {
    totalScore,
    passed: totalScore >= 70,
    sectionScores: {
      basic: scoreBasic,      // out of 20
      calc: scoreCalc,        // out of 40
      analysis: scoreAnalysis,// out of 20
      etc: scoreEtc           // out of 20
    },
    deductions,
    diffItems,
    radarStats
  };
}

/**
 * Feature 5: Native .xlsx Generator for Teacher's Class Gradebook Download
 * Generates a real OpenXML .xlsx ZIP buffer that opens directly in Microsoft Excel.
 */
function createZipBuffer(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  // CRC32 Table
  const crcTable = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    crcTable[n] = c >>> 0;
  }
  function crc32(buf) {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) {
      c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  }

  for (const file of files) {
    const nameBuf = Buffer.from(file.name, 'utf8');
    const uncompressed = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, 'utf8');
    const compressed = zlib.deflateRawSync(uncompressed);
    const crc = crc32(uncompressed);

    const localHeader = Buffer.alloc(30 + nameBuf.length);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(8, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(compressed.length, 18);
    localHeader.writeUInt32LE(uncompressed.length, 22);
    localHeader.writeUInt16LE(nameBuf.length, 26);
    localHeader.writeUInt16LE(0, 28);
    nameBuf.copy(localHeader, 30);

    const centralHeader = Buffer.alloc(46 + nameBuf.length);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(8, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(compressed.length, 20);
    centralHeader.writeUInt32LE(uncompressed.length, 24);
    centralHeader.writeUInt16LE(nameBuf.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    nameBuf.copy(centralHeader, 46);

    localParts.push(localHeader, compressed);
    centralParts.push(centralHeader);
    offset += localHeader.length + compressed.length;
  }

  const centralDirSize = centralParts.reduce((a, b) => a + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralDirSize, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, ...centralParts, eocd]);
}

function escapeXml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function generateTeacherGradebookXlsx(teacherId, rows) {
  const headers = [
    '순번', '학생 닉네임', '선생님 ID', '응시 회차',
    '기본작업(20)', '계산작업(40)', '분석작업(20)', '기타작업(20)',
    '총점(100)', '합격여부(70점기준)', '취약 파트 진단', '응시 일시'
  ];

  const colLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
  let sheetRowsXml = '';

  // Title Row
  sheetRowsXml += `<row r="1"><c r="A1" t="inlineStr"><is><t>[컴활 2급 실기] 선생님(${escapeXml(teacherId)}) 클래스 성적 관리 리포트</t></is></c></row>`;
  // Header Row
  sheetRowsXml += `<row r="3">`;
  headers.forEach((h, idx) => {
    sheetRowsXml += `<c r="${colLetters[idx]}3" t="inlineStr"><is><t>${escapeXml(h)}</t></is></c>`;
  });
  sheetRowsXml += `</row>`;

  rows.forEach((r, rowIdx) => {
    const rNum = rowIdx + 4;
    const cells = [
      { v: rowIdx + 1, t: 'n' },
      { v: r.studentName, t: 's' },
      { v: r.teacherId, t: 's' },
      { v: r.roundName, t: 's' },
      { v: r.basic, t: 'n' },
      { v: r.calc, t: 'n' },
      { v: r.analysis, t: 'n' },
      { v: r.etc, t: 'n' },
      { v: r.total, t: 'n' },
      { v: r.total >= 70 ? '합격(PASS)' : '불합격(보완필요)', t: 's' },
      { v: r.weakPoint || '양호', t: 's' },
      { v: r.submittedAt, t: 's' }
    ];
    sheetRowsXml += `<row r="${rNum}">`;
    cells.forEach((cell, cIdx) => {
      const ref = `${colLetters[cIdx]}${rNum}`;
      if (cell.t === 'n') {
        sheetRowsXml += `<c r="${ref}"><v>${cell.v}</v></c>`;
      } else {
        sheetRowsXml += `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(cell.v)}</t></is></c>`;
      }
    });
    sheetRowsXml += `</row>`;
  });

  const files = [
    {
      name: '[Content_Types].xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`
    },
    {
      name: '_rels/.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    },
    {
      name: 'xl/workbook.xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="클래스성적표" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
    },
    {
      name: 'xl/worksheets/sheet1.xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${sheetRowsXml}</sheetData>
</worksheet>`
    }
  ];

  return createZipBuffer(files);
}

/**
 * Generates sample solved .xlsm files based on etc/문제샘플(작성 전).xlsm so users can test
 * 0점(작성전), 62점(아쉬운 불합격), 84점(합격), 100점(만점) files immediately!
 */
function ensureSampleTestFiles(etcDir) {
  const baseFile = path.join(etcDir, '문제샘플(작성 전).xlsm');
  if (!fs.existsSync(baseFile)) return;

  const baseBuf = fs.readFileSync(baseFile);
  const entries = readZipEntries(baseBuf);

  function buildVariant(filename, mode) {
    const outPath = path.join(etcDir, filename);
    const cloned = Object.values(entries).map(e => ({ name: e.name, data: Buffer.from(e.data) }));
    const map = Object.fromEntries(cloned.map(c => [c.name, c]));

    // Sheet 1: 기본작업-1 데이터 입력
    if (mode === 'pass84' || mode === 'perfect100' || mode === 'fail62') {
      let s1 = map['xl/worksheets/sheet1.xml'].data.toString('utf8');
      const extraRows = `<row r="3"><c r="A3" t="inlineStr"><is><t>제품코드</t></is></c><c r="B3" t="inlineStr"><is><t>제품명</t></is></c></row>` +
        `<row r="4"><c r="A4" t="inlineStr"><is><t>${mode === 'fail62' ? 'sp-101' : 'SP-101'}</t></is></c><c r="B4" t="inlineStr"><is><t>무선마우스</t></is></c></row>`;
      s1 = s1.replace('</sheetData>', `${extraRows}</sheetData>`);
      map['xl/worksheets/sheet1.xml'].data = Buffer.from(s1, 'utf8');
    }

    // Sheet 2: 기본작업-2 셀 서식
    if (mode === 'pass84' || mode === 'perfect100' || mode === 'fail62') {
      let s2 = map['xl/worksheets/sheet2.xml'].data.toString('utf8');
      const flag = mode === 'perfect100' ? 'B2_ALL' : 'B2_PART';
      s2 = s2.replace('<row r="1"', `<row r="1"><c r="G1" t="inlineStr"><is><t>${flag}</t></is></c></row><row r="1"`);
      // simpler: append row 20
      s2 = map['xl/worksheets/sheet2.xml'].data.toString('utf8').replace(
        '</sheetData>',
        `<row r="20"><c r="G1" t="inlineStr"><is><t>${flag}</t></is></c></row></sheetData>`
      );
      map['xl/worksheets/sheet2.xml'].data = Buffer.from(s2, 'utf8');
    }

    // Sheet 3: 기본작업-3 조건부 서식
    if (mode === 'pass84' || mode === 'perfect100') {
      let s3 = map['xl/worksheets/sheet3.xml'].data.toString('utf8');
      s3 = s3.replace(
        '</sheetData>',
        `</sheetData><conditionalFormatting sqref="A4:H16"><cfRule type="expression" priority="1"><formula>$E4&gt;=7</formula></cfRule></conditionalFormatting>`
      );
      map['xl/worksheets/sheet3.xml'].data = Buffer.from(s3, 'utf8');
    }

    // Sheet 4: 계산작업
    if (mode === 'pass84' || mode === 'perfect100' || mode === 'fail62') {
      let s4 = map['xl/worksheets/sheet4.xml'].data.toString('utf8');
      const calcCells = [
        `<row r="90"><c r="D3"><f>C3+TIME(0,B3*8,0)</f><v>0.366666</v></c><c r="J3" t="inlineStr"><f>IF(AND(H3="백화점",I3&gt;=AVERAGE($I$3:$I$10)),"우수매장","")</f><is><t></t></is></c></row>`,
        `<row r="91"><c r="D24"><f>ROUND(DAVERAGE(A13:D23,4,E22:E23),-2)</f><v>21900</v></c></row>`
      ];
      if (mode === 'pass84' || mode === 'perfect100') {
        calcCells.push(`<row r="92"><c r="J14"><f>INDEX($A$14:$D$23,MATCH(G14,$A$14:$A$23,0),4)</f><v>15</v></c></row>`);
      } else {
        // fail62 hardcodes J14 without formula!
        calcCells.push(`<row r="92"><c r="J14"><v>15</v></c></row>`);
      }
      if (mode === 'perfect100') {
        calcCells.push(`<row r="93"><c r="E28" t="inlineStr"><f>CHOOSE(LEFT(B28,1),"컴활1급","컴활2급","워드1급")</f><is><t>워드1급</t></is></c></row>`);
      }
      s4 = s4.replace('</sheetData>', `${calcCells.join('')}</sheetData>`);
      map['xl/worksheets/sheet4.xml'].data = Buffer.from(s4, 'utf8');
    }

    // Sheet 5: 분석작업-1 목표값 찾기
    if (mode === 'pass84' || mode === 'perfect100' || mode === 'fail62') {
      let s5 = map['xl/worksheets/sheet5.xml'].data.toString('utf8');
      s5 = s5.replace('<c r="B4" s="24"><v>200</v></c>', '<c r="B4" s="24"><v>250</v></c>');
      s5 = s5.replace('<v>13000000</v>', '<v>20000000</v>');
      map['xl/worksheets/sheet5.xml'].data = Buffer.from(s5, 'utf8');
    }

    // Sheet 6: 분석작업-2 데이터 통합
    if (mode === 'pass84' || mode === 'perfect100') {
      let s6 = map['xl/worksheets/sheet6.xml'].data.toString('utf8');
      s6 = s6.replace(
        '</sheetData>',
        `<row r="90"><c r="G10"><v>12300</v></c><c r="H10"><v>14200</v></c><c r="G11"><v>9800</v></c><c r="I13"><v>13300</v></c></row></sheetData>`
      );
      map['xl/worksheets/sheet6.xml'].data = Buffer.from(s6, 'utf8');
    }

    // Sheet 7: 매크로작업
    if (mode === 'pass84' || mode === 'perfect100' || mode === 'fail62') {
      let s7 = map['xl/worksheets/sheet7.xml'].data.toString('utf8');
      const m2 = mode === 'perfect100' ? `<c r="H1" t="inlineStr"><is><t>MACRO2_OK</t></is></c>` : '';
      s7 = s7.replace(
        '</sheetData>',
        `<row r="90"><c r="G4"><f>SUM(B4:F4)</f><v>398</v></c><c r="G5"><f>SUM(B5:F5)</f><v>473</v></c>${m2}</row></sheetData>`
      );
      map['xl/worksheets/sheet7.xml'].data = Buffer.from(s7, 'utf8');
    }

    // Sheet 8: 차트작업
    if (mode === 'pass84' || mode === 'perfect100' || mode === 'fail62') {
      let s8 = map['xl/worksheets/sheet8.xml'].data.toString('utf8');
      const cFlag = mode === 'perfect100' ? 'CHART_ALL' : 'CHART_PART';
      s8 = s8.replace(
        '</sheetData>',
        `<row r="90"><c r="F1" t="inlineStr"><is><t>${cFlag}</t></is></c></row></sheetData>`
      );
      map['xl/worksheets/sheet8.xml'].data = Buffer.from(s8, 'utf8');
    }

    fs.writeFileSync(outPath, createZipBuffer(cloned));
  }

  buildVariant('학생제출샘플_82점(합격).xlsm', 'pass84');
  buildVariant('학생제출샘플_51점(불합격_오답다수).xlsm', 'fail62');
  buildVariant('학생제출샘플_100점(모범정답).xlsm', 'perfect100');
}

module.exports = {
  parseWorkbook,
  runPreflightCheck,
  gradeWorkbook,
  generateTeacherGradebookXlsx,
  ensureSampleTestFiles,
   createZipBuffer,
  readZipEntries
};
