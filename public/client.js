// ============================================================================
// Dual-Mode Engine + Mock Exam Problem Bank + Real 8-Sheet .xlsx Problem Generator
// ============================================================================

let currentFileBuffer = null;
let currentFileName = null;
let latestGradeData = null;
let selectedDeductionIdx = 0;
let timerSeconds = 40 * 60;
let timerInterval = null;
let generatedCount = 0;
let currentExamSpec = null;
let showAnswersAll = false;
let filterUnmasteredOnly = false;
let wrongNotesList = [
  {
    id: 'wn-1',
    student_nickname: '짱아',
    teacher_id: 'Zzang_AA',
    exam_code: '대표기출따라하기',
    problem_no: '계산작업 (5)',
    task_name: '[표5] 자격증 수강등록 현황 [E28:E37]',
    deduction: -8,
    user_input: '=CHOOSE(LEFT(B28,1),"컴활1급","컴퓨터활용2급","워드1급")',
    correct_formula: '=CHOOSE(LEFT(B28,1),"컴활1급","컴활2급","워드1급")',
    error_reason: 'CHOOSE 함수 내부 반환 텍스트 오타 ("컴퓨터활용2급" → 정답: "컴활2급")로 인해 결과값 불일치 전액 감점',
    coach_tip: '함수 구조가 완벽해도 큰따옴표 안의 출력 문자열에 글자 하나만 달라도 8점 전체가 감점됩니다!',
    student_memo: '문제 지시사항의 큰따옴표("컴활2급") 글자를 토씨 하나 틀리지 않고 그대로 적기!',
    is_mastered: false
  },
  {
    id: 'wn-2',
    student_nickname: '짱아',
    teacher_id: 'Zzang_AA',
    exam_code: '대표기출따라하기',
    problem_no: '기본작업-3',
    task_name: '조건부 서식 [A4:H16] (부분점수 없음)',
    deduction: -5,
    user_input: '=$E$4>=7 (적용범위: $A$3:$H$16)',
    correct_formula: '=$E4>=7 (적용범위: $A$4:$H$16)',
    error_reason: '열 이름 앞에만 $를 붙이는 혼합참조($E4) 대신 절대참조($E$4)를 사용했고, 적용 범위에 제목 행(3행)을 포함함',
    coach_tip: '조건부 서식은 데이터 첫 행(A4:H16)만 드래그한 뒤 F4 키를 두 번 눌러 =$E4>=7 혼합참조로 만들어야 합니다!',
    student_memo: '조건부 서식 범위 잡을 때 제목 행(3행) 절대 포함하지 말 것! F4 두 번 눌러서 $E4 만들기!',
    is_mastered: false
  },
  {
    id: 'wn-3',
    student_nickname: '짱아',
    teacher_id: 'Zzang_AA',
    exam_code: '기본모의고사 01회',
    problem_no: '분석작업-2',
    task_name: '도서분류별 부분합 (부분점수 없음)',
    deduction: -10,
    user_input: '2차 평균 부분합 실행 시 [새로운 값으로 대치] 체크 유지함',
    correct_formula: '1차 합계 실행 후 2차 평균 실행 시 [새로운 값으로 대치] 체크 해제 필수',
    error_reason: '2차 부분합에서 [새로운 값으로 대치]를 해제하지 않아 1차 합계 결과가 덮어써짐 (10점 전액 감점)',
    coach_tip: '부분합을 두 개 이상 구할 때는 두 번째 부분합 창에서 반드시 [새로운 값으로 대치] 체크를 풀어야 두 줄이 모두 남습니다!',
    student_memo: '부분합 1번 정렬 → 2번 1차 부분합 → 3번 2차 부분합(대치 해제!) 공식처럼 외우기',
    is_mastered: true
  },
  {
    id: 'wn-4',
    student_nickname: '짱아',
    teacher_id: 'Zzang_AA',
    exam_code: '대표기출따라하기',
    problem_no: '매크로작업 (2)',
    task_name: '서식 매크로 및 도형 연결 (부분점수 적용)',
    deduction: -5,
    user_input: '[A3:G3] 주황 채우기 매크로 기록 완료, [I6:J7] 직사각형에 매크로 미지정',
    correct_formula: '직사각형 도형("서식") 생성 후 우클릭 → [매크로 지정] → "서식" 선택',
    error_reason: '매크로 기록은 정상이나 도형 컨트롤에 매크로 연결(macro 속성)이 누락됨',
    coach_tip: 'Alt 키를 누르고 도형을 셀에 맞춰 그린 뒤, 곧바로 도형 우클릭 → [매크로 지정]까지 해야 5점 만점입니다!',
    student_memo: '도형 그리고 텍스트 적은 뒤 우클릭 [매크로 지정] 후 빈 셀 클릭해 작동 확인하기!',
    is_mastered: false
  }
];

// ============================================================================
// 0. Supabase Direct REST Client Configuration (Project: ptargvvfuhcmobhhzxlu)
// ============================================================================
const SUPABASE_URL = 'https://ptargvvfuhcmobhhzxlu.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB0YXJndnZmdWhjbW9iaGh6eGx1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTM0NjMyNDQsImV4cCI6MjEwOTAzOTI0NH0.gxSeFNiA8GKJzGrsfhOZz8lGGmQYsrdCfxQw19j2QWg';

async function sbFetch(path, options = {}) {
  const headers = {
    apikey: SUPABASE_ANON_KEY,
    Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
    'Content-Type': 'application/json',
    ...(options.headers || {})
  };
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers
  });
  if (!res.ok) {
    throw new Error(`Supabase HTTP ${res.status}`);
  }
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

async function syncFromSupabase() {
  const badge = document.getElementById('supabaseStatusBadge');
  try {
    // 1. Fetch exams, submissions, student rankings, teacher rankings, wrong notes in parallel
    const [exams, submissions, stuRank, tchrRank, dbWrongNotes] = await Promise.all([
      sbFetch('comlicense_exams?select=*&order=created_at.asc'),
      sbFetch('comlicense_submissions?select=*&order=submitted_at.desc&limit=50'),
      sbFetch('comlicense_student_rankings_view?select=*&limit=10'),
      sbFetch('comlicense_teacher_rankings_view?select=*&limit=5'),
      sbFetch('comlicense_wrong_notes?select=*&order=created_at.desc&limit=30')
    ]);

    if (Array.isArray(dbWrongNotes) && dbWrongNotes.length > 0) {
      wrongNotesList = dbWrongNotes;
      renderWrongNotes();
    }

    if (Array.isArray(exams) && exams.length > 0) {
      const roundSelect = document.getElementById('roundSelect');
      exams.forEach(ex => {
        if (!MOCK_EXAM_BANK[ex.exam_code]) {
          MOCK_EXAM_BANK[ex.exam_code] = {
            id: ex.exam_code,
            title: ex.title,
            themeName: ex.theme_name,
            basic: ex.basic_questions || [],
            calc: ex.calc_questions || [],
            analysis: ex.analysis_questions || [],
            etc: ex.etc_questions || [],
            sampleRows: ex.sample_rows || []
          };
          const opt = document.createElement('option');
          opt.value = ex.exam_code;
          opt.textContent = ex.title;
          roundSelect.appendChild(opt);
        }
      });
    }

    if (Array.isArray(submissions) && submissions.length > 0) {
      localState.submissions = submissions.map(s => ({
        studentName: s.student_nickname,
        teacherId: s.teacher_id,
        roundName: s.exam_code,
        basic: s.score_basic,
        calc: s.score_calc,
        analysis: s.score_analysis,
        etc: s.score_etc,
        total: s.total_score,
        weakPoint: s.weak_point || '양호',
        submittedAt: (s.submitted_at || '').replace('T', ' ').slice(0, 19),
        radarStats: s.radar_stats || {}
      }));
      localState.participantCount = 2100 + submissions.length;
    }

    if (Array.isArray(stuRank) && stuRank.length > 0) {
      localState.studentRankings = stuRank.map(r => ({
        name: r.masked_name,
        rawName: r.raw_name,
        score: 1700 + Number(r.cumulative_score || 0)
      }));
      renderRankingBars('studentRankingChart', localState.studentRankings, 'name', 2200);
    }

    if (Array.isArray(tchrRank) && tchrRank.length > 0) {
      localState.teacherRankings = tchrRank.map(t => ({
        id: t.masked_id,
        rawId: t.raw_id,
        score: 450 + Number(t.activity_score || 0)
      }));
      renderRankingBars('teacherRankingChart', localState.teacherRankings, 'id', 850);
    }

    if (badge) {
      badge.textContent = `🟢 Supabase DB 실시간 연동됨 (${submissions ? submissions.length : 0}건 저장됨)`;
    }
    return true;
  } catch (err) {
    if (badge) {
      badge.textContent = '🟡 로컬 독립 모드 (오프라인 캐시 작동중)';
    }
    return false;
  }
}

async function saveExamToSupabase(newExam) {
  try {
    await sbFetch('comlicense_exams', {
      method: 'POST',
      headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
      body: JSON.stringify({
        exam_code: newExam.id,
        publisher: 'youngjin',
        exam_year: 2026,
        subject: 'excel',
        level: 2,
        title: newExam.title,
        theme_name: newExam.themeName,
        is_ai_generated: true,
        basic_questions: newExam.basic,
        calc_questions: newExam.calc,
        analysis_questions: newExam.analysis,
        etc_questions: newExam.etc,
        sample_rows: newExam.sampleRows
      })
    });
  } catch (_) {}
}

async function saveSubmissionToSupabase(payload) {
  try {
    // Ensure exam_code exists in comlicense_exams first
    const spec = MOCK_EXAM_BANK[payload.roundName];
    if (spec && payload.roundName.startsWith('[AI생성')) {
      await saveExamToSupabase(spec);
    }
    await sbFetch('comlicense_submissions', {
      method: 'POST',
      headers: { Prefer: 'return=minimal' },
      body: JSON.stringify({
        exam_code: payload.roundName,
        student_nickname: payload.studentName,
        teacher_id: payload.teacherId,
        file_name: payload.fileName,
        score_basic: payload.gradeResult.sectionScores.basic,
        score_calc: payload.gradeResult.sectionScores.calc,
        score_analysis: payload.gradeResult.sectionScores.analysis,
        score_etc: payload.gradeResult.sectionScores.etc,
        total_score: payload.gradeResult.totalScore,
        weak_point: payload.weakPoint,
        preflight_checks: payload.preflight?.checks || [],
        deductions: payload.gradeResult.deductions || [],
        diff_items: payload.gradeResult.diffItems || [],
        radar_stats: payload.gradeResult.radarStats || {}
      })
    });
    await syncFromSupabase();
  } catch (_) {}
}

// ============================================================================
// 1. Built-in 5 Mock Exam Problem Bank (2026 MS Office 2021 Syllabus)
// ============================================================================
const MOCK_EXAM_BANK = {
  '대표기출따라하기': {
    id: '대표기출따라하기',
    title: '대표기출따라하기 (정육식당·셔틀버스·스마트폰)',
    themeName: '상공 정육식당 & 셔틀버스 운행 & 스마트폰 매출',
    basic: [
      {
        no: '기본작업-1 (5점 / 부분점수 없음)',
        inst: '[기본작업-1] 시트의 [A3:E15] 영역에 다음 상공산업 4월 매출현황 데이터를 오타·띄어쓰기·영문 대소문자 오류 없이 정확히 입력하시오.',
        ans: 'A3:E3(제품코드, 제품명, 규격, 단가, 수량) / A4(SP-101, 무선마우스, 블루투스 5.0, 25000, 120) 등 정확 입력'
      },
      {
        no: '기본작업-2 (10점 / 5문항 각 2점 부분점수)',
        inst: '① [A1:F1] 병합하고 가운데 맞춤, 맑은 고딕 16pt 굵게\n② [C4:C15] 포장일자를 사용자 지정 표시 형식(yyyy"년" mm"월" dd"일")으로 설정\n③ [E4:E15] 판매량 뒤에 "kg" 표시(#,##0"kg")\n④ [A3:F3] 채우기 색(표준 색-노랑) 및 가운데 맞춤\n⑤ [A3:F15] 모든 테두리 적용 후 바깥쪽 굵은 상자 테두리 지정',
        ans: '셀 서식(Ctrl+1) → 표시 형식: yyyy"년" mm"월" dd"일" 및 #,##0"kg" / 테두리 및 표준 색-노랑 적용'
      },
      {
        no: '기본작업-3 (5점 / 부분점수 없음)',
        inst: '[기본작업-3] 시트의 상공 포장이사 예약 현황([A4:H16])에서 "이사량(톤)"이 7 이상인 행 전체에 대해 글꼴 색 "파랑", 스타일 "굵게"가 적용되도록 수식을 이용한 조건부 서식을 작성하시오.',
        ans: '범위: $A$4:$H$16 (제목행 A3 제외) | 조건부 서식 수식: =$E4>=7'
      }
    ],
    calc: [
      {
        no: '계산작업 1번 (8점) - [표1] 셔틀버스 운행 시간표',
        inst: '출발시간([C3:C10])에 정류장수([B3:B10])×8분을 더하여 도착시간([D3:D10])을 계산하시오. (TIME 함수 사용)',
        ans: '=C3+TIME(0, B3*8, 0)'
      },
      {
        no: '계산작업 2번 (8점) - [표2] 영업부서별 실적표',
        inst: '매장위치([H3:H10])가 "백화점"이면서 판매실적([I3:I10])이 전체 평균 이상이면 "우수매장", 아니면 공백("")을 비고([J3:J10])에 표시하시오. (IF, AND, AVERAGE 함수 사용)',
        ans: '=IF(AND(H3="백화점", I3>=AVERAGE($I$3:$I$10)), "우수매장", "")'
      },
      {
        no: '계산작업 3번 (8점) - [표3] 영화예매현황',
        inst: '장르가 "드라마"인 영화들의 예매수량 평균을 구하여 백의 자리에서 반올림한 값을 [D24] 셀에 계산하시오. 조건은 [E22:E23]에 입력하시오. (ROUND, DAVERAGE 함수 사용)',
        ans: '=ROUND(DAVERAGE(A13:D23, 4, E22:E23), -2)  [조건 E22: 장르, E23: 드라마]'
      },
      {
        no: '계산작업 4번 (8점) - [표4] 아파트 당첨현황',
        inst: '[표3]의 영역을 참조하여 추첨자 코드에 대응하는 결과값을 [J14:J24]에 계산하시오. (INDEX, MATCH 함수 사용)',
        ans: '=INDEX($A$14:$D$23, MATCH(G14, $A$14:$A$23, 0), 4)'
      },
      {
        no: '계산작업 5번 (8점) - [표5] 자격증 수강등록 현황',
        inst: '수강코드([B28:B37])의 첫 글자가 1이면 "컴활1급", 2이면 "컴활2급", 3이면 "워드1급"을 수강과목([E28:E37])에 표시하시오. (CHOOSE, LEFT 함수 사용)',
        ans: '=CHOOSE(LEFT(B28, 1), "컴활1급", "컴활2급", "워드1급")'
      }
    ],
    analysis: [
      {
        no: '분석작업-1 (10점 / 부분점수 없음) - 목표값 찾기',
        inst: '4월 스마트폰 매출 현황([A3:E6])에서 순이익([E6])이 20,000,000이 되려면 판매수량([B4])이 얼마가 되어야 하는지 [목표값 찾기] 기능을 사용하여 계산하시오.',
        ans: '[데이터] → [가상 분석] → [목표값 찾기] | 수식 셀: E6, 찾는 값: 20000000, 값을 바꿀 셀: $B$4 (결과: 250)'
      },
      {
        no: '분석작업-2 (10점 / 부분점수 없음) - 데이터 통합',
        inst: '서울([A2:D6]), 경기([F2:I6]), 인천([A9:D13]) 영업실적 표의 1월~3월 실적을 합산하여 [표4] 수도권 영업실적([F9:I13]) 영역에 [데이터 통합] 기능으로 계산하시오.',
        ans: '[F9:I13] 범위 선택 → [데이터 통합] → 함수: 합계, 참조 추가 후 사용할 레이블 [첫 행], [왼쪽 열] 모두 체크'
      }
    ],
    etc: [
      {
        no: '매크로작업 1·2번 (10점 / 각 5점, 부분점수 있음)',
        inst: '① [G4:G13] 영역에 국어~과학탐구 총점(SUM)을 계산하는 매크로("총점")를 기록하고 [I3:J4] 양식 단추에 지정하시오.\n② [A3:G3] 채우기 색(표준 색-주황) 매크로("서식")를 기록하고 [I6:J7] 직사각형 도형에 지정하시오.',
        ans: '매크로1 수식: =SUM(B4:F4) / [개발 도구] → [매크로 기록] 후 양식 단추 및 직사각형 우클릭 → [매크로 지정]'
      },
      {
        no: '차트작업 1~5번 (10점 / 각 2점, 부분점수 있음)',
        inst: '① 데이터 범위에서 [젠가] 계열 제외\n② [3월] 계열을 "표식이 있는 꺾은선형"으로 변경 후 보조축 지정\n③ 차트 제목 "1분기 보드게임 매출 현황" (굴림, 16pt, 굵게)\n④ [루미큐브] 3월 요소에만 데이터 레이블(값) 위쪽 표시\n⑤ 범례 위치 "아래쪽" 및 차트 영역 "둥근 모서리" 설정',
        ans: '차트 데이터 선택(젠가 해제) → 계열 차트 종류 변경(3월 보조축) → 차트 영역 서식(둥근 모서리)'
      }
    ],
    sampleRows: [
      ['종류', '등급', '포장일자', '가격(200g)', '판매량(kg)', '총판매금액'],
      ['꽃등심', '1+', '2025-02-03', 22500, 85, 9562500],
      ['갈비살', '1등급', '2025-02-05', 21000, 88, 9240000],
      ['안심', '1+', '2025-02-03', 19800, 54, 5346000],
      ['채끝살', '2등급', '2025-02-07', 8200, 12, 492000]
    ]
  },

  '기본모의고사 01회': {
    id: '기본모의고사 01회',
    title: '기본모의고사 01회 (상공전자 인사급여·도서판매)',
    themeName: '상공전자 인사급여 및 베스트셀러 도서판매 현황',
    basic: [
      {
        no: '기본작업-1 (5점 / 부분점수 없음)',
        inst: '[기본작업-1] 시트의 [A3:E14] 영역에 상공전자 신입사원 연수 평가표를 오타 없이 정확히 입력하시오.',
        ans: '사번(SG-2026-01 ~ SG-2026-10) 대소문자 및 직무명 띄어쓰기 100% 일치 필수'
      },
      {
        no: '기본작업-2 (10점 / 5문항 각 2점 부분점수)',
        inst: '① [A1:F1] 병합 후 가운데 맞춤, 돋움 15pt 굵게, 밑줄(이중 실선)\n② [D4:D15] 기본급에 천 단위 구분 기호 및 "원" 표시(#,##0"원")\n③ [E4:E15] 셀 값이 0일 때 "-"로 표시하는 사용자 지정 서식\n④ [A3:F3] 배경색(표준 색-연한 녹색) 지정\n⑤ [F4:F15] 메모 삽입("인사팀 확인필") 및 항상 표시',
        ans: '사용자 지정 서식: #,##0"원";-#,##0"원";"-" / 메모 삽입 후 우클릭 → [메모 표시/숨기기]'
      },
      {
        no: '기본작업-3 (5점 / 부분점수 없음) - 고급 필터',
        inst: '[A3:G20] 영역에서 부서가 "개발팀"이거나 실적이 85점 이상인 데이터만 추출하여 [A25] 셀부터 표시하시오. (조건 범위: [B22:C24])',
        ans: 'B22(부서), C22(실적) / B23(개발팀), C24(>=85) — OR 조건이므로 서로 다른 행에 입력!'
      }
    ],
    calc: [
      {
        no: '계산작업 1번 (8점) - 직급수당 계산',
        inst: '사원코드([A3:A10])의 마지막 2글자를 추출하여 [수당기준표]([G3:I6])에서 직급수당을 찾아 표시하시오. (VLOOKUP, RIGHT, VALUE 함수 사용)',
        ans: '=VLOOKUP(VALUE(RIGHT(A3, 2)), $G$3:$I$6, 3, FALSE)'
      },
      {
        no: '계산작업 2번 (8점) - 승진심사 대상 판정',
        inst: '근속연수([C3:C10])가 5 이상이거나 고과점수([D3:D10])가 90 이상이면 "심사대상", 아니면 공백("")을 표시하시오. (IF, OR 함수 사용)',
        ans: '=IF(OR(C3>=5, D3>=90), "심사대상", "")'
      },
      {
        no: '계산작업 3번 (8점) - 영업1팀 급여 총액',
        inst: '[표3]에서 소속이 "영업1팀"인 사원들의 급여 합계를 [D24]에 구하시오. (DSUM 함수 사용, 조건범위 [E22:E23])',
        ans: '=DSUM(A13:D23, 4, E22:E23)'
      },
      {
        no: '계산작업 4번 (8점) - 실적 순위별 포상금',
        inst: '판매실적([C14:C22])의 내림차순 순위가 1위면 "금상", 2위면 "은상", 3위면 "동상", 나머지는 공백("")으로 표시하시오. (IFERROR, CHOOSE, RANK.EQ 함수 사용)',
        ans: '=IFERROR(CHOOSE(RANK.EQ(C14, $C$14:$C$22), "금상", "은상", "동상"), "")'
      },
      {
        no: '계산작업 5번 (8점) - 입사연월일 조합',
        inst: '주민번호 앞자리([B28:B36])에서 연, 월, 일을 추출하여 날짜 형태로 변환하시오. (DATE, LEFT, MID 함수 사용)',
        ans: '=DATE(1900+LEFT(B28,2), MID(B28,3,2), MID(B28,5,2))'
      }
    ],
    analysis: [
      {
        no: '분석작업-1 (10점 / 부분점수 없음) - 피벗 테이블',
        inst: '[A2:F25] 데이터를 이용하여 [A28] 셀에 피벗 테이블을 작성하시오. (행: 부서명, 열: 직급, 값: 기본급의 평균 및 상여금의 합계, 빈 셀은 "*" 표시)',
        ans: '[삽입] → [피벗 테이블] → 기존 워크시트(A28) → 피벗 테이블 옵션(빈 셀 표시: *)'
      },
      {
        no: '분석작업-2 (10점 / 부분점수 없음) - 부분합',
        inst: '[도서분류] 기준 오름차순 정렬 후, 도서분류별 [판매수량]의 합계와 [판매금액]의 평균을 순서대로 부분합으로 구하시오.',
        ans: '[데이터] → [정렬(도서분류 오름차순)] → [부분합 1차: 합계] → [부분합 2차: 평균(새로운 값으로 대치 해제!)]'
      }
    ],
    etc: [
      {
        no: '매크로작업 1·2번 (10점 / 부분점수 있음)',
        inst: '① [F4:F13]에 평균(=AVERAGE) 계산 매크로("평균계산") 기록 및 빗면 도형 연결\n② [D4:E13] 통화 서식 지정 매크로("통화서식") 기록 및 양식 단추 연결',
        ans: '빗면 도형 및 양식 단추에 각각 우클릭 → [매크로 지정] 연결'
      },
      {
        no: '차트작업 1~5번 (10점 / 부분점수 있음)',
        inst: '① 차트 종류를 "누적 세로 막대형"으로 변경\n② 차트 제목을 [A1] 셀과 연동\n③ 세로(값) 축 주 단위를 50,000으로 고정\n④ 데이터 테이블(범례 표지 포함)을 차트 하단에 표시\n⑤ 그림 영역 채우기 색(질감-파랑 박엽지) 적용',
        ans: '차트 제목 선택 후 수식줄에 =시트명!$A$1 입력 / 축 서식 → 주 단위(50000)'
      }
    ],
    sampleRows: [
      ['사번', '성명', '부서', '근속연수', '고과점수', '기본급'],
      ['SG-2601', '김민수', '영업1팀', 6, 92, 3450000],
      ['SG-2602', '이서연', '개발팀', 3, 88, 3100000],
      ['SG-2603', '박도현', '영업1팀', 7, 95, 3900000],
      ['SG-2604', '최유진', '인사팀', 2, 79, 2850000]
    ]
  },

  '기본모의고사 02회': {
    id: '기본모의고사 02회',
    title: '기본모의고사 02회 (대한물류 배송현황·해외수출)',
    themeName: '대한물류 화물운송 및 분기별 해외수출 실적',
    basic: [
      {
        no: '기본작업-1 (5점)',
        inst: '[A3:E15] 영역에 대한물류 권역별 화물 접수 내역을 정확히 입력하시오.',
        ans: '운송장번호 영문·하이픈(-) 및 숫자 정확히 일치 입력'
      },
      {
        no: '기본작업-2 (10점)',
        inst: '① [A1:G1] 병합 및 가운데 맞춤, 굴림 16pt 굵게\n② [E4:E15] 중량(톤)에 소수점 한 자리 표시(0.0"톤")\n③ [C3:C15] 열 너비 15로 설정\n④ [A3:G15] 사용자 지정 정렬(권역별 오름차순)\n⑤ 시트 탭 색상을 "표준 색-파랑"으로 변경',
        ans: '표시 형식: 0.0"톤" / 시트 탭 우클릭 → [탭 색] → 표준 색-파랑'
      },
      {
        no: '기본작업-3 (5점) - 조건부 서식',
        inst: '[A4:G16]에서 배송지역([C4])이 "제주"이거나 운임비([F4])가 150,000 이상인 행 전체에 채우기 색 "노랑"을 적용하시오.',
        ans: '조건부 서식 수식: =OR($C4="제주", $F4>=150000)'
      }
    ],
    calc: [
      {
        no: '계산작업 1번 (8점) - 최다운임 품목명 검색',
        inst: '운임비가 가장 높은 화물의 품목명을 [B4:F12]에서 찾아 표시하시오. (INDEX, MATCH, MAX 함수 사용)',
        ans: '=INDEX(B4:B12, MATCH(MAX(F4:F12), F4:F12, 0))'
      },
      {
        no: '계산작업 2번 (8점) - 다중 조건 운임 합계',
        inst: '권역이 "수도권"이면서 차량톤수가 5톤 이상인 화물의 운임비 합계를 구하시오. (SUMIFS 함수 사용)',
        ans: '=SUMIFS(F4:F15, C4:C15, "수도권", E4:E15, ">=5")'
      },
      {
        no: '계산작업 3번 (8점) - 우수 배송 건수 집계',
        inst: '배송소요시간이 24시간 이하이고 고객평점이 4.5 이상인 건수를 구해 뒤에 "건"을 붙이시오. (COUNTIFS 함수 및 & 연산자 사용)',
        ans: '=COUNTIFS(D14:D25, "<=24", E14:E25, ">=4.5") & "건"'
      },
      {
        no: '계산작업 4번 (8점) - 계약 유지 개월수',
        inst: '계약시작일([B28])과 종료일([C28]) 사이의 총 개월수를 계산하시오. (DATEDIF 함수 사용)',
        ans: '=DATEDIF(B28, C28, "m")'
      },
      {
        no: '계산작업 5번 (8점) - 관세 반올림 계산',
        inst: '수출금액이 평균 이상이면 수출금액의 8%, 아니면 5%를 계산하되 십의 자리에서 올림하시오. (ROUNDUP, IF, AVERAGE 함수 사용)',
        ans: '=ROUNDUP(B3*IF(B3>=AVERAGE($B$3:$B$10), 0.08, 0.05), -2)'
      }
    ],
    analysis: [
      {
        no: '분석작업-1 (10점) - 시나리오 관리자',
        inst: '유류할증료율([B4])이 15%로 인상될 때("할증인상")와 5%로 인하될 때("할증인하") 총운송원가([E8])의 변화를 시나리오 요약 보고서로 작성하시오.',
        ans: '[데이터] → [가상 분석] → [시나리오 관리자] → 시나리오 2개 추가 후 [요약(결과 셀: E8)]'
      },
      {
        no: '분석작업-2 (10점) - 데이터 통합',
        inst: '1~3분기 수출실적 표를 이용해 연간 합계를 [F10:I15]에 데이터 통합하시오.',
        ans: '[데이터 통합] → 함수: 합계, 첫 행·왼쪽 열 체크'
      }
    ],
    etc: [
      {
        no: '매크로작업 1·2번 (10점)',
        inst: '① [G4:G12] 운임합계 매크로 기록 및 타원 도형 연결\n② [A3:G3] 글꼴 굵게·기울임꼴 매크로 기록 및 단추 연결',
        ans: '매크로 기록 및 도형 매크로 지정'
      },
      {
        no: '차트작업 1~5번 (10점)',
        inst: '① [항공운송] 계열 차트 종류를 꺾은선형으로 변경 및 보조축 지정\n② 차트 제목 "권역별 물류 운송량" 입력\n③ 최고값 요소에 데이터 레이블 표시\n④ 눈금선 파선 지정\n⑤ 차트 스타일 5번 적용',
        ans: '차트 디자인 및 서식 탭에서 5가지 요소 순차 적용'
      }
    ],
    sampleRows: [
      ['운송장번호', '품목명', '권역', '배송시간', '차량톤수', '운임비'],
      ['DL-101', '반도체부품', '수도권', 12, 5, 240000],
      ['DL-102', '냉동식품', '영남권', 18, 8, 380000],
      ['DL-103', '의료기기', '수도권', 9, 5, 210000]
    ]
  },

  '기출문제 01회': {
    id: '기출문제 01회',
    title: '상시기출복원 01회 (문화센터 수강·호텔예약)',
    themeName: '행복문화센터 강좌 수강 및 오션호텔 객실예약 현황',
    basic: [
      {
        no: '기본작업-1 (5점)',
        inst: '[A3:E15] 영역에 문화센터 겨울학기 수강 신청자 명단을 정확히 입력하시오.',
        ans: '회원코드 및 강좌명 오타 없이 입력'
      },
      {
        no: '기본작업-2 (10점)',
        inst: '① [A1:F1] 병합 및 가운데 맞춤, 맑은 고딕 18pt\n② [D4:D15] 수강료 천 단위 기호 및 "원"\n③ [B4:B15] 글자 뒤에 "회원" 표시(@ "회원")\n④ [A3:F3] 굵은 아래쪽 테두리\n⑤ 행 높이 22 지정',
        ans: '텍스트 표시 형식: @"회원" / 숫자 표시 형식: #,##0"원"'
      },
      {
        no: '기본작업-3 (5점) - 자동 필터',
        inst: '[A3:F18]에서 강좌분류가 "어학"이고 수강료가 상위 30% 이상인 데이터만 자동 필터로 표시하시오.',
        ans: '[데이터] → [필터] → 강좌분류(어학) & 수강료(숫자 필터 → 상위 10 → 30 항목이 아닌 30% 선택)'
      }
    ],
    calc: [
      {
        no: '계산작업 1번 (8점) - 할인율 조견표 참조',
        inst: '객실코드([B3:B10])의 오른쪽 2글자를 이용해 [할인율표]([G2:J3])에서 할인율을 찾아 표시하시오. (HLOOKUP, RIGHT 함수 사용)',
        ans: '=HLOOKUP(RIGHT(B3, 2), $G$2:$J$3, 2, FALSE)'
      },
      {
        no: '계산작업 2번 (8점) - 요일별 주말할증 판정',
        inst: '체크인일자([C3:C10])가 토요일 또는 일요일이면 "주말요금", 아니면 "평일요금"을 표시하시오. (IF, WEEKDAY 함수 사용)',
        ans: '=IF(WEEKDAY(C3, 2)>=6, "주말요금", "평일요금")'
      },
      {
        no: '계산작업 3번 (8점) - VIP 회원 평균 결제액',
        inst: '등급이 "VIP"인 고객의 결제금액 평균을 구해 백 원 단위 미만 버림(절사)하시오. (TRUNC, DAVERAGE 함수 사용)',
        ans: '=TRUNC(DAVERAGE(A13:E23, 5, G13:G14), -2)'
      },
      {
        no: '계산작업 4번 (8점) - 조식 포함 여부 코드 변환',
        inst: '예약번호 4번째 글자가 "B"이면 "조식포함", "N"이면 "불포함"으로 표시하시오. (IF, MID 함수 사용)',
        ans: '=IF(MID(A14, 4, 1)="B", "조식포함", "불포함")'
      },
      {
        no: '계산작업 5번 (8점) - 최다 신청 강좌 정원 비율',
        inst: '전체 강좌 중 신청인원이 30명 이상인 강좌의 비율(%)을 구하시오. (COUNTIF, COUNTA 함수 사용)',
        ans: '=COUNTIF(E28:E37, ">=30") / COUNTA(E28:E37)'
      }
    ],
    analysis: [
      {
        no: '분석작업-1 (10점) - 부분합',
        inst: '[객실등급] 기준 내림차순 정렬 후, 객실등급별 [숙박일수]의 평균과 [결제금액]의 최대값을 부분합으로 구하시오.',
        ans: '[정렬(객실등급 내림차순)] → [부분합 1차: 평균] → [부분합 2차: 최대값]'
      },
      {
        no: '분석작업-2 (10점) - 목표값 찾기',
        inst: '호텔 총매출액([F12])이 50,000,000원이 되려면 스위트룸 객실단가([C6])가 얼마여야 하는지 목표값 찾기를 수행하시오.',
        ans: '수식 셀: F12, 찾는 값: 50000000, 값을 바꿀 셀: $C$6'
      }
    ],
    etc: [
      {
        no: '매크로작업 1·2번 (10점)',
        inst: '① [F4:F12] 결제금액(=숙박일수*일일요금) 매크로 기록 및 단추 연결\n② [A3:F12] 테두리 해제 매크로 기록 및 도형 연결',
        ans: '매크로 기록 및 컨트롤 매크로 지정'
      },
      {
        no: '차트작업 1~5번 (10점)',
        inst: '① 차트 제목 "월별 객실 가동률 현황"\n② [스위트룸] 계열 보조축 꺾은선형 변경\n③ 범례 위쪽 배치\n④ 세로축 최대값 100 고정\n⑤ 바깥쪽 그림자 효과 지정',
        ans: '차트 요소 추가 및 축 서식 설정'
      }
    ],
    sampleRows: [
      ['예약번호', '객실코드', '체크인일자', '숙박일수', '결제금액'],
      ['RV-B-101', 'RM-DX', '2026-07-18', 3, 450000],
      ['RV-N-102', 'RM-ST', '2026-07-19', 2, 520000],
      ['RV-B-103', 'RM-PR', '2026-07-21', 4, 880000]
    ]
  },

  '기출문제 02회': {
    id: '기출문제 02회',
    title: '상시기출복원 02회 (스마트팜 수확량·장학금 - 2026 신유형)',
    themeName: '그린스마트팜 작물 수확량 및 대학 장학금 사정표',
    basic: [
      {
        no: '기본작업-1 (5점)',
        inst: '[A3:E15] 영역에 스마트팜 동별 온도·습도 및 일일 수확량 기록을 입력하시오.',
        ans: '오타 및 소수점 자릿수 정확 일치 입력'
      },
      {
        no: '기본작업-2 (10점)',
        inst: '① [A1:F1] 병합 및 가운데 맞춤, 맑은 고딕 16pt\n② [E4:E15] 수확량 천 단위 기호 및 "박스"(#,##0"박스")\n③ [A3:F3] 채우기 색(표준 색-연한 파랑)\n④ [A4:A15] 가운데 맞춤\n⑤ 시트 보호(내용 및 개체 체크)',
        ans: '[검토] → [시트 보호] 적용'
      },
      {
        no: '기본작업-3 (5점) - 조건부 서식',
        inst: '수확량([E4:E16])이 전체 평균보다 큰 행 전체에 대해 글꼴 굵게, 채우기 색 "연한 녹색"을 적용하시오.',
        ans: '조건부 서식 수식: =$E4>AVERAGE($E$4:$E$16)'
      }
    ],
    calc: [
      {
        no: '계산작업 1번 (8점) - 2026 신유형 다중 조건 장학등급 (IFS)',
        inst: '학점평균([D3:D10])이 4.3 이상이면 "전액장학", 4.0 이상이면 "반액장학", 3.5 이상이면 "성적장학", 그 외에는 "일반"을 표시하시오. (IFS 함수 사용)',
        ans: '=IFS(D3>=4.3, "전액장학", D3>=4.0, "반액장학", D3>=3.5, "성적장학", TRUE, "일반")'
      },
      {
        no: '계산작업 2번 (8점) - 2026 신유형 작물코드 분기 (SWITCH)',
        inst: '작물코드([B3:B10])의 첫 글자가 "T"이면 "토마토", "S"이면 "딸기", "P"이면 "파프리카"로 표시하시오. (SWITCH, LEFT 함수 사용)',
        ans: '=SWITCH(LEFT(B3, 1), "T", "토마토", "S", "딸기", "P", "파프리카")'
      },
      {
        no: '계산작업 3번 (8점) - 출하 예정 영업일 계산 (WORKDAY)',
        inst: '파종일자([C14:C22])로부터 주말을 제외한 15영업일 뒤의 출하예정일을 계산하시오. (WORKDAY 함수 사용)',
        ans: '=WORKDAY(C14, 15)'
      },
      {
        no: '계산작업 4번 (8점) - 최고·최저 수확량 차이 (LARGE, SMALL)',
        inst: '2번째로 많은 수확량과 2번째로 적은 수확량의 차이를 [E24]에 계산하시오. (LARGE, SMALL 함수 사용)',
        ans: '=LARGE(E14:E23, 2) - SMALL(E14:E23, 2)'
      },
      {
        no: '계산작업 5번 (8점) - 정기점검 만료월 말일 (EOMONTH)',
        inst: '설치일자([B28:B35])로부터 6개월 후가 속한 달의 마지막 날짜(말일)를 계산하시오. (EOMONTH 함수 사용)',
        ans: '=EOMONTH(B28, 6)'
      }
    ],
    analysis: [
      {
        no: '분석작업-1 (10점) - 피벗 테이블',
        inst: '스마트팜 수확 데이터([A2:F22])를 이용하여 [A25] 셀에 피벗 테이블을 생성하고, 파종일자를 "월" 단위로 그룹화하시오.',
        ans: '피벗 테이블 생성 후 날짜 필드 우클릭 → [그룹] → "월" 선택'
      },
      {
        no: '분석작업-2 (10점) - 데이터 통합',
        inst: '1농장~3농장의 작물별 수확량 평균을 [F9:I13]에 데이터 통합으로 계산하시오.',
        ans: '[데이터 통합] → 함수: 평균(Average) 선택, 첫 행·왼쪽 열 체크'
      }
    ],
    etc: [
      {
        no: '매크로작업 1·2번 (10점)',
        inst: '① [F4:F12] 장학금 지급액 계산 매크로 기록 및 단추 지정\n② [D4:D12] 소수점 둘째 자리(0.00) 서식 매크로 기록 및 도형 지정',
        ans: '매크로 기록 및 도형 연결'
      },
      {
        no: '차트작업 1~5번 (10점)',
        inst: '① 차트 종류 "묶은 가로 막대형" 변경\n② 차트 제목 "농장별 스마트팜 수확 실적"\n③ 범례 아래쪽 배치\n④ 최대값 계열 요소에만 데이터 레이블 표시\n⑤ 그림 영역 테두리 실선 적용',
        ans: '차트 디자인 및 서식 설정'
      }
    ],
    sampleRows: [
      ['농장코드', '작물코드', '파종일자', '재배온도', '수확량(박스)', '품질점수'],
      ['FM-01', 'T-101', '2026-03-02', 23.5, 420, 4.4],
      ['FM-02', 'S-202', '2026-03-05', 19.0, 510, 4.1],
      ['FM-03', 'P-301', '2026-03-08', 25.0, 380, 3.8]
    ]
  }
};

// ============================================================================
// 2. Smart Problem Generator Engine (Creates new Mock Exam + Real 8-Sheet .xlsx)
// ============================================================================
const GENERATOR_THEMES = {
  mart: {
    name: '한빛마트 지점별 매출·재고 현황',
    org: '한빛유통',
    itemCol: '상품명',
    catCol: '카테고리',
    metricCol: '판매수량',
    priceCol: '판매금액',
    items: [
      ['HB-101', '유기농우유', '유제품', 85, 297500, '강남점'],
      ['HB-102', '프리미엄사과', '청과', 120, 1080000, '분당점'],
      ['HB-103', '한우등심세트', '정육', 45, 2475000, '강남점'],
      ['HB-104', '제주감귤박스', '청과', 95, 760000, '수원점'],
      ['HB-105', '그릭요거트', '유제품', 140, 630000, '인천점']
    ]
  },
  hospital: {
    name: '서울메디컬 진료·입원비 정산 현황',
    org: '서울메디컬센터',
    itemCol: '환자명',
    catCol: '진료과목',
    metricCol: '입원일수',
    priceCol: '진료비총액',
    items: [
      ['MD-201', '김태연', '내과', 7, 680000, '특실'],
      ['MD-202', '이준혁', '정형외과', 14, 1450000, '일반실'],
      ['MD-203', '박소담', '안과', 3, 420000, '특실'],
      ['MD-204', '정해인', '내과', 10, 950000, '일반실'],
      ['MD-205', '고윤정', '피부과', 2, 310000, '특실']
    ]
  },
  academy: {
    name: '미래IT아카데미 수강생 성적·장학금 현황',
    org: '미래IT교육원',
    itemCol: '수강생명',
    catCol: '수강과정',
    metricCol: '실기점수',
    priceCol: '수강료',
    items: [
      ['IT-301', '강하늘', '컴활2급실기', 92, 180000, '주말반'],
      ['IT-302', '배수지', '파이썬데이터', 88, 250000, '평일반'],
      ['IT-303', '남주혁', '컴활2급실기', 76, 180000, '평일반'],
      ['IT-304', '김지원', '웹프론트엔드', 95, 320000, '주말반'],
      ['IT-305', '박보검', '컴활2급실기', 64, 180000, '주말반']
    ]
  },
  logistics: {
    name: '대한특송 화물운송·배차 관리 현황',
    org: '대한글로벌물류',
    itemCol: '화물품목',
    catCol: '배송권역',
    metricCol: '운송거리(km)',
    priceCol: '운송요금',
    items: [
      ['LG-401', '노트북모니터', '수도권', 45, 350000, '5톤트럭'],
      ['LG-402', '신선식자재', '영남권', 320, 890000, '냉동탑차'],
      ['LG-403', '의류박스', '충청권', 140, 420000, '1톤트럭'],
      ['LG-404', '정밀반도체', '수도권', 60, 750000, '무진동차'],
      ['LG-405', '친환경가구', '호남권', 280, 680000, '5톤트럭']
    ]
  },
  hotel: {
    name: '오션뷰호텔 객실예약·멤버십 매출 현황',
    org: '오션뷰리조트',
    itemCol: '예약자명',
    catCol: '객실타입',
    metricCol: '숙박박수',
    priceCol: '결제금액',
    items: [
      ['HT-501', '유재석', '스위트룸', 4, 1280000, 'VIP'],
      ['HT-502', '장도연', '디럭스룸', 2, 460000, 'GOLD'],
      ['HT-503', '손석구', '프리미엄', 5, 1650000, 'VIP'],
      ['HT-504', '한소희', '디럭스룸', 3, 690000, 'SILVER'],
      ['HT-505', '이정재', '스위트룸', 6, 1920000, 'VIP']
    ]
  }
};

function generateCustomMockExam() {
  generatedCount += 1;
  const themeKeySel = document.getElementById('genThemeSelect')?.value || 'random';
  const funcPack = document.getElementById('genFuncPackSelect')?.value || 'standard';
  const analysisEl = document.getElementById('genAnalysisSelect');
  const analysisPack = analysisEl ? analysisEl.value : ['goal_consolidate', 'pivot_subtotal', 'scenario_sort'][generatedCount % 3];

  const themeKeys = Object.keys(GENERATOR_THEMES);
  const chosenThemeKey = themeKeySel === 'random'
    ? themeKeys[Math.floor(Math.random() * themeKeys.length)]
    : themeKeySel;
  const theme = GENERATOR_THEMES[chosenThemeKey];

  // Randomize numbers so every generation produces a unique exam
  const thresholdMetric = Math.floor(Math.random() * 40) + 50;
  const targetGoal = (Math.floor(Math.random() * 15) + 25) * 1000000;
  const bonusPct = (Math.floor(Math.random() * 5) + 6);

  const examId = `[AI생성 #${generatedCount}] ${theme.name.split(' ')[0]} 모의고사`;

  const funcSets = {
    standard: [
      {
        no: '계산작업 1번 (8점) - 조견표 참조 단가 계산 (VLOOKUP, LEFT)',
        inst: `[표1]에서 코드([A4:A8])의 앞 2글자를 추출하여 [기준단가표]([H4:J7])에서 기준금액을 찾아 표시하시오. (VLOOKUP, LEFT 함수 사용)`,
        ans: `=VLOOKUP(LEFT(A4, 2), $H$4:$J$7, 2, FALSE)`
      },
      {
        no: '계산작업 2번 (8점) - 우수 실적 판정 (IF, AND, AVERAGE)',
        inst: `[표2]에서 ${theme.metricCol}([D4:D8])이 ${thresholdMetric} 이상이면서 ${theme.priceCol}([E4:E8])이 전체 평균 이상이면 "우수대상", 아니면 공백("")을 표시하시오. (IF, AND, AVERAGE 함수 사용)`,
        ans: `=IF(AND(D4>=${thresholdMetric}, E4>=AVERAGE($E$4:$E$8)), "우수대상", "")`
      },
      {
        no: '계산작업 3번 (8점) - 조건부 평균 반올림 (ROUND, DAVERAGE)',
        inst: `[표3]에서 ${theme.catCol}이 "${theme.items[0][2]}"인 항목들의 ${theme.priceCol} 평균을 구해 천의 자리에서 반올림하시오. (ROUND, DAVERAGE 함수 사용)`,
        ans: `=ROUND(DAVERAGE(A13:E18, 5, G13:G14), -3)`
      },
      {
        no: '계산작업 4번 (8점) - 순위별 등급 부여 (CHOOSE, RANK.EQ)',
        inst: `${theme.priceCol}([E4:E8])이 높은 순서대로 1위면 "S등급", 2위면 "A등급", 3위면 "B등급", 그 외에는 "일반"을 표시하시오. (IFERROR, CHOOSE, RANK.EQ 함수 사용)`,
        ans: `=IFERROR(CHOOSE(RANK.EQ(E4, $E$4:$E$8), "S등급", "A등급", "B등급"), "일반")`
      },
      {
        no: '계산작업 5번 (8점) - 조건별 장려금 합산 (SUMIF)',
        inst: `구분([F4:F8])이 "${theme.items[0][5]}"인 데이터의 ${theme.priceCol} 합계를 구하시오. (SUMIF 함수 사용)`,
        ans: `=SUMIF(F4:F8, "${theme.items[0][5]}", E4:E8)`
      }
    ],
    lookup: [
      {
        no: '계산작업 1번 (8점) - 최대 실적 항목명 역추적 (INDEX, MATCH, MAX)',
        inst: `${theme.priceCol}([E4:E8])이 가장 높은 데이터의 ${theme.itemCol}([B4:B8])을 찾아 표시하시오. (INDEX, MATCH, MAX 함수 사용)`,
        ans: `=INDEX(B4:B8, MATCH(MAX(E4:E8), E4:E8, 0))`
      },
      {
        no: '계산작업 2번 (8점) - 가로형 할인율표 참조 (HLOOKUP, RIGHT)',
        inst: `코드([A4:A8])의 마지막 1글자를 숫자로 변환하여 [할인율표]([H2:K3])에서 할인율을 찾아 ${theme.priceCol}에 곱하시오. (HLOOKUP, RIGHT, VALUE 함수 사용)`,
        ans: `=E4 * HLOOKUP(VALUE(RIGHT(A4, 1)), $H$2:$K$3, 2, FALSE)`
      },
      {
        no: '계산작업 3번 (8점) - 코드 순번별 담당부서 (CHOOSE, MID)',
        inst: `코드([A4:A8])의 4번째 글자가 1이면 "1본부", 2이면 "2본부", 3이면 "3본부"로 표시하시오. (CHOOSE, MID 함수 사용)`,
        ans: `=CHOOSE(MID(A4, 4, 1), "1본부", "2본부", "3본부")`
      },
      {
        no: '계산작업 4번 (8점) - 오류 방지 단가 검색 (IFERROR, VLOOKUP)',
        inst: `상품코드를 참조표에서 찾되, 코드가 없으면 "미등록코드"라고 표시하시오. (IFERROR, VLOOKUP 함수 사용)`,
        ans: `=IFERROR(VLOOKUP(A4, $H$4:$J$10, 3, FALSE), "미등록코드")`
      },
      {
        no: '계산작업 5번 (8점) - 2차원 교차표 값 추출 (INDEX, MATCH 중첩)',
        inst: `행 번호와 열 번호를 모두 MATCH 함수로 찾아 교차 구간의 수수료를 추출하시오. (INDEX, MATCH 함수 사용)`,
        ans: `=INDEX($H$4:$J$7, MATCH(B4, $G$4:$G$7, 0), MATCH(C4, $H$3:$J$3, 0))`
      }
    ],
    new2026: [
      {
        no: '계산작업 1번 (8점) - [2026 신유형] 다중 구간 등급 판정 (IFS)',
        inst: `${theme.metricCol}([D4:D8])이 100 이상이면 "최우수", 70 이상이면 "우수", 40 이상이면 "보통", 그 외에는 "노력요함"을 표시하시오. (IFS 함수 사용)`,
        ans: `=IFS(D4>=100, "최우수", D4>=70, "우수", D4>=40, "보통", TRUE, "노력요함")`
      },
      {
        no: '계산작업 2번 (8점) - [2026 신유형] 코드별 명칭 분기 (SWITCH, LEFT)',
        inst: `코드([A4:A8])의 앞 2글자를 비교하여 "${theme.items[0][0].slice(0,2)}"이면 "${theme.org} 본사", 그 외에는 "지점"으로 표시하시오. (SWITCH, LEFT 함수 사용)`,
        ans: `=SWITCH(LEFT(A4, 2), "${theme.items[0][0].slice(0,2)}", "${theme.org} 본사", "지점")`
      },
      {
        no: '계산작업 3번 (8점) - 영업일 기준 정산일 계산 (WORKDAY)',
        inst: `접수일([C4:C8])로부터 주말을 제외한 10영업일 후의 정산예정일을 구하시오. (WORKDAY 함수 사용)`,
        ans: `=WORKDAY(C4, 10)`
      },
      {
        no: '계산작업 4번 (8점) - 계약 만료월 말일 계산 (EOMONTH)',
        inst: `기준일([C4:C8])로부터 3개월 후가 속한 달의 말일 날짜를 구하시오. (EOMONTH 함수 사용)`,
        ans: `=EOMONTH(C4, 3)`
      },
      {
        no: '계산작업 5번 (8점) - 최빈값 및 중앙값 차이 (MODE.SNGL, MEDIAN)',
        inst: `실적 데이터의 최빈값과 중앙값의 차이를 계산하시오. (MODE.SNGL, MEDIAN 함수 사용)`,
        ans: `=MODE.SNGL(D4:D15) - MEDIAN(D4:D15)`
      }
    ],
    stat: [
      {
        no: '계산작업 1번 (8점) - 다중 조건 합계 (SUMIFS)',
        inst: `${theme.catCol}이 "${theme.items[0][2]}"이면서 ${theme.metricCol}이 ${thresholdMetric} 이상인 ${theme.priceCol} 합계를 구하시오. (SUMIFS 함수 사용)`,
        ans: `=SUMIFS(E4:E8, C4:C8, "${theme.items[0][2]}", D4:D8, ">=${thresholdMetric}")`
      },
      {
        no: '계산작업 2번 (8점) - 다중 조건 건수 및 단위 결합 (COUNTIFS, &)',
        inst: `${theme.priceCol}이 500,000 이상이고 구분이 "${theme.items[0][5]}"인 건수를 구해 뒤에 "건"을 붙이시오. (COUNTIFS 함수, & 연산자 사용)`,
        ans: `=COUNTIFS(E4:E8, ">=500000", F4:F8, "${theme.items[0][5]}") & "건"`
      },
      {
        no: '계산작업 3번 (8점) - 조건부 평균 정밀 계산 (ROUND, AVERAGEIFS)',
        inst: `${theme.metricCol}이 50 이상인 데이터들의 ${theme.priceCol} 평균을 구해 백의 자리에서 반올림하시오. (ROUND, AVERAGEIFS 함수 사용)`,
        ans: `=ROUND(AVERAGEIFS(E4:E8, D4:D8, ">=50"), -2)`
      },
      {
        no: '계산작업 4번 (8점) - 상위 2위와 하위 2위 편차 (LARGE, SMALL)',
        inst: `두 번째로 큰 ${theme.priceCol}과 두 번째로 작은 ${theme.priceCol}의 차이를 계산하시오. (LARGE, SMALL 함수 사용)`,
        ans: `=LARGE(E4:E8, 2) - SMALL(E4:E8, 2)`
      },
      {
        no: '계산작업 5번 (8점) - 내림차순 순위 표시 (RANK.EQ, &)',
        inst: `${theme.priceCol}([E4:E8])의 전체 순위를 구하여 "1위" 형태로 표시하시오. (RANK.EQ 함수 및 & 연산자 사용)`,
        ans: `=RANK.EQ(E4, $E$4:$E$8, 0) & "위"`
      }
    ]
  };

  const analysisSets = {
    goal_consolidate: [
      {
        no: '분석작업-1 (10점 / 부분점수 없음) - 목표값 찾기',
        inst: `${theme.org} 총순이익([E6])이 ${targetGoal.toLocaleString()}원이 되려면 목표 ${theme.metricCol}([B4])이 얼마가 되어야 하는지 [목표값 찾기]로 계산하시오.`,
        ans: `[데이터] → [가상 분석] → [목표값 찾기] | 수식 셀: E6, 찾는 값: ${targetGoal}, 값을 바꿀 셀: $B$4`
      },
      {
        no: '분석작업-2 (10점 / 부분점수 없음) - 데이터 통합',
        inst: `1~3지점의 월별 ${theme.priceCol} 표를 합산하여 [F9:I13] 영역에 [데이터 통합] 기능으로 계산하시오.`,
        ans: `[F9:I13] 드래그 → [데이터 통합] → 함수: 합계, 첫 행·왼쪽 열 모두 체크`
      }
    ],
    pivot_subtotal: [
      {
        no: '분석작업-1 (10점 / 부분점수 없음) - 피벗 테이블',
        inst: `[A3:F18] 데이터를 이용해 [A22] 셀에 피벗 테이블을 작성하시오. (행: ${theme.catCol}, 열: 구분, 값: ${theme.priceCol}의 합계 및 ${theme.metricCol}의 평균)`,
        ans: `[삽입] → [피벗 테이블] → 위치: 기존 워크시트(A22) → 행/열/값 배치`
      },
      {
        no: '분석작업-2 (10점 / 부분점수 없음) - 부분합',
        inst: `[${theme.catCol}] 기준 오름차순 정렬 후, ${theme.catCol}별 [${theme.metricCol}]의 합계와 [${theme.priceCol}]의 최대값을 순서대로 부분합으로 구하시오.`,
        ans: `[정렬(${theme.catCol} 오름차순)] → [부분합 1차: 합계] → [부분합 2차: 최대값(새로운 값으로 대치 해제)]`
      }
    ],
    scenario_sort: [
      {
        no: '분석작업-1 (10점 / 부분점수 없음) - 시나리오 관리자',
        inst: `인센티브율([B5])이 ${bonusPct + 5}%로 인상될 때("최대인상")와 ${Math.max(1, bonusPct - 3)}%로 인하될 때("최소인하") 총지급액([E8])의 변동을 시나리오 요약 보고서로 작성하시오.`,
        ans: `[가상 분석] → [시나리오 관리자] → 변경 셀: $B$5, 결과 셀: $E$8 요약 생성`
      },
      {
        no: '분석작업-2 (10점 / 부분점수 없음) - 다중 기준 정렬 및 자동필터',
        inst: `1차 기준 [${theme.catCol}] 오름차순, 2차 기준 [${theme.priceCol}] 내림차순으로 정렬하시오.`,
        ans: `[데이터] → [정렬] → 기준 추가하여 1차 오름차순, 2차 내림차순 설정`
      }
    ]
  };

  const newExam = {
    id: examId,
    title: `${examId} (${theme.name})`,
    themeName: theme.name,
    basic: [
      {
        no: '기본작업-1 (5점 / 부분점수 없음)',
        inst: `[기본작업-1] 시트의 [A3:F8] 영역에 제시된 ${theme.name} 기초 데이터를 오타·대소문자 오류 없이 입력하시오.`,
        ans: `${theme.items[0][0]} ~ ${theme.items[4][0]} 코드 및 ${theme.itemCol} 텍스트 100% 일치 입력`
      },
      {
        no: '기본작업-2 (10점 / 5문항 각 2점 부분점수)',
        inst: `① [A1:F1] 병합하고 가운데 맞춤, 맑은 고딕 16pt 굵게\n② [D4:D8] ${theme.metricCol}에 사용자 지정 단위 표시(#,##0"건")\n③ [E4:E8] ${theme.priceCol}에 천 단위 기호 및 "원"(#,##0"원")\n④ [A3:F3] 채우기 색(표준 색-노랑) 및 굵게\n⑤ [A3:F8] 모든 테두리 및 바깥쪽 굵은 상자 테두리 적용`,
        ans: `셀 서식(Ctrl+1) 표시 형식 및 테두리·채우기 5개 항목 순차 적용`
      },
      {
        no: '기본작업-3 (5점 / 부분점수 없음) - 조건부 서식',
        inst: `[A4:F8] 영역에서 ${theme.metricCol}([D4])이 ${thresholdMetric} 이상인 행 전체에 글꼴 색 "파랑", 스타일 "굵게"를 적용하는 수식 조건부 서식을 작성하시오.`,
        ans: `적용 범위: $A$4:$F$8 | 조건부 서식 수식: =$D4>=${thresholdMetric}`
      }
    ],
    calc: funcSets[funcPack] || funcSets.standard,
    analysis: analysisSets[analysisPack] || analysisSets.goal_consolidate,
    etc: [
      {
        no: '매크로작업 1·2번 (10점 / 각 5점, 부분점수 있음)',
        inst: `① [G4:G8] 영역에 부가세 포함 금액(=E4*1.${bonusPct < 10 ? '0' + bonusPct : bonusPct}) 계산 매크로("정산계산") 기록 및 양식 단추 지정\n② [A3:G3] 채우기 색(표준 색-주황) 매크로("표머리글서식") 기록 및 도형 지정`,
        ans: `매크로 기록 후 양식 단추 및 도형 우클릭 → [매크로 지정]`
      },
      {
        no: '차트작업 1~5번 (10점 / 각 2점, 부분점수 있음)',
        inst: `① 차트 제목을 "${theme.name}"으로 변경(굴림 16pt 굵게)\n② [${theme.priceCol}] 계열을 "표식이 있는 꺾은선형"으로 변경 후 보조축 지정\n③ 최대값 요소에만 데이터 레이블(값) 표시\n④ 범례를 "아래쪽"으로 이동\n⑤ 차트 영역 테두리 "둥근 모서리" 적용`,
        ans: `계열 차트 종류 변경(보조축) + 차트 제목 + 데이터 레이블 + 범례 + 둥근 모서리`
      }
    ],
    sampleRows: [
      ['코드', theme.itemCol, theme.catCol, theme.metricCol, theme.priceCol, '구분'],
      ...theme.items
    ]
  };

  MOCK_EXAM_BANK[examId] = newExam;
  saveExamToSupabase(newExam);

  // Add to left card dropdown and select it
  const roundSelect = document.getElementById('roundSelect');
  const opt = document.createElement('option');
  opt.value = examId;
  opt.textContent = newExam.title;
  roundSelect.prepend(opt);
  roundSelect.value = examId;

  renderExamPaper(newExam);
  return newExam;
}

function renderExamPaper(examSpec) {
  currentExamSpec = examSpec;
  document.getElementById('examPaperTitle').textContent = `${examSpec.title} — 실기 시험 문제지 (제한시간: 40분 / 합격기준: 70점 이상)`;
  document.getElementById('examPaperBadge').textContent = examSpec.id.startsWith('[AI생성') ? '✨ AI 신규 생성 완료' : '📘 기출·모의고사 수록본';

  const container = document.getElementById('examPaperContainer');
  const ansClass = showAnswersAll ? 'exam-q-ans show' : 'exam-q-ans';

  function renderDomainBox(title, badgeText, items) {
    return `
      <div class="exam-domain-box">
        <div class="exam-domain-title">
          <span>${title}</span>
          <span style="font-size:11.5px;background:#EFF6FF;color:#1E40AF;border:1px solid #93C5FD;padding:2px 9px;border-radius:999px;">${badgeText}</span>
        </div>
        ${items.map(item => `
          <div class="exam-q-item">
            <div style="font-weight:500;color:#0F172A;margin-bottom:4px;"><strong>📌 ${item.no}</strong></div>
            <div style="color:#1E293B;white-space:pre-line;line-height:1.55;">${item.inst}</div>
            <div class="${ansClass}"><strong>✅ 모범 정답 공식/해설:</strong> ${item.ans}</div>
          </div>
        `).join('')}
      </div>
    `;
  }

  container.innerHTML = [
    renderDomainBox('1️⃣ 문제1: 기본작업 (시트 1~3)', '배점 20점 (5+10+5)', examSpec.basic),
    renderDomainBox('2️⃣ 문제2: 계산작업 (시트 4 - 함수 5문항)', '배점 40점 (각 8점 × 5)', examSpec.calc),
    renderDomainBox('3️⃣ 문제3: 분석작업 (시트 5~6)', '배점 20점 (각 10점 × 2)', examSpec.analysis),
    renderDomainBox('4️⃣ 문제4: 기타작업 (시트 7 매크로 · 시트 8 차트)', '배점 20점 (10+10)', examSpec.etc)
  ].join('');
}

// ============================================================================
// Pure JS Real 8-Sheet OpenXML (.xlsx) Generator for Any Selected/Generated Exam!
// ============================================================================
function buildReal8SheetXlsxBlob(examSpec, includeAnswers = false) {
  const sheetNames = [
    '기본작업-1',
    '기본작업-2',
    '기본작업-3',
    '계산작업',
    '분석작업-1',
    '분석작업-2',
    '매크로작업',
    '차트작업'
  ];

  function esc(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  const colLetters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
  function rowsToSheetXml(titleText, tableRows, extraNoteRows = []) {
    let xml = `<row r="1"><c r="A1" t="inlineStr"><is><t>${esc(titleText)}</t></is></c></row>`;
    tableRows.forEach((r, rIdx) => {
      const rowNum = rIdx + 3;
      xml += `<row r="${rowNum}">`;
      r.forEach((val, cIdx) => {
        const ref = `${colLetters[cIdx]}${rowNum}`;
        if (typeof val === 'number') {
          xml += `<c r="${ref}"><v>${val}</v></c>`;
        } else {
          xml += `<c r="${ref}" t="inlineStr"><is><t>${esc(val)}</t></is></c>`;
        }
      });
      xml += `</row>`;
    });
    extraNoteRows.forEach((note, nIdx) => {
      const rowNum = tableRows.length + 5 + nIdx;
      xml += `<row r="${rowNum}"><c r="A${rowNum}" t="inlineStr"><is><t>${esc(note)}</t></is></c></row>`;
    });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${xml}</sheetData>
</worksheet>`;
  }

  const files = [
    {
      name: '[Content_Types].xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${sheetNames.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n  ')}
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
    ${sheetNames.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('\n    ')}
  </sheets>
</workbook>`
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetNames.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('\n  ')}
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    },
    {
      name: 'xl/styles.xml',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><sz val="11"/><name val="맑은 고딕"/><family val="2"/><charset val="129"/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`
    }
  ];

  // Populate all 8 sheets with the exam's sample data & instructions (and answers if includeAnswers=true)
  sheetNames.forEach((sName, idx) => {
    let notes = [];
    if (idx === 0) notes = [examSpec.basic[0].inst, ...(includeAnswers ? ['[정답] ' + examSpec.basic[0].ans] : [])];
    else if (idx === 1) notes = [examSpec.basic[1].inst, ...(includeAnswers ? ['[정답] ' + examSpec.basic[1].ans] : [])];
    else if (idx === 2) notes = [examSpec.basic[2].inst, ...(includeAnswers ? ['[정답] ' + examSpec.basic[2].ans] : [])];
    else if (idx === 3) {
      notes = examSpec.calc.map(c => `${c.no}: ${c.inst}${includeAnswers ? '  ==> 정답수식: ' + c.ans : ''}`);
    } else if (idx === 4) notes = [examSpec.analysis[0].inst, ...(includeAnswers ? ['[정답] ' + examSpec.analysis[0].ans] : [])];
    else if (idx === 5) notes = [examSpec.analysis[1].inst, ...(includeAnswers ? ['[정답] ' + examSpec.analysis[1].ans] : [])];
    else if (idx === 6) notes = [examSpec.etc[0].inst, ...(includeAnswers ? ['[정답] ' + examSpec.etc[0].ans] : [])];
    else if (idx === 7) notes = [examSpec.etc[1].inst, ...(includeAnswers ? ['[정답] ' + examSpec.etc[1].ans] : [])];

    const rows = (idx === 0 && !includeAnswers)
      ? [examSpec.sampleRows[0]] // Basic-1 problem file leaves rows blank for typing practice!
      : examSpec.sampleRows;

    files.push({
      name: `xl/worksheets/sheet${idx + 1}.xml`,
      data: rowsToSheetXml(`${examSpec.themeName} - [${sName}]`, rows, notes)
    });
  });

  // Pack as ZIP (Store method 0 with CRC32 so zero external libraries needed in browser)
  const encoder = new TextEncoder();
  const crcTable = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    crcTable[n] = c >>> 0;
  }
  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = crcTable[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  const localChunks = [];
  const centralChunks = [];
  let offset = 0;

  for (const f of files) {
    const nameBytes = encoder.encode(f.name);
    const dataBytes = encoder.encode(f.data);
    const crc = crc32(dataBytes);

    const lh = new Uint8Array(30 + nameBytes.length);
    const lhv = new DataView(lh.buffer);
    lhv.setUint32(0, 0x04034b50, true);
    lhv.setUint16(4, 20, true);
    lhv.setUint16(6, 0, true);
    lhv.setUint16(8, 0, true); // Store (uncompressed)
    lhv.setUint32(14, crc, true);
    lhv.setUint32(18, dataBytes.length, true);
    lhv.setUint32(22, dataBytes.length, true);
    lhv.setUint16(26, nameBytes.length, true);
    lh.set(nameBytes, 30);

    const ch = new Uint8Array(46 + nameBytes.length);
    const chv = new DataView(ch.buffer);
    chv.setUint32(0, 0x02014b50, true);
    chv.setUint16(4, 20, true);
    chv.setUint16(6, 20, true);
    chv.setUint16(10, 0, true);
    chv.setUint32(16, crc, true);
    chv.setUint32(20, dataBytes.length, true);
    chv.setUint32(24, dataBytes.length, true);
    chv.setUint16(28, nameBytes.length, true);
    chv.setUint32(42, offset, true);
    ch.set(nameBytes, 46);

    localChunks.push(lh, dataBytes);
    centralChunks.push(ch);
    offset += lh.length + dataBytes.length;
  }

  const cdSize = centralChunks.reduce((a, b) => a + b.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, offset, true);

  return new Blob([...localChunks, ...centralChunks, eocd], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

function downloadExamWorkbook(includeAnswers = false) {
  const spec = currentExamSpec || MOCK_EXAM_BANK['대표기출따라하기'];
  const blob = buildReal8SheetXlsxBlob(spec, includeAnswers);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const suffix = includeAnswers ? '모범정답포함' : '수험생문제용';
  const cleanTitle = spec.id.replace(/[^\w가-힣-]/g, '_');
  a.href = url;
  a.download = `컴활2급실기_${cleanTitle}_${suffix}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ============================================================================
// Existing Grading State & Presets
// ============================================================================
const localState = {
  participantCount: 2103,
  averageScore: 39.01,
  studentRankings: [
    { name: '짱******', rawName: '짱아', score: 2040 },
    { name: '박******', rawName: '박지호', score: 1890 },
    { name: 's******', rawName: 'sky2026', score: 1865 },
    { name: 'e******', rawName: 'excelKing', score: 1780 },
    { name: '김******', rawName: '김나경', score: 1735 },
    { name: 'g******', rawName: 'goodpass', score: 1520 },
    { name: '김******', rawName: '김영민', score: 1510 },
    { name: 's******', rawName: 'star99', score: 1480 },
    { name: 'S******', rawName: 'Sujin', score: 1450 },
    { name: '임******', rawName: '임홍진', score: 1385 }
  ],
  teacherRankings: [
    { id: 'Z******', rawId: 'Zzang_AA', score: 775 },
    { id: '0******', rawId: '0123academy', score: 645 },
    { id: 'j******', rawId: 'jjang_t', score: 350 },
    { id: 'i******', rawId: 'itmaster', score: 285 },
    { id: '아******', rawId: '아이티학원', score: 245 }
  ],
  distributions: {
    total: [
      { label: '0~10점미만', value: 49.3, color: '#1E3A8A' },
      { label: '10~20점미만', value: 1.8, color: '#1E40AF' },
      { label: '20~30점미만', value: 1.5, color: '#2563EB' },
      { label: '30~40점미만', value: 2.1, color: '#38BDF8' },
      { label: '40~50점미만', value: 2.9, color: '#60A5FA' },
      { label: '50~60점미만', value: 3.4, color: '#93C5FD' },
      { label: '60~70점미만', value: 5.7, color: '#D97706' },
      { label: '70~80점미만', value: 9.5, color: '#059669' },
      { label: '80~90점미만', value: 13.8, color: '#2563EB' },
      { label: '90~100점미만', value: 6.2, color: '#1E40AF' },
      { label: '100', value: 3.8, color: '#1E3A8A' }
    ],
    p1: [
      { label: '0~5점미만', value: 48.9, color: '#1E3A8A' },
      { label: '5~10점미만', value: 1.1, color: '#1E40AF' },
      { label: '10~15점미만', value: 11.6, color: '#60A5FA' },
      { label: '15~20점미만', value: 24.5, color: '#2563EB' },
      { label: '20', value: 13.9, color: '#38BDF8' }
    ],
    p2: [
      { label: '0~5점미만', value: 54.4, color: '#1E3A8A' },
      { label: '10~15점미만', value: 3.2, color: '#1E40AF' },
      { label: '15~20점미만', value: 2.8, color: '#60A5FA' },
      { label: '25~30점미만', value: 5.5, color: '#93C5FD' },
      { label: '30~35점미만', value: 13.8, color: '#2563EB' },
      { label: '40', value: 20.3, color: '#38BDF8' }
    ],
    p3: [
      { label: '0~5점미만', value: 51.1, color: '#1E3A8A' },
      { label: '10~15점미만', value: 10.3, color: '#60A5FA' },
      { label: '20', value: 38.6, color: '#2563EB' }
    ],
    p4: [
      { label: '0~5점미만', value: 51.5, color: '#1E3A8A' },
      { label: '5~10점미만', value: 2.1, color: '#1E40AF' },
      { label: '10~15점미만', value: 2.8, color: '#93C5FD' },
      { label: '15~20점미만', value: 7.2, color: '#60A5FA' },
      { label: '20', value: 36.4, color: '#2563EB' }
    ]
  },
  submissions: [
    {
      studentName: '짱아',
      teacherId: 'Zzang_AA',
      roundName: '대표기출따라하기',
      basic: 18,
      calc: 32,
      analysis: 20,
      etc: 16,
      total: 86,
      weakPoint: '계산작업 (CHOOSE·LEFT)',
      submittedAt: '2026-11-18 14:20:11',
      radarStats: { '데이터입력·서식': 87, '조건부서식·필터': 100, '함수(계산작업)': 80, '데이터분석(피벗·통합)': 100, '매크로·VBA': 80, '차트작업': 80 }
    },
    {
      studentName: '짱아',
      teacherId: 'Zzang_AA',
      roundName: '기본모의고사 01회',
      basic: 16,
      calc: 24,
      analysis: 10,
      etc: 18,
      total: 68,
      weakPoint: '분석작업 (데이터 통합)',
      submittedAt: '2026-11-17 19:05:42',
      radarStats: { '데이터입력·서식': 80, '조건부서식·필터': 80, '함수(계산작업)': 60, '데이터분석(피벗·통합)': 50, '매크로·VBA': 90, '차트작업': 90 }
    },
    {
      studentName: '짱아',
      teacherId: 'Zzang_AA',
      roundName: '기출문제 01회',
      basic: 20,
      calc: 32,
      analysis: 20,
      etc: 18,
      total: 90,
      weakPoint: '차트작업 (보조축 서식)',
      submittedAt: '2026-11-18 18:12:30',
      radarStats: { '데이터입력·서식': 100, '조건부서식·필터': 100, '함수(계산작업)': 80, '데이터분석(피벗·통합)': 100, '매크로·VBA': 100, '차트작업': 80 }
    },
    {
      studentName: '박지호',
      teacherId: 'Zzang_AA',
      roundName: '대표기출따라하기',
      basic: 16,
      calc: 24,
      analysis: 20,
      etc: 14,
      total: 74,
      weakPoint: '계산작업 (INDEX·MATCH)',
      submittedAt: '2026-11-18 15:40:00',
      radarStats: { '데이터입력·서식': 75, '조건부서식·필터': 100, '함수(계산작업)': 60, '데이터분석(피벗·통합)': 100, '매크로·VBA': 70, '차트작업': 70 }
    },
    {
      studentName: '김나경',
      teacherId: 'Zzang_AA',
      roundName: '대표기출따라하기',
      basic: 11,
      calc: 16,
      analysis: 10,
      etc: 14,
      total: 51,
      weakPoint: '기본작업-1 오타 및 함수 수식 누락',
      submittedAt: '2026-11-18 16:15:22',
      radarStats: { '데이터입력·서식': 60, '조건부서식·필터': 20, '함수(계산작업)': 40, '데이터분석(피벗·통합)': 50, '매크로·VBA': 50, '차트작업': 80 }
    },
    {
      studentName: '최수지',
      teacherId: 'Zzang_AA',
      roundName: '대표기출따라하기',
      basic: 20,
      calc: 40,
      analysis: 20,
      etc: 18,
      total: 98,
      weakPoint: '전 영역 우수 (안정권)',
      submittedAt: '2026-11-18 17:02:09',
      radarStats: { '데이터입력·서식': 100, '조건부서식·필터': 100, '함수(계산작업)': 100, '데이터분석(피벗·통합)': 100, '매크로·VBA': 100, '차트작업': 80 }
    },
    {
      studentName: '민준호',
      teacherId: 'Zzang_AA',
      roundName: '대표기출따라하기',
      basic: 13,
      calc: 16,
      analysis: 10,
      etc: 10,
      total: 49,
      weakPoint: '계산작업 및 분석작업-2 집중 보완 필요',
      submittedAt: '2026-11-18 17:45:50',
      radarStats: { '데이터입력·서식': 60, '조건부서식·필터': 60, '함수(계산작업)': 40, '데이터분석(피벗·통합)': 50, '매크로·VBA': 50, '차트작업': 50 }
    }
  ]
};

const PRESET_SCENARIOS = {
  raw0: {
    fileName: '문제샘플(작성 전).xlsm',
    preflight: {
      ok: true,
      checks: [
        { label: '필수 8개 시트 구조 검사', status: 'pass', detail: '8개 정규 시트(기본1~3, 계산, 분석1~2, 매크로, 차트) 모두 정상 확인됨' },
        { label: '기본작업-1 데이터 입력 감지', status: 'warn', detail: '데이터 미입력 상태 (1개 제목 셀만 존재 - 5점 감점 주의)' },
        { label: '계산작업 수식(=) 작성 검사', status: 'fail', detail: '5대 대표 계산 영역 중 0개 영역에 수식(=) 입력됨 (전 문항 미작성)' },
        { label: '매크로(VBA) 및 도형 단추 검사', status: 'warn', detail: '매크로 모듈 미감지 (작성 전 원본 파일입니다)' }
      ]
    },
    gradeResult: buildPresetGradeResult('raw0')
  },
  fail51: {
    fileName: '학생제출샘플_51점(불합격_오답다수).xlsm',
    preflight: {
      ok: true,
      checks: [
        { label: '필수 8개 시트 구조 검사', status: 'pass', detail: '8개 정규 시트(기본1~3, 계산, 분석1~2, 매크로, 차트) 모두 정상 확인됨' },
        { label: '기본작업-1 데이터 입력 감지', status: 'pass', detail: '입력된 데이터 셀 확인됨 (단, 영문 대소문자 오타 주의)' },
        { label: '계산작업 수식(=) 작성 검사', status: 'warn', detail: '주의: [J14] 셀에 수식(=) 없이 결과값(15)만 직접 타이핑한 문항 1건 감지 (0점 처리 대상)' },
        { label: '매크로(VBA) 및 도형 단추 검사', status: 'pass', detail: '매크로1(총점) 계산 기록 확인됨 / 매크로2(서식) 누락' }
      ]
    },
    gradeResult: buildPresetGradeResult('fail51')
  },
  pass82: {
    fileName: '학생제출샘플_82점(합격).xlsm',
    preflight: {
      ok: true,
      checks: [
        { label: '필수 8개 시트 구조 검사', status: 'pass', detail: '8개 정규 시트(기본1~3, 계산, 분석1~2, 매크로, 차트) 모두 정상 확인됨' },
        { label: '기본작업-1 데이터 입력 감지', status: 'pass', detail: '외부 데이터 [A3:E15] 오타 없이 정상 입력됨' },
        { label: '계산작업 수식(=) 작성 검사', status: 'pass', detail: '5개 함수 문제 중 4개 영역에 정규 수식(=) 작성 확인됨' },
        { label: '매크로(VBA) 및 도형 단추 검사', status: 'pass', detail: '매크로 기록 및 양식 단추 연결 확인됨' }
      ]
    },
    gradeResult: buildPresetGradeResult('pass82')
  },
  perfect100: {
    fileName: '학생제출샘플_100점(모범정답).xlsm',
    preflight: {
      ok: true,
      checks: [
        { label: '필수 8개 시트 구조 검사', status: 'pass', detail: '8개 정규 시트(기본1~3, 계산, 분석1~2, 매크로, 차트) 모두 정상 확인됨' },
        { label: '기본작업-1 데이터 입력 감지', status: 'pass', detail: '외부 데이터 전체 셀 100% 정확 입력 확인됨' },
        { label: '계산작업 수식(=) 작성 검사', status: 'pass', detail: '5대 계산 영역 모두 지정 함수 수식(=) 정상 작성됨' },
        { label: '매크로(VBA) 및 도형 단추 검사', status: 'pass', detail: '총점·서식 매크로 2종 및 도형 연결 100% 정상 확인됨' }
      ]
    },
    gradeResult: buildPresetGradeResult('perfect100')
  }
};

function buildPresetGradeResult(mode) {
  const allDeductions = [
    {
      id: 'basic-1',
      modes: ['raw0', 'fail51'],
      problemNo: '기본작업-1',
      taskName: '데이터 입력 (부분점수 없음)',
      deduction: -5,
      detail: mode === 'fail51'
        ? '[A4] 셀 영문 대소문자 불일치 (입력값: "sp-101", 정답: "SP-101") - 하나라도 틀리면 0점 처리'
        : '[A3:E15] 외부 데이터 미입력 (부분점수 없음, 5점 전액 감점)',
      videoTimestamp: '00:45',
      videoTitle: '기본작업-1: 오타 없이 1분 만에 데이터 입력하는 팁',
      coachTip: '기본작업-1은 띄어쓰기 한 칸이나 영문 대소문자(SP-101 vs sp-101)만 달라도 5점 전체가 감점됩니다. 입력 후 눈으로 꼭 대조하세요!',
      sheetName: '기본작업-1',
      cells: [
        { ref: 'A3', userVal: mode === 'fail51' ? '제품코드' : '(빈 셀)', ansVal: '제품코드', match: mode === 'fail51' },
        { ref: 'A4', userVal: mode === 'fail51' ? 'sp-101 (소문자 오타)' : '(빈 셀)', ansVal: 'SP-101 (대문자 필수)', match: false },
        { ref: 'B4', userVal: mode === 'fail51' ? '무선마우스' : '(빈 셀)', ansVal: '무선마우스', match: mode === 'fail51' }
      ]
    },
    {
      id: 'basic-2-4',
      modes: ['raw0', 'fail51', 'pass82'],
      problemNo: '기본작업-2 (4)',
      taskName: '셀 서식 (부분점수 적용)',
      deduction: -2,
      detail: '[A3:F3] 셀 음영(채우기 색: 표준 색-노랑) 및 가운데 맞춤 미적용 (5문항 중 4번 감점)',
      videoTimestamp: '03:20',
      videoTitle: '기본작업-2: 셀 서식 & 사용자 지정 표시형식 완벽 정리 (각 2점 부분점수)',
      coachTip: '기본작업-2(셀 서식)는 5문항 각각 2점씩 부분점수가 있습니다! 테마 색-노랑이 아니라 [표준 색-노랑]을 선택해야 감점되지 않아요.',
      sheetName: '기본작업-2',
      cells: [
        { ref: 'A1:F1', userVal: mode === 'raw0' ? '미병합' : '병합 + 맑은 고딕 16pt 굵게', ansVal: '병합 + 맑은 고딕 16pt 굵게', match: mode !== 'raw0' },
        { ref: 'C4:C15', userVal: mode === 'raw0' ? '2025-02-03' : '2025년 02월 03일', ansVal: 'yyyy"년" mm"월" dd"일"', match: mode !== 'raw0' },
        { ref: 'A3:F3 채우기', userVal: '테마 색(연한 노랑) 또는 없음', ansVal: '표준 색 - 노랑 (#FFFF00)', match: false }
      ]
    },
    {
      id: 'basic-2-5',
      modes: ['raw0', 'fail51', 'pass82'],
      problemNo: '기본작업-2 (5)',
      taskName: '셀 서식 (부분점수 적용)',
      deduction: -2,
      detail: '[A3:F15] 모든 테두리 적용 후 바깥쪽 굵은 상자 테두리 미적용',
      videoTimestamp: '04:50',
      videoTitle: '기본작업-2 (5번): 안쪽 실선 + 바깥쪽 굵은 상자 테두리 10초 완성',
      coachTip: '테두리는 [모든 테두리]를 먼저 누른 뒤, 바로 이어서 [굵은 상자 테두리]를 눌러야 안쪽 실선과 바깥쪽 굵은 선이 동시에 적용됩니다!',
      sheetName: '기본작업-2',
      cells: [
        { ref: 'A3:F15 안쪽', userVal: '가는 실선', ansVal: '가는 실선 (모든 테두리)', match: true },
        { ref: 'A3:F15 바깥쪽', userVal: '가는 실선 (미변경)', ansVal: '굵은 상자 테두리', match: false }
      ]
    },
    {
      id: 'basic-3',
      modes: ['raw0', 'fail51'],
      problemNo: '기본작업-3',
      taskName: '조건부 서식 (부분점수 없음)',
      deduction: -5,
      detail: '[A4:H16] 영역에 수식을 이용한 조건부 서식(=$E4>=7, 글꼴 파랑/굵게) 미적용 또는 절대참조($E$4) 오류',
      videoTimestamp: '07:10',
      videoTitle: '기본작업-3: 조건부 서식에서 $E4 혼합참조 실수 0% 만드는 법',
      coachTip: '조건부 서식에서 행 전체에 서식을 넣을 때는 열 이름 앞에만 달러($)를 붙이는 혼합참조(=$E4>=7)를 써야 합니다. 제목 행(A3:H3)을 범위에 포함하면 0점 처리되니 주의하세요!',
      sheetName: '기본작업-3',
      cells: [
        { ref: '적용 범위', userVal: mode === 'fail51' ? '=$A$3:$H$16 (제목행 포함 오류)' : '(미적용)', ansVal: '=$A$4:$H$16 (데이터 행만 선택)', match: false },
        { ref: '조건부 수식', userVal: mode === 'fail51' ? '=$E$4>=7 (절대참조 오류)' : '(미작성)', ansVal: '=$E4>=7 (열 고정 혼합참조)', match: false }
      ]
    },
    {
      id: 'calc-1',
      modes: ['raw0'],
      problemNo: '계산작업 (1)',
      taskName: '[표1] 셔틀버스 운행 시간표 [D3:D10]',
      deduction: -8,
      detail: '[D3:D10] 필수 함수(TIME) 미사용 (출발시간 + 정류장수×8분 도착시간 구하기)',
      videoTimestamp: '11:45',
      videoTitle: '계산작업 1번: TIME 함수를 활용한 셔틀버스 도착시간 구하기',
      coachTip: '출발시간(C3)에 정류장당 8분을 더할 때는 =C3+TIME(0, B3*8, 0)처럼 작성하면 60분이 넘어도 자동으로 1시간이 올라갑니다!',
      sheetName: '계산작업',
      cells: [
        { ref: 'D3 수식', userVal: '(수식 미입력)', ansVal: '=C3+TIME(0, B3*8, 0)', match: false },
        { ref: 'D3 결과값', userVal: '(빈 셀)', ansVal: '08:48', match: false }
      ]
    },
    {
      id: 'calc-2',
      modes: ['raw0'],
      problemNo: '계산작업 (2)',
      taskName: '[표2] 영업부서별 실적표 [J3:J10]',
      deduction: -8,
      detail: '[J3:J10] 필수 함수(IF, AND, AVERAGE) 미사용 ("백화점"이면서 평균 이상이면 "우수매장")',
      videoTimestamp: '15:20',
      videoTitle: '계산작업 2번: IF, AND, AVERAGE 중첩 함수와 절대참조($) 핵심',
      coachTip: '제시된 함수(IF, AND, AVERAGE)를 모두 사용해야 하며, 평균 범위($I$3:$I$10)에 반드시 F4 키로 절대참조($)를 걸어야 합니다!',
      sheetName: '계산작업',
      cells: [
        { ref: 'J3 수식', userVal: '(수식 미입력)', ansVal: '=IF(AND(H3="백화점",I3>=AVERAGE($I$3:$I$10)),"우수매장","")', match: false },
        { ref: 'J5 (강남영업)', userVal: '(빈 셀)', ansVal: '"우수매장" (백화점 & 5580 >= 4738.75)', match: false }
      ]
    },
    {
      id: 'calc-3',
      modes: ['raw0'],
      problemNo: '계산작업 (3)',
      taskName: '[표3] 영화예매현황 [D24]',
      deduction: -8,
      detail: '[D24] 필수 함수(ROUND, DAVERAGE) 미사용 (드라마 예매수량 평균을 백의 자리에서 반올림)',
      videoTimestamp: '18:50',
      videoTitle: '계산작업 3번: 데이터베이스 함수(DAVERAGE)와 ROUND 자릿수(-2) 공식',
      coachTip: '백의 자리에서 반올림하여 천의 자리까지 표시할 때는 ROUND의 두 번째 인수를 -2로 지정합니다. (십의 자리는 -1, 백의 자리는 -2, 천의 자리는 -3!)',
      sheetName: '계산작업',
      cells: [
        { ref: 'E22:E23 조건', userVal: '<조건> 미작성', ansVal: 'E22: 장르 / E23: 드라마', match: false },
        { ref: 'D24 수식', userVal: '(수식 미입력)', ansVal: '=ROUND(DAVERAGE(A13:D23, 4, E22:E23), -2)', match: false }
      ]
    },
    {
      id: 'calc-4',
      modes: ['raw0', 'fail51'],
      problemNo: '계산작업 (4)',
      taskName: '[표4] 아파트 당첨현황 [J14:J24]',
      deduction: -8,
      detail: '[J14] 필수 함수(INDEX, MATCH) 수식 없이 결과값(15)만 직접 입력됨 (수식 미사용 시 0점 처리)',
      videoTimestamp: '21:30',
      videoTitle: '계산작업 4번: INDEX + MATCH 함수 조합 100% 이해하기',
      coachTip: '결과값이 맞아도 문제에 제시된 함수를 모두 사용하지 않거나 숫자만 직접 타이핑하면 8점 전체가 감점됩니다! MATCH의 세 번째 인수에는 정확히 일치(0)를 꼭 적어주세요.',
      sheetName: '계산작업',
      cells: [
        { ref: 'J14 수식', userVal: '15 (수식 없이 값만 입력됨)', ansVal: '=INDEX($A$14:$D$23, MATCH(G14, $A$14:$A$23, 0), 4)', match: false },
        { ref: '채점 판정', userVal: '필수 함수(INDEX, MATCH) 누락', ansVal: 'INDEX, MATCH 모두 포함 필수 (-8점)', match: false }
      ]
    },
    {
      id: 'calc-5',
      modes: ['raw0', 'fail51', 'pass82'],
      problemNo: '계산작업 (5)',
      taskName: '[표5] 자격증 수강등록 현황 [E28:E37]',
      deduction: -8,
      detail: '[E28:E37] 필수 함수(CHOOSE, LEFT) 텍스트 오타 ("컴퓨터활용2급" → 정답: "컴활2급")로 결과값 불일치',
      videoTimestamp: '23:40',
      videoTitle: '계산작업 5번: CHOOSE + LEFT/MID 문자열 추출 함수 단골 유형',
      coachTip: '함수식이 완벽해도 큰따옴표 안의 출력 문자열("컴활1급", "컴활2급", "워드1급")에 글자 하나만 달라도 결과값 불일치로 8점이 감점됩니다!',
      sheetName: '계산작업',
      cells: [
        { ref: 'E28 수식', userVal: '=CHOOSE(LEFT(B28,1),"컴활1급","컴퓨터활용2급","워드1급")', ansVal: '=CHOOSE(LEFT(B28,1),"컴활1급","컴활2급","워드1급")', match: false },
        { ref: 'E29 결과값', userVal: '컴퓨터활용2급 (오타)', ansVal: '컴활2급', match: false }
      ]
    },
    {
      id: 'analysis-1',
      modes: ['raw0'],
      problemNo: '분석작업-1',
      taskName: '목표값 찾기 (부분점수 없음)',
      deduction: -10,
      detail: '[E6] 순이익이 20,000,000이 되도록 하는 [B4] 판매수량 목표값 찾기 미수행 (현재 B4=200, 정답=250)',
      videoTimestamp: '25:30',
      videoTitle: '분석작업-1: 목표값 찾기 30초 만에 10점 만점 받기',
      coachTip: '분석작업은 부분점수가 없습니다! [데이터] → [가상 분석] → [목표값 찾기]에서 수식 셀(E6), 찾는 값(20000000), 값을 바꿀 셀($B$4)을 차례대로 넣으세요.',
      sheetName: '분석작업-1',
      cells: [
        { ref: 'B4 (판매수량)', userVal: '200', ansVal: '250 (목표값 찾기 완료 후)', match: false },
        { ref: 'E6 (순이익)', userVal: '13,000,000', ansVal: '20,000,000', match: false }
      ]
    },
    {
      id: 'analysis-2',
      modes: ['raw0', 'fail51'],
      problemNo: '분석작업-2',
      taskName: '데이터 통합 (부분점수 없음)',
      deduction: -10,
      detail: '[표4] 수도권 영업실적 [F9:I13] 영역에 서울·경기·인천 영업실적 합계 데이터 통합 미수행 또는 첫 행/왼쪽 열 체크 누락',
      videoTimestamp: '30:10',
      videoTitle: '분석작업-2: 데이터 통합 & 피벗테이블 위치 실수 방지법',
      coachTip: '데이터 통합에서 사용할 레이블의 [첫 행]과 [왼쪽 열] 체크박스를 둘 다 켜지 않으면 숫자가 밀려서 10점 전체가 감점됩니다!',
      sheetName: '분석작업-2',
      cells: [
        { ref: 'G10 (SMT-S1 1월)', userVal: '(빈 셀)', ansVal: '12300 (서울 5800 + 경기 3800 + 인천 2700)', match: false },
        { ref: 'H10 (SMT-S1 2월)', userVal: '(빈 셀)', ansVal: '14200 (서울 6800 + 경기 4800 + 인천 2600)', match: false },
        { ref: 'I13 (SMT-I2 3월)', userVal: '(빈 셀)', ansVal: '13300 (서울 6000 + 경기 3900 + 인천 3400)', match: false }
      ]
    },
    {
      id: 'macro-2',
      modes: ['raw0', 'fail51', 'pass82'],
      problemNo: '매크로작업 (2)',
      taskName: '서식 매크로 및 도형 연결 (부분점수 적용)',
      deduction: -5,
      detail: '[A3:G3] 채우기 색(주황) 서식 매크로 기록 후 [I6:J7] 직사각형 도형("서식")에 매크로 지정 연결 누락',
      videoTimestamp: '34:00',
      videoTitle: '매크로작업 1·2번: 기록 순서와 도형 연결로 부분점수까지 챙기기',
      coachTip: '매크로를 다 기록하고도 도형에 우클릭 → [매크로 지정]을 깜빡하는 수험생이 정말 많습니다! 다행히 매크로 1번(5점)은 맞아서 부분점수를 확보했어요.',
      sheetName: '매크로작업',
      cells: [
        { ref: 'G4:G13 (매크로1)', userVal: '=SUM(B4:F4) (398점 정상)', ansVal: '=SUM(B4:F4) (398점)', match: mode !== 'raw0' },
        { ref: 'I6:J7 도형(매크로2)', userVal: '매크로 미지정 (연결 누락)', ansVal: '"서식" 매크로 연결 완료', match: false }
      ]
    },
    {
      id: 'chart-5',
      modes: ['raw0', 'fail51', 'pass82'],
      problemNo: '차트작업 (5)',
      taskName: '차트 영역 및 범례 수정 (부분점수 적용)',
      deduction: -2,
      detail: '범례 위치를 [아래쪽]으로 변경했으나 차트 영역 테두리 [둥근 모서리] 미적용 (5문항 중 5번 감점)',
      videoTimestamp: '37:20',
      videoTitle: '차트작업 1~5번: 2점씩 차곡차곡 쌓는 차트 세부 설정',
      coachTip: '차트작업은 5문항 각 2점씩 부분점수가 적용됩니다! 차트 바깥쪽 빈 공간을 더블클릭 → [차트 영역 서식] → [테두리] 맨 아래에 있는 [둥근 모서리] 체크박스를 꼭 확인하세요.',
      sheetName: '차트작업',
      cells: [
        { ref: '데이터 계열 (1번)', userVal: '젠가 제외 완료', ansVal: '젠가 제외 4개 계열', match: mode !== 'raw0' },
        { ref: '3월 계열 (2번)', userVal: '표식이 있는 꺾은선형 (보조축)', ansVal: '표식이 있는 꺾은선형 (보조축)', match: mode !== 'raw0' },
        { ref: '차트 영역 테두리 (5번)', userVal: '직각 모서리 (기본값)', ansVal: '둥근 모서리 체크 완료', match: false }
      ]
    }
  ];

  const filtered = allDeductions.filter(d => d.modes.includes(mode));
  const scoresByMode = {
    raw0: { basic: 0, calc: 0, analysis: 0, etc: 0, total: 0 },
    fail51: { basic: 6, calc: 24, analysis: 10, etc: 13, total: 53 },
    pass82: { basic: 16, calc: 32, analysis: 20, etc: 13, total: 81 },
    perfect100: { basic: 20, calc: 40, analysis: 20, etc: 20, total: 100 }
  };
  const sc = scoresByMode[mode] || scoresByMode.pass82;

  return {
    totalScore: sc.total,
    passed: sc.total >= 70,
    sectionScores: {
      basic: sc.basic,
      calc: sc.calc,
      analysis: sc.analysis,
      etc: sc.etc
    },
    deductions: filtered,
    diffItems: filtered.map(f => ({
      id: f.id,
      sheetName: f.sheetName,
      title: `${f.problemNo} - ${f.taskName}`,
      cells: f.cells
    })),
    radarStats: {
      '데이터입력·서식': Math.round((sc.basic / 20) * 100),
      '조건부서식·필터': mode === 'pass82' || mode === 'perfect100' ? 100 : 25,
      '함수(계산작업)': Math.round((sc.calc / 40) * 100),
      '데이터분석(피벗·통합)': Math.round((sc.analysis / 20) * 100),
      '매크로·VBA': mode === 'perfect100' ? 100 : (mode === 'raw0' ? 0 : 60),
      '차트작업': mode === 'perfect100' ? 100 : (mode === 'raw0' ? 0 : 80)
    }
  };
}

async function inspectXlsmBufferClient(arrayBuffer, fileName) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  if (bytes.length < 22 || view.getUint32(0, true) !== 0x04034b50) {
    return {
      ok: false,
      fileName,
      checks: [{ label: '엑셀 패키지 무결성 검사', status: 'fail', detail: '유효한 엑셀(.xlsm/.xlsx) ZIP 구조가 아닙니다.' }]
    };
  }

  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  const totalEntries = eocd !== -1 ? view.getUint16(eocd + 10, true) : 0;
  const cdOffset = eocd !== -1 ? view.getUint32(eocd + 16, true) : 0;

  const entryNames = [];
  let ptr = cdOffset;
  const decoder = new TextDecoder('utf-8');
  for (let i = 0; i < totalEntries; i++) {
    if (ptr + 46 > bytes.length || view.getUint32(ptr, true) !== 0x02014b50) break;
    const fnLen = view.getUint16(ptr + 28, true);
    const exLen = view.getUint16(ptr + 30, true);
    const cmLen = view.getUint16(ptr + 32, true);
    const name = decoder.decode(bytes.subarray(ptr + 46, ptr + 46 + fnLen));
    entryNames.push(name);
    ptr += 46 + fnLen + exLen + cmLen;
  }

  const sheetFiles = entryNames.filter(n => n.startsWith('xl/worksheets/sheet') && n.endsWith('.xml'));
  let matchedPreset = 'pass82';
  if (fileName.includes('작성 전') || fileName.includes('0점') || fileName.includes('수험생문제용')) matchedPreset = 'raw0';
  else if (fileName.includes('51점') || fileName.includes('불합격')) matchedPreset = 'fail51';
  else if (fileName.includes('100점') || fileName.includes('모범정답')) matchedPreset = 'perfect100';

  return {
    ok: sheetFiles.length >= 8,
    matchedPreset,
    checks: [
      {
        label: '필수 8개 시트 구조 검사',
        status: sheetFiles.length >= 8 ? 'pass' : 'fail',
        detail: `워크시트 ${sheetFiles.length}개(기본1~3, 계산, 분석1~2, 매크로, 차트) 정상 확인됨`
      },
      ...PRESET_SCENARIOS[matchedPreset].preflight.checks.slice(1)
    ]
  };
}

document.addEventListener('DOMContentLoaded', async () => {
  initUrlParams();
  initTimer();
  initUploadAndPresets();
  initModalsAndButtons();
  initProblemGenerator();
  initSampleGuideAndWrongNote();
  initUpgradeStepRibbon();

  renderRankingBars('studentRankingChart', localState.studentRankings, 'name', 2200);
  renderRankingBars('teacherRankingChart', localState.teacherRankings, 'id', 850);
  renderAllPieCharts(localState.distributions, null);

  renderWrongNotes();
  await syncFromSupabase();
  renderExamPaper(MOCK_EXAM_BANK['대표기출따라하기']);
  await runQuickDemo('fail51', null, false);
});

function initSampleGuideAndWrongNote() {
  // Top nav buttons
  const btnNavSample = document.getElementById('btnNavSampleGuide');
  if (btnNavSample) {
    btnNavSample.addEventListener('click', () => {
      document.getElementById('sampleGuideSection').scrollIntoView({ behavior: 'smooth' });
    });
  }
  const btnNavWrong = document.getElementById('btnNavWrongNote');
  if (btnNavWrong) {
    btnNavWrong.addEventListener('click', () => {
      document.getElementById('wrongNoteSection').scrollIntoView({ behavior: 'smooth' });
    });
  }

  // Sample scorecard quick loaders
  const btnPass81 = document.getElementById('btnLoadSamplePass81');
  if (btnPass81) {
    btnPass81.addEventListener('click', async () => {
      await runQuickDemo('pass82', null, false);
      document.getElementById('sectionScoreCard').scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  }
  const btnFail53 = document.getElementById('btnLoadSampleFail53');
  if (btnFail53) {
    btnFail53.addEventListener('click', async () => {
      await runQuickDemo('fail51', null, false);
      document.getElementById('sectionScoreCard').scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  }

  // Live Formula Mini-Tester
  const btnTestFormula = document.getElementById('btnTestLiveFormula');
  if (btnTestFormula) {
    btnTestFormula.addEventListener('click', () => {
      const val = (document.getElementById('liveFormulaInput').value || '').trim();
      const upper = val.toUpperCase();
      const fb = document.getElementById('liveFormulaFeedback');
      if (!val.startsWith('=')) {
        fb.style.color = '#E11D48';
        fb.innerHTML = '🔴 <strong>오답 (0점):</strong> 엑셀 수식은 반드시 등호(<code>=</code>)로 시작해야 합니다!';
        return;
      }
      const missing = ['IF', 'AND', 'AVERAGE'].filter(fn => !upper.includes(fn));
      if (missing.length > 0) {
        fb.style.color = '#E11D48';
        fb.innerHTML = `🔴 <strong>오답 (0점):</strong> 문제에서 지시한 필수 함수(<strong>${missing.join(', ')}</strong>)가 누락되었습니다!`;
        return;
      }
      if (!upper.includes('$I$3:$I$10') && !upper.includes('I$3:I$10')) {
        fb.style.color = '#D97706';
        fb.innerHTML = '🟡 <strong>주의 (감점 위험):</strong> AVERAGE 범위(<code>$I$3:$I$10</code>)에 절대참조(<code>$</code>)를 지정하지 않으면 아래 행으로 채우기 드래그할 때 범위가 밀려 오답 처리됩니다!';
        return;
      }
      if (!val.includes('백화점') || !val.includes('우수매장')) {
        fb.style.color = '#E11D48';
        fb.innerHTML = '🔴 <strong>오답 (0점):</strong> 조건 텍스트(<code>"백화점"</code>) 또는 출력 텍스트(<code>"우수매장"</code>) 오타를 확인해 주세요!';
        return;
      }
      fb.style.color = '#059669';
      fb.innerHTML = '🟢 <strong>정답 인정 (+8점 획득!):</strong> 필수 함수(IF, AND, AVERAGE) 3개 모두 포함, 절대참조(<code>$I$3:$I$10</code>) 및 출력 문자열이 완벽히 일치합니다!';
    });
  }

  // Wrong-Answer Note controls
  const btnFilter = document.getElementById('btnFilterUnmastered');
  if (btnFilter) {
    btnFilter.addEventListener('click', () => {
      filterUnmasteredOnly = !filterUnmasteredOnly;
      btnFilter.textContent = filterUnmasteredOnly ? '📋 전체 오답 모두 보기' : '🔍 미암기 오답만 보기';
      renderWrongNotes();
    });
  }

  const btnSyncAll = document.getElementById('btnSyncCurrentToWrongNote');
  if (btnSyncAll) {
    btnSyncAll.addEventListener('click', () => {
      syncCurrentDeductionsToWrongNotes(true);
    });
  }

  const btnAddSelected = document.getElementById('btnAddSelectedToWrongNote');
  if (btnAddSelected) {
    btnAddSelected.addEventListener('click', () => {
      addSingleDeductionToWrongNote(selectedDeductionIdx);
      document.getElementById('wrongNoteSection').scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
  }

  const btnPrintWrong = document.getElementById('btnPrintWrongNoteOnly');
  if (btnPrintWrong) {
    btnPrintWrong.addEventListener('click', () => {
      document.body.classList.add('print-wrong-note-only');
      window.print();
      setTimeout(() => document.body.classList.remove('print-wrong-note-only'), 500);
    });
  }

  // Exam paper font switcher & exam paper print
  const fontSel = document.getElementById('examFontSelect');
  if (fontSel) {
    fontSel.addEventListener('change', () => {
      const wrapper = document.getElementById('examPaperSheetWrapper');
      wrapper.className = `exam-paper-sheet ${fontSel.value}`;
      const labelMap = {
        'exam-font-malgun': '맑은 고딕 11pt (Office 2021 표준)',
        'exam-font-gulim': '굴림 / 굴림체 (상공회의소 실기 지시서)',
        'exam-font-batang': '바탕 / 바탕체 (국가기술자격 인쇄 시험지)'
      };
      document.getElementById('examPaperMeta').textContent = `적용 서체: ${labelMap[fontSel.value] || '맑은 고딕'} | 총 8개 시트`;
    });
  }

  const btnPrintExam = document.getElementById('btnPrintExamPaperOnly');
  if (btnPrintExam) {
    btnPrintExam.addEventListener('click', () => {
      document.body.classList.add('print-exam-paper-only');
      window.print();
      setTimeout(() => document.body.classList.remove('print-exam-paper-only'), 500);
    });
  }
}

function addSingleDeductionToWrongNote(idx) {
  if (!latestGradeData || !latestGradeData.gradeResult) return;
  const gr = latestGradeData.gradeResult;
  const d = gr.deductions[idx];
  if (!d) return;

  const diff = gr.diffItems.find(item => item.id === d.id) || gr.diffItems[0];
  const userIn = diff && diff.cells && diff.cells[0] ? diff.cells[0].userVal : '(수식/서식 누락)';
  const ansIn = diff && diff.cells && diff.cells[0] ? diff.cells[0].ansVal : d.detail;

  const newEntry = {
    id: `wn-${Date.now()}-${idx}`,
    student_nickname: latestGradeData.studentName || '짱아',
    teacher_id: latestGradeData.teacherId || 'Zzang_AA',
    exam_code: latestGradeData.roundName || '대표기출따라하기',
    problem_no: d.problemNo,
    task_name: d.taskName,
    deduction: d.deduction,
    user_input: userIn,
    correct_formula: ansIn,
    error_reason: d.detail,
    coach_tip: d.coachTip,
    student_memo: '반드시 다시 풀어보고 완벽히 암기하기!',
    is_mastered: false
  };

  const exists = wrongNotesList.some(w => w.exam_code === newEntry.exam_code && w.problem_no === newEntry.problem_no);
  if (!exists) {
    wrongNotesList.unshift(newEntry);
    try {
      sbFetch('comlicense_wrong_notes', {
        method: 'POST',
        headers: { Prefer: 'return=minimal' },
        body: JSON.stringify({
          student_nickname: newEntry.student_nickname,
          teacher_id: newEntry.teacher_id,
          exam_code: newEntry.exam_code,
          problem_no: newEntry.problem_no,
          task_name: newEntry.task_name,
          deduction: newEntry.deduction,
          user_input: newEntry.user_input,
          correct_formula: newEntry.correct_formula,
          error_reason: newEntry.error_reason,
          coach_tip: newEntry.coach_tip,
          student_memo: newEntry.student_memo,
          is_mastered: false
        })
      });
    } catch (_) {}
  }
  renderWrongNotes();
}

function syncCurrentDeductionsToWrongNotes(scrollToNote = false) {
  if (!latestGradeData || !latestGradeData.gradeResult) return;
  const deductions = latestGradeData.gradeResult.deductions || [];
  deductions.forEach((_, idx) => addSingleDeductionToWrongNote(idx));
  if (scrollToNote) {
    document.getElementById('wrongNoteSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function renderWrongNotes() {
  const container = document.getElementById('wrongNoteListContainer');
  const badge = document.getElementById('wrongNoteCountBadge');
  if (!container) return;

  const studentName = document.getElementById('studentNameInput')?.value || '짱아';
  const teacherId = document.getElementById('teacherIdInput')?.value || 'Zzang_AA';
  document.getElementById('wrongNoteStudentLabel').textContent = studentName;
  document.getElementById('wrongNoteTeacherLabel').textContent = teacherId;

  const visibleNotes = filterUnmasteredOnly
    ? wrongNotesList.filter(w => !w.is_mastered)
    : wrongNotesList;

  if (badge) {
    const unmasteredCount = wrongNotesList.filter(w => !w.is_mastered).length;
    badge.textContent = `총 ${wrongNotesList.length}문항 (미암기 ${unmasteredCount}문항)`;
  }

  if (visibleNotes.length === 0) {
    container.innerHTML = `
      <div style="text-align:center;padding:28px;background:#F8FAFC;border:1px dashed #93C5FD;border-radius:10px;color:#059669;">
        <strong>🎉 모든 오답 문항의 복습을 완료했습니다!</strong> (전체 보기로 전환하거나 새로운 채점을 진행해 보세요)
      </div>
    `;
    return;
  }

  container.innerHTML = visibleNotes.map((w, idx) => `
    <div class="wrong-note-card ${w.is_mastered ? 'mastered' : ''}" data-wn-id="${w.id}">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:6px;">
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
          <span style="background:#1E3A8A;color:#FFFFFF;font-size:11.5px;padding:3px 9px;border-radius:6px;"><strong>${w.exam_code}</strong></span>
          <strong style="font-size:14.5px;color:#0F172A;">[${w.problem_no}] ${w.task_name}</strong>
          <span style="color:#E11D48;font-size:12.5px;"><strong>(${w.deduction}점 감점)</strong></span>
        </div>
        <label style="display:flex;align-items:center;gap:6px;font-size:12.5px;cursor:pointer;color:${w.is_mastered ? '#059669' : '#1E40AF'};">
          <input type="checkbox" class="wn-master-check" data-id="${w.id}" ${w.is_mastered ? 'checked' : ''} />
          <strong>${w.is_mastered ? '✅ 복습·암기 완료' : '⬜ 복습 완료 체크'}</strong>
        </label>
      </div>

      <div style="font-size:12.5px;color:#475569;margin-bottom:8px;">
        <strong>❗ 감점 원인:</strong> ${w.error_reason}
      </div>

      <div class="wrong-note-grid">
        <div style="background:#FFF1F2;border:1px solid #FECDD3;padding:9px 12px;border-radius:8px;font-size:12px;">
          <div style="color:#E11D48;margin-bottom:3px;"><strong>❌ 내가 작성한 오답 수식/서식</strong></div>
          <code style="color:#0F172A;font-family:Consolas,monospace;word-break:break-all;">${w.user_input}</code>
        </div>
        <div style="background:#ECFDF5;border:1px solid #A7F3D0;padding:9px 12px;border-radius:8px;font-size:12px;">
          <div style="color:#059669;margin-bottom:3px;"><strong>⭕ 반드시 외울 모범 정답 공식</strong></div>
          <code style="color:#0F172A;font-family:Consolas,monospace;word-break:break-all;">${w.correct_formula}</code>
        </div>
      </div>

      <div style="background:#EFF6FF;border-left:4px solid #2563EB;padding:8px 12px;border-radius:6px;font-size:12px;color:#1E3A8A;margin-bottom:8px;">
        💡 <strong>핵심 암기 포인트:</strong> ${w.coach_tip}
      </div>

      <div style="display:flex;align-items:center;gap:8px;">
        <span style="font-size:12px;color:#1E40AF;white-space:nowrap;"><strong>✏️ 나의 복습 메모:</strong></span>
        <input type="text" class="memo-input wn-memo-input" data-id="${w.id}" value="${(w.student_memo || '').replace(/"/g, '&quot;')}" placeholder="이 문제를 다시 틀리지 않기 위한 나만의 핵심 메모를 적어보세요!" />
      </div>
    </div>
  `).join('');

  container.querySelectorAll('.wn-master-check').forEach(chk => {
    chk.addEventListener('change', (e) => {
      const id = e.target.getAttribute('data-id');
      const target = wrongNotesList.find(x => String(x.id) === String(id));
      if (target) {
        target.is_mastered = e.target.checked;
        renderWrongNotes();
      }
    });
  });

  container.querySelectorAll('.wn-memo-input').forEach(inp => {
    inp.addEventListener('change', (e) => {
      const id = e.target.getAttribute('data-id');
      const target = wrongNotesList.find(x => String(x.id) === String(id));
      if (target) {
        target.student_memo = e.target.value;
      }
    });
  });
}

function initProblemGenerator() {
  const roundSelect = document.getElementById('roundSelect');
  roundSelect.addEventListener('change', () => {
    const val = roundSelect.value;
    const spec = MOCK_EXAM_BANK[val] || MOCK_EXAM_BANK['대표기출따라하기'];
    renderExamPaper(spec);
  });

  document.getElementById('btnViewExamPaper').addEventListener('click', () => {
    const val = roundSelect.value;
    renderExamPaper(MOCK_EXAM_BANK[val] || MOCK_EXAM_BANK['대표기출따라하기']);
    document.getElementById('problemGeneratorSection').scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('btnOpenGeneratorQuick').addEventListener('click', () => {
    const newExam = generateCustomMockExam();
    const sec = document.getElementById('problemGeneratorSection');
    sec.scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('btnNavGenerator').addEventListener('click', () => {
    document.getElementById('problemGeneratorSection').scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('btnGenerateNewExam').addEventListener('click', () => {
    generateCustomMockExam();
  });

  document.getElementById('btnToggleAnswers').addEventListener('click', () => {
    showAnswersAll = !showAnswersAll;
    document.querySelectorAll('.exam-q-ans').forEach(el => {
      el.classList.toggle('show', showAnswersAll);
    });
  });

  document.getElementById('btnDownloadCurrentExamXlsx').addEventListener('click', () => {
    downloadExamWorkbook(false);
  });

  document.getElementById('btnDownloadCurrentAnsXlsx').addEventListener('click', () => {
    downloadExamWorkbook(true);
  });
}

function initUrlParams() {
  const params = new URLSearchParams(window.location.search);
  const co = params.get('co') || 'youngjin';
  const year = params.get('year') || '2026';
  const subject = params.get('subject') || 'excel';
  const level = params.get('level') || '2';
  const badge = document.getElementById('urlParamBadge');
  if (badge) {
    badge.textContent = `co=${co} · year=${year} · subject=${subject} · level=${level}급`;
  }
}

function initTimer() {
  const timerBtn = document.getElementById('examTimerBtn');
  const timerDisplay = document.getElementById('timerDisplay');
  const paceHint = document.getElementById('timerPaceHint');
  timerBtn.addEventListener('click', () => {
    if (timerInterval) {
      clearInterval(timerInterval);
      timerInterval = null;
      timerBtn.classList.remove('running');
    } else {
      timerBtn.classList.add('running');
      timerInterval = setInterval(() => {
        if (timerSeconds > 0) {
          timerSeconds--;
          const m = Math.floor(timerSeconds / 60);
          const s = timerSeconds % 60;
          timerDisplay.textContent = `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
          if (paceHint) {
            if (m >= 35) paceHint.textContent = '(기본작업 5분 구간)';
            else if (m >= 17) paceHint.textContent = '(계산작업 18분 구간)';
            else if (m >= 9) paceHint.textContent = '(분석작업 8분 구간)';
            else paceHint.textContent = '(매크로·차트 마무리 구간!)';
          }
        } else {
          clearInterval(timerInterval);
          timerInterval = null;
          alert('⏰ 40분 시험 시간이 종료되었습니다! 작성한 파일을 제출해 주세요.');
        }
      }, 1000);
    }
  });
}

function initUploadAndPresets() {
  const dropzone = document.getElementById('uploadDropzone');
  const fileInput = document.getElementById('fileInput');
  const fileLabel = document.getElementById('selectedFileLabel');

  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
  dropzone.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      await handleSelectedFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', async (e) => {
    if (e.target.files && e.target.files[0]) {
      await handleSelectedFile(e.target.files[0]);
    }
  });

  async function handleSelectedFile(file) {
    currentFileName = file.name;
    fileLabel.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    currentFileBuffer = await file.arrayBuffer();

    const clientPf = await inspectXlsmBufferClient(currentFileBuffer, file.name);
    renderPreflightBox(clientPf);
  }

  document.querySelectorAll('.preset-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const preset = btn.getAttribute('data-preset');
      await runQuickDemo(preset);
    });
  });

  document.getElementById('btnStartGrade').addEventListener('click', async () => {
    if (currentFileBuffer) {
      const clientPf = await inspectXlsmBufferClient(currentFileBuffer, currentFileName || 'upload.xlsm');
      await runQuickDemo(clientPf.matchedPreset || 'pass82', currentFileName);
    } else {
      await runQuickDemo('pass82');
    }
  });
}

async function runQuickDemo(preset, customFileName = null, persistToDb = true) {
  const studentName = document.getElementById('studentNameInput').value || '짱아';
  const teacherId = document.getElementById('teacherIdInput').value || 'Zzang_AA';
  const roundName = document.getElementById('roundSelect').value;

  const scenario = PRESET_SCENARIOS[preset] || PRESET_SCENARIOS.pass82;
  const fName = customFileName || scenario.fileName;
  document.getElementById('selectedFileLabel').textContent = fName;

  localState.participantCount += 1;
  const gr = scenario.gradeResult;
  const nowStr = new Date().toISOString().replace('T', ' ').slice(0, 19);
  const topWeak = gr.deductions.length > 0
    ? `${gr.deductions[0].problemNo} (${gr.deductions[0].taskName})`
    : '전 영역 완벽(만점)';

  localState.submissions.unshift({
    studentName,
    teacherId,
    roundName,
    basic: gr.sectionScores.basic,
    calc: gr.sectionScores.calc,
    analysis: gr.sectionScores.analysis,
    etc: gr.sectionScores.etc,
    total: gr.totalScore,
    weakPoint: topWeak,
    submittedAt: nowStr,
    radarStats: gr.radarStats
  });

  if (persistToDb) {
    saveSubmissionToSupabase({
      fileName: fName,
      studentName,
      teacherId,
      roundName,
      weakPoint: topWeak,
      preflight: scenario.preflight,
      gradeResult: gr
    });
  }

  applyGradeResponse({
    fileName: fName,
    studentName,
    teacherId,
    roundName,
    preflight: scenario.preflight,
    gradeResult: gr,
    participantCount: localState.participantCount,
    averageScore: localState.averageScore,
    studentRankings: localState.studentRankings,
    teacherRankings: localState.teacherRankings
  });
}

function renderPreflightBox(preflight) {
  const box = document.getElementById('preflightBox');
  const badge = document.getElementById('preflightStatusBadge');
  const listEl = document.getElementById('preflightChecklist');

  if (!preflight) return;
  const hasWarnOrFail = preflight.checks.some(c => c.status !== 'pass');
  box.className = hasWarnOrFail ? 'preflight-box warn' : 'preflight-box';
  badge.textContent = hasWarnOrFail ? '⚠️ 주의항목 발견' : '✅ 8개 시트·수식 정상';
  badge.style.background = hasWarnOrFail ? '#fef08a' : '#dcfce7';

  listEl.innerHTML = preflight.checks.map(c => {
    const icon = c.status === 'pass' ? '🟢' : (c.status === 'warn' ? '🟡' : '🔴');
    return `<div class="preflight-item"><span>${icon}</span><div><strong>${c.label}:</strong> ${c.detail}</div></div>`;
  }).join('');
}

function applyGradeResponse(data) {
  latestGradeData = data;
  renderPreflightBox(data.preflight);

  const gr = data.gradeResult;
  const sec = gr.sectionScores;

  document.getElementById('meterBasic').style.width = `${(sec.basic / 20) * 100}%`;
  document.getElementById('scoreBasicText').textContent = `${sec.basic} / 20점`;

  document.getElementById('meterCalc').style.width = `${(sec.calc / 40) * 100}%`;
  document.getElementById('scoreCalcText').textContent = `${sec.calc} / 40점`;

  document.getElementById('meterAnalysis').style.width = `${(sec.analysis / 20) * 100}%`;
  document.getElementById('scoreAnalysisText').textContent = `${sec.analysis} / 20점`;

  document.getElementById('meterEtc').style.width = `${(sec.etc / 20) * 100}%`;
  document.getElementById('scoreEtcText').textContent = `${sec.etc} / 20점`;

  document.getElementById('totalScoreNum').textContent = gr.totalScore;
  const passPill = document.getElementById('passStatusPill');
  const banner = document.getElementById('totalScoreBanner');
  if (gr.totalScore >= 70) {
    passPill.textContent = `🎉 축하합니다! 합격 기준(70점)을 통과했습니다 (+${gr.totalScore - 70}점 여유)`;
    banner.style.background = 'linear-gradient(135deg, #1E3A8A 0%, #2563EB 52%, #059669 100%)';
  } else {
    passPill.textContent = `💪 조금만 더! 합격(70점)까지 ${70 - gr.totalScore}점 남았습니다`;
    banner.style.background = 'linear-gradient(135deg, #1E40AF 0%, #2563EB 52%, #38BDF8 100%)';
  }

  document.getElementById('statParticipants').textContent = data.participantCount;

  const tbody = document.getElementById('deductionTableBody');
  if (gr.deductions.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:35px;color:#15803d;font-weight:700;">🏆 감점 내역이 없습니다! 8개 시트 모든 문항을 완벽하게 작성하셨습니다 (100점 만점).</td></tr>`;
    document.getElementById('aiCoachBubble').innerHTML = `🎉 <strong>인생 코치의 칭찬 한마디:</strong> 완벽합니다! 실제 시험장에서도 지금처럼만 풀면 100% 합격입니다.`;
    document.getElementById('diffTableBody').innerHTML = `<tr><td colspan="4" style="text-align:center;color:#15803d;font-weight:700;">모든 셀의 수식과 서식이 모범 정답과 100% 일치합니다.</td></tr>`;
  } else {
    tbody.innerHTML = gr.deductions.map((d, idx) => `
      <tr data-idx="${idx}" class="${idx === 0 ? 'selected' : ''}">
        <td style="font-weight:700;color:#1e293b;">${d.problemNo}</td>
        <td style="color:#475569;">${d.taskName}</td>
        <td class="deduct-badge">${d.deduction}</td>
        <td>${d.detail}</td>
        <td>
          <button type="button" class="timestamp-btn row-video-btn" data-idx="${idx}">🎬 ${d.videoTimestamp}</button>
          <button type="button" class="minimap-btn row-diff-btn" data-idx="${idx}">🔬 셀비교</button>
        </td>
      </tr>
    `).join('');

    tbody.querySelectorAll('tr').forEach(tr => {
      tr.addEventListener('click', () => {
        tbody.querySelectorAll('tr').forEach(r => r.classList.remove('selected'));
        tr.classList.add('selected');
        const idx = parseInt(tr.getAttribute('data-idx'), 10);
        selectDeductionItem(idx);
      });
    });

    tbody.querySelectorAll('.row-video-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const idx = parseInt(btn.getAttribute('data-idx'), 10);
        selectDeductionItem(idx);
        const d = gr.deductions[idx];
        openVideoModalAt(d.videoTimestamp, d.videoTitle, d.coachTip);
      });
    });

    selectDeductionItem(0);
  }

  renderRankingBars('studentRankingChart', data.studentRankings, 'name', 2200);
  renderRankingBars('teacherRankingChart', data.teacherRankings, 'id', 850);
  renderAllPieCharts(localState.distributions, gr);

  document.getElementById('radarTargetLabel').textContent = `${data.studentName} 수험생`;
  drawRadarChart(gr.radarStats);
}

function selectDeductionItem(idx) {
  selectedDeductionIdx = idx;
  if (!latestGradeData || !latestGradeData.gradeResult) return;
  const gr = latestGradeData.gradeResult;
  const d = gr.deductions[idx];
  if (!d) return;

  document.getElementById('aiCoachBubble').innerHTML = `💡 <strong>[${d.problemNo}] 인생 코치의 1:1 오답 처방전:</strong> ${d.coachTip}`;

  const jumpBtn = document.getElementById('btnDiffVideoJump');
  jumpBtn.textContent = `🎬 ${d.problemNo} 해설 바로보기 (${d.videoTimestamp})`;
  jumpBtn.onclick = () => openVideoModalAt(d.videoTimestamp, d.videoTitle, d.coachTip);

  let diff = gr.diffItems.find(item => item.id === d.id);
  if (!diff && gr.diffItems.length > 0) diff = gr.diffItems[0];

  const diffBody = document.getElementById('diffTableBody');
  if (diff && diff.cells) {
    diffBody.innerHTML = diff.cells.map(c => `
      <tr>
        <td style="font-weight:700;">${diff.sheetName}!${c.ref}</td>
        <td class="${c.match ? 'cell-correct' : 'cell-wrong'}">${c.userVal}</td>
        <td class="cell-correct">${c.ansVal}</td>
        <td style="font-weight:700;color:${c.match ? '#15803d' : '#dc2626'};">${c.match ? '⭕ 일치' : '❌ 불일치'}</td>
      </tr>
    `).join('');
  }
}

function renderRankingBars(containerId, items, labelKey, maxVal) {
  const el = document.getElementById(containerId);
  // Theme B Royal Blue steps (#1E3A8A -> #2563EB -> #38BDF8 -> #60A5FA -> #93C5FD -> #CBD5E1)
  const colors = [
    'linear-gradient(90deg, #1E40AF 0%, #2563EB 60%, #38BDF8 100%)',
    'linear-gradient(90deg, #2563EB 0%, #60A5FA 100%)',
    'linear-gradient(90deg, #38BDF8 0%, #93C5FD 100%)',
    '#60A5FA',
    '#93C5FD',
    '#CBD5E1',
    '#CBD5E1',
    '#CBD5E1',
    '#CBD5E1',
    '#CBD5E1'
  ];
  el.innerHTML = items.map((item, idx) => {
    const pct = Math.min(100, Math.round((item.score / maxVal) * 100));
    return `
      <div class="bar-row">
        <div style="text-align:right;color:#475569;font-family:var(--font-medium);">${item[labelKey]}</div>
        <div class="bar-track">
          <div class="bar-fill" style="width:${pct}%;background:${colors[idx] || '#CBD5E1'};"></div>
        </div>
        <div style="font-family:var(--font-medium);color:#0F172A;font-size:12px;">${item.score.toLocaleString()}</div>
      </div>
    `;
  }).join('');
}

function getBucketIndexForScore(buckets, score, maxScore) {
  if (score === null || score === undefined) return 0;
  if (score >= maxScore) return buckets.length - 1;
  if (maxScore === 100) {
    return Math.min(buckets.length - 2, Math.floor(score / 10));
  }
  const ratio = score / maxScore;
  return Math.min(buckets.length - 2, Math.floor(ratio * (buckets.length - 1)));
}

function renderAllPieCharts(distributions, gradeResult) {
  const totalScore = gradeResult ? gradeResult.totalScore : 0;
  const sBasic = gradeResult ? gradeResult.sectionScores.basic : 0;
  const sCalc = gradeResult ? gradeResult.sectionScores.calc : 0;
  const sAnalysis = gradeResult ? gradeResult.sectionScores.analysis : 0;
  const sEtc = gradeResult ? gradeResult.sectionScores.etc : 0;

  drawPieWithLegend('pieTotal', 'legendTotal', distributions.total, getBucketIndexForScore(distributions.total, totalScore, 100));
  drawPieWithLegend('pieP1', 'legendP1', distributions.p1, getBucketIndexForScore(distributions.p1, sBasic, 20));
  drawPieWithLegend('pieP2', 'legendP2', distributions.p2, getBucketIndexForScore(distributions.p2, sCalc, 40));
  drawPieWithLegend('pieP3', 'legendP3', distributions.p3, getBucketIndexForScore(distributions.p3, sAnalysis, 20));
  drawPieWithLegend('pieP4', 'legendP4', distributions.p4, getBucketIndexForScore(distributions.p4, sEtc, 20));
}

function drawPieWithLegend(canvasId, legendId, items, highlightIdx) {
  const canvas = document.getElementById(canvasId);
  const legend = document.getElementById(legendId);
  if (!canvas || !legend) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width;
  const h = canvas.height;
  ctx.clearRect(0, 0, w, h);

  const cx = w * 0.48;
  const cy = h * 0.5;
  const r = Math.min(w, h) * 0.39;
  const total = items.reduce((a, b) => a + b.value, 0);

  let startAngle = -Math.PI / 2;

  items.forEach((item, idx) => {
    const sliceAngle = (item.value / total) * Math.PI * 2;
    const midAngle = startAngle + sliceAngle / 2;
    const isHighlighted = idx === highlightIdx;
    const offset = isHighlighted ? 7 : 0;
    const ox = cx + Math.cos(midAngle) * offset;
    const oy = cy + Math.sin(midAngle) * offset;

    ctx.save();
    ctx.beginPath();
    ctx.moveTo(ox, oy + 6);
    ctx.arc(ox, oy + 6, r, startAngle, startAngle + sliceAngle);
    ctx.closePath();
    ctx.fillStyle = 'rgba(15, 23, 42, 0.22)';
    ctx.fill();
    ctx.restore();

    ctx.beginPath();
    ctx.moveTo(ox, oy);
    ctx.arc(ox, oy, r, startAngle, startAngle + sliceAngle);
    ctx.closePath();
    ctx.fillStyle = item.color;
    ctx.fill();
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = 1.2;
    ctx.stroke();

    if (item.value >= 9) {
      const tx = ox + Math.cos(midAngle) * (r * 0.62);
      const ty = oy + Math.sin(midAngle) * (r * 0.62);
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 10.5px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(`${item.value}%`, tx, ty);
    }

    startAngle += sliceAngle;
  });

  legend.innerHTML = items.map((item, idx) => {
    const isMine = idx === highlightIdx;
    return `
      <li class="${isMine ? 'my-bucket' : ''}">
        <span class="legend-dot" style="background:${item.color};"></span>
        <span>${item.label}${isMine ? '[내점수]' : ''}</span>
      </li>
    `;
  }).join('');
}

function drawRadarChart(radarStats) {
  const canvas = document.getElementById('radarCanvas');
  if (!canvas || !radarStats) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width;
  const h = canvas.height;
  ctx.clearRect(0, 0, w, h);

  const keys = Object.keys(radarStats);
  const cx = w / 2;
  const cy = h / 2 + 4;
  const maxR = 64;

  [0.25, 0.5, 0.75, 1].forEach(level => {
    ctx.beginPath();
    keys.forEach((_, i) => {
      const angle = (Math.PI * 2 * i) / keys.length - Math.PI / 2;
      const x = cx + Math.cos(angle) * maxR * level;
      const y = cy + Math.sin(angle) * maxR * level;
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.closePath();
    ctx.strokeStyle = '#cbd5e1';
    ctx.lineWidth = 1;
    ctx.stroke();
  });

  ctx.beginPath();
  keys.forEach((k, i) => {
    const val = (radarStats[k] || 0) / 100;
    const angle = (Math.PI * 2 * i) / keys.length - Math.PI / 2;
    const x = cx + Math.cos(angle) * maxR * val;
    const y = cy + Math.sin(angle) * maxR * val;
    if (i === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.closePath();
  ctx.fillStyle = 'rgba(37, 99, 235, 0.28)';
  ctx.fill();
  ctx.strokeStyle = '#2563eb';
  ctx.lineWidth = 2;
  ctx.stroke();

  // Labels
  ctx.fillStyle = '#0F172A';
  ctx.font = '500 10.5px "G마켓 산스 TTF Medium", sans-serif';
  keys.forEach((k, i) => {
    const angle = (Math.PI * 2 * i) / keys.length - Math.PI / 2;
    const x = cx + Math.cos(angle) * (maxR + 22);
    const y = cy + Math.sin(angle) * (maxR + 14);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const shortLabel = k.split('(')[0];
    ctx.fillText(`${shortLabel}(${radarStats[k]}%)`, x, y);
  });

  const sorted = Object.entries(radarStats).sort((a, b) => a[1] - b[1]);
  const weakest = sorted[0];
  const strongest = sorted[sorted.length - 1];
  const presEl = document.getElementById('radarPrescription');
  if (presEl) {
    presEl.innerHTML = `
      <p style="color:#0F172A;margin-bottom:5px;"><strong>📊 6대 핵심 역량 진단 리포트</strong></p>
      <p style="margin-bottom:4px;color:#059669;">🟢 <strong>최강점 영역:</strong> ${strongest[0]} (${strongest[1]}%)</p>
      <p style="margin-bottom:6px;color:#E11D48;">🔴 <strong>집중 보완 영역:</strong> ${weakest[0]} (${weakest[1]}%)</p>
      <div style="background:#FFFFFF;border:1px solid #93C5FD;padding:9px 11px;border-radius:8px;color:#1E3A8A;font-size:12px;">
        <strong>🎯 인생 코치의 맞춤 전략:</strong> 현재 <strong>${weakest[0]}</strong> 파트만 보완하면 평균 +10~16점 상승하여 안정적인 80점대 합격권에 진입할 수 있습니다!
      </div>
    `;
  }
}

function initModalsAndButtons() {
  document.getElementById('btnDownloadRound').addEventListener('click', () => {
    downloadExamWorkbook(false);
  });

  document.getElementById('btnDownloadAllZip').addEventListener('click', () => {
    downloadExamWorkbook(true);
  });

  document.getElementById('btnPrint').addEventListener('click', () => window.print());

  document.getElementById('btnToggleStats').addEventListener('click', () => {
    document.getElementById('statsSection').scrollIntoView({ behavior: 'smooth' });
  });

  const videoModal = document.getElementById('videoModal');
  document.getElementById('btnOpenVideoModal').addEventListener('click', () => {
    openVideoModalAt('11:45', '계산작업 1번: TIME 함수를 활용한 도착시간 계산', '출발시간(C3) + TIME(0, B3*8, 0) 수식을 작성하고 채우기 핸들로 드래그합니다.');
  });
  document.getElementById('btnCloseVideoModal').addEventListener('click', () => videoModal.classList.remove('open'));

  document.querySelectorAll('#videoChapterButtons button').forEach(btn => {
    btn.addEventListener('click', () => {
      openVideoModalAt(
        btn.getAttribute('data-ts'),
        btn.getAttribute('data-title'),
        '해당 타임스탬프 구간으로 이동했습니다. 문제 지시사항의 함수와 절대참조($) 위치를 함께 확인하세요!'
      );
    });
  });

  const helpModal = document.getElementById('helpModal');
  document.getElementById('btnHelpModal').addEventListener('click', () => helpModal.classList.add('open'));
  document.getElementById('btnCloseHelpModal').addEventListener('click', () => helpModal.classList.remove('open'));

  document.getElementById('btnLookupStudent').addEventListener('click', () => lookupStudentHistory());
  document.getElementById('btnLookupTeacher').addEventListener('click', () => lookupTeacherReport());
}

function openVideoModalAt(timestamp, title, guideText) {
  document.getElementById('videoTimestampBadge').textContent = `⏱️ ${timestamp} 구간 재생 중`;
  document.getElementById('videoChapterHeadline').textContent = title;
  document.getElementById('videoChapterGuide').innerHTML = guideText;
  document.getElementById('videoModal').classList.add('open');
}

async function lookupStudentHistory() {
  const nickname = (document.getElementById('lookupStudentInput').value || '짱아').trim();
  let records = localState.submissions.filter(s => s.studentName.toLowerCase() === nickname.toLowerCase());
  if (records.length === 0) records = localState.submissions.slice(0, 3);

  const radarKeys = ['데이터입력·서식', '조건부서식·필터', '함수(계산작업)', '데이터분석(피벗·통합)', '매크로·VBA', '차트작업'];
  const avgRadar = {};
  for (const k of radarKeys) {
    avgRadar[k] = Math.round(records.reduce((acc, r) => acc + (r.radarStats?.[k] ?? 75), 0) / records.length);
  }

  document.getElementById('radarTargetLabel').textContent = `${nickname} 수험생 누적 평균`;
  drawRadarChart(avgRadar);

  const panel = document.getElementById('studentHistoryPanel');
  panel.style.display = 'block';
  panel.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:6px;">
      <h4 style="color:#0f172a;">📒 [${nickname}] 수험생의 회차별 채점 이력 및 오답 노트 (총 ${records.length}회 응시)</h4>
      <span style="font-size:12px;color:#2563eb;font-weight:700;">상단 6각형 레이더 차트에 누적 역량이 반영되었습니다</span>
    </div>
    <table class="excel-diff-grid">
      <thead>
        <tr>
          <th>응시 회차</th>
          <th>기본(20)</th>
          <th>계산(40)</th>
          <th>분석(20)</th>
          <th>기타(20)</th>
          <th>총점(100)</th>
          <th>합격 판정</th>
          <th>주요 취약 파트 기록</th>
          <th>응시 일시</th>
        </tr>
      </thead>
      <tbody>
        ${records.map(r => `
          <tr>
            <td style="font-weight:700;">${r.roundName}</td>
            <td>${r.basic}</td>
            <td>${r.calc}</td>
            <td>${r.analysis}</td>
            <td>${r.etc}</td>
            <td style="font-weight:800;color:${r.total >= 70 ? '#15803d' : '#dc2626'};">${r.total}점</td>
            <td>${r.total >= 70 ? '🟢 합격권' : '🔴 보완필요'}</td>
            <td>${r.weakPoint}</td>
            <td style="color:#64748b;">${r.submittedAt}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}

async function lookupTeacherReport() {
  const teacherId = (document.getElementById('lookupTeacherInput').value || 'Zzang_AA').trim();
  if (teacherId.length < 4) {
    alert('선생님 ID를 4자 이상 입력해 주세요. (예: Zzang_AA)');
    return;
  }
  let records = localState.submissions.filter(s => s.teacherId.toLowerCase() === teacherId.toLowerCase());
  if (records.length === 0) records = localState.submissions;

  const passCount = records.filter(r => r.total >= 70).length;
  const passRate = Math.round((passCount / records.length) * 100);
  const avgTotal = (records.reduce((a, b) => a + b.total, 0) / records.length).toFixed(1);

  const panel = document.getElementById('teacherReportPanel');
  panel.style.display = 'block';
  panel.innerHTML = `
    <div style="display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:10px;margin-bottom:12px;">
      <div>
        <h4 style="color:#0f172a;font-size:14.5px;">🏫 선생님(${teacherId}) 클래스 성적 관리 대시보드</h4>
        <p style="font-size:12px;color:#475569;">제출 인원: <strong>${records.length}명</strong> | 평균 점수: <strong>${avgTotal}점</strong> | 합격률(70점↑): <strong>${passRate}% (${passCount}명)</strong></p>
      </div>
      <button type="button" id="btnDownloadTeacherExcel" class="btn-grade" style="display:inline-flex;align-items:center;gap:6px;">
        📥 학급 성적표 엑셀(.csv) 다운로드
      </button>
    </div>
    <table class="excel-diff-grid">
      <thead>
        <tr>
          <th>학생 닉네임</th>
          <th>응시 회차</th>
          <th>기본(20)</th>
          <th>계산(40)</th>
          <th>분석(20)</th>
          <th>기타(20)</th>
          <th>총점</th>
          <th>합격 신호등</th>
          <th>맞춤 지도 포인트</th>
        </tr>
      </thead>
      <tbody>
        ${records.map(r => {
          const signal = r.total >= 80 ? '🟢 안정권' : (r.total >= 70 ? '🟡 턱걸이합격' : '🔴 집중케어');
          return `
            <tr>
              <td style="font-weight:700;">${r.studentName}</td>
              <td>${r.roundName}</td>
              <td>${r.basic}</td>
              <td>${r.calc}</td>
              <td>${r.analysis}</td>
              <td>${r.etc}</td>
              <td style="font-weight:800;">${r.total}점</td>
              <td>${signal}</td>
              <td>${r.weakPoint}</td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table>
  `;

  document.getElementById('btnDownloadTeacherExcel').addEventListener('click', () => {
    downloadTeacherGradebookClient(teacherId, records);
  });
}

function downloadTeacherGradebookClient(teacherId, records) {
  const headers = ['순번', '학생 닉네임', '선생님 ID', '응시 회차', '기본작업(20)', '계산작업(40)', '분석작업(20)', '기타작업(20)', '총점(100)', '합격여부', '취약 파트 진단', '응시 일시'];
  const csvRows = [
    headers.join(','),
    ...records.map((r, i) => [
      i + 1,
      `"${r.studentName}"`,
      `"${r.teacherId}"`,
      `"${r.roundName}"`,
      r.basic,
      r.calc,
      r.analysis,
      r.etc,
      r.total,
      `"${r.total >= 70 ? '합격(PASS)' : '불합격(보완필요)'}"`,
      `"${r.weakPoint}"`,
      `"${r.submittedAt}"`
    ].join(','))
  ];
  const bom = '\uFEFF';
  const blob = new Blob([bom + csvRows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `컴활2급실기_선생님(${teacherId})_학급성적표.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function initUpgradeStepRibbon() {
  const pills = document.querySelectorAll('.upgrade-pill');
  pills.forEach(pill => {
    pill.addEventListener('click', async () => {
      pills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      const step = pill.getAttribute('data-step');

      if (step === '1') {
        const el = document.getElementById('preflightBox');
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.style.outline = '3px solid #2563eb';
        setTimeout(() => (el.style.outline = 'none'), 2000);
      } else if (step === '2') {
        const el = document.getElementById('visualDiffPanel');
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.style.outline = '3px solid #2563eb';
        setTimeout(() => (el.style.outline = 'none'), 2000);
      } else if (step === '3') {
        openVideoModalAt('15:20', '계산작업 2번: [표2] IF, AND, AVERAGE 중첩 함수 완벽 공략', '제시된 3개 함수를 모두 사용하고 평균 범위($I$3:$I$10)를 절대참조로 고정하는 핵심 포인트를 영상으로 확인하세요!');
      } else if (step === '4') {
        await lookupStudentHistory();
        const el = document.getElementById('radarUpgradeBox');
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.style.outline = '3px solid #2563eb';
        setTimeout(() => (el.style.outline = 'none'), 2000);
      } else if (step === '5') {
        await lookupTeacherReport();
        const el = document.getElementById('teacherReportPanel');
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.style.outline = '3px solid #16a34a';
        setTimeout(() => (el.style.outline = 'none'), 2000);
      } else if (step === '6') {
        const el = document.getElementById('problemGeneratorSection');
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        el.style.outline = '3px solid #2563EB';
        setTimeout(() => (el.style.outline = 'none'), 2000);
      } else if (step === '7') {
        const el = document.getElementById('sampleGuideSection');
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        el.style.outline = '3px solid #2563EB';
        setTimeout(() => (el.style.outline = 'none'), 2000);
      } else if (step === '8') {
        const el = document.getElementById('wrongNoteSection');
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        el.style.outline = '3px solid #2563EB';
        setTimeout(() => (el.style.outline = 'none'), 2000);
      }
    });
  });
}
