const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const {
  runPreflightCheck,
  gradeWorkbook,
  generateTeacherGradebookXlsx,
  ensureSampleTestFiles,
  createZipBuffer
} = require('./grader');

const PORT = process.env.PORT || 3000;
const ETC_DIR = path.join(__dirname, 'etc');
const PUBLIC_DIR = path.join(__dirname, 'public');

// Generate sample test .xlsm variants in etc/ on startup
ensureSampleTestFiles(ETC_DIR);

// In-memory store pre-populated with realistic comlicense.co.kr data matching screenshots
const state = {
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
      roundName: '기출문제 01회',
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
      roundName: '기출문제 02회',
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

function maskName(str) {
  if (!str) return '익******';
  return str.charAt(0) + '******';
}

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(data));
}

function collectBodyBuffer(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  try {
    // 1. GET /api/stats
    if (pathname === '/api/stats' && req.method === 'GET') {
      return sendJson(res, 200, {
        participantCount: state.participantCount,
        averageScore: state.averageScore,
        studentRankings: state.studentRankings.slice(0, 10),
        teacherRankings: state.teacherRankings.slice(0, 5),
        distributions: {
          total: [
            { label: '0~10점미만', value: 49.3, color: '#2b6cb0' },
            { label: '10~20점미만', value: 1.8, color: '#c53030' },
            { label: '20~30점미만', value: 1.5, color: '#dd6b20' },
            { label: '30~40점미만', value: 2.1, color: '#2f855a' },
            { label: '40~50점미만', value: 2.9, color: '#805ad5' },
            { label: '50~60점미만', value: 3.4, color: '#319795' },
            { label: '60~70점미만', value: 5.7, color: '#d53f8c' },
            { label: '70~80점미만', value: 9.5, color: '#38a169' },
            { label: '80~90점미만', value: 13.8, color: '#9b2c2c' },
            { label: '90~100점미만', value: 6.2, color: '#2c5282' },
            { label: '100', value: 3.8, color: '#6b46c1' }
          ],
          p1: [
            { label: '0~5점미만', value: 48.9, color: '#2b6cb0' },
            { label: '5~10점미만', value: 1.1, color: '#c53030' },
            { label: '10~15점미만', value: 11.6, color: '#dd6b20' },
            { label: '15~20점미만', value: 24.5, color: '#2f855a' },
            { label: '20', value: 13.9, color: '#805ad5' }
          ],
          p2: [
            { label: '0~5점미만', value: 54.4, color: '#2b6cb0' },
            { label: '10~15점미만', value: 3.2, color: '#c53030' },
            { label: '15~20점미만', value: 2.8, color: '#dd6b20' },
            { label: '25~30점미만', value: 5.5, color: '#2f855a' },
            { label: '30~35점미만', value: 13.8, color: '#805ad5' },
            { label: '40', value: 20.3, color: '#319795' }
          ],
          p3: [
            { label: '0~5점미만', value: 51.1, color: '#2b6cb0' },
            { label: '10~15점미만', value: 10.3, color: '#c53030' },
            { label: '20', value: 38.6, color: '#dd6b20' }
          ],
          p4: [
            { label: '0~5점미만', value: 51.5, color: '#2b6cb0' },
            { label: '5~10점미만', value: 2.1, color: '#c53030' },
            { label: '10~15점미만', value: 2.8, color: '#dd6b20' },
            { label: '15~20점미만', value: 7.2, color: '#2f855a' },
            { label: '20', value: 36.4, color: '#805ad5' }
          ]
        }
      });
    }

    // 2. POST /api/preflight (Feature 1: 브라우저 즉시 사전 검사)
    if (pathname === '/api/preflight' && req.method === 'POST') {
      const bodyBuf = await collectBodyBuffer(req);
      const fileName = decodeURIComponent(req.headers['x-file-name'] || '업로드파일.xlsm');
      const result = runPreflightCheck(bodyBuf, fileName);
      return sendJson(res, 200, result);
    }

    // 3. POST /api/grade (Actual File Upload Grading)
    if (pathname === '/api/grade' && req.method === 'POST') {
      const bodyBuf = await collectBodyBuffer(req);
      const fileName = decodeURIComponent(req.headers['x-file-name'] || '업로드파일.xlsm');
      const studentName = decodeURIComponent(req.headers['x-student-name'] || '오미진').trim() || '오미진';
      const teacherId = decodeURIComponent(req.headers['x-teacher-id'] || 'youngjin01').trim() || 'youngjin01';
      const roundName = decodeURIComponent(req.headers['x-round-name'] || '대표기출따라하기');

      const preflight = runPreflightCheck(bodyBuf, fileName);
      const gradeResult = gradeWorkbook(bodyBuf);

      // Record submission
      state.participantCount += 1;
      const nowStr = new Date().toISOString().replace('T', ' ').slice(0, 19);
      const topWeak = gradeResult.deductions.length > 0
        ? `${gradeResult.deductions[0].problemNo} (${gradeResult.deductions[0].taskName})`
        : '전 영역 완벽(만점)';

      state.submissions.unshift({
        studentName,
        teacherId,
        roundName,
        basic: gradeResult.sectionScores.basic,
        calc: gradeResult.sectionScores.calc,
        analysis: gradeResult.sectionScores.analysis,
        etc: gradeResult.sectionScores.etc,
        total: gradeResult.totalScore,
        weakPoint: topWeak,
        submittedAt: nowStr,
        radarStats: gradeResult.radarStats
      });

      // Update cumulative rankings
      const existingStu = state.studentRankings.find(s => s.rawName === studentName);
      if (existingStu) {
        existingStu.score += gradeResult.totalScore;
      } else {
        state.studentRankings.push({
          name: maskName(studentName),
          rawName: studentName,
          score: 1400 + gradeResult.totalScore
        });
      }
      state.studentRankings.sort((a, b) => b.score - a.score);

      if (teacherId && teacherId.length >= 4) {
        const existingT = state.teacherRankings.find(t => t.rawId === teacherId);
        if (existingT) {
          existingT.score += Math.round(gradeResult.totalScore / 2);
        } else {
          state.teacherRankings.push({
            id: maskName(teacherId),
            rawId: teacherId,
            score: 300 + Math.round(gradeResult.totalScore / 2)
          });
        }
        state.teacherRankings.sort((a, b) => b.score - a.score);
      }

      return sendJson(res, 200, {
        fileName,
        studentName,
        teacherId,
        roundName,
        preflight,
        gradeResult,
        participantCount: state.participantCount,
        averageScore: state.averageScore,
        studentRankings: state.studentRankings.slice(0, 10),
        teacherRankings: state.teacherRankings.slice(0, 5)
      });
    }

    // 4. POST /api/quick-demo (1-click grading with local sample files in etc/)
    if (pathname === '/api/quick-demo' && req.method === 'POST') {
      const bodyBuf = await collectBodyBuffer(req);
      const payload = JSON.parse(bodyBuf.toString('utf8') || '{}');
      const preset = payload.preset || 'pass82';
      let targetFile = '학생제출샘플_82점(합격).xlsm';
      if (preset === 'raw0') targetFile = '문제샘플(작성 전).xlsm';
      else if (preset === 'fail51') targetFile = '학생제출샘플_51점(불합격_오답다수).xlsm';
      else if (preset === 'perfect100') targetFile = '학생제출샘플_100점(모범정답).xlsm';

      const filePath = path.join(ETC_DIR, targetFile);
      const fileBuf = fs.readFileSync(filePath);
      const studentName = (payload.studentName || '오미진').trim();
      const teacherId = (payload.teacherId || 'youngjin01').trim();
      const roundName = payload.roundName || '대표기출따라하기';

      const preflight = runPreflightCheck(fileBuf, targetFile);
      const gradeResult = gradeWorkbook(fileBuf);

      state.participantCount += 1;
      const nowStr = new Date().toISOString().replace('T', ' ').slice(0, 19);
      const topWeak = gradeResult.deductions.length > 0
        ? `${gradeResult.deductions[0].problemNo} (${gradeResult.deductions[0].taskName})`
        : '전 영역 완벽(만점)';

      state.submissions.unshift({
        studentName,
        teacherId,
        roundName,
        basic: gradeResult.sectionScores.basic,
        calc: gradeResult.sectionScores.calc,
        analysis: gradeResult.sectionScores.analysis,
        etc: gradeResult.sectionScores.etc,
        total: gradeResult.totalScore,
        weakPoint: topWeak,
        submittedAt: nowStr,
        radarStats: gradeResult.radarStats
      });

      return sendJson(res, 200, {
        fileName: targetFile,
        studentName,
        teacherId,
        roundName,
        preflight,
        gradeResult,
        participantCount: state.participantCount,
        averageScore: state.averageScore,
        studentRankings: state.studentRankings.slice(0, 10),
        teacherRankings: state.teacherRankings.slice(0, 5)
      });
    }

    // 5. GET /api/student-history (Student nickname lookup + Feature 4 Radar Chart)
    if (pathname === '/api/student-history' && req.method === 'GET') {
      const nickname = (parsedUrl.query.nickname || '오미진').trim();
      let records = state.submissions.filter(s => s.studentName.toLowerCase() === nickname.toLowerCase());
      if (records.length === 0) {
        records = state.submissions.slice(0, 3);
      }
      // Aggregate average radar stats
      const radarKeys = ['데이터입력·서식', '조건부서식·필터', '함수(계산작업)', '데이터분석(피벗·통합)', '매크로·VBA', '차트작업'];
      const avgRadar = {};
      for (const k of radarKeys) {
        avgRadar[k] = Math.round(records.reduce((acc, r) => acc + (r.radarStats?.[k] ?? 75), 0) / records.length);
      }
      return sendJson(res, 200, {
        nickname,
        records,
        avgRadar
      });
    }

    // 6. GET /api/teacher-report (Teacher ID lookup + Class Dashboard)
    if (pathname === '/api/teacher-report' && req.method === 'GET') {
      const teacherId = (parsedUrl.query.teacherId || 'youngjin01').trim();
      let records = state.submissions.filter(s => s.teacherId.toLowerCase() === teacherId.toLowerCase());
      if (records.length === 0) {
        records = state.submissions;
      }
      const passCount = records.filter(r => r.total >= 70).length;
      const avgTotal = (records.reduce((a, b) => a + b.total, 0) / records.length).toFixed(1);
      return sendJson(res, 200, {
        teacherId,
        totalStudents: records.length,
        passCount,
        passRate: Math.round((passCount / records.length) * 100),
        avgTotal,
        records
      });
    }

    // 7. GET /api/teacher-export (Feature 5: Real .xlsx Gradebook Download)
    if (pathname === '/api/teacher-export' && req.method === 'GET') {
      const teacherId = (parsedUrl.query.teacherId || 'youngjin01').trim();
      let records = state.submissions.filter(s => s.teacherId.toLowerCase() === teacherId.toLowerCase());
      if (records.length === 0) records = state.submissions;

      const xlsxBuf = generateTeacherGradebookXlsx(teacherId, records);
      const downloadName = encodeURIComponent(`컴활2급실기_선생님(${teacherId})_학급성적표.xlsx`);
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${downloadName}`,
        'Content-Length': xlsxBuf.length
      });
      return res.end(xlsxBuf);
    }

    // 8. GET /api/download-problem & /api/download-all-zip
    if (pathname === '/api/download-problem' && req.method === 'GET') {
      const file = parsedUrl.query.variant || '문제샘플(작성 전).xlsm';
      const safeFile = path.basename(file);
      const fullPath = path.join(ETC_DIR, safeFile);
      if (!fs.existsSync(fullPath)) {
        res.writeHead(404);
        return res.end('File not found');
      }
      const buf = fs.readFileSync(fullPath);
      res.writeHead(200, {
        'Content-Type': 'application/vnd.ms-excel.sheet.macroEnabled.12',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(safeFile)}`,
        'Content-Length': buf.length
      });
      return res.end(buf);
    }

    if (pathname === '/api/download-all-zip' && req.method === 'GET') {
      const xlsmFiles = fs.readdirSync(ETC_DIR).filter(f => f.endsWith('.xlsm'));
      const zipEntries = xlsmFiles.map(f => ({
        name: f,
        data: fs.readFileSync(path.join(ETC_DIR, f))
      }));
      const zipBuf = createZipBuffer(zipEntries);
      res.writeHead(200, {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent('2026_이기적_컴활2급실기_전체문제파일.zip')}`,
        'Content-Length': zipBuf.length
      });
      return res.end(zipBuf);
    }

    // Serve Main App for / and /chk_page
    if (pathname === '/' || pathname === '/chk_page') {
      const indexHtml = fs.readFileSync(path.join(PUBLIC_DIR, 'index.html'), 'utf8');
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(indexHtml);
    }

    // Serve static files from public/ or etc/
    const decodedPath = decodeURIComponent(pathname);
    const candidatePublic = path.join(PUBLIC_DIR, decodedPath);
    const candidateEtc = path.join(ETC_DIR, path.basename(decodedPath));

    let targetPath = null;
    if (fs.existsSync(candidatePublic) && fs.statSync(candidatePublic).isFile()) {
      targetPath = candidatePublic;
    } else if (fs.existsSync(candidateEtc) && fs.statSync(candidateEtc).isFile()) {
      targetPath = candidateEtc;
    }

    if (targetPath) {
      const ext = path.extname(targetPath).toLowerCase();
      const mimeMap = {
        '.html': 'text/html; charset=utf-8',
        '.js': 'application/javascript; charset=utf-8',
        '.css': 'text/css; charset=utf-8',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.svg': 'image/svg+xml'
      };
      res.writeHead(200, { 'Content-Type': mimeMap[ext] || 'application/octet-stream' });
      return res.end(fs.readFileSync(targetPath));
    }

    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Not Found');
  } catch (err) {
    console.error(err);
    sendJson(res, 500, { error: err.message });
  }
});

server.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}/chk_page?co=youngjin&year=2026&subject=excel&level=2`);
});
