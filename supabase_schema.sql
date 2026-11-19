-- ============================================================================
-- comlicense 프로젝트 Supabase(PostgreSQL 17) 데이터베이스 스키마 및 초기 시드
-- Project ID: ptargvvfuhcmobhhzxlu (https://ptargvvfuhcmobhhzxlu.supabase.co)
-- ============================================================================

-- 1. 컴퓨터활용능력 2급 실기 회차별 문제은행 및 AI 생성 모의고사 테이블
CREATE TABLE IF NOT EXISTS public.comlicense_exams (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    exam_code TEXT UNIQUE NOT NULL,
    publisher TEXT NOT NULL DEFAULT 'youngjin',
    exam_year INT NOT NULL DEFAULT 2026,
    subject TEXT NOT NULL DEFAULT 'excel',
    level INT NOT NULL DEFAULT 2,
    title TEXT NOT NULL,
    theme_name TEXT NOT NULL,
    is_ai_generated BOOLEAN NOT NULL DEFAULT FALSE,
    time_limit_minutes INT NOT NULL DEFAULT 40,
    passing_score INT NOT NULL DEFAULT 70,
    basic_questions JSONB NOT NULL DEFAULT '[]'::jsonb,
    calc_questions JSONB NOT NULL DEFAULT '[]'::jsonb,
    analysis_questions JSONB NOT NULL DEFAULT '[]'::jsonb,
    etc_questions JSONB NOT NULL DEFAULT '[]'::jsonb,
    sample_rows JSONB NOT NULL DEFAULT '[]'::jsonb,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- 2. 수험생 채점 제출 이력 및 5대 고도화 분석 리포트 테이블
CREATE TABLE IF NOT EXISTS public.comlicense_submissions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    exam_code TEXT NOT NULL REFERENCES public.comlicense_exams(exam_code),
    student_nickname TEXT NOT NULL DEFAULT '짱아',
    teacher_id TEXT NOT NULL DEFAULT 'Zzang_AA',
    file_name TEXT NOT NULL,
    score_basic INT NOT NULL CHECK (score_basic BETWEEN 0 AND 20),
    score_calc INT NOT NULL CHECK (score_calc BETWEEN 0 AND 40),
    score_analysis INT NOT NULL CHECK (score_analysis BETWEEN 0 AND 20),
    score_etc INT NOT NULL CHECK (score_etc BETWEEN 0 AND 20),
    total_score INT NOT NULL CHECK (total_score BETWEEN 0 AND 100),
    passed BOOLEAN GENERATED ALWAYS AS (total_score >= 70) STORED,
    weak_point TEXT DEFAULT '',
    preflight_checks JSONB NOT NULL DEFAULT '[]'::jsonb,
    deductions JSONB NOT NULL DEFAULT '[]'::jsonb,
    diff_items JSONB NOT NULL DEFAULT '[]'::jsonb,
    radar_stats JSONB NOT NULL DEFAULT '{}'::jsonb,
    submitted_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- 3. 조회 성능 최적화 인덱스
CREATE INDEX IF NOT EXISTS idx_comlicense_submissions_student ON public.comlicense_submissions(student_nickname);
CREATE INDEX IF NOT EXISTS idx_comlicense_submissions_teacher ON public.comlicense_submissions(teacher_id);
CREATE INDEX IF NOT EXISTS idx_comlicense_submissions_exam_code ON public.comlicense_submissions(exam_code);
CREATE INDEX IF NOT EXISTS idx_comlicense_submissions_submitted_at ON public.comlicense_submissions(submitted_at DESC);

-- 4. 실시간 집계 뷰 (security_invoker = true 적용)
CREATE OR REPLACE VIEW public.comlicense_student_rankings_view
WITH (security_invoker = true) AS
SELECT
    student_nickname AS raw_name,
    LEFT(student_nickname, 1) || '******' AS masked_name,
    COUNT(*)::INT AS submission_count,
    SUM(total_score)::INT AS cumulative_score,
    MAX(total_score)::INT AS best_score,
    ROUND(AVG(total_score), 1) AS avg_score
FROM public.comlicense_submissions
GROUP BY student_nickname
ORDER BY SUM(total_score) DESC;

CREATE OR REPLACE VIEW public.comlicense_teacher_rankings_view
WITH (security_invoker = true) AS
SELECT
    teacher_id AS raw_id,
    LEFT(teacher_id, 1) || '******' AS masked_id,
    COUNT(*)::INT AS student_submissions,
    SUM(ROUND(total_score / 2.0))::INT AS activity_score,
    ROUND(AVG(total_score), 1) AS class_avg_score,
    ROUND((COUNT(*) FILTER (WHERE total_score >= 70)::NUMERIC / NULLIF(COUNT(*), 0)) * 100, 1) AS pass_rate
FROM public.comlicense_submissions
WHERE LENGTH(TRIM(teacher_id)) >= 4
GROUP BY teacher_id
ORDER BY SUM(ROUND(total_score / 2.0)) DESC;

CREATE OR REPLACE VIEW public.comlicense_round_stats_view
WITH (security_invoker = true) AS
SELECT
    exam_code,
    COUNT(*)::INT AS participant_count,
    ROUND(AVG(total_score), 2) AS average_score,
    COUNT(*) FILTER (WHERE total_score >= 70)::INT AS pass_count
FROM public.comlicense_submissions
GROUP BY exam_code;

-- 5. RLS(행 수준 보안) 정책
ALTER TABLE public.comlicense_exams ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.comlicense_submissions ENABLE ROW LEVEL SECURITY;

-- 6. 웹사이트 테마(Theme B 로열 블루) · 폰트(G마켓 산스) · 브랜딩 환경설정 테이블
CREATE TABLE IF NOT EXISTS public.comlicense_site_settings (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    setting_key TEXT UNIQUE NOT NULL,
    theme_code TEXT NOT NULL DEFAULT 'Theme B · 로열 블루',
    font_heading TEXT NOT NULL DEFAULT 'G마켓 산스 TTF Medium',
    font_body TEXT NOT NULL DEFAULT 'G마켓 산스 TTF Light',
    bg_color TEXT NOT NULL DEFAULT '#FFFFFF',
    primary_color TEXT NOT NULL DEFAULT '#2563EB',
    grad_primary TEXT NOT NULL DEFAULT 'linear-gradient(135deg, #1E40AF 0%, #2563EB 52%, #38BDF8 100%)',
    default_teacher_id TEXT NOT NULL DEFAULT 'Zzang_AA',
    default_student_nickname TEXT NOT NULL DEFAULT '짱아',
    study_cafe_title TEXT NOT NULL DEFAULT 'Zzang_AA''s스터디카페',
    contact_email TEXT NOT NULL DEFAULT 'test@hanafostest.com',
    theme_tokens JSONB NOT NULL DEFAULT '{}'::jsonb,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE public.comlicense_site_settings ENABLE ROW LEVEL SECURITY;

-- 7. 수험생 개인별 맞춤 오답노트 & 복습 메모·암기 체크 테이블
CREATE TABLE IF NOT EXISTS public.comlicense_wrong_notes (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    student_nickname TEXT NOT NULL DEFAULT '짱아',
    teacher_id TEXT NOT NULL DEFAULT 'Zzang_AA',
    exam_code TEXT NOT NULL DEFAULT '대표기출따라하기',
    problem_no TEXT NOT NULL,
    task_name TEXT NOT NULL,
    deduction INT NOT NULL DEFAULT -8,
    user_input TEXT NOT NULL DEFAULT '',
    correct_formula TEXT NOT NULL DEFAULT '',
    error_reason TEXT NOT NULL DEFAULT '',
    coach_tip TEXT NOT NULL DEFAULT '',
    student_memo TEXT NOT NULL DEFAULT '',
    is_mastered BOOLEAN NOT NULL DEFAULT FALSE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_comlicense_wrong_notes_student ON public.comlicense_wrong_notes(student_nickname);
ALTER TABLE public.comlicense_wrong_notes ENABLE ROW LEVEL SECURITY;
