/**
 * GitHub Direct Repository Creator & Multi-File Publisher (Zero git.exe dependency)
 * Automatically reads GITHUB_TOKEN / ghp_... from .env or command-line argument,
 * and strictly excludes .env files from being uploaded to GitHub.
 */
const fs = require('fs');
const path = require('path');

function loadTokenFromEnvFile() {
  const envPaths = [
    path.join(__dirname, '.env'),
    path.join(__dirname, '.env.local')
  ];
  for (const p of envPaths) {
    if (fs.existsSync(p)) {
      const content = fs.readFileSync(p, 'utf8');
      const match = content.match(/(?:GITHUB_TOKEN|GITHUB_PERSONAL_ACCESS_TOKEN|GH_TOKEN|GITHUB_PAT)\s*=\s*["']?([^\s"']+)["']?/i)
        || content.match(/(ghp_[A-Za-z0-9_]+|github_pat_[A-Za-z0-9_]+)/);
      if (match && match[1]) {
        return match[1].trim();
      }
    }
  }
  return null;
}

const TOKEN = process.argv[2]
  || process.env.GITHUB_PERSONAL_ACCESS_TOKEN
  || process.env.GITHUB_TOKEN
  || loadTokenFromEnvFile();
const REPO_NAME = process.argv[3] || 'comlicense';

if (!TOKEN) {
  console.error('토큰을 찾을 수 없습니다. c:\\agapp_project\\comlicense\\.env 파일에 GITHUB_TOKEN=ghp_... 를 넣거나 인자로 전달해 주세요.');
  process.exit(1);
}

async function ghRequest(endpoint, method = 'GET', body = null) {
  const res = await fetch(`https://api.github.com${endpoint}`, {
    method,
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'comlicense-uploader',
      ...(body ? { 'Content-Type': 'application/json' } : {})
    },
    ...(body ? { body: JSON.stringify(body) } : {})
  });
  const text = await res.text();
  const data = text ? JSON.parse(text) : null;
  if (!res.ok && res.status !== 422) {
    throw new Error(`GitHub API Error (${res.status}): ${text}`);
  }
  return { status: res.status, data };
}

function collectFiles(dir, baseDir = dir) {
  let results = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    // 보안: .env 파일 및 시스템 폴더는 절대 깃허브에 업로드하지 않음
    if (
      entry.name === 'node_modules' ||
      entry.name === '.git' ||
      entry.name === '.vercel' ||
      entry.name.startsWith('.env')
    ) {
      continue;
    }
    const fullPath = path.join(dir, entry.name);
    const relPath = path.relative(baseDir, fullPath).replace(/\\/g, '/');
    if (entry.isDirectory()) {
      results = results.concat(collectFiles(fullPath, baseDir));
    } else {
      results.push({ fullPath, relPath });
    }
  }
  return results;
}

async function main() {
  const { data: user } = await ghRequest('/user');
  const owner = user.login;
  console.log(`✅ GitHub 인증 성공: ${owner}`);

  // 1. Create repository if it doesn't exist
  const createRes = await ghRequest('/user/repos', 'POST', {
    name: REPO_NAME,
    description: "Zzang_AA's 컴퓨터활용능력 2급 실기 스마트 웹 자동채점 & 모의고사 생성 플랫폼",
    private: false,
    auto_init: true
  });
  if (createRes.status === 201) {
    console.log(`✅ 새 저장소 생성 완료: https://github.com/${owner}/${REPO_NAME}`);
    await new Promise(r => setTimeout(r, 2000));
  } else {
    console.log(`ℹ️ 기존 저장소 연결됨: https://github.com/${owner}/${REPO_NAME}`);
  }

  // 2. Upload project files via Contents API
  const files = collectFiles(__dirname);
  for (const file of files) {
    const contentBase64 = fs.readFileSync(file.fullPath).toString('base64');
    let sha = undefined;
    try {
      const existing = await fetch(`https://api.github.com/repos/${owner}/${REPO_NAME}/contents/${encodeURI(file.relPath)}`, {
        headers: {
          Authorization: `Bearer ${TOKEN}`,
          Accept: 'application/vnd.github+json',
          'User-Agent': 'comlicense-uploader'
        }
      });
      if (existing.ok) {
        const exJson = await existing.json();
        sha = exJson.sha;
      }
    } catch (_) {}

    await ghRequest(`/repos/${owner}/${REPO_NAME}/contents/${encodeURI(file.relPath)}`, 'PUT', {
      message: `feat: upload ${file.relPath}`,
      content: contentBase64,
      ...(sha ? { sha } : {})
    });
    console.log(`📤 업로드 완료: ${file.relPath}`);
  }

  console.log(`\n🎉 전체 파일 업로드 성공! 저장소 주소: https://github.com/${owner}/${REPO_NAME}`);
}

main().catch(err => {
  console.error(err.message);
  process.exit(1);
});
