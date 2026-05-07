# Dr. Reju-All Amazon Dashboard

Vercel 정적 호스팅 + Serverless Function (Google Sheets API + 서비스 계정)

## 구조
```
dashboard/
├── index.html         # 프론트 (Vercel 정적 서빙)
├── api/data.js        # Serverless Function — 시트에서 JSON 반환
├── package.json       # googleapis 의존성
├── vercel.json        # Vercel 설정
└── Code.gs            # (선택) Apps Script 호환 백엔드
```

## 배포 절차

### 1. Google Cloud — 서비스 계정 생성

1. [console.cloud.google.com](https://console.cloud.google.com) → 새 프로젝트 생성 (또는 기존 사용)
2. **API 및 서비스 > 라이브러리** → "Google Sheets API" 검색 → **사용 설정**
3. **API 및 서비스 > 사용자 인증 정보** → **사용자 인증 정보 만들기 > 서비스 계정**
4. 이름 입력 (예: `dashboard-reader`) → 만들기 → 역할 없이 완료
5. 생성된 서비스 계정 클릭 → **키** 탭 → **키 추가 > 새 키 만들기 > JSON**
6. JSON 파일이 다운로드됩니다 (절대 공개 금지)

### 2. 시트 공유

1. Google Sheets 열기 (`10d21g2iUkqb2uRVEw9ZCSQgmOAqKVyuIiOSNBozl1ok`)
2. **공유** 버튼 클릭
3. 다운로드한 JSON 파일 안의 `client_email` 값 복사 (예: `dashboard-reader@xxx.iam.gserviceaccount.com`)
4. 그 이메일을 **뷰어** 권한으로 추가
5. 시트는 **비공개로 유지** — 외부 노출 없음

### 3. GitHub 푸시

```bash
cd C:\Users\USER\Desktop\MD\dashboard
git init
git add .
git commit -m "Initial dashboard"
git branch -M main
git remote add origin https://github.com/<YOUR_USER>/<REPO>.git
git push -u origin main
```

### 4. Vercel 배포

1. [vercel.com](https://vercel.com) 로그인 → **Add New > Project**
2. GitHub 저장소 선택 → **Import**
3. **Environment Variables** 추가 (Deploy 전):
   - **Name**: `GOOGLE_SERVICE_ACCOUNT_JSON`
   - **Value**: 다운로드한 JSON 파일 **전체 내용**을 복사 붙여넣기
4. **Deploy** 클릭

배포 완료 후 `https://<project>.vercel.app` 접속 → 데이터 자동 로드.

## 동작 방식

- 프론트는 `/api/data`를 fetch → 같은 도메인이라 CORS 없음
- Serverless Function이 서비스 계정으로 시트 읽기 → JSON 반환
- 응답에 `s-maxage=60` 헤더 → Vercel Edge에서 60초 캐시
- 함수 인스턴스 메모리 캐시도 60초 유지 (콜드 스타트 외엔 ~50ms)

## 캐시 강제 갱신
URL 끝에 `?fresh=1` 추가 → 캐시 무시하고 새로 가져옴.

## 보안
- 서비스 계정 JSON은 **Vercel 환경변수에만** 저장 (Git에 커밋 금지 — `.gitignore` 처리됨)
- API 키 같은 클라이언트 노출 없음
- 시트 자체는 비공개 유지

## 로컬 테스트 (선택)
```bash
npm i -g vercel
vercel dev
```
환경변수는 `.env.local`에 `GOOGLE_SERVICE_ACCOUNT_JSON='{"type":"service_account",...}'` 형태로.
