# WordPress Bot with Gemini

Google Gemini AI를 활용하여 SEO 최적화된 WordPress 블로그 포스트를 자동 생성하고 업로드하는 GUI 애플리케이션입니다.

## 설치 방법

### 1. 필수 패키지 설치

```bash
pip install -r requirements.txt
```

### 2. 환경 변수 설정

프로젝트 루트 디렉토리에 `.env` 파일을 생성하고 다음 내용을 추가하세요:

```env
# Gemini API Key
GEMINI_API_KEY=your_gemini_api_key_here

# WordPress 설정
WORDPRESS_URL=https://your-wordpress-site.com
WORDPRESS_USERNAME=your_username
WORDPRESS_SECRET=your_app_password_or_token

# 인증 방식 (application_password 또는 bearer)
WORDPRESS_AUTH_METHOD=application_password
```

## 실행 방법

### GUI 모드 (기본)

```bash
python wodpress-bot.py
```

또는

```bash
python3 wodpress-bot.py
```

## 사용 방법

### 1. 프로그램 실행
터미널에서 위 명령어를 실행하면 GUI 창이 열립니다.

### 2. 설정 확인
- **환경 변수 설정 시**: API 키와 토큰 필드는 비어있어도 됩니다 (자동으로 환경 변수에서 읽음)
- **환경 변수 미설정 시**: GUI에 직접 입력해야 합니다

### 3. 포스트 생성
1. **Post Title**: 블로그 포스트 제목 입력
2. **Blog Topic**: 블로그 주제 입력
3. **Target Keyword**: 타겟 키워드 입력
4. **Word Count (min)**: 최소 단어 수 입력 (기본값: 1500)
5. **첨부 파일** (선택): 참고 자료로 사용할 파일 추가
   - 지원 형식: `.txt`, `.md`, `.docx`, `.xlsx`, `.xls`, `.csv`, `.png`, `.jpg`, `.jpeg`
6. **Build Prompt** 버튼: 프롬프트 자동 생성 (선택)
7. **Generate & Create Draft** 버튼: 포스트 생성 및 WordPress에 업로드

### 4. WordPress 포스트 가져오기
- **WordPress 포스트 가져오기** 버튼: 기존 포스트를 가져와서 편집 가능

## 주요 기능

- ✅ **AI 기반 콘텐츠 생성**: Google Gemini 2.0 Flash 모델 사용
- ✅ **SEO 최적화**: Rank Math SEO 100점 목표로 최적화
- ✅ **다양한 파일 형식 지원**: 텍스트, Word, Excel, CSV, 이미지 OCR
- ✅ **자동 메타데이터 생성**: SEO 제목, 메타 설명, 포커스 키워드 자동 추출
- ✅ **WordPress 자동 업로드**: REST API를 통한 드래프트 자동 생성

## 환경 변수 vs GUI 입력

### 환경 변수 사용 (권장)
- `.env` 파일에 민감 정보 저장
- GUI에 표시되지 않아 보안 강화
- Git에 커밋되지 않음

### GUI 직접 입력
- 환경 변수가 없을 때 사용
- 설정 파일에 저장됨 (민감 정보는 저장되지 않음)

## 문제 해결

### 패키지 설치 오류
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### 환경 변수 로드 오류
- `.env` 파일이 프로젝트 루트에 있는지 확인
- 파일 이름이 정확히 `.env`인지 확인 (확장자 없음)

### WordPress 연결 오류
- WordPress URL이 올바른지 확인
- Application Password 또는 Bearer Token이 유효한지 확인
- 방화벽 설정 확인

## 파일 구조

```
wordpress_LLM/
├── wodpress-bot.py          # 메인 프로그램
├── requirements.txt         # 필수 패키지 목록
├── .env                     # 환경 변수 (Git에 커밋되지 않음)
├── wodpress_settings.json   # 설정 파일 (비민감 정보만 저장)
├── .gitignore              # Git 제외 파일 목록
└── README.md               # 이 파일
```

## 보안 주의사항

⚠️ **중요**: 
- `.env` 파일은 절대 Git에 커밋하지 마세요
- API 키와 토큰은 외부에 공유하지 마세요
- 환경 변수를 사용하는 것이 보안상 안전합니다

