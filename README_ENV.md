# 환경 변수 설정 가이드

이 프로젝트는 보안을 위해 민감한 정보(API 키, 토큰 등)를 환경 변수로 관리합니다.

## 설정 방법

1. 프로젝트 루트 디렉토리에 `.env` 파일을 생성하세요.

2. 다음 내용을 복사하여 `.env` 파일에 붙여넣고 실제 값으로 수정하세요:

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

3. `.env` 파일은 Git에 커밋되지 않습니다 (`.gitignore`에 포함됨)

## 주의사항

- `.env` 파일은 절대 Git에 커밋하지 마세요
- API 키와 토큰은 외부에 공유하지 마세요
- 환경 변수가 설정되어 있으면 GUI 입력값보다 우선적으로 사용됩니다

