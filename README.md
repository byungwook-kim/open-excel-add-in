# Open Excel Add-in MVP

실험용 Excel taskpane add-in입니다. 현재 범위는 다음 두 가지입니다.

- 현재 workbook의 range를 읽어서 textarea에 TSV 텍스트로 보여주기
- textarea의 TSV 텍스트를 range에 다시 쓰기

추가로 **Experimental ChatGPT Sign-in** 버튼이 있지만, 이 부분은 **작동 보장 기능이 아니라 OAuth 가능성을 보는 스파이크**입니다. 무료 사용 entitlement나 실제 ChatGPT 모델 호출 가능 여부를 증명하지는 않습니다.

## 프로젝트 구조

```text
assets/
  icon-16.png
  icon-32.png
  icon-80.png
src/taskpane/
  pages/
    taskpane.html
    auth-start.html
    auth-callback.html
  scripts/
    taskpane.js
    auth-start.js
    auth-callback.js
    rangeText.js
    rangeText.test.js
  styles/
    taskpane.css
manifest.xml
webpack.config.js
package.json
```

- `pages/`: taskpane 및 OAuth dialog용 HTML
- `scripts/`: Office.js 진입점과 TSV 파서 로직
- `styles/`: taskpane 스타일

## 시작하기

```powershell
npm install
```

## 로컬 개발

### 1) dev server 실행

```powershell
npm run dev-server
```

이 명령은 `https://localhost:3000` 에 webpack dev server를 띄웁니다. 처음 실행 시 Office add-in용 localhost 인증서를 신뢰하도록 설정할 수 있습니다.

핵심 페이지:

- `https://localhost:3000/taskpane.html`
- `https://localhost:3000/auth-start.html`
- `https://localhost:3000/auth-callback.html`

### 2) add-in sideload

```powershell
npm run start
```

이 명령은 `manifest.xml`을 기준으로 Excel desktop sideload를 시도합니다.

중지:

```powershell
npm run stop
```

## Excel에 설치하기 (Sideload)

로컬에서 개발 중인 add-in을 Excel에 설치하는 방법은 두 가지가 있습니다.

### 방법 1: npm run start (권장)

가장 간단한 방법입니다.

```powershell
npm run start
```

이 명령은 다음을 수행합니다:

1. `npm run dev-server`를 백그라운드에서 실행 (실행 중이 아니면)
2. `manifest.xml`을 읽어서 Excel desktop을 찾아 sideload 시도
3. 성공하면 Excel에 "Open Excel Add-in"이 ribbon에 표시됨

실행 중지:

```powershell
npm run stop
```

### 방법 2: 수동 sideload

`npm run start`가 안 되거나 더 구체적인 제어 필요 시 수동으로 합니다.

#### 2-1) dev server 수동 실행

```powershell
npm run dev-server
```

#### 2-2) manifest.xml 파일 경로 확인

빌드된 manifest 사용 시:

```powershell
# dist/manifest.xml의 절대 경로를 기록합니다
# 예: C:\Users\User\Desktop\open-excel-add-in\dist\manifest.xml
```

#### 2-3) Excel에서 수동 sideload

1. Excel을 엽니다.
2. 파일 → 옵션 → 개발자 탭을 활성화합니다 (없다면 "개발자" 체크).
3. 개발자 탭 → "Add-in 관리"를 누릅니다.
4. "찾아보기"를 클릭하고 위에서 기록한 `manifest.xml` 파일을 선택합니다.
5. 확인을 누르면 ribbon에 add-in이 나타납니다.

### 설치 후 실행

1. Excel ribbon에 "Open Excel Add-in" 버튼이 보입니다.
2. 버튼을 누르면 오른쪽에 taskpane이 열립니다.
3. 사용 가능 환경:
   - Excel 2016 이상 (Windows/Mac)
   - Excel for Microsoft 365
   - Excel on the web

### 문제 발생 시 체크리스트

- [ ] dev server가 `https://localhost:3000`에서 실행 중인가?
- [ ] manifest.xml의 SourceLocation이 HTTPS로 시작하는가?
- [ ] 아이콘 이미지가 `https://localhost:3000/assets/icon-32.png`에서 접근 가능한가?
- [ ] Excel이 "개발자" 탭을 활성화했는가?

## 테스트

### 단위 테스트

```powershell
npm test
```

현재 테스트는 `src/taskpane/scripts/rangeText.test.js`에서 TSV 파서/포매터를 검증합니다.

검증 항목:

- 단일 값 파싱
- TSV 행렬 파싱
- jagged row를 직사각형으로 정규화
- 값 포매팅
- sheet prefix가 포함된 주소 파싱
- 좌상단 anchor 추출

### 빌드 검증

```powershell
npm run build
```

성공하면 `dist/`에 다음이 생성됩니다.

- `taskpane.html`, `taskpane.js`, `taskpane.css`
- `auth-start.html`, `auth-start.js`
- `auth-callback.html`, `auth-callback.js`
- `manifest.xml`
- `assets/*.png`

### manifest 검증

```powershell
npm run validate
```

이 명령은 Office add-in manifest 스키마와 URL/아이콘 구성을 검사합니다.

## 수동 테스트

### Read / Write 수동 확인

1. Excel에서 workbook을 엽니다.
2. add-in taskpane를 엽니다.
3. address에 `A1` 또는 `A1:B2` 같은 range를 입력합니다.
4. `Read`를 누르면 현재 값이 textarea에 TSV 형식으로 들어와야 합니다.
5. textarea를 수정한 뒤 `Write`를 누르면 입력한 address를 좌상단 anchor로 사용해 worksheet에 값을 써야 합니다.

예시:

```text
name	age
kim	25
lee	30
```

이를 `A1`에 Write하면 `A1:B3`에 표처럼 써집니다.

### OAuth 스파이크 수동 확인

1. taskpane에서 `Experimental ChatGPT Sign-in`을 누릅니다.
2. Office Dialog API로 OAuth dialog가 열리는지 확인합니다.
3. 성공 시 taskpane 상태 영역과 auth status에 "실험적 로그인 저장됨" 메시지가 보여야 합니다.
4. `Reset token`을 누르면 localStorage에 저장된 실험용 토큰 상태가 삭제되어야 합니다.

## 디버깅 가이드

### 브라우저/호스팅 확인

dev server가 정상인지 가장 먼저 확인합니다.

```powershell
npm run dev-server
```

그리고 브라우저에서 아래 주소가 열리는지 봅니다.

- `https://localhost:3000/taskpane.html`

페이지가 열리지 않으면 대부분 다음 중 하나입니다.

- localhost 인증서 신뢰 문제
- dev server 미실행
- 포트 3000 충돌

### build 문제 디버깅

```powershell
npm run build
```

확인 포인트:

- `webpack.config.js`의 `entry` 경로가 실제 `scripts/` 구조와 일치하는지
- CopyWebpackPlugin 경로가 `pages/`, `styles/`를 정확히 가리키는지
- `dist/manifest.xml`과 `dist/taskpane.html`이 생성되는지

### manifest 문제 디버깅

```powershell
npm run validate
```

주로 깨지는 지점:

- `SourceLocation`이 HTTPS가 아닌 경우
- 아이콘 URL 누락
- localhost URL 오타

### Read / Write 문제 디버깅

문제가 생기면 먼저 `status` 패널 메시지를 봅니다.

핵심 코드 위치:

- `src/taskpane/scripts/taskpane.js`
  - `readRange()`
  - `writeRange()`
- `src/taskpane/scripts/rangeText.js`

체크 포인트:

- address가 비어 있지 않은지
- `Sheet1!A1:B2` 같은 sheet-qualified address를 잘 넣었는지
- jagged TSV는 자동으로 직사각형으로 맞춰지지만, 셀 내부 탭/줄바꿈은 lossless round-trip이 아님

### OAuth 문제 디버깅

핵심 파일:

- `src/taskpane/scripts/auth-start.js`
- `src/taskpane/scripts/auth-callback.js`

체크 포인트:

- dialog가 실제로 열리는지
- redirect가 `https://localhost:3000/auth-callback.html`로 돌아오는지
- callback에서 `code`, `state`, `verifier`가 모두 맞는지
- token exchange 응답이 200인지

중요:

- 이 OAuth는 **실험적 스파이크**입니다.
- ChatGPT 무료 사용이나 downstream 모델 호출 entitlement를 보장하지 않습니다.
- localStorage 토큰 저장은 prototype 용도일 뿐 production-safe 설계가 아닙니다.

## 현재 검증 상태

이미 확인한 항목:

- `npm install` 성공
- `npm test` 통과
- `npm run build` 성공
- `npm run validate` 성공
- `https://localhost:3000/taskpane.html` HTTP 200 응답 확인

아직 직접 증명하지 않은 항목:

- 실제 Excel host 안에서 end-to-end OAuth 성공
- OAuth 후 무료 사용 entitlement 확보

## 권장 다음 단계

1. Excel에서 sideload 후 read/write를 실제로 확인
2. OAuth dialog가 실제 Office host에서 끝까지 도는지 확인
3. 무료 사용이 목적이면 OAuth 성공과 entitlement를 별도로 검증
4. 제품화 단계로 갈 경우 localStorage 토큰 구조 제거 및 재설계

## 배포 방법

사용자에게 npm을 설치하게 하지 않으려면, 다음 방법 중 하나를 선택합니다.

### 방법 1: Office Store (AppSource) 배포

공식 Microsoft 스토어에 올려서 모든 사용자가 Excel에서 직접 검색해서 설치하게 합니다.

장점:
- 설치가 매우 간단 (Excel → 삽입 → 내 Add-in → 스토어에서 검색)
- Microsoft'strust가짐
- 자동 업데이트

준비물:
- Microsoft Partner Center 계정
- 매니페스트 검증 통과 (`npm run validate`)
- 아이콘/스크린샷 같은 스토어용 에셋
- 개인 정보 취급방침 URL

절차:
1. Partner Center에서 "새 제출" 생성
2. 매니페스트와 에셋 업로드
3. Microsoft 검수 대기 (보통 2-5일)
4. 승인되면 스토어에 노출

참고: 현재 MVP 단계에서는 이 방법을 바로 쓰기엔 인증/Gratuity/entitlement 문제가 해결되지 않은 상태입니다.

### 방법 2: 중앙 집중식 배포 (Microsoft 365용)

기업이나 학교 같은 조직에서 관리자가 조직 구성원에게 일괄적으로 설치하게 합니다.

장점:
- 관리자가 한번에 배포
- 사용자별 설치 불필요
- 조직 내 업데이트 가능

준비물:
- Microsoft 365 관리자 권한
- SharePoint나 Azure에서 add-in 파일 호스팅

절차:
1. 관리자센터 → 배포 → Add-in 추가
2. 매니페스트 또는 파일 URL 업로드
3. 사용자 선택 후 배포

참고: 이 방법은 조직 라이선스가 있는 환경에만 해당합니다.

### 방법 3: 수동 sideload (가장 간단, 현재 MVP에 적합)

**무료로 배포**하려면 GitHub Pages를 사용하면 됩니다. 별도 서버 없이 무료 HTTPS 호스팅이 됩니다.

#### 3-1) GitHub Pages로 무료 배포

**배포된 주소**: `https://byungwook-kim.github.io/open-excel-add-in/`

```powershell
# 1. GitHub에 저장소 만들고 로컬 폴더와 연결 (이미 함)
git init
git add .
git commit -m "Initial commit"

# 2. GitHub Pages 활성화 (저장소 → Settings → Pages → Source: "main branch" → Save)

# 3. npm으로 배포 ( gh-pages 패키지 필요)
npm run deploy
```

이 명령은 `dist/` 폴더 내용을 GitHub Pages에 올립니다.

**주의**: 매니페스트의 URL이 `localhost:3000`으로 되어 있어서, 배포 후 매니페스트의 URL을 실제 배포 URL로 수정해야 합니다.

예를 들어, GitHub Pages URL이 `https://byungwook-kim.github.io/open-excel-add-in/`이라면:

`manifest.xml`에서:
```xml
<IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png" />
```
를
```xml
<IconUrl DefaultValue="https://byungwook-kim.github.io/open-excel-add-in/assets/icon-32.png" />
```
로 모두 변경합니다. (이미 수정됨)

#### 3-2) 사용자에게 배포

1. 수정된 `manifest.xml` 파일을 사용자에게 전달
2. 사용자는 Excel에서 수동 sideload (위 "Excel에 설치하기" 참고)

이 방식의 장점:
- 무료 (GitHub Pages 무료)
- HTTPS 자동 제공
- npm 없이Users도 설치 가능

### 방법 4: Outlook/Teams와 함께 번들 (Advanced)

다른 Microsoft 365 제품과 함께 묶어서 배포합니다.

이 방법은 현재 MVP 범위를 벗어나며, 별도 통합 작업이 필요합니다.

## 배포 시 주의사항

1. **HTTPS 필수**: 매니페스트의 SourceLocation은 반드시 HTTPS여야 합니다.
   - 하지만 인증서를 직접 관리할 필요 없음
   - 무료 호스팅 서비스(GitHub Pages, Netlify, Vercel, Cloudflare Pages)가 SSL/TLS를 자동으로 제공함
   - 자체 서버 시 Let's Encrypt가 무료로 인증서를 발급해 줌

2. **매니페스트 URL**: 배포 시 manifest.xml의 모든 URL을 실제 호스팅 주소로 변경해야 합니다.
   - `localhost:3000` → 실제 배포 URL (예: `https://yourname.github.io/open-excel-add-in/`)
   - 아이콘, SourceLocation, SupportUrl 모두 포함

3. **HTTPS 자동 제공**: 무료 호스팅 서비스가 SSL/TLS를 자동으로 제공함
   - 사용자가 인증서를 직접 신뢰할 필요 없음
   - GitHub Pages, Netlify 같은 서비스는 HTTPS가 기본

3. **매니페스트 ID**: 배포 시 새 ID를 생성해야 합니다 (현재는 테스트용 ID).

4. **OAuth 경고**: 현재 OAuth는 실험적 스파이크입니다.
   - 프로덕션 배포 전엔 실제 entitlement/Gratuity를 검증해야 합니다.

## 현재 권장: 방법 3 (수동 sideload)

MVP 단계에서는 방법 3 (수동 sideload)이 가장 실용적입니다:

1. `npm run build`로 `dist/` 생성
2. GitHub Pages나 Netlify에 무료 호스팅
3. 사용자에게 GitHub/Pages URL과 매니페스트 sideload 방법 안내

이후 Office Store 검수 통과가 목표라면, 인증/entitlement 문제를 해결한 뒤 Partner Center에서 제출합니다.
