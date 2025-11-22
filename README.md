# PowerPoint 대체 텍스트 관리자

PowerPoint 슬라이드의 대체 텍스트(Alt Text)를 효율적으로 관리하고 편집할 수 있는 애드인 프로젝트입니다.

## 📋 목차

- [개요](#개요)
- [프로젝트 구조](#프로젝트-구조)
- [두 가지 구현 방식](#두-가지-구현-방식)
- [빠른 시작](#빠른-시작)
- [주요 기능](#주요-기능)
- [설치 가이드](#설치-가이드)
- [개발 환경](#개발-환경)
- [라이선스](#라이선스)

## 개요

이 프로젝트는 PowerPoint 프레젠테이션의 접근성을 향상시키기 위한 도구입니다. 슬라이드의 모든 요소(도형, 이미지, 텍스트 상자 등)에 대체 텍스트를 쉽게 추가하고 관리할 수 있습니다.

### 왜 대체 텍스트가 중요한가?

- **접근성**: 시각 장애인을 위한 스크린 리더 지원
- **검색 최적화**: 콘텐츠 검색 및 인덱싱 개선
- **규정 준수**: 접근성 표준 (WCAG, ADA) 충족
- **문서 품질**: 전문적인 프레젠테이션 제작

## 프로젝트 구조

```
.
├── README.md                           # 이 파일
├── Samples/
│   ├── hello-world/
│   │   └── powerpoint-hello-world/    # Office.js 기반 애드인
│   │       ├── taskpane.html          # 대체 텍스트 관리 UI
│   │       ├── manifest-localhost.xml # 로컬 개발용 매니페스트
│   │       └── assets/                # 아이콘 리소스
│   └── dynamic-dpi/
│       └── VSTO PowerPointAddIn/      # VSTO 참고 샘플
├── Templates/
│   └── PowerPoint.MVCAddInTemplate/   # MVC 템플릿
├── docs/
│   ├── QUICK_INSTALL.md               # Office.js 빠른 설치
│   ├── SIDELOAD_GUIDE.md              # Office.js 사이드로드 가이드
│   ├── CREATE_VSTO_PROJECT.md         # VSTO 프로젝트 생성 가이드
│   └── ALT_TEXT_MANAGER_GUIDE.md      # 사용자 가이드
└── .gitignore
```

## 두 가지 구현 방식

### 1. Office.js 웹 애드인 (구현 완료)

**위치**: `Samples/hello-world/powerpoint-hello-world/`

**장점**:
- ✅ 크로스 플랫폼 (Windows, Mac, 웹)
- ✅ 최신 Office JavaScript API 사용
- ✅ 웹 기술 (HTML, CSS, JavaScript)
- ✅ 빠른 개발 및 배포
- ✅ 즉시 사용 가능 (이미 구현됨)

**단점**:
- ❌ API 제한으로 실제 Alt Text 속성 직접 수정 불가
- ❌ 이름 필드에 `[ALT: 텍스트]` 형식으로 저장

**사용 방법**: [QUICK_INSTALL.md](QUICK_INSTALL.md) 참조

### 2. VSTO 애드인 (가이드 제공)

**가이드**: [CREATE_VSTO_PROJECT.md](CREATE_VSTO_PROJECT.md)

**장점**:
- ✅ 실제 Alt Text 속성 직접 수정 가능
- ✅ PPAM 파일로 배포 가능
- ✅ 더 강력한 PowerPoint API 접근
- ✅ 오프라인 완전 지원

**단점**:
- ❌ Windows 전용
- ❌ Visual Studio 및 .NET 개발 환경 필요
- ❌ 사용자가 직접 프로젝트 생성 필요

**시작하기**: Visual Studio 2022에서 새 프로젝트 생성 후 가이드 참조

## 빠른 시작

### Office.js 애드인 (권장 - 바로 사용 가능)

#### 전제 조건
- Node.js 및 npm 설치
- PowerPoint (Microsoft 365 또는 2019 이상)

#### 1단계: 서버 시작

```bash
cd Samples/hello-world/powerpoint-hello-world
http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
```

#### 2단계: PowerPoint에서 애드인 로드

1. PowerPoint 실행
2. **개발 도구** 탭 활성화 (파일 > 옵션 > 리본 사용자 지정)
3. **개발 도구** > **추가 기능** > **내 추가 기능 업로드**
4. `manifest-localhost.xml` 파일 선택

#### 3단계: 사용

1. 홈 탭에서 **대체 텍스트 관리자** 버튼 클릭
2. 작업 창에서 **슬라이드 요소 분석** 클릭
3. 대체 텍스트 추가 및 관리

자세한 내용: [QUICK_INSTALL.md](QUICK_INSTALL.md)

### VSTO 애드인 (Visual Studio 필요)

Visual Studio 2022에서 새 프로젝트를 생성해야 합니다.

자세한 가이드: [CREATE_VSTO_PROJECT.md](CREATE_VSTO_PROJECT.md)

## 주요 기능

### 📊 슬라이드 요소 분석
- 현재 슬라이드의 모든 요소 목록 표시
- 각 요소의 이름, 타입, 인덱스 정보
- 대체 텍스트 설정 여부 확인

### ✏️ 선택한 요소 편집
- PowerPoint에서 요소 선택 후 대체 텍스트 입력
- 여러 요소에 동시 적용 가능

### ⚙️ 규칙 기반 자동 설정

#### 1. 접두사 + 요소 이름
```
규칙: "슬라이드 요소 - "
결과: "슬라이드 요소 - Rectangle 1"
```

#### 2. 요소 타입별 템플릿
```
결과: "도형: Rectangle 1", "이미지: Picture 2"
```

#### 3. 사용자 정의 템플릿
```
템플릿: "{type} {index}: {name}"
결과: "도형 1: Rectangle 1"
```

변수:
- `{name}`: 요소 이름
- `{type}`: 요소 타입 (한글)
- `{index}`: 요소 번호 (1부터 시작)

### 🎯 선택적 적용
- **모든 요소에 적용**: 기존 대체 텍스트도 덮어쓰기
- **빈 요소만 적용**: 대체 텍스트가 없는 요소만

## 설치 가이드

### Office.js 애드인

다양한 설치 방법을 제공합니다:

1. **개발자 도구 사용** (권장): [QUICK_INSTALL.md](QUICK_INSTALL.md)
2. **공유 폴더 카탈로그**: [SIDELOAD_GUIDE.md](SIDELOAD_GUIDE.md)
3. **Office Online**: [SIDELOAD_GUIDE.md](SIDELOAD_GUIDE.md)

### VSTO 애드인

Visual Studio에서 프로젝트 생성: [CREATE_VSTO_PROJECT.md](CREATE_VSTO_PROJECT.md)

## 개발 환경

### Office.js 애드인

**필수 사항**:
- Node.js 16 이상
- npm 또는 yarn
- PowerPoint (Microsoft 365, 2019, 2021)

**개발 도구**:
- http-server (전역 설치)
- office-addin-dev-certs (SSL 인증서)

**설치**:
```bash
npm install --global http-server office-addin-dev-certs
npx office-addin-dev-certs install
```

### VSTO 애드인

**필수 사항**:
- Visual Studio 2022
- Office/SharePoint 개발 워크로드
- .NET Framework 4.7.2 이상
- PowerPoint (데스크톱 버전)

## 사용 예시

### 교육 자료용 슬라이드
```
규칙: 사용자 정의 템플릿
템플릿: "슬라이드 요소 - {type}: {name}"
결과: "슬라이드 요소 - 도형: Rectangle 1"
```

### 접근성 강화
```
규칙: 타입별 템플릿
결과: "이미지: Company Logo", "도형: Background"
```

### 순차적 번호 매기기
```
규칙: 사용자 정의 템플릿
템플릿: "{index}번째 {type}"
결과: "1번째 도형", "2번째 이미지"
```

## 문제 해결

### Office.js 애드인

**서버가 실행되지 않는 경우**:
```bash
cd Samples/hello-world/powerpoint-hello-world
http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
```

**인증서 오류**:
1. 브라우저에서 `https://localhost:3000` 접속
2. 인증서 경고 무시하고 계속 진행
3. PowerPoint 재시작

**애드인이 표시되지 않음**:
1. PowerPoint 완전 종료 (작업 관리자에서 POWERPNT.EXE 확인)
2. 서버 실행 확인
3. Manifest 파일 경로 확인

### VSTO 애드인

**Office PIA 오류**:
```
Install-Package Microsoft.Office.Interop.PowerPoint
```

**디버깅 문제**: [CREATE_VSTO_PROJECT.md](CREATE_VSTO_PROJECT.md) 참조

## 기술 스택

### Office.js 애드인
- HTML5 / CSS3
- Vanilla JavaScript
- Office.js API
- HTTPS (localhost:3000)

### VSTO 애드인
- C# / .NET Framework
- Windows Forms
- PowerPoint Interop API
- Visual Studio Tools for Office (VSTO)

## 향후 개선 계획

- [ ] 모든 슬라이드 일괄 처리
- [ ] 대체 텍스트 내보내기/가져오기 (CSV, JSON)
- [ ] 대체 텍스트 유효성 검사
- [ ] 다국어 템플릿 지원
- [ ] AI 기반 이미지 자동 설명 생성
- [ ] Office.js API 업데이트 시 실제 Alt Text 속성 지원

## 기여하기

이슈 및 풀 리퀘스트를 환영합니다!

## 참고 자료

- [Office Add-ins 문서](https://learn.microsoft.com/office/dev/add-ins/)
- [PowerPoint JavaScript API](https://learn.microsoft.com/javascript/api/powerpoint)
- [VSTO 개발 가이드](https://learn.microsoft.com/visualstudio/vsto/)
- [PowerPoint Object Model](https://learn.microsoft.com/office/vba/api/overview/powerpoint/object-model)
- [접근성 표준 (WCAG)](https://www.w3.org/WAI/standards-guidelines/wcag/)

## 라이선스

MIT License

## 저자

개발자: Claude & User
버전: 1.0.0
날짜: 2025-11-22
