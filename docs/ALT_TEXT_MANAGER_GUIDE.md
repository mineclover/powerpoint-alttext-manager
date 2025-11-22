# PowerPoint 대체 텍스트 관리자 사용 가이드

## 개요

이 애드인은 PowerPoint 슬라이드의 대체 텍스트(Alt Text)를 효율적으로 관리할 수 있도록 도와주는 도구입니다.

## 주요 기능

### 1. 슬라이드 요소 분석
- 현재 슬라이드의 모든 요소(도형, 이미지, 텍스트 상자 등)를 목록으로 표시
- 각 요소의 이름, 타입, 인덱스 정보 제공
- 대체 텍스트 설정 여부 확인

### 2. 선택한 요소 편집
- PowerPoint에서 요소를 선택하고 대체 텍스트 입력
- 한 번에 여러 요소에 동일한 대체 텍스트 적용 가능

### 3. 규칙 기반 자동 설정
세 가지 규칙으로 대체 텍스트를 자동으로 생성할 수 있습니다:

#### a) 접두사 + 요소 이름
- 설정한 접두사를 모든 요소 이름 앞에 추가
- 예: "슬라이드 요소 - " + "Rectangle 1" = "슬라이드 요소 - Rectangle 1"

#### b) 요소 타입별 템플릿
- 요소의 타입에 따라 자동으로 템플릿 생성
- 예: "도형: Rectangle 1", "이미지: Picture 2"

#### c) 사용자 정의 템플릿
- 변수를 사용하여 자유롭게 템플릿 작성
- 사용 가능한 변수:
  - `{name}`: 요소 이름
  - `{type}`: 요소 타입 (한글)
  - `{index}`: 요소 인덱스 (1부터 시작)
- 예: "{type} {index}: {name}" → "도형 1: Rectangle 1"

### 4. 선택적 적용
- **모든 요소에 규칙 적용**: 슬라이드의 모든 요소에 규칙 적용 (기존 대체 텍스트도 덮어쓰기)
- **대체 텍스트 없는 요소만 적용**: 대체 텍스트가 없는 요소에만 규칙 적용

## 설치 방법

### 1. 자동 로드 (권장)

manifest 파일이 이미 `C:\Users\kbrainc\AppData\Roaming\Microsoft\AddIns` 폴더에 복사되어 있습니다.

1. **PowerPoint 완전히 종료**
2. **PowerPoint 다시 시작**
3. 새 프레젠테이션을 만들거나 기존 파일 열기
4. **홈** 탭에서 **대체 텍스트 관리자** 버튼 확인

### 2. 수동 업로드

PowerPoint에서 직접 애드인을 업로드할 수도 있습니다:

1. PowerPoint 실행
2. **삽입** 탭 > **추가 기능** > **내 추가 기능**
3. **내 추가 기능 관리** > **내 추가 기능 업로드**
4. `C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world\manifest-localhost.xml` 파일 선택

## 사용 방법

### 기본 사용 흐름

1. **슬라이드 분석**
   - PowerPoint에서 작업할 슬라이드로 이동
   - 애드인 작업 창에서 "슬라이드 요소 분석" 버튼 클릭
   - 모든 요소 목록이 표시됩니다

2. **개별 요소 편집**
   - PowerPoint에서 요소를 선택
   - 작업 창의 "대체 텍스트" 입력란에 텍스트 입력
   - "선택한 요소에 적용" 버튼 클릭

3. **일괄 적용**
   - "규칙 기반 자동 설정" 섹션에서 원하는 규칙 선택
   - 필요한 경우 접두사 또는 템플릿 수정
   - 적용 버튼 클릭:
     - "모든 요소에 규칙 적용": 모든 요소에 적용
     - "대체 텍스트 없는 요소만 적용": 빈 요소만 적용

### 실용 예시

#### 예시 1: 교육 자료용 슬라이드
```
규칙: 사용자 정의 템플릿
템플릿: "슬라이드 요소 - {type}: {name}"
결과: "슬라이드 요소 - 도형: Rectangle 1"
```

#### 예시 2: 접근성 강화
```
규칙: 타입별 템플릿
결과: "이미지: Logo", "도형: Background Shape"
```

#### 예시 3: 순차적 번호 매기기
```
규칙: 사용자 정의 템플릿
템플릿: "{index}번째 {type}"
결과: "1번째 도형", "2번째 이미지"
```

## 기술적 제한사항

현재 PowerPoint JavaScript API의 제한으로 인해:

- 대체 텍스트가 요소의 이름 필드에 `[ALT: 텍스트]` 형식으로 저장됩니다
- PowerPoint의 기본 대체 텍스트 기능과는 별도로 작동합니다
- 추후 API가 업데이트되면 실제 대체 텍스트 속성을 직접 수정하도록 개선 예정

## 문제 해결

### 애드인이 로드되지 않는 경우

1. **서버 확인**
   - 브라우저에서 `https://localhost:3000/taskpane.html` 접속
   - 페이지가 표시되는지 확인 (인증서 경고는 무시하고 진행)

2. **PowerPoint 재시작**
   - PowerPoint를 완전히 종료하고 다시 시작
   - 작업 관리자에서 PowerPoint 프로세스가 모두 종료되었는지 확인

3. **Manifest 파일 확인**
   - `C:\Users\kbrainc\AppData\Roaming\Microsoft\AddIns\manifest-localhost.xml` 파일 존재 확인

4. **인증서 신뢰**
   - `C:\Users\kbrainc\.office-addin-dev-certs\ca.crt` 파일을 더블 클릭
   - "인증서 설치" > "로컬 컴퓨터" > "신뢰할 수 있는 루트 인증 기관"에 설치

### 서버 재시작

서버가 중지된 경우:
```bash
cd "C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world"
http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
```

## 개발 정보

### 프로젝트 구조
```
Samples/hello-world/powerpoint-hello-world/
├── manifest-localhost.xml  # 애드인 매니페스트
├── taskpane.html          # 메인 UI 및 로직
├── localhost.crt          # SSL 인증서
├── localhost.key          # SSL 키
└── assets/                # 아이콘 이미지
```

### 서버 정보
- URL: https://localhost:3000
- 프로토콜: HTTPS (SSL)
- 포트: 3000

### API 참조
- [PowerPoint JavaScript API](https://learn.microsoft.com/javascript/api/powerpoint)
- [Office Add-ins 문서](https://learn.microsoft.com/office/dev/add-ins/)

## 향후 개선 계획

1. 실제 Alt Text 속성 지원 (API 지원 시)
2. 모든 슬라이드 일괄 처리
3. 대체 텍스트 내보내기/가져오기 (CSV, JSON)
4. 대체 텍스트 유효성 검사
5. 다국어 템플릿 지원
6. 이미지 자동 인식 및 설명 생성 (AI 기반)

## 피드백 및 문의

기능 제안이나 버그 리포트는 개발팀에 문의해주세요.
