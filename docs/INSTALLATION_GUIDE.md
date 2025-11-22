# PowerPoint Add-in 설치 가이드

## 현재 상태

로컬 개발 서버가 성공적으로 실행 중입니다:
- 서버 주소: https://localhost:3000
- 프로젝트: powerpoint-hello-world

## Windows에서 PowerPoint 애드인 설치하기

### 방법 1: 네트워크 공유 폴더를 사용한 사이드로드 (권장)

1. **공유 폴더 만들기**
   - 새 폴더를 만듭니다 (예: `C:\AddinManifests`)
   - 폴더를 우클릭하고 **속성** > **공유** 탭 선택
   - **공유** 버튼을 클릭하고 자신을 추가
   - 네트워크 경로를 메모합니다 (예: `\\YourPC\AddinManifests`)

2. **신뢰할 수 있는 카탈로그에 폴더 추가**
   - **파일 탐색기**를 엽니다
   - 주소 표시줄에 입력: `%APPDATA%\Microsoft\Office\16.0\Wef\`
   - 해당 폴더에서 `TrustedCatalogs` 폴더를 찾습니다 (없으면 생성)

   또는 레지스트리를 사용:
   - `regedit`를 실행합니다
   - 다음 경로로 이동:
     ```
     HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs
     ```
   - 새 키를 만들고 GUID 이름을 지정합니다 (예: `{12345678-1234-1234-1234-123456789012}`)
   - 해당 키에 문자열 값 `Url`을 추가하고 공유 폴더 경로를 입력합니다

3. **Manifest 파일 복사**
   - `manifest-localhost.xml` 파일을 공유 폴더에 복사합니다
   - 파일 위치: `C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world\manifest-localhost.xml`

4. **PowerPoint 재시작**
   - PowerPoint를 완전히 종료하고 다시 시작합니다
   - 새 프레젠테이션을 만들거나 기존 파일을 엽니다

5. **애드인 확인**
   - **홈** 탭 리본에 **Hello world** 버튼이 나타나야 합니다
   - 버튼을 클릭하면 작업 창이 열립니다

### 방법 2: Office Add-in 웹 인터페이스 사용

1. **PowerPoint 열기**
   - PowerPoint를 실행하고 새 프레젠테이션을 만듭니다

2. **애드인 업로드**
   - **삽입** 탭 > **추가 기능** 섹션 > **추가 기능** 클릭
   - **Office 추가 기능** 대화 상자에서 **내 추가 기능** 탭 선택
   - **내 추가 기능 관리** 드롭다운 클릭
   - **내 추가 기능 업로드** 선택
   - `manifest-localhost.xml` 파일을 찾아서 업로드

3. **애드인 사용**
   - 업로드가 완료되면 **홈** 탭에 **Hello world** 버튼이 나타납니다
   - 버튼을 클릭하여 작업 창을 엽니다
   - 슬라이드에 텍스트 상자를 선택하고 **Say Hello** 버튼을 클릭하면 "Hello world!"가 삽입됩니다

## 문제 해결

### 서버가 실행 중인지 확인
브라우저에서 다음 URL을 열어보세요:
```
https://localhost:3000/taskpane.html
```
인증서 경고가 나타나면 "고급" > "안전하지 않음으로 이동"을 클릭하여 계속 진행합니다.

### 서버 재시작이 필요한 경우
1. 현재 실행 중인 서버를 중지합니다
2. 다음 명령을 실행합니다:
```bash
cd "Samples/hello-world/powerpoint-hello-world"
http-server -S -C localhost.crt -K localhost.key --cors . -p 3000
```

### 인증서 오류가 발생하는 경우
인증서가 신뢰되지 않는 경우:
1. 브라우저에서 `https://localhost:3000`을 열고 인증서를 신뢰하도록 설정
2. Windows의 신뢰할 수 있는 루트 인증 기관에 CA 인증서를 추가:
   - `C:\Users\kbrainc\.office-addin-dev-certs\ca.crt` 파일을 더블 클릭
   - "인증서 설치" 클릭
   - "로컬 컴퓨터" 선택
   - "모든 인증서를 다음 저장소에 저장" 선택
   - "신뢰할 수 있는 루트 인증 기관" 선택

## 애드인 기능

이 Hello World 애드인은 다음 기능을 제공합니다:
- PowerPoint 슬라이드에서 선택한 텍스트 영역에 "Hello world!" 텍스트 삽입
- Office JavaScript API 사용 방법 데모

## 다음 단계

애드인이 정상적으로 작동하면:
1. `taskpane.html` 파일을 수정하여 자신만의 기능을 추가할 수 있습니다
2. 서버가 자동으로 변경 사항을 반영합니다 (페이지 새로고침 필요)
3. Manifest 파일을 수정한 경우 PowerPoint를 재시작해야 합니다

## 프로젝트 정보

- 프로젝트 위치: `C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world`
- Manifest 파일: `manifest-localhost.xml`
- 메인 HTML: `taskpane.html`
- 서버 포트: 3000 (HTTPS)
