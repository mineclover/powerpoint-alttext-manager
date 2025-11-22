# Office.js 애드인 사이드로드 가이드

## Office.js 애드인 vs VSTO 애드인

현재 만든 애드인은 **Office.js (웹 애드인)** 방식입니다:
- ✅ 크로스 플랫폼 (Windows, Mac, Web)
- ✅ 최신 Office JavaScript API 사용
- ✅ 웹 기술 (HTML, CSS, JavaScript)
- ❌ PPAM/PPA 파일 형식이 아님 (XML Manifest 사용)

**VSTO 애드인 (PPAM/PPA)**은 다른 방식입니다:
- Windows 전용
- .NET 기반
- COM 추가 기능

## 데스크톱 PowerPoint에서 사이드로드하기

### 방법 1: 개발 도구를 사용한 사이드로드 (권장)

#### 1단계: 개발 탭 활성화

1. PowerPoint 실행
2. **파일** > **옵션** 클릭
3. 왼쪽에서 **리본 사용자 지정** 선택
4. 오른쪽에서 **개발 도구** 체크박스 선택
5. **확인** 클릭

#### 2단계: 애드인 추가

1. PowerPoint에서 **개발 도구** 탭 클릭
2. **추가 기능** 클릭
3. **내 추가 기능** 탭 선택
4. 오른쪽 상단 **...** (더보기) 메뉴 클릭
5. **내 추가 기능 업로드** 선택
6. 찾아보기 클릭하고 manifest 파일 선택:
   ```
   C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world\manifest-localhost.xml
   ```
7. **업로드** 클릭

### 방법 2: 공유 폴더 카탈로그 사용

#### 1단계: 공유 폴더 만들기

1. 폴더 생성: `C:\AddInManifests`
2. 폴더 우클릭 > **속성** > **공유** 탭
3. **공유** 버튼 클릭
4. 사용자 추가 (자신의 계정)
5. 네트워크 경로 확인 (예: `\\DESKTOP-XXX\AddInManifests`)

#### 2단계: 신뢰할 수 있는 카탈로그로 등록

**옵션 A: 파일 탐색기 방법**
1. 탐색기 주소창에 입력: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
2. 파일 생성: `TrustedCatalogSettings.txt`
3. 내용:
   ```
   \\DESKTOP-XXX\AddInManifests
   ```

**옵션 B: 레지스트리 방법**
1. `Win + R` > `regedit` 실행
2. 경로 이동:
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs
   ```
3. 새 키 생성 (이름: GUID 형식, 예: `{12345678-1234-1234-1234-123456789012}`)
4. 해당 키에 문자열 값 `Url` 추가
5. 값 데이터: `\\DESKTOP-XXX\AddInManifests` (네트워크 경로)

#### 3단계: Manifest 복사

```bash
copy "C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world\manifest-localhost.xml" "C:\AddInManifests\"
```

#### 4단계: PowerPoint 재시작

PowerPoint를 완전히 종료하고 다시 시작하면 애드인이 자동으로 로드됩니다.

### 방법 3: AppSource 배포 (프로덕션용)

실제 배포 시에는 Microsoft AppSource에 게시할 수 있습니다.

## 문제 해결

### "이 앱을 신뢰할 수 없습니다" 오류

1. 브라우저에서 `https://localhost:3000` 접속
2. 인증서 경고 무시하고 계속 진행
3. PowerPoint 재시작

### 애드인이 표시되지 않음

1. PowerPoint 완전 종료 확인 (작업 관리자에서 POWERPNT.EXE 종료)
2. 서버가 실행 중인지 확인:
   ```bash
   curl -k https://localhost:3000/taskpane.html
   ```
3. Manifest 파일의 URL이 올바른지 확인

### 개발 탭이 없는 경우

Microsoft 365 구독이 필요할 수 있습니다. Office 2019 이상에서만 지원됩니다.

## Office Web Apps에서 테스트

가장 간단한 방법:

1. https://office.live.com 접속
2. PowerPoint Online 실행
3. **삽입** > **추가 기능**
4. **내 추가 기능** > **내 추가 기능 업로드**
5. Manifest 파일 업로드

**주의:** 웹에서는 localhost 서버에 접근할 수 없으므로, GitHub Pages나 다른 공개 호스팅이 필요합니다.

## 현재 서버 상태

- 서버: https://localhost:3000
- Manifest: `C:\woo-work\workflow\Office-Add-in-samples\Samples\hello-world\powerpoint-hello-world\manifest-localhost.xml`

## VSTO 애드인으로 전환이 필요한가요?

Office.js 애드인이 더 현대적이고 권장되는 방식입니다. 하지만 VSTO (PPAM/PPA)가 필요하다면:
- Visual Studio 필요
- .NET Framework 기반
- Windows 전용

말씀해주시면 VSTO 방식으로도 만들어드릴 수 있습니다.
