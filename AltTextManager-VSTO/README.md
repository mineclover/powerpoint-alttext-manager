# PowerPoint 대체 텍스트 관리자 - VSTO 애드인

## ✅ 프로젝트 완성!

**제대로 된 C# Windows Forms 기반 VSTO 애드인**입니다.

🎉 **현재 상태**:
- ✅ 빌드 완료 (Visual Studio 2025)
- ✅ PowerPoint에 설치 및 활성화 완료
- ✅ 리본 UI 추가 (대체 텍스트 관리자 버튼)

📖 **상세 현황**: [PROJECT_STATUS.md](PROJECT_STATUS.md)

## 📦 포함된 파일

- `ThisAddIn.cs` - VSTO 애드인 진입점
- `AltTextManagerForm.cs` - Windows Forms UI (실제 창)
- `AltTextManagerForm.Designer.cs` - UI 디자인 코드
- `Properties/AssemblyInfo.cs` - 어셈블리 정보
- `AltTextManager.csproj` - 프로젝트 파일

## 🎯 주요 기능

### Windows Forms UI
- ✅ ListView - 슬라이드 요소 목록 (실시간 표시)
- ✅ TextBox - 대체 텍스트 입력
- ✅ Button - 적용, 자동생성, 삭제 등
- ✅ ComboBox - 템플릿 규칙 선택
- ✅ CheckBox - 빈 요소만 적용 옵션

### PowerPoint Interop
- ✅ 현재 슬라이드 요소 탐색
- ✅ 대체 텍스트 읽기/쓰기
- ✅ 도형 타입 인식 (이미지, 도형, 표 등)

### 자동 생성 기능
- 접두사 + 요소 이름
- 타입별 템플릿
- 사용자 정의 템플릿

## 🔨 빌드 방법

**상세한 빌드 가이드**: [BUILD_GUIDE.md](BUILD_GUIDE.md)

### 빠른 시작 (Visual Studio 설치 시)

```powershell
# PowerShell에서 실행
cd P:\Project\pptx-addon\AltTextManager-VSTO
.\build.ps1
```

### 수동 빌드

#### Visual Studio에서
1. `AltTextManager.csproj` 열기
2. **빌드 → 솔루션 빌드** (Ctrl+Shift+B)
3. `bin\Debug\AltTextManager.vsto` 생성 확인

#### 명령줄에서 (Developer Command Prompt)
```cmd
msbuild AltTextManager.csproj /p:Configuration=Debug
```

## 📦 설치 방법

### 방법 1: 자동 설치 (ClickOnce) - 권장

빌드 후 생성된 설치 파일 실행:
```
bin\Debug\AltTextManager.vsto
```

### 방법 2: 수동 설치 (개발자 모드)

PowerShell 스크립트 사용 (관리자 권한):
```powershell
.\install-manual.ps1
```

### 방법 3: 제거

```powershell
.\uninstall.ps1
```

## 🚀 사용 방법

1. PowerPoint 실행
2. **개발 도구 탭** → **COM 추가 기능**
3. **AltTextManager** 체크
4. 메뉴에서 대체 텍스트 관리자 열기

또는 VBA에서:

```vba
Sub OpenAltTextManager()
    Globals.ThisAddIn.ShowAltTextManager()
End Sub
```

## 🔧 필수 요구사항

- Windows 10/11
- PowerPoint 2013 이상
- .NET Framework 4.7.2
- VSTO Runtime

## 📝 기술 스택

- **언어**: C# (.NET Framework 4.7.2)
- **UI**: Windows Forms
- **Office Interop**: PowerPoint Primary Interop Assembly
- **애드인 모델**: VSTO (Visual Studio Tools for Office)

## 🎨 UI 스크린샷

```
┌────────────────────────────────────────────────────────┐
│  PowerPoint 대체 텍스트 관리자                    [X]  │
├────────────────────────────────────────────────────────┤
│                                                         │
│  현재 슬라이드 요소 목록:            [새로고침]       │
│  ┌──────────────────────────────────────────────────┐ │
│  │ # │ 이름      │ 타입    │ 대체 텍스트          │ │
│  │ 1 │ Rect 1    │ 도형    │ (없음)               │ │
│  │ 2 │ Picture 2 │ 이미지  │ 회사 로고            │ │
│  └──────────────────────────────────────────────────┘ │
│                                                         │
│  대체 텍스트 입력:                                     │
│  ┌──────────────────────────────────────────────────┐ │
│  │                                                   │ │
│  │                                                   │ │
│  └──────────────────────────────────────────────────┘ │
│                                                         │
│  [선택 항목에 적용]   [모두 적용]                     │
│                                                         │
│  자동 생성:                                            │
│  규칙: [접두사 + 요소 이름 ▼]  [자동 생성]           │
│  ☐ 빈 요소만 적용                                     │
│  예: 슬라이드 요소 - Rectangle 1                      │
│                                                         │
│  [모두 삭제]                              [닫기]      │
│                                                         │
│  상태: 5개의 요소를 찾았습니다                         │
└────────────────────────────────────────────────────────┘
```

## 🆚 VBA/PPAM vs VSTO 비교

| 항목 | VBA/PPAM | VSTO (이 프로젝트) |
|------|----------|-------------------|
| UI | UserForm (제한적) | **Windows Forms (완전한 GUI)** |
| 언어 | VBA | **C# (.NET)** |
| 배포 | 수동 빌드 필요 | **설치 프로그램** |
| 디버깅 | 어려움 | **Visual Studio 지원** |
| 버전 관리 | 불편 | **Git 친화적** |
| 유지보수 | 어려움 | **쉬움** |

## 🐛 문제 해결

### "매니페스트를 읽을 수 없습니다"
- DLL 경로가 정확한지 확인
- Manifest 파일이 있는지 확인 (`AltTextManager.dll.manifest`)

### "VSTO Runtime이 설치되지 않았습니다"
- [VSTO Runtime 다운로드](https://www.microsoft.com/download/details.aspx?id=48217)

### "형식을 로드할 수 없습니다"
- .NET Framework 4.7.2 설치 확인
- PowerPoint 재시작

## 📄 라이선스

MIT License

## 🎉 완성!

이제 **제대로 된 VSTO 애드인**이 완성되었습니다!

- ✅ VBA가 아닌 **C# 코드**
- ✅ UserForm이 아닌 **Windows Forms**
- ✅ 수동 빌드가 아닌 **설치 프로그램**
- ✅ 매크로가 아닌 **실제 애드인**
