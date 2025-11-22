# VSTO 대체 텍스트 관리자 프로젝트 만들기

## 1단계: Visual Studio에서 새 프로젝트 생성

### 프로젝트 생성

1. **Visual Studio 2022** 실행

2. **새 프로젝트 만들기** 클릭

3. 검색창에 **"PowerPoint VSTO"** 입력

4. **PowerPoint VSTO 추가 기능** 템플릿 선택
   - 템플릿이 안 보이면: **도구** > **도구 및 기능 가져오기** > **Office/SharePoint 개발** 워크로드 설치 필요

5. 프로젝트 설정:
   - 프로젝트 이름: `AltTextManagerAddIn`
   - 위치: `C:\woo-work\workflow\Office-Add-in-samples\`
   - 솔루션 이름: `AltTextManagerAddIn`

6. **만들기** 클릭

### Office/SharePoint 개발 워크로드 설치 (필요한 경우)

템플릿이 보이지 않으면:

1. Visual Studio Installer 실행
2. Visual Studio 2022 옆의 **수정** 클릭
3. **워크로드** 탭에서 **Office/SharePoint 개발** 체크
4. **수정** 클릭하여 설치

## 2단계: 프로젝트 구조 이해

생성된 프로젝트에는 다음 파일들이 있습니다:

```
AltTextManagerAddIn/
├── ThisAddIn.cs              # 애드인 진입점
├── ThisAddIn.Designer.cs     # 자동 생성 코드
├── Properties/
│   ├── AssemblyInfo.cs
│   └── Settings.settings
└── AltTextManagerAddIn.csproj
```

## 3단계: 코드 작성

### ThisAddIn.cs 수정

```csharp
using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace AltTextManagerAddIn
{
    public partial class ThisAddIn
    {
        private Office.CommandBarButton altTextButton;
        private AltTextTaskPane taskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 리본 메뉴에 버튼 추가
            AddRibbonButton();
        }

        private void AddRibbonButton()
        {
            try
            {
                Office.CommandBar commandBar = this.Application.CommandBars["Standard"];

                altTextButton = (Office.CommandBarButton)commandBar.Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);

                altTextButton.Caption = "대체 텍스트 관리자";
                altTextButton.Tag = "AltTextManager";
                altTextButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(AltTextButton_Click);
            }
            catch (Exception ex)
            {
                MessageBox.Show("버튼 추가 실패: " + ex.Message);
            }
        }

        private void AltTextButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ShowTaskPane();
        }

        private void ShowTaskPane()
        {
            if (taskPane == null)
            {
                taskPane = new AltTextTaskPane();
                var customTaskPane = this.CustomTaskPanes.Add(taskPane, "대체 텍스트 관리자");
                customTaskPane.Width = 350;
            }

            foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
            {
                if (pane.Title == "대체 텍스트 관리자")
                {
                    pane.Visible = !pane.Visible;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO에서 생성한 코드

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
```

## 4단계: 작업 창 UserControl 추가

1. **프로젝트 우클릭** > **추가** > **사용자 정의 컨트롤**
2. 이름: `AltTextTaskPane.cs`
3. **추가** 클릭

### AltTextTaskPane.cs 디자인

디자이너에서 다음 컨트롤 추가:

- Label: "현재 슬라이드 분석"
- Button (btnAnalyze): "슬라이드 요소 분석"
- ListBox (lstShapes): 요소 목록 표시
- Label: "대체 텍스트"
- TextBox (txtAltText): 대체 텍스트 입력 (Multiline)
- Button (btnSetAltText): "선택한 요소에 적용"
- GroupBox: "규칙 기반 자동 설정"
  - ComboBox (cmbRuleType): 규칙 선택
  - TextBox (txtPrefix): 접두사
  - Button (btnApplyAll): "모든 요소에 적용"
  - Button (btnApplyEmpty): "빈 요소만 적용"

### AltTextTaskPane.cs 코드

```csharp
using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AltTextManagerAddIn
{
    public partial class AltTextTaskPane : UserControl
    {
        public AltTextTaskPane()
        {
            InitializeComponent();
            InitializeControls();
        }

        private void InitializeControls()
        {
            cmbRuleType.Items.AddRange(new string[] {
                "접두사 + 요소 이름",
                "요소 타입별 템플릿",
                "사용자 정의 템플릿"
            });
            cmbRuleType.SelectedIndex = 0;
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            AnalyzeCurrentSlide();
        }

        private void AnalyzeCurrentSlide()
        {
            try
            {
                lstShapes.Items.Clear();

                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;

                if (pres.Slides.Count == 0)
                {
                    MessageBox.Show("슬라이드가 없습니다.");
                    return;
                }

                PowerPoint.Slide slide = (PowerPoint.Slide)pres.Slides[app.ActiveWindow.Selection.SlideRange.SlideIndex];

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string altText = shape.AlternativeText ?? "(없음)";
                    string item = $"{shape.Name} ({shape.Type}) - Alt: {altText}";
                    lstShapes.Items.Add(item);
                    lstShapes.Tag = slide; // 슬라이드 참조 저장
                }

                MessageBox.Show($"{slide.Shapes.Count}개의 요소를 찾았습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("분석 오류: " + ex.Message);
            }
        }

        private void btnSetAltText_Click(object sender, EventArgs e)
        {
            SetAltTextForSelected();
        }

        private void SetAltTextForSelected()
        {
            try
            {
                string altText = txtAltText.Text.Trim();
                if (string.IsNullOrEmpty(altText))
                {
                    MessageBox.Show("대체 텍스트를 입력하세요.");
                    return;
                }

                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection selection = app.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("요소를 선택하세요.");
                    return;
                }

                int count = 0;
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    shape.AlternativeText = altText;
                    count++;
                }

                MessageBox.Show($"{count}개 요소에 대체 텍스트를 설정했습니다.");
                AnalyzeCurrentSlide(); // 목록 새로고침
            }
            catch (Exception ex)
            {
                MessageBox.Show("설정 오류: " + ex.Message);
            }
        }

        private void btnApplyAll_Click(object sender, EventArgs e)
        {
            ApplyRules(false);
        }

        private void btnApplyEmpty_Click(object sender, EventArgs e)
        {
            ApplyRules(true);
        }

        private void ApplyRules(bool onlyEmpty)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;
                PowerPoint.Slide slide = (PowerPoint.Slide)pres.Slides[app.ActiveWindow.Selection.SlideRange.SlideIndex];

                int count = 0;
                int index = 1;

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (onlyEmpty && !string.IsNullOrEmpty(shape.AlternativeText))
                    {
                        index++;
                        continue;
                    }

                    string altText = "";
                    string ruleType = cmbRuleType.SelectedItem?.ToString() ?? "";

                    switch (ruleType)
                    {
                        case "접두사 + 요소 이름":
                            altText = txtPrefix.Text + shape.Name;
                            break;

                        case "요소 타입별 템플릿":
                            altText = $"{GetShapeTypeName(shape.Type)}: {shape.Name}";
                            break;

                        case "사용자 정의 템플릿":
                            altText = $"{GetShapeTypeName(shape.Type)} {index}: {shape.Name}";
                            break;
                    }

                    shape.AlternativeText = altText;
                    count++;
                    index++;
                }

                MessageBox.Show($"{count}개 요소에 규칙을 적용했습니다.");
                AnalyzeCurrentSlide(); // 목록 새로고침
            }
            catch (Exception ex)
            {
                MessageBox.Show("규칙 적용 오류: " + ex.Message);
            }
        }

        private string GetShapeTypeName(Microsoft.Office.Core.MsoShapeType type)
        {
            switch (type)
            {
                case Microsoft.Office.Core.MsoShapeType.msoAutoShape:
                    return "도형";
                case Microsoft.Office.Core.MsoShapeType.msoPicture:
                    return "그림";
                case Microsoft.Office.Core.MsoShapeType.msoTextBox:
                    return "텍스트 상자";
                case Microsoft.Office.Core.MsoShapeType.msoTable:
                    return "표";
                case Microsoft.Office.Core.MsoShapeType.msoChart:
                    return "차트";
                case Microsoft.Office.Core.MsoShapeType.msoGroup:
                    return "그룹";
                case Microsoft.Office.Core.MsoShapeType.msoLine:
                    return "선";
                default:
                    return type.ToString();
            }
        }
    }
}
```

## 5단계: 빌드 및 디버그

1. **빌드** > **솔루션 빌드** (Ctrl+Shift+B)

2. **디버그** > **디버깅 시작** (F5)
   - PowerPoint가 자동으로 실행됩니다
   - 애드인이 로드됩니다

3. PowerPoint에서 툴바의 "대체 텍스트 관리자" 버튼 클릭

## 6단계: 배포 (PPAM 파일 생성)

### 게시 설정

1. **프로젝트 우클릭** > **게시**

2. 게시 위치 설정:
   ```
   C:\woo-work\workflow\Office-Add-in-samples\AltTextManager-Deploy\
   ```

3. **지금 게시** 클릭

### 수동 설치

1. 빌드 출력 폴더로 이동:
   ```
   C:\woo-work\workflow\Office-Add-in-samples\AltTextManagerAddIn\bin\Debug\
   ```

2. `.dll` 파일과 `.vsto` 파일 확인

3. PowerPoint에서:
   - **파일** > **옵션** > **추가 기능**
   - **관리**: **PowerPoint 추가 기능** 선택
   - **이동** 클릭
   - **새로 추가** 클릭
   - `.ppam` 또는 `.vsto` 파일 선택

## 문제 해결

### "Office PIA가 설치되지 않음" 오류

NuGet 패키지 설치:
```
Install-Package Microsoft.Office.Interop.PowerPoint
```

### 디버그 시 PowerPoint가 시작되지 않음

1. **프로젝트 속성** > **디버그**
2. **시작 작업**: **외부 프로그램 시작** 선택
3. 경로 입력:
   ```
   C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE
   ```

## 참고 자료

- [VSTO 개발 문서](https://learn.microsoft.com/visualstudio/vsto/office-solutions-development-overview-vsto)
- [PowerPoint Object Model](https://learn.microsoft.com/office/vba/api/overview/powerpoint/object-model)
