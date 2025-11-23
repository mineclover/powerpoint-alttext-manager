using System;
using System.Collections.Generic;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AltTextManager
{
    public partial class AltTextManagerForm : Form
    {
        private PowerPoint.Application pptApp;
        private List<ShapeInfo> currentShapes;

        public AltTextManagerForm(PowerPoint.Application application)
        {
            InitializeComponent();
            this.pptApp = application;
            this.currentShapes = new List<ShapeInfo>();
        }

        private void AltTextManagerForm_Load(object sender, EventArgs e)
        {
            // 폼 로드 시 초기화
            cboTemplateType.Items.Add("접두사 + 요소 이름");
            cboTemplateType.Items.Add("타입별 템플릿");
            cboTemplateType.Items.Add("사용자 정의 템플릿");
            cboTemplateType.SelectedIndex = 0;

            RefreshShapeList();
        }

        private void RefreshShapeList()
        {
            lstShapes.Items.Clear();
            currentShapes.Clear();

            try
            {
                if (pptApp.ActivePresentation == null)
                {
                    lblStatus.Text = "활성 프레젠테이션이 없습니다";
                    return;
                }

                if (pptApp.ActiveWindow.Selection.SlideRange.Count == 0)
                {
                    lblStatus.Text = "슬라이드를 선택하세요";
                    return;
                }

                PowerPoint.Slide slide = pptApp.ActiveWindow.Selection.SlideRange[1];

                int index = 1;
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    string shapeType = GetShapeTypeName(shape.Type);
                    string altText = shape.AlternativeText ?? "";

                    ShapeInfo info = new ShapeInfo
                    {
                        Index = index,
                        Name = shape.Name,
                        Type = shapeType,
                        AltText = altText,
                        Shape = shape
                    };

                    currentShapes.Add(info);

                    ListViewItem item = new ListViewItem(index.ToString());
                    item.SubItems.Add(shape.Name);
                    item.SubItems.Add(shapeType);
                    item.SubItems.Add(string.IsNullOrEmpty(altText) ? "(없음)" : altText);
                    item.Tag = info;

                    lstShapes.Items.Add(item);

                    index++;
                }

                lblStatus.Text = $"{currentShapes.Count}개의 요소를 찾았습니다";
            }
            catch (Exception ex)
            {
                lblStatus.Text = "오류: " + ex.Message;
            }
        }

        private string GetShapeTypeName(Microsoft.Office.Core.MsoShapeType shapeType)
        {
            switch (shapeType)
            {
                case Microsoft.Office.Core.MsoShapeType.msoAutoShape:
                    return "도형";
                case Microsoft.Office.Core.MsoShapeType.msoPicture:
                    return "이미지";
                case Microsoft.Office.Core.MsoShapeType.msoTextBox:
                    return "텍스트 상자";
                case Microsoft.Office.Core.MsoShapeType.msoPlaceholder:
                    return "자리 표시자";
                case Microsoft.Office.Core.MsoShapeType.msoTable:
                    return "표";
                case Microsoft.Office.Core.MsoShapeType.msoChart:
                    return "차트";
                case Microsoft.Office.Core.MsoShapeType.msoGroup:
                    return "그룹";
                case Microsoft.Office.Core.MsoShapeType.msoLine:
                    return "선";
                case Microsoft.Office.Core.MsoShapeType.msoMedia:
                    return "미디어";
                default:
                    return "기타";
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshShapeList();
            txtAltText.Clear();
        }

        private void lstShapes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstShapes.SelectedItems.Count > 0)
            {
                ShapeInfo info = (ShapeInfo)lstShapes.SelectedItems[0].Tag;
                txtAltText.Text = info.AltText;
            }
        }

        private void btnApplySelected_Click(object sender, EventArgs e)
        {
            if (lstShapes.SelectedItems.Count == 0)
            {
                MessageBox.Show("적용할 요소를 선택하세요", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                ShapeInfo info = (ShapeInfo)lstShapes.SelectedItems[0].Tag;
                info.Shape.AlternativeText = txtAltText.Text;

                MessageBox.Show("대체 텍스트가 적용되었습니다", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshShapeList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"적용 중 오류가 발생했습니다:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnApplyAll_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtAltText.Text))
            {
                MessageBox.Show("적용할 대체 텍스트를 입력하세요", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult result = MessageBox.Show(
                "모든 요소에 동일한 대체 텍스트를 적용하시겠습니까?",
                "확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.No)
                return;

            try
            {
                int count = 0;
                bool onlyEmpty = chkOnlyEmpty.Checked;

                foreach (ShapeInfo info in currentShapes)
                {
                    if (onlyEmpty && !string.IsNullOrEmpty(info.AltText))
                        continue;

                    info.Shape.AlternativeText = txtAltText.Text;
                    count++;
                }

                MessageBox.Show($"{count}개 요소에 대체 텍스트가 적용되었습니다", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshShapeList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"적용 중 오류가 발생했습니다:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAutoGenerate_Click(object sender, EventArgs e)
        {
            int templateType = cboTemplateType.SelectedIndex;
            bool onlyEmpty = chkOnlyEmpty.Checked;

            try
            {
                string template = "";

                switch (templateType)
                {
                    case 0: // 접두사 + 요소 이름
                        using (var inputForm = new InputDialog("접두사를 입력하세요:", "접두사 입력", "슬라이드 요소 - "))
                        {
                            if (inputForm.ShowDialog() != DialogResult.OK)
                                return;
                            template = inputForm.InputText + "{name}";
                        }
                        break;

                    case 1: // 타입별 템플릿
                        template = "{type}: {name}";
                        break;

                    case 2: // 사용자 정의 템플릿
                        using (var inputForm = new InputDialog(
                            "템플릿을 입력하세요:\n\n사용 가능한 변수:\n{name} = 요소 이름\n{type} = 요소 타입\n{index} = 순번",
                            "사용자 정의 템플릿",
                            "{type} {index}: {name}"))
                        {
                            if (inputForm.ShowDialog() != DialogResult.OK)
                                return;
                            template = inputForm.InputText;
                        }
                        break;
                }

                ApplyTemplate(template, onlyEmpty);

                MessageBox.Show("대체 텍스트가 자동 생성되었습니다", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshShapeList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"자동 생성 중 오류가 발생했습니다:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyTemplate(string template, bool onlyEmpty)
        {
            foreach (ShapeInfo info in currentShapes)
            {
                if (onlyEmpty && !string.IsNullOrEmpty(info.AltText))
                    continue;

                string altText = template
                    .Replace("{name}", info.Name)
                    .Replace("{type}", info.Type)
                    .Replace("{index}", info.Index.ToString());

                info.Shape.AlternativeText = altText;
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "모든 대체 텍스트를 삭제하시겠습니까?",
                "확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result == DialogResult.No)
                return;

            try
            {
                foreach (ShapeInfo info in currentShapes)
                {
                    info.Shape.AlternativeText = "";
                }

                MessageBox.Show("모든 대체 텍스트가 삭제되었습니다", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshShapeList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"삭제 중 오류가 발생했습니다:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cboTemplateType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cboTemplateType.SelectedIndex)
            {
                case 0:
                    lblTemplateSample.Text = "예: 슬라이드 요소 - Rectangle 1";
                    break;
                case 1:
                    lblTemplateSample.Text = "예: 도형: Rectangle 1";
                    break;
                case 2:
                    lblTemplateSample.Text = "예: {type} {index}: {name}";
                    break;
            }
        }

        // 도형 정보를 담는 클래스
        private class ShapeInfo
        {
            public int Index { get; set; }
            public string Name { get; set; }
            public string Type { get; set; }
            public string AltText { get; set; }
            public PowerPoint.Shape Shape { get; set; }
        }
    }
}
