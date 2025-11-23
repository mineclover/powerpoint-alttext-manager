using System;
using Microsoft.Office.Tools.Ribbon;

namespace AltTextManager
{
    public partial class AltTextRibbon
    {
        private void AltTextRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // 리본 로드 시 초기화
        }

        private void btnShowManager_Click(object sender, RibbonControlEventArgs e)
        {
            // 대체 텍스트 관리자 폼 열기
            Globals.ThisAddIn.ShowAltTextManager();
        }
    }
}
