using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace AltTextManager
{
    public partial class ThisAddIn
    {
        private AltTextManagerForm managerForm;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 애드인 로드 시 초기화
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 애드인 종료 시 정리
            if (managerForm != null && !managerForm.IsDisposed)
            {
                managerForm.Close();
                managerForm.Dispose();
            }
        }

        /// <summary>
        /// 대체 텍스트 관리자 폼 표시
        /// </summary>
        public void ShowAltTextManager()
        {
            try
            {
                if (managerForm == null || managerForm.IsDisposed)
                {
                    managerForm = new AltTextManagerForm(this.Application);
                }

                if (!managerForm.Visible)
                {
                    managerForm.Show();
                }
                else
                {
                    managerForm.BringToFront();
                    managerForm.Activate();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"대체 텍스트 관리자를 열 수 없습니다.\n\n오류: {ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
