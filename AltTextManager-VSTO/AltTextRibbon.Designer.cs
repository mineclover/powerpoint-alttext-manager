namespace AltTextManager
{
    partial class AltTextRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AltTextRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpAltText = this.Factory.CreateRibbonGroup();
            this.btnShowManager = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpAltText.SuspendLayout();
            this.SuspendLayout();
            //
            // tab1
            //
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpAltText);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            //
            // grpAltText
            //
            this.grpAltText.Items.Add(this.btnShowManager);
            this.grpAltText.Label = "대체 텍스트";
            this.grpAltText.Name = "grpAltText";
            //
            // btnShowManager
            //
            this.btnShowManager.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnShowManager.Label = "대체 텍스트 관리자";
            this.btnShowManager.Name = "btnShowManager";
            this.btnShowManager.ShowImage = true;
            this.btnShowManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowManager_Click);
            //
            // AltTextRibbon
            //
            this.Name = "AltTextRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AltTextRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpAltText.ResumeLayout(false);
            this.grpAltText.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAltText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowManager;
    }

    partial class ThisRibbonCollection
    {
        internal AltTextRibbon AltTextRibbon
        {
            get { return this.GetRibbon<AltTextRibbon>(); }
        }
    }
}
