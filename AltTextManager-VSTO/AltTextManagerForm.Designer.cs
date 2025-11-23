namespace AltTextManager
{
    partial class AltTextManagerForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.lstShapes = new System.Windows.Forms.ListView();
            this.colIndex = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAltText = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lblAltText = new System.Windows.Forms.Label();
            this.txtAltText = new System.Windows.Forms.TextBox();
            this.btnApplySelected = new System.Windows.Forms.Button();
            this.btnApplyAll = new System.Windows.Forms.Button();
            this.lblAutoGenerate = new System.Windows.Forms.Label();
            this.cboTemplateType = new System.Windows.Forms.ComboBox();
            this.btnAutoGenerate = new System.Windows.Forms.Button();
            this.chkOnlyEmpty = new System.Windows.Forms.CheckBox();
            this.lblTemplateSample = new System.Windows.Forms.Label();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();
            //
            // lblTitle
            //
            this.lblTitle.AutoSize = true;
            this.lblTitle.Location = new System.Drawing.Point(12, 15);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(141, 12);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "현재 슬라이드 요소 목록:";
            //
            // btnRefresh
            //
            this.btnRefresh.Location = new System.Drawing.Point(490, 10);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(80, 23);
            this.btnRefresh.TabIndex = 1;
            this.btnRefresh.Text = "새로고침";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            //
            // lstShapes
            //
            this.lstShapes.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colIndex,
            this.colName,
            this.colType,
            this.colAltText});
            this.lstShapes.FullRowSelect = true;
            this.lstShapes.GridLines = true;
            this.lstShapes.HideSelection = false;
            this.lstShapes.Location = new System.Drawing.Point(12, 40);
            this.lstShapes.MultiSelect = false;
            this.lstShapes.Name = "lstShapes";
            this.lstShapes.Size = new System.Drawing.Size(558, 180);
            this.lstShapes.TabIndex = 2;
            this.lstShapes.UseCompatibleStateImageBehavior = false;
            this.lstShapes.View = System.Windows.Forms.View.Details;
            this.lstShapes.SelectedIndexChanged += new System.EventHandler(this.lstShapes_SelectedIndexChanged);
            //
            // colIndex
            //
            this.colIndex.Text = "#";
            this.colIndex.Width = 40;
            //
            // colName
            //
            this.colName.Text = "이름";
            this.colName.Width = 180;
            //
            // colType
            //
            this.colType.Text = "타입";
            this.colType.Width = 100;
            //
            // colAltText
            //
            this.colAltText.Text = "대체 텍스트";
            this.colAltText.Width = 220;
            //
            // lblAltText
            //
            this.lblAltText.AutoSize = true;
            this.lblAltText.Location = new System.Drawing.Point(12, 235);
            this.lblAltText.Name = "lblAltText";
            this.lblAltText.Size = new System.Drawing.Size(105, 12);
            this.lblAltText.TabIndex = 3;
            this.lblAltText.Text = "대체 텍스트 입력:";
            //
            // txtAltText
            //
            this.txtAltText.Location = new System.Drawing.Point(12, 255);
            this.txtAltText.Multiline = true;
            this.txtAltText.Name = "txtAltText";
            this.txtAltText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtAltText.Size = new System.Drawing.Size(558, 80);
            this.txtAltText.TabIndex = 4;
            //
            // btnApplySelected
            //
            this.btnApplySelected.Location = new System.Drawing.Point(12, 345);
            this.btnApplySelected.Name = "btnApplySelected";
            this.btnApplySelected.Size = new System.Drawing.Size(130, 30);
            this.btnApplySelected.TabIndex = 5;
            this.btnApplySelected.Text = "선택 항목에 적용";
            this.btnApplySelected.UseVisualStyleBackColor = true;
            this.btnApplySelected.Click += new System.EventHandler(this.btnApplySelected_Click);
            //
            // btnApplyAll
            //
            this.btnApplyAll.Location = new System.Drawing.Point(150, 345);
            this.btnApplyAll.Name = "btnApplyAll";
            this.btnApplyAll.Size = new System.Drawing.Size(100, 30);
            this.btnApplyAll.TabIndex = 6;
            this.btnApplyAll.Text = "모두 적용";
            this.btnApplyAll.UseVisualStyleBackColor = true;
            this.btnApplyAll.Click += new System.EventHandler(this.btnApplyAll_Click);
            //
            // lblAutoGenerate
            //
            this.lblAutoGenerate.AutoSize = true;
            this.lblAutoGenerate.Location = new System.Drawing.Point(12, 395);
            this.lblAutoGenerate.Name = "lblAutoGenerate";
            this.lblAutoGenerate.Size = new System.Drawing.Size(69, 12);
            this.lblAutoGenerate.TabIndex = 7;
            this.lblAutoGenerate.Text = "자동 생성:";
            //
            // cboTemplateType
            //
            this.cboTemplateType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboTemplateType.FormattingEnabled = true;
            this.cboTemplateType.Location = new System.Drawing.Point(12, 415);
            this.cboTemplateType.Name = "cboTemplateType";
            this.cboTemplateType.Size = new System.Drawing.Size(250, 20);
            this.cboTemplateType.TabIndex = 8;
            this.cboTemplateType.SelectedIndexChanged += new System.EventHandler(this.cboTemplateType_SelectedIndexChanged);
            //
            // btnAutoGenerate
            //
            this.btnAutoGenerate.Location = new System.Drawing.Point(270, 413);
            this.btnAutoGenerate.Name = "btnAutoGenerate";
            this.btnAutoGenerate.Size = new System.Drawing.Size(100, 23);
            this.btnAutoGenerate.TabIndex = 9;
            this.btnAutoGenerate.Text = "자동 생성";
            this.btnAutoGenerate.UseVisualStyleBackColor = true;
            this.btnAutoGenerate.Click += new System.EventHandler(this.btnAutoGenerate_Click);
            //
            // chkOnlyEmpty
            //
            this.chkOnlyEmpty.AutoSize = true;
            this.chkOnlyEmpty.Location = new System.Drawing.Point(12, 445);
            this.chkOnlyEmpty.Name = "chkOnlyEmpty";
            this.chkOnlyEmpty.Size = new System.Drawing.Size(132, 16);
            this.chkOnlyEmpty.TabIndex = 10;
            this.chkOnlyEmpty.Text = "빈 요소만 적용";
            this.chkOnlyEmpty.UseVisualStyleBackColor = true;
            //
            // lblTemplateSample
            //
            this.lblTemplateSample.AutoSize = true;
            this.lblTemplateSample.ForeColor = System.Drawing.SystemColors.GrayText;
            this.lblTemplateSample.Location = new System.Drawing.Point(12, 470);
            this.lblTemplateSample.Name = "lblTemplateSample";
            this.lblTemplateSample.Size = new System.Drawing.Size(233, 12);
            this.lblTemplateSample.TabIndex = 11;
            this.lblTemplateSample.Text = "예: 슬라이드 요소 - Rectangle 1";
            //
            // btnClearAll
            //
            this.btnClearAll.Location = new System.Drawing.Point(12, 500);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(90, 25);
            this.btnClearAll.TabIndex = 12;
            this.btnClearAll.Text = "모두 삭제";
            this.btnClearAll.UseVisualStyleBackColor = true;
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            //
            // btnClose
            //
            this.btnClose.Location = new System.Drawing.Point(480, 500);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(90, 25);
            this.btnClose.TabIndex = 13;
            this.btnClose.Text = "닫기";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            //
            // lblStatus
            //
            this.lblStatus.AutoSize = true;
            this.lblStatus.ForeColor = System.Drawing.SystemColors.GrayText;
            this.lblStatus.Location = new System.Drawing.Point(12, 540);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(29, 12);
            this.lblStatus.TabIndex = 14;
            this.lblStatus.Text = "준비";
            //
            // AltTextManagerForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 561);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnClearAll);
            this.Controls.Add(this.lblTemplateSample);
            this.Controls.Add(this.chkOnlyEmpty);
            this.Controls.Add(this.btnAutoGenerate);
            this.Controls.Add(this.cboTemplateType);
            this.Controls.Add(this.lblAutoGenerate);
            this.Controls.Add(this.btnApplyAll);
            this.Controls.Add(this.btnApplySelected);
            this.Controls.Add(this.txtAltText);
            this.Controls.Add(this.lblAltText);
            this.Controls.Add(this.lstShapes);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.lblTitle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AltTextManagerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PowerPoint 대체 텍스트 관리자";
            this.Load += new System.EventHandler(this.AltTextManagerForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.ListView lstShapes;
        private System.Windows.Forms.ColumnHeader colIndex;
        private System.Windows.Forms.ColumnHeader colName;
        private System.Windows.Forms.ColumnHeader colType;
        private System.Windows.Forms.ColumnHeader colAltText;
        private System.Windows.Forms.Label lblAltText;
        private System.Windows.Forms.TextBox txtAltText;
        private System.Windows.Forms.Button btnApplySelected;
        private System.Windows.Forms.Button btnApplyAll;
        private System.Windows.Forms.Label lblAutoGenerate;
        private System.Windows.Forms.ComboBox cboTemplateType;
        private System.Windows.Forms.Button btnAutoGenerate;
        private System.Windows.Forms.CheckBox chkOnlyEmpty;
        private System.Windows.Forms.Label lblTemplateSample;
        private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblStatus;
    }
}
