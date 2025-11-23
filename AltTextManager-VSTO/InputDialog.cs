using System;
using System.Windows.Forms;

namespace AltTextManager
{
    public class InputDialog : Form
    {
        private Label lblPrompt;
        private TextBox txtInput;
        private Button btnOK;
        private Button btnCancel;

        public string InputText => txtInput.Text;

        public InputDialog(string prompt, string title, string defaultValue = "")
        {
            InitializeComponent();
            this.Text = title;
            lblPrompt.Text = prompt;
            txtInput.Text = defaultValue;
        }

        private void InitializeComponent()
        {
            this.lblPrompt = new Label();
            this.txtInput = new TextBox();
            this.btnOK = new Button();
            this.btnCancel = new Button();
            this.SuspendLayout();

            // lblPrompt
            this.lblPrompt.AutoSize = true;
            this.lblPrompt.Location = new System.Drawing.Point(12, 15);
            this.lblPrompt.MaximumSize = new System.Drawing.Size(360, 0);
            this.lblPrompt.Name = "lblPrompt";
            this.lblPrompt.Size = new System.Drawing.Size(100, 12);
            this.lblPrompt.TabIndex = 0;
            this.lblPrompt.Text = "입력:";

            // txtInput
            this.txtInput.Location = new System.Drawing.Point(12, 40);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(360, 21);
            this.txtInput.TabIndex = 1;

            // btnOK
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(216, 75);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "확인";
            this.btnOK.UseVisualStyleBackColor = true;

            // btnCancel
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(297, 75);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "취소";
            this.btnCancel.UseVisualStyleBackColor = true;

            // InputDialog
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(384, 111);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtInput);
            this.Controls.Add(this.lblPrompt);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InputDialog";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "입력";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
