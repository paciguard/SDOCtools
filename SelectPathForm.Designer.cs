
namespace SESDTools
{
    partial class SelectPathForm
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
            this.lblPrintFilePath = new System.Windows.Forms.Label();
            this.txtPrintFilePath = new System.Windows.Forms.TextBox();
            this.btnPrintFilePathOpen = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOutput = new System.Windows.Forms.Button();
            this.labelRecieve = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblPrintFilePath
            // 
            this.lblPrintFilePath.AutoSize = true;
            this.lblPrintFilePath.Location = new System.Drawing.Point(12, 40);
            this.lblPrintFilePath.Name = "lblPrintFilePath";
            this.lblPrintFilePath.Size = new System.Drawing.Size(113, 12);
            this.lblPrintFilePath.TabIndex = 0;
            this.lblPrintFilePath.Text = "Path To Print File";
            // 
            // txtPrintFilePath
            // 
            this.txtPrintFilePath.Enabled = false;
            this.txtPrintFilePath.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtPrintFilePath.Location = new System.Drawing.Point(12, 55);
            this.txtPrintFilePath.Name = "txtPrintFilePath";
            this.txtPrintFilePath.Size = new System.Drawing.Size(320, 21);
            this.txtPrintFilePath.TabIndex = 13;
            // 
            // btnPrintFilePathOpen
            // 
            this.btnPrintFilePathOpen.Location = new System.Drawing.Point(338, 55);
            this.btnPrintFilePathOpen.Name = "btnPrintFilePathOpen";
            this.btnPrintFilePathOpen.Size = new System.Drawing.Size(46, 21);
            this.btnPrintFilePathOpen.TabIndex = 17;
            this.btnPrintFilePathOpen.Text = "Open";
            this.btnPrintFilePathOpen.UseVisualStyleBackColor = true;
            this.btnPrintFilePathOpen.Click += new System.EventHandler(this.BtnPrintFilePathOpen_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(294, 150);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(89, 29);
            this.btnClose.TabIndex = 19;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // btnOutput
            // 
            this.btnOutput.Location = new System.Drawing.Point(199, 150);
            this.btnOutput.Name = "btnOutput";
            this.btnOutput.Size = new System.Drawing.Size(89, 29);
            this.btnOutput.TabIndex = 18;
            this.btnOutput.Text = "Output";
            this.btnOutput.UseVisualStyleBackColor = true;
            this.btnOutput.Click += new System.EventHandler(this.BtnOutput_Click);
            // 
            // labelRecieve
            // 
            this.labelRecieve.AutoSize = true;
            this.labelRecieve.Location = new System.Drawing.Point(12, 150);
            this.labelRecieve.Name = "labelRecieve";
            this.labelRecieve.Size = new System.Drawing.Size(0, 12);
            this.labelRecieve.TabIndex = 20;
            this.labelRecieve.Visible = false;
            // 
            // SelectPathForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LavenderBlush;
            this.ClientSize = new System.Drawing.Size(396, 202);
            this.Controls.Add(this.labelRecieve);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOutput);
            this.Controls.Add(this.btnPrintFilePathOpen);
            this.Controls.Add(this.txtPrintFilePath);
            this.Controls.Add(this.lblPrintFilePath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SelectPathForm";
            this.Text = "Select Path";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblPrintFilePath;
        private System.Windows.Forms.TextBox txtPrintFilePath;
        private System.Windows.Forms.Button btnPrintFilePathOpen;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOutput;
        private System.Windows.Forms.Label labelRecieve;
    }
}