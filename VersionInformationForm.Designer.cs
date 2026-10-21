
namespace SESDTools
{
    partial class VersionInformationForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtDeveloperPhone = new System.Windows.Forms.TextBox();
            this.txtVersionInstructions = new System.Windows.Forms.TextBox();
            this.txtVersionNumber = new System.Windows.Forms.TextBox();
            this.lblDeveloperPhone = new System.Windows.Forms.Label();
            this.lblVersionInstructions = new System.Windows.Forms.Label();
            this.lblVersionNumber = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.AccessibleName = "gbAdminiInformation";
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.BackColor = System.Drawing.Color.LightSkyBlue;
            this.groupBox1.Controls.Add(this.txtDeveloperPhone);
            this.groupBox1.Controls.Add(this.txtVersionInstructions);
            this.groupBox1.Controls.Add(this.txtVersionNumber);
            this.groupBox1.Controls.Add(this.lblDeveloperPhone);
            this.groupBox1.Controls.Add(this.lblVersionInstructions);
            this.groupBox1.Controls.Add(this.lblVersionNumber);
            this.groupBox1.Location = new System.Drawing.Point(1, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(394, 199);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Administrator Information";
            // 
            // txtDeveloperPhone
            // 
            this.txtDeveloperPhone.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtDeveloperPhone.Enabled = false;
            this.txtDeveloperPhone.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtDeveloperPhone.Location = new System.Drawing.Point(139, 149);
            this.txtDeveloperPhone.Name = "txtDeveloperPhone";
            this.txtDeveloperPhone.Size = new System.Drawing.Size(230, 14);
            this.txtDeveloperPhone.TabIndex = 5;
            // 
            // txtVersionInstructions
            // 
            this.txtVersionInstructions.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtVersionInstructions.Enabled = false;
            this.txtVersionInstructions.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtVersionInstructions.Location = new System.Drawing.Point(139, 66);
            this.txtVersionInstructions.Multiline = true;
            this.txtVersionInstructions.Name = "txtVersionInstructions";
            this.txtVersionInstructions.Size = new System.Drawing.Size(230, 71);
            this.txtVersionInstructions.TabIndex = 4;
            // 
            // txtVersionNumber
            // 
            this.txtVersionNumber.BackColor = System.Drawing.SystemColors.Window;
            this.txtVersionNumber.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtVersionNumber.Enabled = false;
            this.txtVersionNumber.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtVersionNumber.Location = new System.Drawing.Point(139, 40);
            this.txtVersionNumber.Name = "txtVersionNumber";
            this.txtVersionNumber.Size = new System.Drawing.Size(230, 14);
            this.txtVersionNumber.TabIndex = 3;
            // 
            // lblDeveloperPhone
            // 
            this.lblDeveloperPhone.AutoSize = true;
            this.lblDeveloperPhone.Location = new System.Drawing.Point(6, 148);
            this.lblDeveloperPhone.Name = "lblDeveloperPhone";
            this.lblDeveloperPhone.Size = new System.Drawing.Size(107, 12);
            this.lblDeveloperPhone.TabIndex = 2;
            this.lblDeveloperPhone.Text = "Developer Phone：";
            // 
            // lblVersionInstructions
            // 
            this.lblVersionInstructions.AutoSize = true;
            this.lblVersionInstructions.Location = new System.Drawing.Point(6, 64);
            this.lblVersionInstructions.Name = "lblVersionInstructions";
            this.lblVersionInstructions.Size = new System.Drawing.Size(131, 12);
            this.lblVersionInstructions.TabIndex = 1;
            this.lblVersionInstructions.Text = "Version Instructions:";
            // 
            // lblVersionNumber
            // 
            this.lblVersionNumber.AutoSize = true;
            this.lblVersionNumber.Location = new System.Drawing.Point(6, 39);
            this.lblVersionNumber.Name = "lblVersionNumber";
            this.lblVersionNumber.Size = new System.Drawing.Size(95, 12);
            this.lblVersionNumber.TabIndex = 0;
            this.lblVersionNumber.Text = "Version Number:";
            // 
            // VersionInformationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(396, 202);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "VersionInformationForm";
            this.Text = "Version InformationForm";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblVersionNumber;
        private System.Windows.Forms.Label lblVersionInstructions;
        private System.Windows.Forms.Label lblDeveloperPhone;
        private System.Windows.Forms.TextBox txtVersionNumber;
        private System.Windows.Forms.TextBox txtVersionInstructions;
        private System.Windows.Forms.TextBox txtDeveloperPhone;
    }
}