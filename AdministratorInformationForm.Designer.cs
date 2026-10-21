
namespace SESDTools
{
    partial class AdministratorInformationForm
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
            this.gbAdministratorInformation = new System.Windows.Forms.GroupBox();
            this.txtMail = new System.Windows.Forms.TextBox();
            this.txtPhone = new System.Windows.Forms.TextBox();
            this.lblMail = new System.Windows.Forms.Label();
            this.lblPhone = new System.Windows.Forms.Label();
            this.gbAdministratorInformation.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbAdministratorInformation
            // 
            this.gbAdministratorInformation.BackColor = System.Drawing.Color.Bisque;
            this.gbAdministratorInformation.Controls.Add(this.txtMail);
            this.gbAdministratorInformation.Controls.Add(this.txtPhone);
            this.gbAdministratorInformation.Controls.Add(this.lblMail);
            this.gbAdministratorInformation.Controls.Add(this.lblPhone);
            this.gbAdministratorInformation.Location = new System.Drawing.Point(3, 2);
            this.gbAdministratorInformation.Name = "gbAdministratorInformation";
            this.gbAdministratorInformation.Size = new System.Drawing.Size(373, 158);
            this.gbAdministratorInformation.TabIndex = 0;
            this.gbAdministratorInformation.TabStop = false;
            this.gbAdministratorInformation.Text = "Administrator Information";
            // 
            // txtMail
            // 
            this.txtMail.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtMail.Enabled = false;
            this.txtMail.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtMail.Location = new System.Drawing.Point(76, 85);
            this.txtMail.Name = "txtMail";
            this.txtMail.Size = new System.Drawing.Size(263, 14);
            this.txtMail.TabIndex = 3;
            // 
            // txtPhone
            // 
            this.txtPhone.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtPhone.Enabled = false;
            this.txtPhone.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtPhone.Location = new System.Drawing.Point(76, 56);
            this.txtPhone.Name = "txtPhone";
            this.txtPhone.Size = new System.Drawing.Size(117, 14);
            this.txtPhone.TabIndex = 2;
            // 
            // lblMail
            // 
            this.lblMail.AutoSize = true;
            this.lblMail.Location = new System.Drawing.Point(35, 84);
            this.lblMail.Name = "lblMail";
            this.lblMail.Size = new System.Drawing.Size(35, 12);
            this.lblMail.TabIndex = 1;
            this.lblMail.Text = "Mail:";
            // 
            // lblPhone
            // 
            this.lblPhone.AutoSize = true;
            this.lblPhone.Cursor = System.Windows.Forms.Cursors.Default;
            this.lblPhone.Location = new System.Drawing.Point(35, 54);
            this.lblPhone.Name = "lblPhone";
            this.lblPhone.Size = new System.Drawing.Size(41, 12);
            this.lblPhone.TabIndex = 0;
            this.lblPhone.Text = "Phone:";
            // 
            // AdministratorInformationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(379, 164);
            this.Controls.Add(this.gbAdministratorInformation);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AdministratorInformationForm";
            this.Text = "AdministratorInformationForm";
            this.gbAdministratorInformation.ResumeLayout(false);
            this.gbAdministratorInformation.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbAdministratorInformation;
        private System.Windows.Forms.TextBox txtMail;
        private System.Windows.Forms.TextBox txtPhone;
        private System.Windows.Forms.Label lblMail;
        private System.Windows.Forms.Label lblPhone;
    }
}