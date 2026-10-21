
namespace SESDTools
{
    partial class AdministratorForm
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
            this.gbGenerateXMLFile = new System.Windows.Forms.GroupBox();
            this.btnXmlFilePathOpen = new System.Windows.Forms.Button();
            this.txtXmlFilePath = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.btnDoneToXmlFile = new System.Windows.Forms.Button();
            this.txtExcelFliePath = new System.Windows.Forms.TextBox();
            this.btnExcelFilePathOpen = new System.Windows.Forms.Button();
            this.lblEXCELFilePath = new System.Windows.Forms.Label();
            this.gbAdminiPasswordUpdate = new System.Windows.Forms.GroupBox();
            this.btnDoneToPasswordUpdate = new System.Windows.Forms.Button();
            this.txtPasswordConfirmation = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.lblPasswordConfirmation = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.gbAdminiInforUpdate = new System.Windows.Forms.GroupBox();
            this.btnDoneToAdminiInforUpdate = new System.Windows.Forms.Button();
            this.txtMail = new System.Windows.Forms.TextBox();
            this.txtPhone = new System.Windows.Forms.TextBox();
            this.lblMail = new System.Windows.Forms.Label();
            this.lblPhone = new System.Windows.Forms.Label();
            this.gbSecureInterfaceUpdate = new System.Windows.Forms.GroupBox();
            this.dgvComputerName = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnDoneToSecureInterfaceUpdate = new System.Windows.Forms.Button();
            this.lblComputerName = new System.Windows.Forms.Label();
            this.gbVersionInforUpdate = new System.Windows.Forms.GroupBox();
            this.txtVersionInstructions = new System.Windows.Forms.TextBox();
            this.btnDoneToVersionInforUpdate = new System.Windows.Forms.Button();
            this.txtDeveloperPhone = new System.Windows.Forms.TextBox();
            this.txtVersionNumber = new System.Windows.Forms.TextBox();
            this.lblDeveloperPhone = new System.Windows.Forms.Label();
            this.lblVersionInstructions = new System.Windows.Forms.Label();
            this.lblVersionNumber = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.gbGenerateXMLFile.SuspendLayout();
            this.gbAdminiPasswordUpdate.SuspendLayout();
            this.gbAdminiInforUpdate.SuspendLayout();
            this.gbSecureInterfaceUpdate.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvComputerName)).BeginInit();
            this.gbVersionInforUpdate.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbGenerateXMLFile
            // 
            this.gbGenerateXMLFile.BackColor = System.Drawing.Color.Khaki;
            this.gbGenerateXMLFile.Controls.Add(this.btnXmlFilePathOpen);
            this.gbGenerateXMLFile.Controls.Add(this.txtXmlFilePath);
            this.gbGenerateXMLFile.Controls.Add(this.label10);
            this.gbGenerateXMLFile.Controls.Add(this.btnDoneToXmlFile);
            this.gbGenerateXMLFile.Controls.Add(this.txtExcelFliePath);
            this.gbGenerateXMLFile.Controls.Add(this.btnExcelFilePathOpen);
            this.gbGenerateXMLFile.Controls.Add(this.lblEXCELFilePath);
            this.gbGenerateXMLFile.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.gbGenerateXMLFile.Location = new System.Drawing.Point(30, 23);
            this.gbGenerateXMLFile.Name = "gbGenerateXMLFile";
            this.gbGenerateXMLFile.Size = new System.Drawing.Size(738, 115);
            this.gbGenerateXMLFile.TabIndex = 0;
            this.gbGenerateXMLFile.TabStop = false;
            this.gbGenerateXMLFile.Text = "Generate XML File";
            // 
            // btnXmlFilePathOpen
            // 
            this.btnXmlFilePathOpen.Location = new System.Drawing.Point(642, 49);
            this.btnXmlFilePathOpen.Name = "btnXmlFilePathOpen";
            this.btnXmlFilePathOpen.Size = new System.Drawing.Size(81, 21);
            this.btnXmlFilePathOpen.TabIndex = 6;
            this.btnXmlFilePathOpen.Text = "Open";
            this.btnXmlFilePathOpen.UseVisualStyleBackColor = true;
            this.btnXmlFilePathOpen.Click += new System.EventHandler(this.BtnXmlFilePathOpen_Click);
            // 
            // txtXmlFilePath
            // 
            this.txtXmlFilePath.Enabled = false;
            this.txtXmlFilePath.Location = new System.Drawing.Point(137, 50);
            this.txtXmlFilePath.Name = "txtXmlFilePath";
            this.txtXmlFilePath.Size = new System.Drawing.Size(492, 21);
            this.txtXmlFilePath.TabIndex = 5;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 55);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(101, 12);
            this.label10.TabIndex = 4;
            this.label10.Text = "To XML File Path";
            // 
            // btnDoneToXmlFile
            // 
            this.btnDoneToXmlFile.Location = new System.Drawing.Point(642, 82);
            this.btnDoneToXmlFile.Name = "btnDoneToXmlFile";
            this.btnDoneToXmlFile.Size = new System.Drawing.Size(81, 21);
            this.btnDoneToXmlFile.TabIndex = 3;
            this.btnDoneToXmlFile.Text = "Done";
            this.btnDoneToXmlFile.UseVisualStyleBackColor = true;
            this.btnDoneToXmlFile.Click += new System.EventHandler(this.BtnDoneToXmlFile_Click);
            // 
            // txtExcelFliePath
            // 
            this.txtExcelFliePath.Enabled = false;
            this.txtExcelFliePath.Location = new System.Drawing.Point(137, 18);
            this.txtExcelFliePath.Name = "txtExcelFliePath";
            this.txtExcelFliePath.Size = new System.Drawing.Size(492, 21);
            this.txtExcelFliePath.TabIndex = 2;
            // 
            // btnExcelFilePathOpen
            // 
            this.btnExcelFilePathOpen.Location = new System.Drawing.Point(642, 17);
            this.btnExcelFilePathOpen.Name = "btnExcelFilePathOpen";
            this.btnExcelFilePathOpen.Size = new System.Drawing.Size(81, 21);
            this.btnExcelFilePathOpen.TabIndex = 1;
            this.btnExcelFilePathOpen.Text = "Open";
            this.btnExcelFilePathOpen.UseVisualStyleBackColor = true;
            this.btnExcelFilePathOpen.Click += new System.EventHandler(this.BtnExcelFilePathOpen_Click);
            // 
            // lblEXCELFilePath
            // 
            this.lblEXCELFilePath.AutoSize = true;
            this.lblEXCELFilePath.Location = new System.Drawing.Point(6, 23);
            this.lblEXCELFilePath.Name = "lblEXCELFilePath";
            this.lblEXCELFilePath.Size = new System.Drawing.Size(95, 12);
            this.lblEXCELFilePath.TabIndex = 0;
            this.lblEXCELFilePath.Text = "EXCEL File Path";
            // 
            // gbAdminiPasswordUpdate
            // 
            this.gbAdminiPasswordUpdate.BackColor = System.Drawing.Color.LightSalmon;
            this.gbAdminiPasswordUpdate.Controls.Add(this.btnDoneToPasswordUpdate);
            this.gbAdminiPasswordUpdate.Controls.Add(this.txtPasswordConfirmation);
            this.gbAdminiPasswordUpdate.Controls.Add(this.txtPassword);
            this.gbAdminiPasswordUpdate.Controls.Add(this.lblPasswordConfirmation);
            this.gbAdminiPasswordUpdate.Controls.Add(this.lblPassword);
            this.gbAdminiPasswordUpdate.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.gbAdminiPasswordUpdate.Location = new System.Drawing.Point(30, 144);
            this.gbAdminiPasswordUpdate.Name = "gbAdminiPasswordUpdate";
            this.gbAdminiPasswordUpdate.Size = new System.Drawing.Size(356, 112);
            this.gbAdminiPasswordUpdate.TabIndex = 1;
            this.gbAdminiPasswordUpdate.TabStop = false;
            this.gbAdminiPasswordUpdate.Text = "Administrator Password Update";
            // 
            // btnDoneToPasswordUpdate
            // 
            this.btnDoneToPasswordUpdate.Location = new System.Drawing.Point(269, 85);
            this.btnDoneToPasswordUpdate.Name = "btnDoneToPasswordUpdate";
            this.btnDoneToPasswordUpdate.Size = new System.Drawing.Size(81, 21);
            this.btnDoneToPasswordUpdate.TabIndex = 4;
            this.btnDoneToPasswordUpdate.Text = "Done";
            this.btnDoneToPasswordUpdate.UseVisualStyleBackColor = true;
            this.btnDoneToPasswordUpdate.Click += new System.EventHandler(this.BtnDoneToPasswordUpdate_Click);
            // 
            // txtPasswordConfirmation
            // 
            this.txtPasswordConfirmation.Location = new System.Drawing.Point(140, 54);
            this.txtPasswordConfirmation.Name = "txtPasswordConfirmation";
            this.txtPasswordConfirmation.Size = new System.Drawing.Size(114, 21);
            this.txtPasswordConfirmation.TabIndex = 3;
            this.txtPasswordConfirmation.UseSystemPasswordChar = true;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(140, 27);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(114, 21);
            this.txtPassword.TabIndex = 2;
            this.txtPassword.UseSystemPasswordChar = true;
            // 
            // lblPasswordConfirmation
            // 
            this.lblPasswordConfirmation.AutoSize = true;
            this.lblPasswordConfirmation.Location = new System.Drawing.Point(6, 60);
            this.lblPasswordConfirmation.Name = "lblPasswordConfirmation";
            this.lblPasswordConfirmation.Size = new System.Drawing.Size(131, 12);
            this.lblPasswordConfirmation.TabIndex = 1;
            this.lblPasswordConfirmation.Text = "Password Confirmation";
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(6, 30);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(53, 12);
            this.lblPassword.TabIndex = 0;
            this.lblPassword.Text = "Password";
            // 
            // gbAdminiInforUpdate
            // 
            this.gbAdminiInforUpdate.BackColor = System.Drawing.Color.Bisque;
            this.gbAdminiInforUpdate.Controls.Add(this.btnDoneToAdminiInforUpdate);
            this.gbAdminiInforUpdate.Controls.Add(this.txtMail);
            this.gbAdminiInforUpdate.Controls.Add(this.txtPhone);
            this.gbAdminiInforUpdate.Controls.Add(this.lblMail);
            this.gbAdminiInforUpdate.Controls.Add(this.lblPhone);
            this.gbAdminiInforUpdate.Location = new System.Drawing.Point(392, 144);
            this.gbAdminiInforUpdate.Name = "gbAdminiInforUpdate";
            this.gbAdminiInforUpdate.Size = new System.Drawing.Size(376, 112);
            this.gbAdminiInforUpdate.TabIndex = 2;
            this.gbAdminiInforUpdate.TabStop = false;
            this.gbAdminiInforUpdate.Text = "Administrator Information Update";
            // 
            // btnDoneToAdminiInforUpdate
            // 
            this.btnDoneToAdminiInforUpdate.Location = new System.Drawing.Point(280, 85);
            this.btnDoneToAdminiInforUpdate.Name = "btnDoneToAdminiInforUpdate";
            this.btnDoneToAdminiInforUpdate.Size = new System.Drawing.Size(81, 21);
            this.btnDoneToAdminiInforUpdate.TabIndex = 6;
            this.btnDoneToAdminiInforUpdate.Text = "Done";
            this.btnDoneToAdminiInforUpdate.UseVisualStyleBackColor = true;
            this.btnDoneToAdminiInforUpdate.Click += new System.EventHandler(this.BtnDoneToAdminiInforUpdate_Click);
            // 
            // txtMail
            // 
            this.txtMail.Location = new System.Drawing.Point(47, 51);
            this.txtMail.Name = "txtMail";
            this.txtMail.Size = new System.Drawing.Size(201, 21);
            this.txtMail.TabIndex = 5;
            // 
            // txtPhone
            // 
            this.txtPhone.Location = new System.Drawing.Point(47, 24);
            this.txtPhone.MaxLength = 11;
            this.txtPhone.Name = "txtPhone";
            this.txtPhone.Size = new System.Drawing.Size(201, 21);
            this.txtPhone.TabIndex = 4;
            // 
            // lblMail
            // 
            this.lblMail.AutoSize = true;
            this.lblMail.Location = new System.Drawing.Point(6, 54);
            this.lblMail.Name = "lblMail";
            this.lblMail.Size = new System.Drawing.Size(29, 12);
            this.lblMail.TabIndex = 3;
            this.lblMail.Text = "Mail";
            // 
            // lblPhone
            // 
            this.lblPhone.AutoSize = true;
            this.lblPhone.Location = new System.Drawing.Point(6, 27);
            this.lblPhone.Name = "lblPhone";
            this.lblPhone.Size = new System.Drawing.Size(35, 12);
            this.lblPhone.TabIndex = 2;
            this.lblPhone.Text = "Phone";
            // 
            // gbSecureInterfaceUpdate
            // 
            this.gbSecureInterfaceUpdate.BackColor = System.Drawing.Color.LightGreen;
            this.gbSecureInterfaceUpdate.Controls.Add(this.dgvComputerName);
            this.gbSecureInterfaceUpdate.Controls.Add(this.btnDoneToSecureInterfaceUpdate);
            this.gbSecureInterfaceUpdate.Controls.Add(this.lblComputerName);
            this.gbSecureInterfaceUpdate.Location = new System.Drawing.Point(30, 262);
            this.gbSecureInterfaceUpdate.Name = "gbSecureInterfaceUpdate";
            this.gbSecureInterfaceUpdate.Size = new System.Drawing.Size(735, 245);
            this.gbSecureInterfaceUpdate.TabIndex = 3;
            this.gbSecureInterfaceUpdate.TabStop = false;
            this.gbSecureInterfaceUpdate.Text = "Secure Interface Update";
            // 
            // dgvComputerName
            // 
            this.dgvComputerName.BackgroundColor = System.Drawing.Color.LightGreen;
            this.dgvComputerName.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvComputerName.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2});
            this.dgvComputerName.Location = new System.Drawing.Point(137, 20);
            this.dgvComputerName.Name = "dgvComputerName";
            this.dgvComputerName.RowTemplate.Height = 23;
            this.dgvComputerName.Size = new System.Drawing.Size(447, 201);
            this.dgvComputerName.TabIndex = 8;
            this.dgvComputerName.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(this.DgvComputerName_RowStateChanged);
            // 
            // Column1
            // 
            this.Column1.HeaderText = "User Name";
            this.Column1.Name = "Column1";
            this.Column1.Width = 200;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Computer Name";
            this.Column2.Name = "Column2";
            this.Column2.Width = 200;
            // 
            // btnDoneToSecureInterfaceUpdate
            // 
            this.btnDoneToSecureInterfaceUpdate.Location = new System.Drawing.Point(642, 218);
            this.btnDoneToSecureInterfaceUpdate.Name = "btnDoneToSecureInterfaceUpdate";
            this.btnDoneToSecureInterfaceUpdate.Size = new System.Drawing.Size(81, 21);
            this.btnDoneToSecureInterfaceUpdate.TabIndex = 7;
            this.btnDoneToSecureInterfaceUpdate.Text = "Done";
            this.btnDoneToSecureInterfaceUpdate.UseVisualStyleBackColor = true;
            this.btnDoneToSecureInterfaceUpdate.Click += new System.EventHandler(this.BtnDoneToSecureInterfaceUpdate_Click);
            // 
            // lblComputerName
            // 
            this.lblComputerName.AutoSize = true;
            this.lblComputerName.Location = new System.Drawing.Point(6, 27);
            this.lblComputerName.Name = "lblComputerName";
            this.lblComputerName.Size = new System.Drawing.Size(83, 12);
            this.lblComputerName.TabIndex = 4;
            this.lblComputerName.Text = "Computer Name";
            // 
            // gbVersionInforUpdate
            // 
            this.gbVersionInforUpdate.BackColor = System.Drawing.Color.LightSkyBlue;
            this.gbVersionInforUpdate.Controls.Add(this.txtVersionInstructions);
            this.gbVersionInforUpdate.Controls.Add(this.btnDoneToVersionInforUpdate);
            this.gbVersionInforUpdate.Controls.Add(this.txtDeveloperPhone);
            this.gbVersionInforUpdate.Controls.Add(this.txtVersionNumber);
            this.gbVersionInforUpdate.Controls.Add(this.lblDeveloperPhone);
            this.gbVersionInforUpdate.Controls.Add(this.lblVersionInstructions);
            this.gbVersionInforUpdate.Controls.Add(this.lblVersionNumber);
            this.gbVersionInforUpdate.Location = new System.Drawing.Point(30, 513);
            this.gbVersionInforUpdate.Name = "gbVersionInforUpdate";
            this.gbVersionInforUpdate.Size = new System.Drawing.Size(734, 157);
            this.gbVersionInforUpdate.TabIndex = 4;
            this.gbVersionInforUpdate.TabStop = false;
            this.gbVersionInforUpdate.Text = "Version Information Update";
            // 
            // txtVersionInstructions
            // 
            this.txtVersionInstructions.Location = new System.Drawing.Point(137, 52);
            this.txtVersionInstructions.Multiline = true;
            this.txtVersionInstructions.Name = "txtVersionInstructions";
            this.txtVersionInstructions.Size = new System.Drawing.Size(447, 50);
            this.txtVersionInstructions.TabIndex = 12;
            // 
            // btnDoneToVersionInforUpdate
            // 
            this.btnDoneToVersionInforUpdate.Location = new System.Drawing.Point(642, 125);
            this.btnDoneToVersionInforUpdate.Name = "btnDoneToVersionInforUpdate";
            this.btnDoneToVersionInforUpdate.Size = new System.Drawing.Size(81, 21);
            this.btnDoneToVersionInforUpdate.TabIndex = 11;
            this.btnDoneToVersionInforUpdate.Text = "Done";
            this.btnDoneToVersionInforUpdate.UseVisualStyleBackColor = true;
            this.btnDoneToVersionInforUpdate.Click += new System.EventHandler(this.BtnDoneToVersionInforUpdate_Click);
            // 
            // txtDeveloperPhone
            // 
            this.txtDeveloperPhone.Location = new System.Drawing.Point(137, 108);
            this.txtDeveloperPhone.Name = "txtDeveloperPhone";
            this.txtDeveloperPhone.Size = new System.Drawing.Size(447, 21);
            this.txtDeveloperPhone.TabIndex = 9;
            // 
            // txtVersionNumber
            // 
            this.txtVersionNumber.Location = new System.Drawing.Point(137, 24);
            this.txtVersionNumber.Name = "txtVersionNumber";
            this.txtVersionNumber.Size = new System.Drawing.Size(447, 21);
            this.txtVersionNumber.TabIndex = 8;
            // 
            // lblDeveloperPhone
            // 
            this.lblDeveloperPhone.AutoSize = true;
            this.lblDeveloperPhone.Location = new System.Drawing.Point(6, 111);
            this.lblDeveloperPhone.Name = "lblDeveloperPhone";
            this.lblDeveloperPhone.Size = new System.Drawing.Size(95, 12);
            this.lblDeveloperPhone.TabIndex = 7;
            this.lblDeveloperPhone.Text = "Developer Phone";
            // 
            // lblVersionInstructions
            // 
            this.lblVersionInstructions.AutoSize = true;
            this.lblVersionInstructions.Location = new System.Drawing.Point(6, 50);
            this.lblVersionInstructions.Name = "lblVersionInstructions";
            this.lblVersionInstructions.Size = new System.Drawing.Size(125, 12);
            this.lblVersionInstructions.TabIndex = 6;
            this.lblVersionInstructions.Text = "Version Instructions";
            // 
            // lblVersionNumber
            // 
            this.lblVersionNumber.AutoSize = true;
            this.lblVersionNumber.Location = new System.Drawing.Point(6, 27);
            this.lblVersionNumber.Name = "lblVersionNumber";
            this.lblVersionNumber.Size = new System.Drawing.Size(89, 12);
            this.lblVersionNumber.TabIndex = 5;
            this.lblVersionNumber.Text = "Version Number";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(672, 680);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(81, 21);
            this.btnClose.TabIndex = 12;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // AdministratorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 716);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.gbVersionInforUpdate);
            this.Controls.Add(this.gbSecureInterfaceUpdate);
            this.Controls.Add(this.gbAdminiInforUpdate);
            this.Controls.Add(this.gbAdminiPasswordUpdate);
            this.Controls.Add(this.gbGenerateXMLFile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AdministratorForm";
            this.Text = "Authority Management";
            this.gbGenerateXMLFile.ResumeLayout(false);
            this.gbGenerateXMLFile.PerformLayout();
            this.gbAdminiPasswordUpdate.ResumeLayout(false);
            this.gbAdminiPasswordUpdate.PerformLayout();
            this.gbAdminiInforUpdate.ResumeLayout(false);
            this.gbAdminiInforUpdate.PerformLayout();
            this.gbSecureInterfaceUpdate.ResumeLayout(false);
            this.gbSecureInterfaceUpdate.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvComputerName)).EndInit();
            this.gbVersionInforUpdate.ResumeLayout(false);
            this.gbVersionInforUpdate.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbGenerateXMLFile;
        private System.Windows.Forms.GroupBox gbAdminiPasswordUpdate;
        private System.Windows.Forms.GroupBox gbAdminiInforUpdate;
        private System.Windows.Forms.GroupBox gbSecureInterfaceUpdate;
        private System.Windows.Forms.GroupBox gbVersionInforUpdate;
        private System.Windows.Forms.Button btnDoneToXmlFile;
        private System.Windows.Forms.TextBox txtExcelFliePath;
        private System.Windows.Forms.Button btnExcelFilePathOpen;
        private System.Windows.Forms.Label lblEXCELFilePath;
        private System.Windows.Forms.TextBox txtPasswordConfirmation;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label lblPasswordConfirmation;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.Button btnDoneToPasswordUpdate;
        private System.Windows.Forms.Button btnDoneToAdminiInforUpdate;
        private System.Windows.Forms.TextBox txtMail;
        private System.Windows.Forms.TextBox txtPhone;
        private System.Windows.Forms.Label lblMail;
        private System.Windows.Forms.Label lblPhone;
        private System.Windows.Forms.Button btnDoneToSecureInterfaceUpdate;
        private System.Windows.Forms.Label lblComputerName;
        private System.Windows.Forms.Label lblDeveloperPhone;
        private System.Windows.Forms.Label lblVersionInstructions;
        private System.Windows.Forms.Label lblVersionNumber;
        private System.Windows.Forms.TextBox txtDeveloperPhone;
        private System.Windows.Forms.TextBox txtVersionNumber;
        private System.Windows.Forms.Button btnDoneToVersionInforUpdate;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnXmlFilePathOpen;
        private System.Windows.Forms.TextBox txtXmlFilePath;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.DataGridView dgvComputerName;
        private System.Windows.Forms.TextBox txtVersionInstructions;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
    }
}