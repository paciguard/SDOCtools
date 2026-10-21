
namespace SESDTools
{
    partial class SumListForm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.cbPGA = new System.Windows.Forms.CheckBox();
            this.cbPPI = new System.Windows.Forms.CheckBox();
            this.cbPCS = new System.Windows.Forms.CheckBox();
            this.cbPNT = new System.Windows.Forms.CheckBox();
            this.cbPUC = new System.Windows.Forms.CheckBox();
            this.cbPDS = new System.Windows.Forms.CheckBox();
            this.cbPSL = new System.Windows.Forms.CheckBox();
            this.dgvSumList = new System.Windows.Forms.DataGridView();
            this.contextMenuStripGDV = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.DeleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSumListDataImport = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.txtSumListFilePath = new System.Windows.Forms.TextBox();
            this.txtXMLFilePath = new System.Windows.Forms.TextBox();
            this.lblSumListFilePath = new System.Windows.Forms.Label();
            this.lblInputDataFileSavePath = new System.Windows.Forms.Label();
            this.btnSumListFilePathOpen = new System.Windows.Forms.Button();
            this.btnXMLFilePathOpen = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.palPublicData = new System.Windows.Forms.Panel();
            this.dateTimeWorkDate = new System.Windows.Forms.DateTimePicker();
            this.cobFlangeStandard = new System.Windows.Forms.ComboBox();
            this.txtApprovedBy = new System.Windows.Forms.TextBox();
            this.txtReviewedBy = new System.Windows.Forms.TextBox();
            this.txtPreparedBy = new System.Windows.Forms.TextBox();
            this.txtSulzerOrderNO = new System.Windows.Forms.TextBox();
            this.txtCustomerPONO = new System.Windows.Forms.TextBox();
            this.txtSite = new System.Windows.Forms.TextBox();
            this.txtProject = new System.Windows.Forms.TextBox();
            this.txtEndUser = new System.Windows.Forms.TextBox();
            this.lblFlangeStandard = new System.Windows.Forms.Label();
            this.lblWorkDate = new System.Windows.Forms.Label();
            this.lblApprovedBy = new System.Windows.Forms.Label();
            this.lblReviewedBy = new System.Windows.Forms.Label();
            this.lblPreparedBy = new System.Windows.Forms.Label();
            this.lblSulzerOrderNO = new System.Windows.Forms.Label();
            this.lblCustomerPONO = new System.Windows.Forms.Label();
            this.lblSite = new System.Windows.Forms.Label();
            this.lblProject = new System.Windows.Forms.Label();
            this.lblEndUser = new System.Windows.Forms.Label();
            this.txtPurchaser = new System.Windows.Forms.TextBox();
            this.lblPurchaser = new System.Windows.Forms.Label();
            this.lblXMLPath = new System.Windows.Forms.Label();
            this.textXMLPath = new System.Windows.Forms.TextBox();
            this.btnXMLPathOpen = new System.Windows.Forms.Button();
            this.btnXMLDataImport = new System.Windows.Forms.Button();
            this.ColItemNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColSeriesNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColTagNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColSize = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColPREPumpPressureClass = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColService = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColQty = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMotorManufacturer = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMotorFrame = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColFrameSize = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColMotorPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColVoltage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColSealFlushFlow = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColCoolingWaterFlow = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColSpeed = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMotorLengthK = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColApiEdition = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColSealPlanA = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColSealPlanB = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColSealPlanC = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColPumpDrainValve = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColWarmUpPiping = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColBearing = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColCompanionFlg = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColCaseBracket = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColBearighousingDrain = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColFlow = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHead = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColEfficiency = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColRatedPower = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColNpshr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMawp = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMawt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColHydrostaticP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColMaterialClass = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColBearingCageMaterial = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColPumpDrainNozzleSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColDryWetShaft = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColWetInput = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColNozzleLoad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColReservedPureLubrication = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColImpellerBalanceHole = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColPREInducerYesNo = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColCoolWaterMainPipeSize = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColDrainValveType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColVoluteCasing = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColImpeller = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColStationaryWearRing = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColCoverCoolingPlate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColThroatBush = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColImpellerKey = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColGasket = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColSpacerSleeve = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColImpellerNut = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColInducer = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColInducerSleeve = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColWearPlate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSumList)).BeginInit();
            this.contextMenuStripGDV.SuspendLayout();
            this.palPublicData.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbPGA
            // 
            this.cbPGA.AutoSize = true;
            this.cbPGA.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.cbPGA.Location = new System.Drawing.Point(6, 4);
            this.cbPGA.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbPGA.Name = "cbPGA";
            this.cbPGA.Size = new System.Drawing.Size(50, 19);
            this.cbPGA.TabIndex = 0;
            this.cbPGA.Text = "PGA";
            this.cbPGA.UseVisualStyleBackColor = true;
            this.cbPGA.CheckedChanged += new System.EventHandler(this.CbPGA_CheckedChanged);
            // 
            // cbPPI
            // 
            this.cbPPI.AutoSize = true;
            this.cbPPI.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.cbPPI.Location = new System.Drawing.Point(129, 4);
            this.cbPPI.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbPPI.Name = "cbPPI";
            this.cbPPI.Size = new System.Drawing.Size(45, 19);
            this.cbPPI.TabIndex = 1;
            this.cbPPI.Text = "PPI";
            this.cbPPI.UseVisualStyleBackColor = true;
            this.cbPPI.CheckedChanged += new System.EventHandler(this.CbPPI_CheckedChanged);
            // 
            // cbPCS
            // 
            this.cbPCS.AutoSize = true;
            this.cbPCS.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.cbPCS.Location = new System.Drawing.Point(245, 4);
            this.cbPCS.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbPCS.Name = "cbPCS";
            this.cbPCS.Size = new System.Drawing.Size(50, 19);
            this.cbPCS.TabIndex = 2;
            this.cbPCS.Text = "PCS";
            this.cbPCS.UseVisualStyleBackColor = true;
            this.cbPCS.CheckedChanged += new System.EventHandler(this.CbPCS_CheckedChanged);
            // 
            // cbPNT
            // 
            this.cbPNT.AutoSize = true;
            this.cbPNT.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.cbPNT.Location = new System.Drawing.Point(482, 4);
            this.cbPNT.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbPNT.Name = "cbPNT";
            this.cbPNT.Size = new System.Drawing.Size(49, 19);
            this.cbPNT.TabIndex = 3;
            this.cbPNT.Text = "PNT";
            this.cbPNT.UseVisualStyleBackColor = true;
            this.cbPNT.CheckedChanged += new System.EventHandler(this.CbPNT_CheckedChanged);
            // 
            // cbPUC
            // 
            this.cbPUC.AutoSize = true;
            this.cbPUC.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.cbPUC.Location = new System.Drawing.Point(356, 4);
            this.cbPUC.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbPUC.Name = "cbPUC";
            this.cbPUC.Size = new System.Drawing.Size(50, 19);
            this.cbPUC.TabIndex = 4;
            this.cbPUC.Text = "PUC";
            this.cbPUC.UseVisualStyleBackColor = true;
            this.cbPUC.CheckedChanged += new System.EventHandler(this.CbPUC_CheckedChanged);
            // 
            // cbPDS
            // 
            this.cbPDS.AutoSize = true;
            this.cbPDS.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.cbPDS.Location = new System.Drawing.Point(596, 4);
            this.cbPDS.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbPDS.Name = "cbPDS";
            this.cbPDS.Size = new System.Drawing.Size(50, 19);
            this.cbPDS.TabIndex = 5;
            this.cbPDS.Text = "PDS";
            this.cbPDS.UseVisualStyleBackColor = true;
            this.cbPDS.CheckedChanged += new System.EventHandler(this.CbPDS_CheckedChanged);
            // 
            // cbPSL
            // 
            this.cbPSL.AutoSize = true;
            this.cbPSL.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.cbPSL.Location = new System.Drawing.Point(712, 4);
            this.cbPSL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cbPSL.Name = "cbPSL";
            this.cbPSL.Size = new System.Drawing.Size(49, 19);
            this.cbPSL.TabIndex = 6;
            this.cbPSL.Text = "PSL";
            this.cbPSL.UseVisualStyleBackColor = true;
            this.cbPSL.CheckedChanged += new System.EventHandler(this.CbPSL_CheckedChanged);
            // 
            // dgvSumList
            // 
            this.dgvSumList.AllowUserToAddRows = false;
            this.dgvSumList.AllowUserToOrderColumns = true;
            this.dgvSumList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvSumList.BackgroundColor = System.Drawing.Color.LightSkyBlue;
            this.dgvSumList.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvSumList.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvSumList.ColumnHeadersHeight = 46;
            this.dgvSumList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColItemNo,
            this.ColSeriesNo,
            this.ColTagNo,
            this.ColType,
            this.ColSize,
            this.ColPREPumpPressureClass,
            this.ColService,
            this.ColQty,
            this.ColMotorManufacturer,
            this.ColMotorFrame,
            this.ColFrameSize,
            this.ColMotorPower,
            this.ColVoltage,
            this.ColSealFlushFlow,
            this.ColCoolingWaterFlow,
            this.ColSpeed,
            this.ColMotorLengthK,
            this.ColApiEdition,
            this.ColSealPlanA,
            this.ColSealPlanB,
            this.ColSealPlanC,
            this.ColPumpDrainValve,
            this.ColWarmUpPiping,
            this.ColBearing,
            this.ColCompanionFlg,
            this.ColCaseBracket,
            this.ColBearighousingDrain,
            this.ColFlow,
            this.ColHead,
            this.ColEfficiency,
            this.ColRatedPower,
            this.ColNpshr,
            this.ColMawp,
            this.ColMawt,
            this.ColHydrostaticP,
            this.ColMaterialClass,
            this.ColBearingCageMaterial,
            this.ColPumpDrainNozzleSize,
            this.ColDryWetShaft,
            this.ColWetInput,
            this.ColNozzleLoad,
            this.ColReservedPureLubrication,
            this.ColImpellerBalanceHole,
            this.ColPREInducerYesNo,
            this.ColCoolWaterMainPipeSize,
            this.ColDrainValveType,
            this.ColVoluteCasing,
            this.ColImpeller,
            this.ColStationaryWearRing,
            this.ColCoverCoolingPlate,
            this.ColThroatBush,
            this.ColImpellerKey,
            this.ColGasket,
            this.ColSpacerSleeve,
            this.ColImpellerNut,
            this.ColInducer,
            this.ColInducerSleeve,
            this.ColWearPlate});
            this.dgvSumList.ContextMenuStrip = this.contextMenuStripGDV;
            this.dgvSumList.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvSumList.EnableHeadersVisualStyles = false;
            this.dgvSumList.Location = new System.Drawing.Point(14, 264);
            this.dgvSumList.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dgvSumList.Name = "dgvSumList";
            this.dgvSumList.RowTemplate.Height = 23;
            this.dgvSumList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvSumList.Size = new System.Drawing.Size(1370, 492);
            this.dgvSumList.TabIndex = 7;
            this.dgvSumList.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DgvSumList_CellMouseDown);
            this.dgvSumList.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.DgvSumList_CellValidating);
            this.dgvSumList.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.DgvSumList_EditingControlShowing);
            // 
            // contextMenuStripGDV
            // 
            this.contextMenuStripGDV.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.DeleteToolStripMenuItem,
            this.AddToolStripMenuItem,
            this.PasteToolStripMenuItem});
            this.contextMenuStripGDV.Name = "contextMenuStripGDV";
            this.contextMenuStripGDV.Size = new System.Drawing.Size(113, 70);
            // 
            // DeleteToolStripMenuItem
            // 
            this.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem";
            this.DeleteToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.DeleteToolStripMenuItem.Text = "删除行";
            this.DeleteToolStripMenuItem.Click += new System.EventHandler(this.DeleteToolStripMenuItem_Click);
            // 
            // AddToolStripMenuItem
            // 
            this.AddToolStripMenuItem.Name = "AddToolStripMenuItem";
            this.AddToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.AddToolStripMenuItem.Text = "插入行";
            this.AddToolStripMenuItem.Click += new System.EventHandler(this.AddToolStripMenuItem_Click);
            // 
            // PasteToolStripMenuItem
            // 
            this.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem";
            this.PasteToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.PasteToolStripMenuItem.Text = "粘贴";
            this.PasteToolStripMenuItem.Click += new System.EventHandler(this.PasteToolStripMenuItem_Click);
            // 
            // btnSumListDataImport
            // 
            this.btnSumListDataImport.Location = new System.Drawing.Point(1036, 19);
            this.btnSumListDataImport.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnSumListDataImport.Name = "btnSumListDataImport";
            this.btnSumListDataImport.Size = new System.Drawing.Size(159, 26);
            this.btnSumListDataImport.TabIndex = 8;
            this.btnSumListDataImport.Text = "SumList Data Import";
            this.btnSumListDataImport.UseVisualStyleBackColor = true;
            this.btnSumListDataImport.Click += new System.EventHandler(this.BtnSumListDataImport_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.FlatAppearance.BorderSize = 0;
            this.btnSave.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.btnSave.Location = new System.Drawing.Point(1169, 808);
            this.btnSave.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(104, 36);
            this.btnSave.TabIndex = 10;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(1280, 808);
            this.btnClose.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(104, 36);
            this.btnClose.TabIndex = 11;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // txtSumListFilePath
            // 
            this.txtSumListFilePath.Enabled = false;
            this.txtSumListFilePath.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtSumListFilePath.Location = new System.Drawing.Point(199, 19);
            this.txtSumListFilePath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtSumListFilePath.Name = "txtSumListFilePath";
            this.txtSumListFilePath.Size = new System.Drawing.Size(747, 21);
            this.txtSumListFilePath.TabIndex = 12;
            // 
            // txtXMLFilePath
            // 
            this.txtXMLFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtXMLFilePath.Enabled = false;
            this.txtXMLFilePath.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtXMLFilePath.Location = new System.Drawing.Point(199, 764);
            this.txtXMLFilePath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtXMLFilePath.Name = "txtXMLFilePath";
            this.txtXMLFilePath.Size = new System.Drawing.Size(747, 21);
            this.txtXMLFilePath.TabIndex = 13;
            // 
            // lblSumListFilePath
            // 
            this.lblSumListFilePath.AutoSize = true;
            this.lblSumListFilePath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSumListFilePath.Location = new System.Drawing.Point(12, 22);
            this.lblSumListFilePath.Name = "lblSumListFilePath";
            this.lblSumListFilePath.Size = new System.Drawing.Size(105, 15);
            this.lblSumListFilePath.TabIndex = 14;
            this.lblSumListFilePath.Text = "SumList File Path";
            // 
            // lblInputDataFileSavePath
            // 
            this.lblInputDataFileSavePath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblInputDataFileSavePath.AutoSize = true;
            this.lblInputDataFileSavePath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInputDataFileSavePath.Location = new System.Drawing.Point(12, 769);
            this.lblInputDataFileSavePath.Name = "lblInputDataFileSavePath";
            this.lblInputDataFileSavePath.Size = new System.Drawing.Size(146, 15);
            this.lblInputDataFileSavePath.TabIndex = 15;
            this.lblInputDataFileSavePath.Text = "Input Data File Save Path";
            // 
            // btnSumListFilePathOpen
            // 
            this.btnSumListFilePathOpen.Location = new System.Drawing.Point(965, 19);
            this.btnSumListFilePathOpen.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnSumListFilePathOpen.Name = "btnSumListFilePathOpen";
            this.btnSumListFilePathOpen.Size = new System.Drawing.Size(64, 26);
            this.btnSumListFilePathOpen.TabIndex = 16;
            this.btnSumListFilePathOpen.Text = "Open";
            this.btnSumListFilePathOpen.UseVisualStyleBackColor = true;
            this.btnSumListFilePathOpen.Click += new System.EventHandler(this.BtnSumListFilePathOpen_Click);
            // 
            // btnXMLFilePathOpen
            // 
            this.btnXMLFilePathOpen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnXMLFilePathOpen.Location = new System.Drawing.Point(965, 764);
            this.btnXMLFilePathOpen.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnXMLFilePathOpen.Name = "btnXMLFilePathOpen";
            this.btnXMLFilePathOpen.Size = new System.Drawing.Size(54, 26);
            this.btnXMLFilePathOpen.TabIndex = 17;
            this.btnXMLFilePathOpen.Text = "Open";
            this.btnXMLFilePathOpen.UseVisualStyleBackColor = true;
            this.btnXMLFilePathOpen.Click += new System.EventHandler(this.BtnXMLFilePathOpen_Click);
            // 
            // btnClear
            // 
            this.btnClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClear.Location = new System.Drawing.Point(1058, 808);
            this.btnClear.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(104, 36);
            this.btnClear.TabIndex = 18;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.BtnClear_Click);
            // 
            // palPublicData
            // 
            this.palPublicData.Controls.Add(this.dateTimeWorkDate);
            this.palPublicData.Controls.Add(this.cobFlangeStandard);
            this.palPublicData.Controls.Add(this.txtApprovedBy);
            this.palPublicData.Controls.Add(this.txtReviewedBy);
            this.palPublicData.Controls.Add(this.txtPreparedBy);
            this.palPublicData.Controls.Add(this.txtSulzerOrderNO);
            this.palPublicData.Controls.Add(this.txtCustomerPONO);
            this.palPublicData.Controls.Add(this.txtSite);
            this.palPublicData.Controls.Add(this.txtProject);
            this.palPublicData.Controls.Add(this.txtEndUser);
            this.palPublicData.Controls.Add(this.lblFlangeStandard);
            this.palPublicData.Controls.Add(this.lblWorkDate);
            this.palPublicData.Controls.Add(this.lblApprovedBy);
            this.palPublicData.Controls.Add(this.lblReviewedBy);
            this.palPublicData.Controls.Add(this.lblPreparedBy);
            this.palPublicData.Controls.Add(this.lblSulzerOrderNO);
            this.palPublicData.Controls.Add(this.lblCustomerPONO);
            this.palPublicData.Controls.Add(this.lblSite);
            this.palPublicData.Controls.Add(this.lblProject);
            this.palPublicData.Controls.Add(this.lblEndUser);
            this.palPublicData.Controls.Add(this.txtPurchaser);
            this.palPublicData.Controls.Add(this.lblPurchaser);
            this.palPublicData.Controls.Add(this.cbPSL);
            this.palPublicData.Controls.Add(this.cbPGA);
            this.palPublicData.Controls.Add(this.cbPPI);
            this.palPublicData.Controls.Add(this.cbPCS);
            this.palPublicData.Controls.Add(this.cbPNT);
            this.palPublicData.Controls.Add(this.cbPUC);
            this.palPublicData.Controls.Add(this.cbPDS);
            this.palPublicData.Location = new System.Drawing.Point(14, 89);
            this.palPublicData.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.palPublicData.Name = "palPublicData";
            this.palPublicData.Size = new System.Drawing.Size(1370, 168);
            this.palPublicData.TabIndex = 19;
            // 
            // dateTimeWorkDate
            // 
            this.dateTimeWorkDate.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            this.dateTimeWorkDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimeWorkDate.Location = new System.Drawing.Point(129, 128);
            this.dateTimeWorkDate.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dateTimeWorkDate.Name = "dateTimeWorkDate";
            this.dateTimeWorkDate.Size = new System.Drawing.Size(291, 21);
            this.dateTimeWorkDate.TabIndex = 54;
            // 
            // cobFlangeStandard
            // 
            this.cobFlangeStandard.FormattingEnabled = true;
            this.cobFlangeStandard.Location = new System.Drawing.Point(566, 128);
            this.cobFlangeStandard.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.cobFlangeStandard.Name = "cobFlangeStandard";
            this.cobFlangeStandard.Size = new System.Drawing.Size(291, 23);
            this.cobFlangeStandard.TabIndex = 38;
            // 
            // txtApprovedBy
            // 
            this.txtApprovedBy.Location = new System.Drawing.Point(990, 96);
            this.txtApprovedBy.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtApprovedBy.Name = "txtApprovedBy";
            this.txtApprovedBy.Size = new System.Drawing.Size(291, 21);
            this.txtApprovedBy.TabIndex = 30;
            // 
            // txtReviewedBy
            // 
            this.txtReviewedBy.Location = new System.Drawing.Point(566, 96);
            this.txtReviewedBy.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtReviewedBy.Name = "txtReviewedBy";
            this.txtReviewedBy.Size = new System.Drawing.Size(291, 21);
            this.txtReviewedBy.TabIndex = 29;
            // 
            // txtPreparedBy
            // 
            this.txtPreparedBy.Location = new System.Drawing.Point(129, 96);
            this.txtPreparedBy.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtPreparedBy.Name = "txtPreparedBy";
            this.txtPreparedBy.Size = new System.Drawing.Size(291, 21);
            this.txtPreparedBy.TabIndex = 28;
            // 
            // txtSulzerOrderNO
            // 
            this.txtSulzerOrderNO.Location = new System.Drawing.Point(990, 66);
            this.txtSulzerOrderNO.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtSulzerOrderNO.Name = "txtSulzerOrderNO";
            this.txtSulzerOrderNO.Size = new System.Drawing.Size(291, 21);
            this.txtSulzerOrderNO.TabIndex = 27;
            // 
            // txtCustomerPONO
            // 
            this.txtCustomerPONO.Location = new System.Drawing.Point(566, 66);
            this.txtCustomerPONO.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtCustomerPONO.Name = "txtCustomerPONO";
            this.txtCustomerPONO.Size = new System.Drawing.Size(291, 21);
            this.txtCustomerPONO.TabIndex = 26;
            // 
            // txtSite
            // 
            this.txtSite.Location = new System.Drawing.Point(129, 66);
            this.txtSite.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtSite.Name = "txtSite";
            this.txtSite.Size = new System.Drawing.Size(291, 21);
            this.txtSite.TabIndex = 25;
            // 
            // txtProject
            // 
            this.txtProject.Location = new System.Drawing.Point(990, 38);
            this.txtProject.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtProject.Name = "txtProject";
            this.txtProject.Size = new System.Drawing.Size(291, 21);
            this.txtProject.TabIndex = 24;
            // 
            // txtEndUser
            // 
            this.txtEndUser.Location = new System.Drawing.Point(566, 38);
            this.txtEndUser.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtEndUser.Name = "txtEndUser";
            this.txtEndUser.Size = new System.Drawing.Size(291, 21);
            this.txtEndUser.TabIndex = 23;
            // 
            // lblFlangeStandard
            // 
            this.lblFlangeStandard.AutoSize = true;
            this.lblFlangeStandard.Location = new System.Drawing.Point(428, 131);
            this.lblFlangeStandard.Name = "lblFlangeStandard";
            this.lblFlangeStandard.Size = new System.Drawing.Size(99, 15);
            this.lblFlangeStandard.TabIndex = 18;
            this.lblFlangeStandard.Text = "Flange Standard";
            // 
            // lblWorkDate
            // 
            this.lblWorkDate.AutoSize = true;
            this.lblWorkDate.Location = new System.Drawing.Point(7, 130);
            this.lblWorkDate.Name = "lblWorkDate";
            this.lblWorkDate.Size = new System.Drawing.Size(62, 15);
            this.lblWorkDate.TabIndex = 17;
            this.lblWorkDate.Text = "Date/日期";
            // 
            // lblApprovedBy
            // 
            this.lblApprovedBy.AutoSize = true;
            this.lblApprovedBy.Location = new System.Drawing.Point(864, 100);
            this.lblApprovedBy.Name = "lblApprovedBy";
            this.lblApprovedBy.Size = new System.Drawing.Size(106, 15);
            this.lblApprovedBy.TabIndex = 16;
            this.lblApprovedBy.Text = "Approved by/批准";
            // 
            // lblReviewedBy
            // 
            this.lblReviewedBy.AutoSize = true;
            this.lblReviewedBy.Location = new System.Drawing.Point(428, 100);
            this.lblReviewedBy.Name = "lblReviewedBy";
            this.lblReviewedBy.Size = new System.Drawing.Size(107, 15);
            this.lblReviewedBy.TabIndex = 15;
            this.lblReviewedBy.Text = "Reviewed by/审核";
            // 
            // lblPreparedBy
            // 
            this.lblPreparedBy.AutoSize = true;
            this.lblPreparedBy.Location = new System.Drawing.Point(5, 100);
            this.lblPreparedBy.Name = "lblPreparedBy";
            this.lblPreparedBy.Size = new System.Drawing.Size(105, 15);
            this.lblPreparedBy.TabIndex = 14;
            this.lblPreparedBy.Text = "Prepared by/编制";
            // 
            // lblSulzerOrderNO
            // 
            this.lblSulzerOrderNO.AutoSize = true;
            this.lblSulzerOrderNO.Location = new System.Drawing.Point(864, 70);
            this.lblSulzerOrderNO.Name = "lblSulzerOrderNO";
            this.lblSulzerOrderNO.Size = new System.Drawing.Size(99, 15);
            this.lblSulzerOrderNO.TabIndex = 13;
            this.lblSulzerOrderNO.Text = "Sulzer Order NO";
            // 
            // lblCustomerPONO
            // 
            this.lblCustomerPONO.AutoSize = true;
            this.lblCustomerPONO.Location = new System.Drawing.Point(428, 70);
            this.lblCustomerPONO.Name = "lblCustomerPONO";
            this.lblCustomerPONO.Size = new System.Drawing.Size(103, 15);
            this.lblCustomerPONO.TabIndex = 12;
            this.lblCustomerPONO.Text = "Customer PO NO";
            // 
            // lblSite
            // 
            this.lblSite.AutoSize = true;
            this.lblSite.Location = new System.Drawing.Point(7, 70);
            this.lblSite.Name = "lblSite";
            this.lblSite.Size = new System.Drawing.Size(84, 15);
            this.lblSite.TabIndex = 11;
            this.lblSite.Text = "Site/安装地点";
            // 
            // lblProject
            // 
            this.lblProject.AutoSize = true;
            this.lblProject.Location = new System.Drawing.Point(864, 41);
            this.lblProject.Name = "lblProject";
            this.lblProject.Size = new System.Drawing.Size(77, 15);
            this.lblProject.TabIndex = 10;
            this.lblProject.Text = "Project/项目";
            // 
            // lblEndUser
            // 
            this.lblEndUser.AutoSize = true;
            this.lblEndUser.Location = new System.Drawing.Point(428, 41);
            this.lblEndUser.Name = "lblEndUser";
            this.lblEndUser.Size = new System.Drawing.Size(90, 15);
            this.lblEndUser.TabIndex = 9;
            this.lblEndUser.Text = "End User:/用户";
            // 
            // txtPurchaser
            // 
            this.txtPurchaser.Location = new System.Drawing.Point(129, 38);
            this.txtPurchaser.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtPurchaser.Name = "txtPurchaser";
            this.txtPurchaser.Size = new System.Drawing.Size(291, 21);
            this.txtPurchaser.TabIndex = 8;
            // 
            // lblPurchaser
            // 
            this.lblPurchaser.AutoSize = true;
            this.lblPurchaser.Location = new System.Drawing.Point(7, 41);
            this.lblPurchaser.Name = "lblPurchaser";
            this.lblPurchaser.Size = new System.Drawing.Size(96, 15);
            this.lblPurchaser.TabIndex = 7;
            this.lblPurchaser.Text = "Purchaser/买方";
            // 
            // lblXMLPath
            // 
            this.lblXMLPath.AutoSize = true;
            this.lblXMLPath.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.lblXMLPath.Location = new System.Drawing.Point(14, 59);
            this.lblXMLPath.Name = "lblXMLPath";
            this.lblXMLPath.Size = new System.Drawing.Size(83, 15);
            this.lblXMLPath.TabIndex = 20;
            this.lblXMLPath.Text = "XML File Path";
            // 
            // textXMLPath
            // 
            this.textXMLPath.Enabled = false;
            this.textXMLPath.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textXMLPath.Location = new System.Drawing.Point(199, 55);
            this.textXMLPath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textXMLPath.Name = "textXMLPath";
            this.textXMLPath.Size = new System.Drawing.Size(747, 21);
            this.textXMLPath.TabIndex = 21;
            // 
            // btnXMLPathOpen
            // 
            this.btnXMLPathOpen.Location = new System.Drawing.Point(965, 54);
            this.btnXMLPathOpen.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnXMLPathOpen.Name = "btnXMLPathOpen";
            this.btnXMLPathOpen.Size = new System.Drawing.Size(64, 26);
            this.btnXMLPathOpen.TabIndex = 22;
            this.btnXMLPathOpen.Text = "Open";
            this.btnXMLPathOpen.UseVisualStyleBackColor = true;
            this.btnXMLPathOpen.Click += new System.EventHandler(this.BtnXMLPathOpen_Click);
            // 
            // btnXMLDataImport
            // 
            this.btnXMLDataImport.Location = new System.Drawing.Point(1036, 54);
            this.btnXMLDataImport.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnXMLDataImport.Name = "btnXMLDataImport";
            this.btnXMLDataImport.Size = new System.Drawing.Size(159, 26);
            this.btnXMLDataImport.TabIndex = 23;
            this.btnXMLDataImport.Text = "XML Data Import";
            this.btnXMLDataImport.UseVisualStyleBackColor = true;
            this.btnXMLDataImport.Click += new System.EventHandler(this.BtnXMLDataImport_Click);
            // 
            // ColItemNo
            // 
            this.ColItemNo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.ColItemNo.Frozen = true;
            this.ColItemNo.HeaderText = "ITEM No/行号";
            this.ColItemNo.MinimumWidth = 100;
            this.ColItemNo.Name = "ColItemNo";
            this.ColItemNo.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.ColItemNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColSeriesNo
            // 
            this.ColSeriesNo.Frozen = true;
            this.ColSeriesNo.HeaderText = "Series No/苏尔寿编号";
            this.ColSeriesNo.MinimumWidth = 150;
            this.ColSeriesNo.Name = "ColSeriesNo";
            this.ColSeriesNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColSeriesNo.Width = 150;
            // 
            // ColTagNo
            // 
            this.ColTagNo.Frozen = true;
            this.ColTagNo.HeaderText = "TAG NO/客户位号";
            this.ColTagNo.MinimumWidth = 100;
            this.ColTagNo.Name = "ColTagNo";
            this.ColTagNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColType
            // 
            this.ColType.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColType.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColType.Frozen = true;
            this.ColType.HeaderText = "TYPE/泵型";
            this.ColType.MinimumWidth = 100;
            this.ColType.Name = "ColType";
            this.ColType.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColSize
            // 
            this.ColSize.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColSize.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColSize.Frozen = true;
            this.ColSize.HeaderText = "SIZE/规格";
            this.ColSize.MinimumWidth = 150;
            this.ColSize.Name = "ColSize";
            this.ColSize.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColSize.Width = 150;
            // 
            // ColPREPumpPressureClass
            // 
            this.ColPREPumpPressureClass.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColPREPumpPressureClass.HeaderText = "PRE泵压力等级";
            this.ColPREPumpPressureClass.MinimumWidth = 150;
            this.ColPREPumpPressureClass.Name = "ColPREPumpPressureClass";
            this.ColPREPumpPressureClass.Width = 150;
            // 
            // ColService
            // 
            this.ColService.HeaderText = "SERVICE/泵名称";
            this.ColService.MinimumWidth = 100;
            this.ColService.Name = "ColService";
            this.ColService.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColQty
            // 
            this.ColQty.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.ColQty.DefaultCellStyle = dataGridViewCellStyle2;
            this.ColQty.HeaderText = "Qty/数量";
            this.ColQty.MinimumWidth = 100;
            this.ColQty.Name = "ColQty";
            this.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColMotorManufacturer
            // 
            this.ColMotorManufacturer.HeaderText = "MOTOR MANUFACTURER/电机供应商";
            this.ColMotorManufacturer.MinimumWidth = 120;
            this.ColMotorManufacturer.Name = "ColMotorManufacturer";
            this.ColMotorManufacturer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColMotorManufacturer.Width = 120;
            // 
            // ColMotorFrame
            // 
            this.ColMotorFrame.HeaderText = "MOTOR  FRAME/电机型号";
            this.ColMotorFrame.MinimumWidth = 100;
            this.ColMotorFrame.Name = "ColMotorFrame";
            this.ColMotorFrame.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColFrameSize
            // 
            this.ColFrameSize.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColFrameSize.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColFrameSize.HeaderText = "FRAME SIZE/基座号";
            this.ColFrameSize.MinimumWidth = 100;
            this.ColFrameSize.Name = "ColFrameSize";
            this.ColFrameSize.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColMotorPower
            // 
            this.ColMotorPower.HeaderText = "MOTOR POWER/电机功率(kW) ";
            this.ColMotorPower.MinimumWidth = 100;
            this.ColMotorPower.Name = "ColMotorPower";
            this.ColMotorPower.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColVoltage
            // 
            this.ColVoltage.HeaderText = "VOLTAGE/电压(V) ";
            this.ColVoltage.MinimumWidth = 100;
            this.ColVoltage.Name = "ColVoltage";
            this.ColVoltage.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColSealFlushFlow
            // 
            this.ColSealFlushFlow.HeaderText = "SEAL FLUSH FLOW/密封冲洗量";
            this.ColSealFlushFlow.MinimumWidth = 150;
            this.ColSealFlushFlow.Name = "ColSealFlushFlow";
            this.ColSealFlushFlow.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColSealFlushFlow.Width = 150;
            // 
            // ColCoolingWaterFlow
            // 
            this.ColCoolingWaterFlow.HeaderText = "COOLING WATER FLOW/外接冷却水冲洗量";
            this.ColCoolingWaterFlow.MinimumWidth = 170;
            this.ColCoolingWaterFlow.Name = "ColCoolingWaterFlow";
            this.ColCoolingWaterFlow.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColCoolingWaterFlow.Width = 170;
            // 
            // ColSpeed
            // 
            this.ColSpeed.HeaderText = "转数(r/min)";
            this.ColSpeed.MinimumWidth = 100;
            this.ColSpeed.Name = "ColSpeed";
            this.ColSpeed.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColMotorLengthK
            // 
            this.ColMotorLengthK.HeaderText = "MOTOR LENGTH K/电机长度K";
            this.ColMotorLengthK.MinimumWidth = 100;
            this.ColMotorLengthK.Name = "ColMotorLengthK";
            this.ColMotorLengthK.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColApiEdition
            // 
            this.ColApiEdition.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColApiEdition.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColApiEdition.HeaderText = "API 682 EDITION/API 682 版本";
            this.ColApiEdition.MinimumWidth = 100;
            this.ColApiEdition.Name = "ColApiEdition";
            this.ColApiEdition.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColSealPlanA
            // 
            this.ColSealPlanA.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColSealPlanA.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColSealPlanA.HeaderText = "SEAL PLAN A/冲洗方案 A";
            this.ColSealPlanA.MinimumWidth = 100;
            this.ColSealPlanA.Name = "ColSealPlanA";
            this.ColSealPlanA.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColSealPlanB
            // 
            this.ColSealPlanB.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColSealPlanB.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColSealPlanB.HeaderText = "SEAL PLAN B/冲洗方案";
            this.ColSealPlanB.MinimumWidth = 100;
            this.ColSealPlanB.Name = "ColSealPlanB";
            this.ColSealPlanB.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColSealPlanC
            // 
            this.ColSealPlanC.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColSealPlanC.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColSealPlanC.HeaderText = "SEAL PLAN C/冲洗方案";
            this.ColSealPlanC.MinimumWidth = 100;
            this.ColSealPlanC.Name = "ColSealPlanC";
            this.ColSealPlanC.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColPumpDrainValve
            // 
            this.ColPumpDrainValve.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColPumpDrainValve.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColPumpDrainValve.HeaderText = "WI/WO PUMP DRAIN VALVE/排残管路阀门";
            this.ColPumpDrainValve.MinimumWidth = 150;
            this.ColPumpDrainValve.Name = "ColPumpDrainValve";
            this.ColPumpDrainValve.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColPumpDrainValve.Width = 150;
            // 
            // ColWarmUpPiping
            // 
            this.ColWarmUpPiping.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColWarmUpPiping.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColWarmUpPiping.HeaderText = "WI/WO WARM-UP PIPING/暖泵管路";
            this.ColWarmUpPiping.MinimumWidth = 150;
            this.ColWarmUpPiping.Name = "ColWarmUpPiping";
            this.ColWarmUpPiping.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColWarmUpPiping.Width = 150;
            // 
            // ColBearing
            // 
            this.ColBearing.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColBearing.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColBearing.HeaderText = "BEARING/轴承配置";
            this.ColBearing.MinimumWidth = 100;
            this.ColBearing.Name = "ColBearing";
            this.ColBearing.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColCompanionFlg
            // 
            this.ColCompanionFlg.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColCompanionFlg.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColCompanionFlg.HeaderText = "COMPANION FLG./配对法兰要求";
            this.ColCompanionFlg.MinimumWidth = 100;
            this.ColCompanionFlg.Name = "ColCompanionFlg";
            this.ColCompanionFlg.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColCaseBracket
            // 
            this.ColCaseBracket.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColCaseBracket.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColCaseBracket.HeaderText = "CASE  WI/WO BRACKET/壳体是否带夹套";
            this.ColCaseBracket.MinimumWidth = 110;
            this.ColCaseBracket.Name = "ColCaseBracket";
            this.ColCaseBracket.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColCaseBracket.Width = 110;
            // 
            // ColBearighousingDrain
            // 
            this.ColBearighousingDrain.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColBearighousingDrain.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColBearighousingDrain.HeaderText = "BEARING HOUSING WI/WO OIL DRAIN/轴承室带排油管路";
            this.ColBearighousingDrain.MinimumWidth = 130;
            this.ColBearighousingDrain.Name = "ColBearighousingDrain";
            this.ColBearighousingDrain.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColBearighousingDrain.Width = 130;
            // 
            // ColFlow
            // 
            this.ColFlow.HeaderText = "FLOW/流量 (m3/h)";
            this.ColFlow.MinimumWidth = 50;
            this.ColFlow.Name = "ColFlow";
            this.ColFlow.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColFlow.Width = 50;
            // 
            // ColHead
            // 
            this.ColHead.HeaderText = "HEAD/扬程 (m)";
            this.ColHead.MinimumWidth = 50;
            this.ColHead.Name = "ColHead";
            this.ColHead.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColHead.Width = 50;
            // 
            // ColEfficiency
            // 
            this.ColEfficiency.HeaderText = "EFFICIENCY/效率(%)";
            this.ColEfficiency.MinimumWidth = 100;
            this.ColEfficiency.Name = "ColEfficiency";
            this.ColEfficiency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColRatedPower
            // 
            this.ColRatedPower.HeaderText = "RATED POWER/额定功率(kW)";
            this.ColRatedPower.MinimumWidth = 100;
            this.ColRatedPower.Name = "ColRatedPower";
            this.ColRatedPower.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColNpshr
            // 
            this.ColNpshr.HeaderText = "NPSHr(m)";
            this.ColNpshr.MinimumWidth = 65;
            this.ColNpshr.Name = "ColNpshr";
            this.ColNpshr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColNpshr.Width = 65;
            // 
            // ColMawp
            // 
            this.ColMawp.HeaderText = "MAWP(Mpa)";
            this.ColMawp.MinimumWidth = 75;
            this.ColMawp.Name = "ColMawp";
            this.ColMawp.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColMawp.Width = 75;
            // 
            // ColMawt
            // 
            this.ColMawt.HeaderText = "MAWT(℃)";
            this.ColMawt.MinimumWidth = 80;
            this.ColMawt.Name = "ColMawt";
            this.ColMawt.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColMawt.Width = 80;
            // 
            // ColHydrostaticP
            // 
            this.ColHydrostaticP.HeaderText = "Hydrostatic P(Mpa)";
            this.ColHydrostaticP.MinimumWidth = 100;
            this.ColHydrostaticP.Name = "ColHydrostaticP";
            this.ColHydrostaticP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColHydrostaticP.Width = 107;
            // 
            // ColMaterialClass
            // 
            this.ColMaterialClass.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColMaterialClass.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColMaterialClass.HeaderText = "MATERIAL CLASS/材料等级";
            this.ColMaterialClass.MinimumWidth = 100;
            this.ColMaterialClass.Name = "ColMaterialClass";
            this.ColMaterialClass.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColBearingCageMaterial
            // 
            this.ColBearingCageMaterial.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColBearingCageMaterial.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColBearingCageMaterial.HeaderText = "BEARING CAGE MATERIAL/轴承保持架材质";
            this.ColBearingCageMaterial.MinimumWidth = 150;
            this.ColBearingCageMaterial.Name = "ColBearingCageMaterial";
            this.ColBearingCageMaterial.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColBearingCageMaterial.Width = 150;
            // 
            // ColPumpDrainNozzleSize
            // 
            this.ColPumpDrainNozzleSize.HeaderText = "PUMP DRAIN NOZZLE SIZE/排残管口口径";
            this.ColPumpDrainNozzleSize.MinimumWidth = 100;
            this.ColPumpDrainNozzleSize.Name = "ColPumpDrainNozzleSize";
            this.ColPumpDrainNozzleSize.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColDryWetShaft
            // 
            this.ColDryWetShaft.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColDryWetShaft.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColDryWetShaft.HeaderText = "DRY/WET SHAFT/干湿轴";
            this.ColDryWetShaft.MinimumWidth = 150;
            this.ColDryWetShaft.Name = "ColDryWetShaft";
            this.ColDryWetShaft.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColDryWetShaft.Width = 150;
            // 
            // ColWetInput
            // 
            this.ColWetInput.HeaderText = "Wet Shaft(自己填写)";
            this.ColWetInput.MinimumWidth = 100;
            this.ColWetInput.Name = "ColWetInput";
            this.ColWetInput.ReadOnly = true;
            // 
            // ColNozzleLoad
            // 
            this.ColNozzleLoad.HeaderText = "NOZZLE LOAD/管口载荷要求";
            this.ColNozzleLoad.MinimumWidth = 100;
            this.ColNozzleLoad.Name = "ColNozzleLoad";
            this.ColNozzleLoad.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColReservedPureLubrication
            // 
            this.ColReservedPureLubrication.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColReservedPureLubrication.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColReservedPureLubrication.HeaderText = "WI/WO RESERVED PURE OIL-MIST LUBRICATION/预留油雾润滑";
            this.ColReservedPureLubrication.MinimumWidth = 150;
            this.ColReservedPureLubrication.Name = "ColReservedPureLubrication";
            this.ColReservedPureLubrication.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColReservedPureLubrication.Width = 150;
            // 
            // ColImpellerBalanceHole
            // 
            this.ColImpellerBalanceHole.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColImpellerBalanceHole.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColImpellerBalanceHole.HeaderText = "IMPELLER WI/WO BALANCE HOLE/叶轮带平衡孔";
            this.ColImpellerBalanceHole.MinimumWidth = 120;
            this.ColImpellerBalanceHole.Name = "ColImpellerBalanceHole";
            this.ColImpellerBalanceHole.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColImpellerBalanceHole.Width = 120;
            // 
            // ColPREInducerYesNo
            // 
            this.ColPREInducerYesNo.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColPREInducerYesNo.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColPREInducerYesNo.HeaderText = "是否带诱导轮";
            this.ColPREInducerYesNo.MinimumWidth = 100;
            this.ColPREInducerYesNo.Name = "ColPREInducerYesNo";
            this.ColPREInducerYesNo.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // ColCoolWaterMainPipeSize
            // 
            this.ColCoolWaterMainPipeSize.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColCoolWaterMainPipeSize.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColCoolWaterMainPipeSize.HeaderText = "冷却水母管口径";
            this.ColCoolWaterMainPipeSize.MinimumWidth = 100;
            this.ColCoolWaterMainPipeSize.Name = "ColCoolWaterMainPipeSize";
            // 
            // ColDrainValveType
            // 
            this.ColDrainValveType.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ColDrainValveType.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.ColDrainValveType.HeaderText = "排残阀类型";
            this.ColDrainValveType.MinimumWidth = 100;
            this.ColDrainValveType.Name = "ColDrainValveType";
            // 
            // ColVoluteCasing
            // 
            this.ColVoluteCasing.HeaderText = "泵体/泵盖";
            this.ColVoluteCasing.Name = "ColVoluteCasing";
            this.ColVoluteCasing.ReadOnly = true;
            this.ColVoluteCasing.Visible = false;
            // 
            // ColImpeller
            // 
            this.ColImpeller.HeaderText = "叶轮";
            this.ColImpeller.Name = "ColImpeller";
            this.ColImpeller.ReadOnly = true;
            this.ColImpeller.Visible = false;
            // 
            // ColStationaryWearRing
            // 
            this.ColStationaryWearRing.HeaderText = "口环";
            this.ColStationaryWearRing.Name = "ColStationaryWearRing";
            this.ColStationaryWearRing.ReadOnly = true;
            this.ColStationaryWearRing.Visible = false;
            // 
            // ColCoverCoolingPlate
            // 
            this.ColCoverCoolingPlate.HeaderText = "泵盖水冷密封板";
            this.ColCoverCoolingPlate.Name = "ColCoverCoolingPlate";
            this.ColCoverCoolingPlate.ReadOnly = true;
            this.ColCoverCoolingPlate.Visible = false;
            // 
            // ColThroatBush
            // 
            this.ColThroatBush.HeaderText = "喉部衬套";
            this.ColThroatBush.Name = "ColThroatBush";
            this.ColThroatBush.ReadOnly = true;
            this.ColThroatBush.Visible = false;
            // 
            // ColImpellerKey
            // 
            this.ColImpellerKey.HeaderText = "叶轮键";
            this.ColImpellerKey.Name = "ColImpellerKey";
            this.ColImpellerKey.ReadOnly = true;
            this.ColImpellerKey.Visible = false;
            // 
            // ColGasket
            // 
            this.ColGasket.HeaderText = "泵体泵盖之间的密封垫";
            this.ColGasket.Name = "ColGasket";
            this.ColGasket.ReadOnly = true;
            this.ColGasket.Visible = false;
            // 
            // ColSpacerSleeve
            // 
            this.ColSpacerSleeve.HeaderText = "轴套";
            this.ColSpacerSleeve.Name = "ColSpacerSleeve";
            this.ColSpacerSleeve.ReadOnly = true;
            this.ColSpacerSleeve.Visible = false;
            // 
            // ColImpellerNut
            // 
            this.ColImpellerNut.HeaderText = "叶轮螺母";
            this.ColImpellerNut.Name = "ColImpellerNut";
            this.ColImpellerNut.ReadOnly = true;
            this.ColImpellerNut.Visible = false;
            // 
            // ColInducer
            // 
            this.ColInducer.HeaderText = "诱导轮";
            this.ColInducer.Name = "ColInducer";
            this.ColInducer.ReadOnly = true;
            this.ColInducer.Visible = false;
            // 
            // ColInducerSleeve
            // 
            this.ColInducerSleeve.HeaderText = "诱导轮衬套";
            this.ColInducerSleeve.Name = "ColInducerSleeve";
            this.ColInducerSleeve.ReadOnly = true;
            this.ColInducerSleeve.Visible = false;
            // 
            // ColWearPlate
            // 
            this.ColWearPlate.HeaderText = "耐磨板";
            this.ColWearPlate.Name = "ColWearPlate";
            this.ColWearPlate.ReadOnly = true;
            this.ColWearPlate.Visible = false;
            // 
            // SumListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(88)))), ((int)(((byte)(162)))));
            this.ClientSize = new System.Drawing.Size(1398, 848);
            this.Controls.Add(this.btnXMLDataImport);
            this.Controls.Add(this.btnXMLPathOpen);
            this.Controls.Add(this.textXMLPath);
            this.Controls.Add(this.lblXMLPath);
            this.Controls.Add(this.palPublicData);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnXMLFilePathOpen);
            this.Controls.Add(this.btnSumListFilePathOpen);
            this.Controls.Add(this.lblInputDataFileSavePath);
            this.Controls.Add(this.lblSumListFilePath);
            this.Controls.Add(this.txtXMLFilePath);
            this.Controls.Add(this.txtSumListFilePath);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnSumListDataImport);
            this.Controls.Add(this.dgvSumList);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SumListForm";
            this.Text = "Basic Data Entry";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SumListForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSumList)).EndInit();
            this.contextMenuStripGDV.ResumeLayout(false);
            this.palPublicData.ResumeLayout(false);
            this.palPublicData.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox cbPGA;
        private System.Windows.Forms.CheckBox cbPPI;
        private System.Windows.Forms.CheckBox cbPCS;
        private System.Windows.Forms.CheckBox cbPNT;
        private System.Windows.Forms.CheckBox cbPUC;
        private System.Windows.Forms.CheckBox cbPDS;
        private System.Windows.Forms.CheckBox cbPSL;
        private System.Windows.Forms.DataGridView dgvSumList;
        private System.Windows.Forms.Button btnSumListDataImport;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.TextBox txtSumListFilePath;
        private System.Windows.Forms.TextBox txtXMLFilePath;
        private System.Windows.Forms.Label lblSumListFilePath;
        private System.Windows.Forms.Label lblInputDataFileSavePath;
        private System.Windows.Forms.Button btnSumListFilePathOpen;
        private System.Windows.Forms.Button btnXMLFilePathOpen;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Panel palPublicData;
        private System.Windows.Forms.TextBox txtApprovedBy;
        private System.Windows.Forms.TextBox txtReviewedBy;
        private System.Windows.Forms.TextBox txtPreparedBy;
        private System.Windows.Forms.TextBox txtSulzerOrderNO;
        private System.Windows.Forms.TextBox txtCustomerPONO;
        private System.Windows.Forms.TextBox txtSite;
        private System.Windows.Forms.TextBox txtProject;
        private System.Windows.Forms.TextBox txtEndUser;
        private System.Windows.Forms.Label lblFlangeStandard;
        private System.Windows.Forms.Label lblWorkDate;
        private System.Windows.Forms.Label lblApprovedBy;
        private System.Windows.Forms.Label lblReviewedBy;
        private System.Windows.Forms.Label lblPreparedBy;
        private System.Windows.Forms.Label lblSulzerOrderNO;
        private System.Windows.Forms.Label lblCustomerPONO;
        private System.Windows.Forms.Label lblSite;
        private System.Windows.Forms.Label lblProject;
        private System.Windows.Forms.Label lblEndUser;
        private System.Windows.Forms.TextBox txtPurchaser;
        private System.Windows.Forms.Label lblPurchaser;
        private System.Windows.Forms.ComboBox cobFlangeStandard;
        private System.Windows.Forms.Label lblXMLPath;
        private System.Windows.Forms.TextBox textXMLPath;
        private System.Windows.Forms.Button btnXMLPathOpen;
        private System.Windows.Forms.Button btnXMLDataImport;
        private System.Windows.Forms.DateTimePicker dateTimeWorkDate;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripGDV;
        private System.Windows.Forms.ToolStripMenuItem DeleteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AddToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem PasteToolStripMenuItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColItemNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSeriesNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColTagNo;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColType;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColSize;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColPREPumpPressureClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColService;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColQty;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMotorManufacturer;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMotorFrame;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColFrameSize;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMotorPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColVoltage;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSealFlushFlow;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColCoolingWaterFlow;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSpeed;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMotorLengthK;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColApiEdition;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColSealPlanA;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColSealPlanB;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColSealPlanC;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColPumpDrainValve;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColWarmUpPiping;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColBearing;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColCompanionFlg;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColCaseBracket;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColBearighousingDrain;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColFlow;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHead;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColEfficiency;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColRatedPower;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColNpshr;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMawp;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColMawt;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColHydrostaticP;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColMaterialClass;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColBearingCageMaterial;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColPumpDrainNozzleSize;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColDryWetShaft;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColWetInput;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColNozzleLoad;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColReservedPureLubrication;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColImpellerBalanceHole;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColPREInducerYesNo;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColCoolWaterMainPipeSize;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColDrainValveType;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColVoluteCasing;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColImpeller;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColStationaryWearRing;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColCoverCoolingPlate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColThroatBush;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColImpellerKey;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColGasket;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSpacerSleeve;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColImpellerNut;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColInducer;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColInducerSleeve;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColWearPlate;
    }
}