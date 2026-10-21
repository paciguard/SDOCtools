using AcAdAmComm;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SESDTools
{
    /// <summary>
    /// File Name       ：SumListForm
    /// Description     ：Data modification screen
    /// Author          ：SY
    /// Create Date     ：2021/07/23 10:43
    /// Modifier        ：
    /// Modifier Date   ：
    /// Revision History：
    /// Version         ：V1.0
    /// </summary>
    public partial class SumListForm : Form
    {
        /// <summary>
        /// Combobox global variable
        /// </summary>
        public DataSet COMBOBOX_DATASAT;
        /// <summary>
        /// data has been saved before the screen is closed
        /// </summary>
        public int CLOSE_FLAG = 0;
        public SumListForm()
        {
            //Center alignment
            this.StartPosition = FormStartPosition.CenterScreen;
            //initialization
            InitializeComponent();
            //Combobox initialization
            InitializeCombobox();
            //datagidview cell name initialization
            InitializeHeaderCellName();
            //datagidview DataError initialization
            this.dgvSumList.DataError += delegate (object sender, DataGridViewDataErrorEventArgs e) { };
        }
        /// <summary>
        /// Combobox Initialization assignment
        /// </summary>
        private void InitializeCombobox()
        {
            //Combobox data file path
            string nodeNamePassword = "xmlDataFilePath";
            string fileName = "COMBOBOXDATATABLE.xml";
            DataSet dataSetCombobox = new DataSet();
            try
            {
                string xmlDataFilePath = XMLHelper.ReadXmlDataFilePathSettings(nodeNamePassword) + "\\" + fileName;
                //Combobox data get
                dataSetCombobox = XmlDataOperation.XMLFileToDataSet(xmlDataFilePath);
            }
            catch(Exception ex)
            {
                //Log write
                LogHelper.GetSingleObj().Error("数据修改画面初始化", ex.Message);
                MessageBox.Show(ex.Message, "ERROR");
                this.Close();
            }
            //Combobox global variable assignment
            COMBOBOX_DATASAT = dataSetCombobox;
        }
        /// <summary>
        /// Datagidview cell name initialization
        /// </summary>
        private void InitializeHeaderCellName()
        {
            dgvSumList.ColumnHeadersHeight = 60;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].HeaderCell.Value = "ITEM No\n行号";
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].HeaderCell.Value = "Series No\n苏尔寿编号";
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].HeaderCell.Value = "TAG NO\n客户位号";
            //TYPE/泵型
            dgvSumList.Columns["ColType"].HeaderCell.Value = "TYPE\n泵型";
            //SIZE/规格
            dgvSumList.Columns["ColSize"].HeaderCell.Value = "SIZE\n规格";
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].HeaderCell.Value = "SERVICE\n泵名称";
            //Qty / 数量
            dgvSumList.Columns["ColQty"].HeaderCell.Value = "Qty\n数量";
            //MOTOR MANUFACTURER / 电机供应商
            dgvSumList.Columns["ColMotorManufacturer"].HeaderCell.Value = "MOTOR MANUFACTURER\n电机供应商";
            //电机型号
            dgvSumList.Columns["ColMotorFrame"].HeaderCell.Value = "MOTOR  FRAME\n电机型号";
            //FRAME SIZE/基座号
            dgvSumList.Columns["ColFrameSize"].HeaderCell.Value = "FRAME SIZE\n基座号";
            //MOTOR POWER/ 电机功率(kW) 
            dgvSumList.Columns["ColMotorPower"].HeaderCell.Value = "MOTOR POWER\n电机功率(kW) ";
            //VOLTAGE/ 电压(V) 
            dgvSumList.Columns["ColVoltage"].HeaderCell.Value = "VOLTAGE\n电压(V) ";
            //转数 r/min
            dgvSumList.Columns["ColSpeed"].HeaderCell.Value = "SPEED\n转数(r/min)";
            //电机长度K
            dgvSumList.Columns["ColMotorLengthK"].HeaderCell.Value = "MOTOR LENGTH K\n电机长度K";
            //API 682 EDITION / API 682 版本
            dgvSumList.Columns["ColApiEdition"].HeaderCell.Value = "API 682 EDITION\nAPI 682 版本";
            //SEAL PLAN A/ 冲洗方案
            dgvSumList.Columns["ColSealPlanA"].HeaderCell.Value = "SEAL PLAN A\n冲洗方案 A";
            //SEAL PLAN B/ 冲洗方案
            dgvSumList.Columns["ColSealPlanB"].HeaderCell.Value = "SEAL PLAN B\n冲洗方案 B";
            //SEAL PLAN C/ 冲洗方案
            dgvSumList.Columns["ColSealPlanC"].HeaderCell.Value = "SEAL PLAN C\n冲洗方案 C";
            //FLOW / 流量 (m3/h)
            dgvSumList.Columns["ColFlow"].HeaderCell.Value = "FLOW\n流量 (m3/h)";
            //HEAD / 扬程 (m)
            dgvSumList.Columns["ColHead"].HeaderCell.Value = "HEAD\n扬程 (m)";
            //EFFICIENCY / 效率(%)
            dgvSumList.Columns["ColEfficiency"].HeaderCell.Value = "EFFICIENCY\n效率(%)";
            //RATED POWER / 额定功率(kW)
            dgvSumList.Columns["ColRatedPower"].HeaderCell.Value = "RATED POWER\n额定功率(kW)";
            //NPSHr  (m)
            dgvSumList.Columns["ColNpshr"].HeaderCell.Value = "NPSHr(m)";
            //MAWP (Mpa)
            dgvSumList.Columns["ColMawp"].HeaderCell.Value = "MAWP(Mpa)";
            //MAWT (℃)
            dgvSumList.Columns["ColMawt"].HeaderCell.Value = "MAWT(℃)";
            //Hydrostatic P (Mpa)
            dgvSumList.Columns["ColHydrostaticP"].HeaderCell.Value = "Hydrostatic P(Mpa)";
            //排残管口口径
            dgvSumList.Columns["ColPumpDrainNozzleSize"].HeaderCell.Value = "PUMP DRAIN NOZZLE SIZE\n排残管口口径";
            //排残管路 阀门
            dgvSumList.Columns["ColPumpDrainValve"].HeaderCell.Value = "WI/WO PUMP DRAIN VALVE\n排残管路阀门";
            //暖泵管路
            dgvSumList.Columns["ColWarmUpPiping"].HeaderCell.Value = "WI/WO WARM-UP PIPING\n暖泵管路";
            //MATERIAL CLASS / 材料等级
            dgvSumList.Columns["ColMaterialClass"].HeaderCell.Value = "MATERIAL CLASS\n材料等级";
            //干湿轴
            dgvSumList.Columns["ColDryWetShaft"].HeaderCell.Value = "DRY/WET SHAFT\n干湿轴";
            //Wet Shaft(自己填写)
            dgvSumList.Columns["ColWetInput"].HeaderCell.Value = "Wet Shaft\n(自己填写)";
            //BEARING / 轴承配置
            dgvSumList.Columns["ColBearing"].HeaderCell.Value = "BEARING\n轴承配置";
            //轴承保持架材质
            dgvSumList.Columns["ColBearingCageMaterial"].HeaderCell.Value = "BEARING CAGE MATERIAL\n轴承保持架材质";
            //COMPANION FLG. / 配对法兰要求
            dgvSumList.Columns["ColCompanionFlg"].HeaderCell.Value = "COMPANION FLG.\n配对法兰要求";
            //壳体是否带夹套
            dgvSumList.Columns["ColCaseBracket"].HeaderCell.Value = "CASE  WI/WO BRACKET\n壳体是否带夹套";
            //管口载荷要求
            dgvSumList.Columns["ColNozzleLoad"].HeaderCell.Value = "NOZZLE LOAD\n管口载荷要求";
            //预留油雾润滑
            dgvSumList.Columns["ColReservedPureLubrication"].HeaderCell.Value = "WI/WO RESERVED PURE OIL-MIST LUBRICATION\n预留油雾润滑";
            //轴承室带排油管路
            dgvSumList.Columns["ColBearighousingDrain"].HeaderCell.Value = "BEARING HOUSING WI/WO OIL DRAIN\n轴承室带排油管路";
            //叶轮带平衡孔
            dgvSumList.Columns["ColImpellerBalanceHole"].HeaderCell.Value = "IMPELLER WI/WO BALANCE HOLE\n叶轮带平衡孔";
            //密封冲洗量
            dgvSumList.Columns["ColSealFlushFlow"].HeaderCell.Value = "SEAL FLUSH FLOW\n密封冲洗量L/min";
            //外接冷却水冲洗量
            dgvSumList.Columns["ColCoolingWaterFlow"].HeaderCell.Value = "COOLING WATER FLOW\n外接冷却水冲洗量L/min";
        }
        /// <summary>
        /// SumList file path selection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSumListFilePathOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "Please select a file";
            dialog.Filter = "Excel File|*.xlsx*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = dialog.FileName;
                txtSumListFilePath.Text = savePath;
            }
        }
        /// <summary>
        /// XML file path selection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnXMLFilePathOpen_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "Please select the input data folder path";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = dialog.SelectedPath;
                txtXMLFilePath.Text = savePath;
            }
        }
        /// <summary>
        /// SumList file import
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSumListDataImport_Click(object sender, EventArgs e)
        {
            //Get the corresponding line from the sumList file
            var strCellNamelist = new List<string>() {"S13","S14","S15","S16","S17", "S18", "S19", "S20","S22", "S23",
                "S24", "S25", "S27", "S28","S29", "S30","S31","S32","S33","S34", "S35","S37","S41","S43", "S45"};
            if (txtSumListFilePath.Text == "")
            {
                //Log write
                LogHelper.GetSingleObj().Error("SumList文件导入", "Please select the excel file!");
                MessageBox.Show("Please select the excel file!", "ERROR");
                return;
            }
            try
            {
                //Button not available
                BtnToEnabled(false);
                //EXCEL file path
                string filePath = txtSumListFilePath.Text;
                //EXCEL file to DataSet
                DataSet dataSet = ExcelDataOperation.SumListExcelToDataSat(filePath, strCellNamelist);
                DataTable dtHead = dataSet.Tables["DataTableHead"];
                DataTable dtBody = dataSet.Tables["DataTableBody"];
                textXMLPath.Clear();
                txtXMLFilePath.Clear();
                //all data clear
                ClearAllControls();
                //frame data binding
                DataBindingFromSumlist(dtHead, dtBody);
            }
            catch(Exception ex)
            {
                //Log write
                LogHelper.GetSingleObj().Error("SumList文件导入", ex.Message);
                //Button available
                BtnToEnabled(true);
                MessageBox.Show("SumList文件导入:"+ ex.Message, "ERROR");
                return;
            }
            //Cursor release
            dgvSumList.ClearSelection();
            //Log write
            LogHelper.GetSingleObj().Info("SumList文件导入", "SumList文件导入成功");
            //Button available
            BtnToEnabled(true);
            MessageBox.Show("Read the sumlist file successfully!");
        }
        /// <summary>
        /// SumList文件读取后绑定数据
        /// </summary>
        /// <param name="dtHead"></param>
        /// <param name="dtBody"></param>
        private void DataBindingFromSumlist(DataTable dtHead, DataTable dtBody)
        {
            //FLANGE STANDARD / 法兰标准
            //get the first record
            string strFlangeStandard = "";
            //DataGridView Combox initialization
            DgvComboxInitialSettings();
            //head assignment
            for (int i = 0; i < dtHead.Rows.Count; i++)
            {
                //Purchaser/买方
                txtPurchaser.Text = dtHead.Rows[i]["S2"].ToString();
                //End User:/用户
                txtEndUser.Text = dtHead.Rows[i]["S3"].ToString();
                //Project/项目
                txtProject.Text = dtHead.Rows[i]["S4"].ToString();
                //txtSite
                txtSite.Text = dtHead.Rows[i]["S5"].ToString();
                //Customer PO NO
                txtCustomerPONO.Text = dtHead.Rows[i]["S6"].ToString();
                //Sulzer Order NO
                txtSulzerOrderNO.Text = dtHead.Rows[i]["S7"].ToString();
            }
            //dgvSumList assignment
            for (int i = 0; i < dtBody.Rows.Count; i++)
            {
                dgvSumList.Rows.Add();
                //行号
                dgvSumList.Rows[i].HeaderCell.Value = string.Format("{0}", i + 1);
                dgvSumList.Rows[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //ITEM No/行号
                dgvSumList.Rows[i].Cells["ColItemNo"].Value = dtBody.Rows[i]["S13"].ToString();
                //Series No/苏尔寿编号
                dgvSumList.Rows[i].Cells["ColSeriesNo"].Value = dtBody.Rows[i]["S14"].ToString();
                //TAG NO/客户位号
                dgvSumList.Rows[i].Cells["ColTagNo"].Value = dtBody.Rows[i]["S15"].ToString();
                //TYPE/泵型ComboBox 赋值 
                dgvSumList.Rows[i].Cells["ColType"].Value = dtBody.Rows[i]["S16"].ToString();
                //依据（TYPE/泵型）ComboBox initialization
                SelecteSizeComboBox(dtBody.Rows[i]["S16"].ToString(), i);
                if (!dtBody.Rows[i]["S16"].ToString().Contains("PR"))
                {
                    //PRE泵压力等级
                    dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                    dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].Value = "";
                    //是否带诱导轮
                    dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].ReadOnly = true;
                    dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].Value = "";
                }
                else
                {
                    //PRE泵压力等级
                    dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].ReadOnly = false;
                    //是否带诱导轮
                    dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].ReadOnly = false;
                }
                if (dtBody.Rows[i]["S16"].ToString().Contains("ZA"))
                {
                    //管口载荷要求
                    dgvSumList.Rows[i].Cells["ColNozzleLoad"].Value = "1";
                }
                else
                {
                    //管口载荷要求
                    dgvSumList.Rows[i].Cells["ColNozzleLoad"].Value = "2";
                }
                if (dtBody.Rows[i]["S16"].ToString().Equals("ZAG") || dtBody.Rows[i]["S16"].ToString().Equals("ZEG"))
                {
                    //壳体是否带夹套
                    dgvSumList.Rows[i].Cells["ColCaseBracket"].Value = true;
                }
                //SIZE/规格 
                dgvSumList.Rows[i].Cells["ColSize"].Value = dtBody.Rows[i]["S17"].ToString();
                //SERVICE / 泵名称
                dgvSumList.Rows[i].Cells["ColService"].Value = dtBody.Rows[i]["S18"].ToString();
                //Qty / 数量
                dgvSumList.Rows[i].Cells["ColQty"].Value = dtBody.Rows[i]["S19"].ToString();
                //MOTOR MANUFACTURER / 电机供应商
                dgvSumList.Rows[i].Cells["ColMotorManufacturer"].Value = dtBody.Rows[i]["S20"].ToString();
                //电机型号
                dgvSumList.Rows[i].Cells["ColMotorFrame"].Value = "";
                //FRAME SIZE/基座号
                dgvSumList.Rows[i].Cells["ColFrameSize"].Value = dtBody.Rows[i]["S22"].ToString();
                //MOTOR POWER/ 电机功率(kW) 
                dgvSumList.Rows[i].Cells["ColMotorPower"].Value = dtBody.Rows[i]["S23"].ToString();
                //VOLTAGE/ 电压(V) 
                dgvSumList.Rows[i].Cells["ColVoltage"].Value = dtBody.Rows[i]["S24"].ToString();
                //转数 r/min
                dgvSumList.Rows[i].Cells["ColSpeed"].Value = dtBody.Rows[i]["S25"].ToString();
                //电机长度K
                dgvSumList.Rows[i].Cells["ColMotorLengthK"].Value = "";
                //API 682 EDITION / API 682 版本
                dgvSumList.Rows[i].Cells["ColApiEdition"].Value = "4TH(默认)";
                //SEAL PLAN A/ 冲洗方案
                dgvSumList.Rows[i].Cells["ColSealPlanA"].Value = "";
                //SEAL PLAN B/ 冲洗方案
                dgvSumList.Rows[i].Cells["ColSealPlanB"].Value = "";
                //SEAL PLAN C/ 冲洗方案
                dgvSumList.Rows[i].Cells["ColSealPlanC"].Value = "";
                //FLOW / 流量 (m3/h)
                dgvSumList.Rows[i].Cells["ColFlow"].Value = dtBody.Rows[i]["S29"].ToString();
                //HEAD / 扬程 (m)
                dgvSumList.Rows[i].Cells["ColHead"].Value = dtBody.Rows[i]["S30"].ToString();
                //EFFICIENCY / 效率(%)
                dgvSumList.Rows[i].Cells["ColEfficiency"].Value = dtBody.Rows[i]["S31"].ToString();
                //RATED POWER / 额定功率(kW)
                dgvSumList.Rows[i].Cells["ColRatedPower"].Value = dtBody.Rows[i]["S32"].ToString();
                //NPSHr  (m)
                dgvSumList.Rows[i].Cells["ColNpshr"].Value = dtBody.Rows[i]["S33"].ToString();
                //MAWP (Mpa)
                dgvSumList.Rows[i].Cells["ColMawp"].Value = dtBody.Rows[i]["S34"].ToString();
                //MAWT (℃)
                dgvSumList.Rows[i].Cells["ColMawt"].Value = dtBody.Rows[i]["S35"].ToString();
                //Hydrostatic P (Mpa)
                dgvSumList.Rows[i].Cells["ColHydrostaticP"].Value = "";
                //FLANGE STANDARD / 法兰标准
                if(i == 0)
                {
                    strFlangeStandard = dtBody.Rows[i]["S37"].ToString();
                }
                //MATERIAL CLASS / 材料等级
                dgvSumList.Rows[i].Cells["ColMaterialClass"].Value = dtBody.Rows[i]["S41"].ToString();
                //BEARING / 轴承配置
                dgvSumList.Rows[i].Cells["ColBearing"].Value = dtBody.Rows[i]["S43"].ToString();
            }
            //FLANGE STANDARD 赋值
            cobFlangeStandard.SelectedValue = strFlangeStandard;

        }
        /// <summary>
        /// Data binding after XML file reading
        /// </summary>
        /// <param name="dtHead"></param>
        /// <param name="dtBody"></param>
        private void DataBindingFromXML(DataTable dtHead, DataTable dtBody)
        {
            //DataGridView Combox Initialization
            DgvComboxInitialSettings();
            string strChecked = "1";
            //head assignment
            for (int i = 0; i < dtHead.Rows.Count; i++)
            {
                if (strChecked.Equals(dtHead.Rows[i]["ColPgaCheck"].ToString()))
                {
                    cbPGA.Checked = true;
                }
                else
                {
                    cbPGA.Checked = false;
                }
                if (strChecked.Equals(dtHead.Rows[i]["ColPpiCheck"].ToString()))
                {
                    cbPPI.Checked = true;
                }
                else
                {
                    cbPPI.Checked = false;
                }
                if (strChecked.Equals(dtHead.Rows[i]["ColPcsCheck"].ToString()))
                {
                    cbPCS.Checked = true;
                }
                else
                {
                    cbPCS.Checked = false;
                }
                if (strChecked.Equals(dtHead.Rows[i]["ColPucCheck"].ToString()))
                {
                    cbPUC.Checked = true;
                }
                else
                {
                    cbPUC.Checked = false;
                }
                if (strChecked.Equals(dtHead.Rows[i]["ColPntCheck"].ToString()))
                {
                    cbPNT.Checked = true;
                }
                else
                {
                    cbPNT.Checked = false;
                }
                if (strChecked.Equals(dtHead.Rows[i]["ColPdsCheck"].ToString()))
                {
                    cbPDS.Checked = true;
                }
                else
                {
                    cbPDS.Checked = false;
                }
                if (strChecked.Equals(dtHead.Rows[i]["ColPslCheck"].ToString()))
                {
                    cbPSL.Checked = true;
                }
                else
                {
                    cbPSL.Checked = false;
                }
                //Purchaser/买方
                txtPurchaser.Text = dtHead.Rows[i]["ColPurchaser"].ToString();
                //End User:/用户
                txtEndUser.Text = dtHead.Rows[i]["ColEndUser"].ToString();
                //Project/项目
                txtProject.Text = dtHead.Rows[i]["ColProject"].ToString();
                //txtSite
                txtSite.Text = dtHead.Rows[i]["ColSite"].ToString();
                //Customer PO NO
                txtCustomerPONO.Text = dtHead.Rows[i]["ColCustomerPoNo"].ToString();
                //Sulzer Order NO
                txtSulzerOrderNO.Text = dtHead.Rows[i]["ColSulzerOrderNo"].ToString();
                //Prepared by/编制
                txtPreparedBy.Text = dtHead.Rows[i]["ColPreparedBy"].ToString();
                //Reviewed by/审核
                txtReviewedBy.Text = dtHead.Rows[i]["ColReviewedBy"].ToString();
                //Approved by/批准
                txtApprovedBy.Text = dtHead.Rows[i]["ColApprovedBy"].ToString();
                //Date/日期
                dateTimeWorkDate.Value = Convert.ToDateTime(dtHead.Rows[i]["ColDate"].ToString());
                //FLANGE STANDARD
                cobFlangeStandard.SelectedValue = dtHead.Rows[i]["ColFlangeStandard"].ToString();
            }
            for (int i = 0; i < dtBody.Rows.Count; i++)
            {
                if(dtBody.Rows[i]["ColMaterialClass"].ToString().Contains("MOD"))
                {
                    dgvSumList.Columns["ColVoluteCasing"].Visible = true;
                    dgvSumList.Columns["ColImpeller"].Visible = true;
                    dgvSumList.Columns["ColStationaryWearRing"].Visible = true;
                    dgvSumList.Columns["ColCoverCoolingPlate"].Visible = true;
                    dgvSumList.Columns["ColThroatBush"].Visible = true;
                    dgvSumList.Columns["ColImpellerKey"].Visible = true;
                    dgvSumList.Columns["ColGasket"].Visible = true;
                    dgvSumList.Columns["ColSpacerSleeve"].Visible = true;
                    dgvSumList.Columns["ColImpellerNut"].Visible = true;
                    dgvSumList.Columns["ColInducer"].Visible = true;
                    dgvSumList.Columns["ColInducerSleeve"].Visible = true;
                    dgvSumList.Columns["ColWearPlate"].Visible = true;
                    break;
                }
            }
            //XML assignment
            for (int i = 0; i < dtBody.Rows.Count; i++)
            {
                dgvSumList.Rows.Add();
                //行号
                dgvSumList.Rows[i].HeaderCell.Value = string.Format("{0}", i + 1);
                dgvSumList.Rows[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //ITEM No/行号
                dgvSumList.Rows[i].Cells["ColItemNo"].Value = dtBody.Rows[i]["ColItemNo"].ToString();
                //Series No/苏尔寿编号
                dgvSumList.Rows[i].Cells["ColSeriesNo"].Value = dtBody.Rows[i]["ColSeriesNo"].ToString();
                //TAG NO/客户位号
                dgvSumList.Rows[i].Cells["ColTagNo"].Value = dtBody.Rows[i]["ColTagNo"].ToString();
                //TYPE/泵型ComboBox 赋值 
                dgvSumList.Rows[i].Cells["ColType"].Value = dtBody.Rows[i]["ColType"].ToString();
                //According to the value obtained from (type / pump type),
                //initialize the (size / specification) combobox first, and then assign a value to it
                SelecteSizeComboBox(dtBody.Rows[i]["ColType"].ToString(), i);
                if (!dtBody.Rows[i]["ColType"].ToString().Contains("PR"))
                {
                    //PRE泵压力等级
                    dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                    dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].Value = "";
                    //是否带诱导轮
                    dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].ReadOnly = true;
                    dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].Value = "";
                    //泵体/泵盖
                    dgvSumList.Rows[i].Cells["ColVoluteCasing"].ReadOnly = false;
                    //叶轮
                    dgvSumList.Rows[i].Cells["ColImpeller"].ReadOnly = false;
                    //口环
                    dgvSumList.Rows[i].Cells["ColStationaryWearRing"].ReadOnly = false;
                    //喉部衬套
                    dgvSumList.Rows[i].Cells["ColThroatBush"].ReadOnly = false;
                    //叶轮键
                    dgvSumList.Rows[i].Cells["ColImpellerKey"].ReadOnly = false;
                    //泵体泵盖之间的密封垫
                    dgvSumList.Rows[i].Cells["ColGasket"].ReadOnly = false;
                    //叶轮螺母
                    dgvSumList.Rows[i].Cells["ColImpellerNut"].ReadOnly = false;
                    if(dtBody.Rows[i]["ColType"].ToString().Equals("ZAO") || dtBody.Rows[i]["ColType"].ToString().Equals("ZEO"))
                    {
                        //耐磨板
                        dgvSumList.Rows[i].Cells["ColWearPlate"].ReadOnly = false;
                    }
                    else
                    {
                        //耐磨板
                        dgvSumList.Rows[i].Cells["ColWearPlate"].ReadOnly = true;
                    }
                }
                else
                {
                    //PRE泵压力等级
                    dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].ReadOnly = false;
                    dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].Value = dtBody.Rows[i]["ColPREPumpPressureClass"].ToString();
                    //是否带诱导轮
                    dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].ReadOnly = false;
                    if (!String.IsNullOrEmpty(dtBody.Rows[i]["ColPREInducerYesNo"].ToString()))
                    {
                        dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].Value = Boolean.Parse(dtBody.Rows[i]["ColPREInducerYesNo"].ToString());
                    }
                    //泵体/泵盖
                    dgvSumList.Rows[i].Cells["ColVoluteCasing"].ReadOnly = false;
                    //叶轮
                    dgvSumList.Rows[i].Cells["ColImpeller"].ReadOnly = false;
                    //口环
                    dgvSumList.Rows[i].Cells["ColStationaryWearRing"].ReadOnly = false;
                    //泵盖水冷密封板
                    dgvSumList.Rows[i].Cells["ColCoverCoolingPlate"].ReadOnly = false;
                    //喉部衬套
                    dgvSumList.Rows[i].Cells["ColThroatBush"].ReadOnly = false;
                    //泵体泵盖之间的密封垫
                    dgvSumList.Rows[i].Cells["ColGasket"].ReadOnly = false;
                    //轴套
                    dgvSumList.Rows[i].Cells["ColSpacerSleeve"].ReadOnly = false;
                    //叶轮螺母
                    dgvSumList.Rows[i].Cells["ColImpellerNut"].ReadOnly = false;
                    //诱导轮
                    dgvSumList.Rows[i].Cells["ColInducer"].ReadOnly = false;
                    //诱导轮衬套
                    dgvSumList.Rows[i].Cells["ColInducerSleeve"].ReadOnly = false;
                    if(dtBody.Rows[i]["ColType"].ToString().Equals("PRESF"))
                    {
                        //耐磨板
                        dgvSumList.Rows[i].Cells["ColWearPlate"].ReadOnly = false;
                    }
                    else
                    {
                        //耐磨板
                        dgvSumList.Rows[i].Cells["ColWearPlate"].ReadOnly = true;
                    }
                }
                //SIZE/规格 
                dgvSumList.Rows[i].Cells["ColSize"].Value = dtBody.Rows[i]["ColSize"].ToString();
                //SERVICE / 泵名称
                dgvSumList.Rows[i].Cells["ColService"].Value = dtBody.Rows[i]["ColService"].ToString();
                //Qty / 数量
                dgvSumList.Rows[i].Cells["ColQty"].Value = dtBody.Rows[i]["ColQty"].ToString();
                //MOTOR MANUFACTURER / 电机供应商
                dgvSumList.Rows[i].Cells["ColMotorManufacturer"].Value = dtBody.Rows[i]["ColMotorManufacturer"].ToString();
                //电机型号
                dgvSumList.Rows[i].Cells["ColMotorFrame"].Value = dtBody.Rows[i]["ColMotorFrame"].ToString();
                //FRAME SIZE/基座号
                dgvSumList.Rows[i].Cells["ColFrameSize"].Value = dtBody.Rows[i]["ColFrameSize"].ToString();
                //MOTOR POWER/ 电机功率(kW) 
                dgvSumList.Rows[i].Cells["ColMotorPower"].Value = dtBody.Rows[i]["ColMotorPower"].ToString();
                //VOLTAGE/ 电压(V) 
                dgvSumList.Rows[i].Cells["ColVoltage"].Value = dtBody.Rows[i]["ColVoltage"].ToString();
                //转数 r/min
                dgvSumList.Rows[i].Cells["ColSpeed"].Value = dtBody.Rows[i]["ColSpeed"].ToString();
                //电机长度K
                dgvSumList.Rows[i].Cells["ColMotorLengthK"].Value = dtBody.Rows[i]["ColMotorLengthK"].ToString();
                //API 682 EDITION / API 682 版本
                dgvSumList.Rows[i].Cells["ColApiEdition"].Value = dtBody.Rows[i]["ColApiEdition"].ToString();
                //SEAL PLAN A/ 冲洗方案
                dgvSumList.Rows[i].Cells["ColSealPlanA"].Value = dtBody.Rows[i]["ColSealPlanA"].ToString();
                //SEAL PLAN B/ 冲洗方案
                dgvSumList.Rows[i].Cells["ColSealPlanB"].Value = dtBody.Rows[i]["ColSealPlanB"].ToString();
                //SEAL PLAN C/ 冲洗方案
                dgvSumList.Rows[i].Cells["ColSealPlanC"].Value = dtBody.Rows[i]["ColSealPlanC"].ToString();
                //FLOW / 流量 (m3/h)
                dgvSumList.Rows[i].Cells["ColFlow"].Value = dtBody.Rows[i]["ColFlow"].ToString();
                //HEAD / 扬程 (m)
                dgvSumList.Rows[i].Cells["ColHead"].Value = dtBody.Rows[i]["ColHead"].ToString();
                //EFFICIENCY / 效率(%)
                dgvSumList.Rows[i].Cells["ColEfficiency"].Value = dtBody.Rows[i]["ColEfficiency"].ToString();
                //RATED POWER / 额定功率(kW)
                dgvSumList.Rows[i].Cells["ColRatedPower"].Value = dtBody.Rows[i]["ColRatedPower"].ToString();
                //NPSHr  (m)
                dgvSumList.Rows[i].Cells["ColNpshr"].Value = dtBody.Rows[i]["ColNpshr"].ToString();
                //MAWP (Mpa)
                dgvSumList.Rows[i].Cells["ColMawp"].Value = dtBody.Rows[i]["ColMawp"].ToString();
                //MAWT (℃)
                dgvSumList.Rows[i].Cells["ColMawt"].Value = dtBody.Rows[i]["ColMawt"].ToString();
                //Hydrostatic P (Mpa)
                dgvSumList.Rows[i].Cells["ColHydrostaticP"].Value = dtBody.Rows[i]["ColHydrostaticP"].ToString();
                //排残管口口径
                dgvSumList.Rows[i].Cells["ColPumpDrainNozzleSize"].Value = dtBody.Rows[i]["ColPumpDrainNozzleSize"].ToString();
                //排残管路 阀门
                dgvSumList.Rows[i].Cells["ColPumpDrainValve"].Value = dtBody.Rows[i]["ColPumpDrainValve"].ToString();
                if (!String.IsNullOrEmpty(dtBody.Rows[i]["ColWarmUpPiping"].ToString()))
                {
                    //暖泵管路
                    dgvSumList.Rows[i].Cells["ColWarmUpPiping"].Value = Boolean.Parse(dtBody.Rows[i]["ColWarmUpPiping"].ToString());
                }
                //MATERIAL CLASS / 材料等级
                dgvSumList.Rows[i].Cells["ColMaterialClass"].Value = dtBody.Rows[i]["ColMaterialClass"].ToString();
                //干湿轴
                dgvSumList.Rows[i].Cells["ColDryWetShaft"].Value = dtBody.Rows[i]["ColDryWetShaft"].ToString();
                if(dtBody.Rows[i]["ColDryWetShaft"].ToString().Equals("Wet"))
                {
                    //Wet Shaft(自己填写)
                    dgvSumList.Rows[i].Cells["ColWetInput"].ReadOnly = false;
                }
                else
                {
                    //Wet Shaft(自己填写)
                    dgvSumList.Rows[i].Cells["ColWetInput"].ReadOnly = true;
                }
                //Wet Shaft(自己填写)
                dgvSumList.Rows[i].Cells["ColWetInput"].Value = dtBody.Rows[i]["ColWetInput"].ToString();
                //BEARING / 轴承配置
                dgvSumList.Rows[i].Cells["ColBearing"].Value = dtBody.Rows[i]["ColBearing"].ToString();
                //轴承保持架材质
                dgvSumList.Rows[i].Cells["ColBearingCageMaterial"].Value = dtBody.Rows[i]["ColBearingCageMaterial"].ToString();
                if (!String.IsNullOrEmpty(dtBody.Rows[i]["ColCompanionFlg"].ToString()))
                {
                    //COMPANION FLG. / 配对法兰要求
                    dgvSumList.Rows[i].Cells["ColCompanionFlg"].Value = Boolean.Parse(dtBody.Rows[i]["ColCompanionFlg"].ToString());
                }
                if (!String.IsNullOrEmpty(dtBody.Rows[i]["ColCaseBracket"].ToString()))
                {
                    //壳体是否带夹套
                    dgvSumList.Rows[i].Cells["ColCaseBracket"].Value = Boolean.Parse(dtBody.Rows[i]["ColCaseBracket"].ToString());
                }
                //管口载荷要求
                dgvSumList.Rows[i].Cells["ColNozzleLoad"].Value = dtBody.Rows[i]["ColNozzleLoad"].ToString();
                if (!String.IsNullOrEmpty(dtBody.Rows[i]["ColReservedPureLubrication"].ToString()))
                {
                    //预留油雾润滑
                    dgvSumList.Rows[i].Cells["ColReservedPureLubrication"].Value = Boolean.Parse(dtBody.Rows[i]["ColReservedPureLubrication"].ToString());
                }
                if (!String.IsNullOrEmpty(dtBody.Rows[i]["ColBearighousingDrain"].ToString()))
                {
                    //轴承室带排油管路
                    dgvSumList.Rows[i].Cells["ColBearighousingDrain"].Value = Boolean.Parse(dtBody.Rows[i]["ColBearighousingDrain"].ToString());
                }
                if (!String.IsNullOrEmpty(dtBody.Rows[i]["ColImpellerBalanceHole"].ToString()))
                {
                    //叶轮带平衡孔
                    dgvSumList.Rows[i].Cells["ColImpellerBalanceHole"].Value = Boolean.Parse(dtBody.Rows[i]["ColImpellerBalanceHole"].ToString());
                }
                //密封冲洗量
                dgvSumList.Rows[i].Cells["ColSealFlushFlow"].Value = dtBody.Rows[i]["ColSealFlushFlow"].ToString();
                //外接冷却水冲洗量
                dgvSumList.Rows[i].Cells["ColCoolingWaterFlow"].Value = dtBody.Rows[i]["ColCoolingWaterFlow"].ToString();
                //冷却水母管口径
                dgvSumList.Rows[i].Cells["ColCoolWaterMainPipeSize"].Value = dtBody.Rows[i]["ColCoolWaterMainPipeSize"].ToString();
                //排残阀类型
                dgvSumList.Rows[i].Cells["ColDrainValveType"].Value = dtBody.Rows[i]["ColDrainValveType"].ToString();
                //泵体/泵盖
                dgvSumList.Rows[i].Cells["ColVoluteCasing"].Value = dtBody.Rows[i]["ColVoluteCasing"].ToString();
                //叶轮
                dgvSumList.Rows[i].Cells["ColImpeller"].Value = dtBody.Rows[i]["ColImpeller"].ToString();
                //口环
                dgvSumList.Rows[i].Cells["ColStationaryWearRing"].Value = dtBody.Rows[i]["ColStationaryWearRing"].ToString();
                //泵盖水冷密封板
                dgvSumList.Rows[i].Cells["ColCoverCoolingPlate"].Value = dtBody.Rows[i]["ColCoverCoolingPlate"].ToString();
                //喉部衬套
                dgvSumList.Rows[i].Cells["ColThroatBush"].Value = dtBody.Rows[i]["ColThroatBush"].ToString();
                //叶轮键
                dgvSumList.Rows[i].Cells["ColImpellerKey"].Value = dtBody.Rows[i]["ColImpellerKey"].ToString();
                //泵体泵盖之间的密封垫
                dgvSumList.Rows[i].Cells["ColGasket"].Value = dtBody.Rows[i]["ColGasket"].ToString();
                //轴套
                dgvSumList.Rows[i].Cells["ColSpacerSleeve"].Value = dtBody.Rows[i]["ColSpacerSleeve"].ToString();
                //叶轮螺母
                dgvSumList.Rows[i].Cells["ColImpellerNut"].Value = dtBody.Rows[i]["ColImpellerNut"].ToString();
                //诱导轮
                dgvSumList.Rows[i].Cells["ColInducer"].Value = dtBody.Rows[i]["ColInducer"].ToString();
                //诱导轮衬套
                dgvSumList.Rows[i].Cells["ColInducerSleeve"].Value = dtBody.Rows[i]["ColInducerSleeve"].ToString();
                //耐磨板
                dgvSumList.Rows[i].Cells["ColWearPlate"].Value = dtBody.Rows[i]["ColWearPlate"].ToString();

            }
        }

        /// <summary>
        /// DataGridView Combox Initialization
        /// </summary>
        private void DgvComboxInitialSettings()
        {
            //FLANGE STANDARD Initialization
            cobFlangeStandard.DataSource = COMBOBOX_DATASAT.Tables["FLANGESTANDARDTABLE"];
            cobFlangeStandard.DisplayMember = "FLANGESTANDARD";
            cobFlangeStandard.ValueMember = "FLANGESTANDARD";
            cobFlangeStandard.SelectedValue = "";
            //TYPE / 泵型 Initialization
            DataGridViewComboBoxColumn dgvComboBoxColumnType = dgvSumList.Columns["ColType"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnType.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnType.DataSource = COMBOBOX_DATASAT.Tables["TYPETABLE"];
            dgvComboBoxColumnType.DisplayMember = "TYPE";
            dgvComboBoxColumnType.ValueMember = "TYPE";
            dgvComboBoxColumnType.DataPropertyName = "TYPE";
            //FRAME SIZE/基座号
            DataGridViewComboBoxColumn dgvComboBoxColumnFrameSize = dgvSumList.Columns["ColFrameSize"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnFrameSize.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnFrameSize.DataSource = COMBOBOX_DATASAT.Tables["FRAMESIZETABLE"];
            dgvComboBoxColumnFrameSize.DisplayMember = "FRAMESIZE";
            dgvComboBoxColumnFrameSize.ValueMember = "FRAMESIZE";
            dgvComboBoxColumnFrameSize.DataPropertyName = "FRAMESIZE";
            //API 682 EDITION / API 682 版本
            DataGridViewComboBoxColumn dgvComboBoxColumnApiEdition = dgvSumList.Columns["ColApiEdition"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnApiEdition.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnApiEdition.DataSource = COMBOBOX_DATASAT.Tables["APIEDITIONTABLE"];
            dgvComboBoxColumnApiEdition.DisplayMember = "APIEDITION";
            dgvComboBoxColumnApiEdition.ValueMember = "APIEDITION";
            dgvComboBoxColumnApiEdition.DataPropertyName = "APIEDITION";
            //SEAL PLAN A/ 冲洗方案
            DataGridViewComboBoxColumn dgvComboBoxColumnSealPlanA = dgvSumList.Columns["ColSealPlanA"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnSealPlanA.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnSealPlanA.DataSource = COMBOBOX_DATASAT.Tables["SEALPLANATABLE"];
            dgvComboBoxColumnSealPlanA.DisplayMember = "SEALPLANA";
            dgvComboBoxColumnSealPlanA.ValueMember = "SEALPLANA";
            dgvComboBoxColumnSealPlanA.DataPropertyName = "SEALPLANA";
            //SEAL PLAN B/ 冲洗方案
            DataGridViewComboBoxColumn dgvComboBoxColumnSealPlanB = dgvSumList.Columns["ColSealPlanB"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnSealPlanB.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnSealPlanB.DataSource = COMBOBOX_DATASAT.Tables["SEALPLANBTABLE"];
            dgvComboBoxColumnSealPlanB.DisplayMember = "SEALPLANB";
            dgvComboBoxColumnSealPlanB.ValueMember = "SEALPLANB";
            dgvComboBoxColumnSealPlanB.DataPropertyName = "SEALPLANB";
            //SEAL PLAN C/ 冲洗方案
            DataGridViewComboBoxColumn dgvComboBoxColumnSealPlanC = dgvSumList.Columns["ColSealPlanC"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnSealPlanC.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnSealPlanC.DataSource = COMBOBOX_DATASAT.Tables["SEALPLANCTABLE"];
            dgvComboBoxColumnSealPlanC.DisplayMember = "SEALPLANC";
            dgvComboBoxColumnSealPlanC.ValueMember = "SEALPLANC";
            dgvComboBoxColumnSealPlanC.DataPropertyName = "SEALPLANC";
            //MATERIAL CLASS / 材料等级
            DataGridViewComboBoxColumn dgvComboBoxColumnMaterialClass = dgvSumList.Columns["ColMaterialClass"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnMaterialClass.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnMaterialClass.DataSource = COMBOBOX_DATASAT.Tables["MATERIALCLASSTABLE"];
            dgvComboBoxColumnMaterialClass.DisplayMember = "MATERIALCLASS";
            dgvComboBoxColumnMaterialClass.ValueMember = "MATERIALCLASS";
            dgvComboBoxColumnMaterialClass.DataPropertyName = "MATERIALCLASS";
            //干湿轴
            DataGridViewComboBoxColumn dgvComboBoxColumnGSZ = dgvSumList.Columns["ColDryWetShaft"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnGSZ.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnGSZ.DataSource = ChoicePDValve.GetWetDryChoices();
            dgvComboBoxColumnGSZ.DisplayMember = "PDName";
            dgvComboBoxColumnGSZ.ValueMember = "PDValue";
            //BEARING / 轴承配置
            DataGridViewComboBoxColumn dgvComboBoxColumnBearing = dgvSumList.Columns["ColBearing"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnBearing.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnBearing.DataSource = COMBOBOX_DATASAT.Tables["BEARINGTABLE"];
            dgvComboBoxColumnBearing.DisplayMember = "BEARING";
            dgvComboBoxColumnBearing.ValueMember = "BEARING";
            dgvComboBoxColumnBearing.DataPropertyName = "BEARING";
            //轴承保持架材质
            DataGridViewComboBoxColumn dgvComboBoxColumnZCBCJCZ = dgvSumList.Columns["ColBearingCageMaterial"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnZCBCJCZ.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnZCBCJCZ.DataSource = ChoicePDValve.GetChoiceCageValve();
            dgvComboBoxColumnZCBCJCZ.DisplayMember = "PDName";
            dgvComboBoxColumnZCBCJCZ.ValueMember = "PDValue";
            //排残管路阀门
            DataGridViewComboBoxColumn dgvComboBoxColumnPCGLFM = dgvSumList.Columns["ColPumpDrainValve"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnPCGLFM.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnPCGLFM.DataSource = ChoicePDValve.GetChoicePDValve();
            dgvComboBoxColumnPCGLFM.DisplayMember = "PDName";
            dgvComboBoxColumnPCGLFM.ValueMember = "PDValue";
            //暖泵管路
            DataGridViewComboBoxColumn dgvComboBoxColumnWarmUpPiping = dgvSumList.Columns["ColWarmUpPiping"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnWarmUpPiping.ValueType = typeof(Boolean);
            dgvComboBoxColumnWarmUpPiping.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnWarmUpPiping.DataSource = Choice.GetChoices();
            dgvComboBoxColumnWarmUpPiping.DisplayMember = "Name";
            dgvComboBoxColumnWarmUpPiping.ValueMember = "Value";
            //COMPANION FLG. / 配对法兰要求
            DataGridViewComboBoxColumn dgvComboBoxColumnCompanionFlag = dgvSumList.Columns["ColCompanionFlg"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnCompanionFlag.ValueType = typeof(Boolean);
            dgvComboBoxColumnCompanionFlag.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnCompanionFlag.DataSource = Choice.GetChoices();
            dgvComboBoxColumnCompanionFlag.DisplayMember = "Name";
            dgvComboBoxColumnCompanionFlag.ValueMember = "Value";
            //壳体是否带夹套/CASE  WI/WO BRACKET
            DataGridViewComboBoxColumn dgvComboBoxColumnCaseWiWoBracket = dgvSumList.Columns["ColCaseBracket"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnCaseWiWoBracket.ValueType = typeof(Boolean);
            dgvComboBoxColumnCaseWiWoBracket.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnCaseWiWoBracket.DataSource = Choice.GetChoices();
            dgvComboBoxColumnCaseWiWoBracket.DisplayMember = "Name";
            dgvComboBoxColumnCaseWiWoBracket.ValueMember = "Value";
            //预留油雾润滑/WI/WO RESERVED PURE OIL-MIST LUBRICATION
            DataGridViewComboBoxColumn dgvComboBoxColumnReservedPureLubrication = dgvSumList.Columns["ColReservedPureLubrication"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnReservedPureLubrication.ValueType = typeof(Boolean);
            dgvComboBoxColumnReservedPureLubrication.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnReservedPureLubrication.DataSource = Choice.GetChoices();
            dgvComboBoxColumnReservedPureLubrication.DisplayMember = "Name";
            dgvComboBoxColumnReservedPureLubrication.ValueMember = "Value";
            //轴承室带排油管路/BEARING HOUSING WI/WO OIL DRAIN 
            DataGridViewComboBoxColumn dgvComboBoxColumnBearingOilDrain = dgvSumList.Columns["ColBearighousingDrain"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnBearingOilDrain.ValueType = typeof(Boolean);
            dgvComboBoxColumnBearingOilDrain.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnBearingOilDrain.DataSource = Choice.GetChoices();
            dgvComboBoxColumnBearingOilDrain.DisplayMember = "Name";
            dgvComboBoxColumnBearingOilDrain.ValueMember = "Value";
            //叶轮带平衡孔/IMPELLER WI/WO BALANCE HOLE
            DataGridViewComboBoxColumn dgvComboBoxColumnImpellerBalanceHole = dgvSumList.Columns["ColImpellerBalanceHole"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnImpellerBalanceHole.ValueType = typeof(Boolean);
            dgvComboBoxColumnImpellerBalanceHole.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnImpellerBalanceHole.DataSource = Choice.GetChoices();
            dgvComboBoxColumnImpellerBalanceHole.DisplayMember = "Name";
            dgvComboBoxColumnImpellerBalanceHole.ValueMember = "Value";
            //PRE泵压力等级
            DataGridViewComboBoxColumn dgvComboBoxColumnPREPumpPressureClass = dgvSumList.Columns["ColPREPumpPressureClass"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnPREPumpPressureClass.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnPREPumpPressureClass.DataSource = COMBOBOX_DATASAT.Tables["PREPUMPPRESSURECLASSTABLE"];
            dgvComboBoxColumnPREPumpPressureClass.DisplayMember = "PRESSURECLASS";
            dgvComboBoxColumnPREPumpPressureClass.ValueMember = "PRESSURECLASS";
            dgvComboBoxColumnPREPumpPressureClass.DataPropertyName = "PRESSURECLASS";
            //是否带诱导轮
            DataGridViewComboBoxColumn dgvComboBoxColumnPREInducerYesNo = dgvSumList.Columns["ColPREInducerYesNo"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnPREInducerYesNo.ValueType = typeof(Boolean);
            dgvComboBoxColumnPREInducerYesNo.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnPREInducerYesNo.DataSource = Choice.GetChoices();
            dgvComboBoxColumnPREInducerYesNo.DisplayMember = "Name";
            dgvComboBoxColumnPREInducerYesNo.ValueMember = "Value";
            //冷却水母管口径
            DataGridViewComboBoxColumn dgvComboBoxColumnCoolWaterMainPipeSize = dgvSumList.Columns["ColCoolWaterMainPipeSize"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnCoolWaterMainPipeSize.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnCoolWaterMainPipeSize.DataSource = COMBOBOX_DATASAT.Tables["COOLWATERMAINPIPESIZETABLE"];
            dgvComboBoxColumnCoolWaterMainPipeSize.DisplayMember = "MAINPIPESIZE";
            dgvComboBoxColumnCoolWaterMainPipeSize.ValueMember = "MAINPIPESIZE";
            dgvComboBoxColumnCoolWaterMainPipeSize.DataPropertyName = "MAINPIPESIZE";
            //排残阀类型
            DataGridViewComboBoxColumn dgvComboBoxColumnDrainValveType = dgvSumList.Columns["ColDrainValveType"] as DataGridViewComboBoxColumn;
            dgvComboBoxColumnDrainValveType.DefaultCellStyle.BackColor = Color.FromName("window");
            dgvComboBoxColumnDrainValveType.DataSource = COMBOBOX_DATASAT.Tables["DRAINVALVETYPETABLE"];
            dgvComboBoxColumnDrainValveType.DisplayMember = "DRAINVALVETYPE";
            dgvComboBoxColumnDrainValveType.ValueMember = "DRAINVALVETYPE";
            dgvComboBoxColumnDrainValveType.DataPropertyName = "DRAINVALVETYPE";
        }
        /// <summary>
        /// Head data to dataTable
        /// </summary>
        /// <returns></returns>
        private DataTable GetHeadDataToTable()
        {
            //Head
            DataTable dataTableHead = new DataTable();
            dataTableHead.TableName = "DataTableHead";
            //24 line add
            DataColumn columnPga = new DataColumn("ColPgaCheck");
            dataTableHead.Columns.Add(columnPga);
            DataColumn columnPpi = new DataColumn("ColPpiCheck");
            dataTableHead.Columns.Add(columnPpi);
            DataColumn columnPcs = new DataColumn("ColPcsCheck");
            dataTableHead.Columns.Add(columnPcs);
            DataColumn columnPuc = new DataColumn("ColPucCheck");
            dataTableHead.Columns.Add(columnPuc);
            DataColumn columnPnt = new DataColumn("ColPntCheck");
            dataTableHead.Columns.Add(columnPnt);
            DataColumn columnPds = new DataColumn("ColPdsCheck");
            dataTableHead.Columns.Add(columnPds);
            DataColumn columnPsl = new DataColumn("ColPslCheck");
            dataTableHead.Columns.Add(columnPsl);
            //Purchaser/买方
            DataColumn columnPurchaser = new DataColumn("ColPurchaser");
            dataTableHead.Columns.Add(columnPurchaser);
            //End User:/用户
            DataColumn columnEndUser = new DataColumn("ColEndUser");
            dataTableHead.Columns.Add(columnEndUser);
            //Project/项目
            DataColumn columnProject = new DataColumn("ColProject");
            dataTableHead.Columns.Add(columnProject);
            //Site/安装地点
            DataColumn columnSite = new DataColumn("ColSite");
            dataTableHead.Columns.Add(columnSite);
            //Customer PO NO
            DataColumn columnCustomerPoNo = new DataColumn("ColCustomerPoNo");
            dataTableHead.Columns.Add(columnCustomerPoNo);
            //Sulzer Order NO
            DataColumn columnSulzerOrderNo = new DataColumn("ColSulzerOrderNo");
            dataTableHead.Columns.Add(columnSulzerOrderNo);
            //Prepared by/ 编制
            DataColumn columnPreparedBy = new DataColumn("ColPreparedBy");
            dataTableHead.Columns.Add(columnPreparedBy);
            //Reviewed by/审核
            DataColumn columnReviewedBy = new DataColumn("ColReviewedBy");
            dataTableHead.Columns.Add(columnReviewedBy);
            //Approved by/批准
            DataColumn columnApprovedBy = new DataColumn("ColApprovedBy");
            dataTableHead.Columns.Add(columnApprovedBy);
            //Date/日期
            DataColumn columnDate = new DataColumn("ColDate");
            dataTableHead.Columns.Add(columnDate);
            //FLANGE STANDARD
            DataColumn columnFlangeStandard = new DataColumn("ColFlangeStandard");
            dataTableHead.Columns.Add(columnFlangeStandard);

            DataRow dataRowHead = dataTableHead.NewRow();
            string strChecked = "1";
            string strUnChecked = "0";
            if (cbPGA.Checked == true)
            {
                dataRowHead[0] = strChecked;
            }
            else
            {
                dataRowHead[0] = strUnChecked;
            }
            if (cbPPI.Checked == true)
            {
                dataRowHead[1] = strChecked;
            }
            else
            {
                dataRowHead[1] = strUnChecked;
            }
            if (cbPCS.Checked == true)
            {
                dataRowHead[2] = strChecked;
            }
            else
            {
                dataRowHead[2] = strUnChecked;
            }
            if (cbPUC.Checked == true)
            {
                dataRowHead[3] = strChecked;
            }
            else
            {
                dataRowHead[3] = strUnChecked;
            }
            if (cbPNT.Checked == true)
            {
                dataRowHead[4] = strChecked;
            }
            else
            {
                dataRowHead[4] = strUnChecked;
            }
            if (cbPDS.Checked == true)
            {
                dataRowHead[5] = strChecked;
            }
            else
            {
                dataRowHead[5] = strUnChecked;
            }
            if (cbPSL.Checked == true)
            {
                dataRowHead[6] = strChecked;
            }
            else
            {
                dataRowHead[6] = strUnChecked;
            }
            //Purchaser/买方
            dataRowHead[7] = txtPurchaser.Text;
            //End User:/用户
            dataRowHead[8] = txtEndUser.Text;
            //Project/项目
            dataRowHead[9] = txtProject.Text;
            //Site/安装地点
            dataRowHead[10] = txtSite.Text;
            //Customer PO NO
            dataRowHead[11] = txtCustomerPONO.Text;
            //Sulzer Order NO
            dataRowHead[12] = txtSulzerOrderNO.Text;
            //Prepared by/ 编制
            dataRowHead[13] = txtPreparedBy.Text;
            //Reviewed by/审核
            dataRowHead[14] = txtReviewedBy.Text;
            //Approved by/批准
            dataRowHead[15] = txtApprovedBy.Text;
            //Date/日期
            dataRowHead[16] = dateTimeWorkDate.Value.ToString();
            //FLANGE STANDARD
            dataRowHead[17] = cobFlangeStandard.SelectedValue;
            dataTableHead.Rows.Add(dataRowHead);

            return dataTableHead;
        }
        /// <summary>
        /// dataGridView data to dataTable
        /// </summary>
        /// <param name="dgv">dataGridView</param>
        /// <returns></returns>
        private DataTable GetDgvToTable(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            // column force conversion
            for (int count = 0; count < dgv.Columns.Count; count++)
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }
            // for line
            for (int count = 0; count < dgv.Rows.Count; count++)
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        /// <summary>
        /// Save button pressed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSave_Click(object sender, EventArgs e)
        {
            //The file path is empty
            if (txtXMLFilePath.Text == "")
            {
                //Log write
                LogHelper.GetSingleObj().Error("修改数据画面保存", "Please select the input data file save path!");
                MessageBox.Show("Please select the input data file save path!", "ERROR");
                return;
            }
            //dgvSumList data
            if (this.dgvSumList.RowCount == 0)
            {
                //Log write
                LogHelper.GetSingleObj().Error("修改数据画面保存", "Please import data!");
                MessageBox.Show("There is no data to save!", "ERROR");
                return;
            }
            //check
            string strErrorChk = DataCheck();
            if (!string.IsNullOrEmpty(strErrorChk))
            {
                //Log write
                LogHelper.GetSingleObj().Error("修改数据画面保存", strErrorChk);
                MessageBox.Show(strErrorChk, "ERROR");
                return;
            }
            try
            {
                //Button not available
                BtnToEnabled(false);
                //XML file save path
                string xmlFilePath = txtXMLFilePath.Text;
                //head data to dataTable
                DataTable dataTableHead = GetHeadDataToTable();
                dataTableHead.TableName = "DataTableHead";
                //dataGridView data to dataTable
                //Body
                DataTable dataTableBody = GetDgvToTable(this.dgvSumList);
                dataTableBody.TableName = "DataTableBody";
                //dataSet
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(dataTableHead.Copy());
                dataSet.Tables.Add(dataTableBody.Copy());
                DateTime dtTime = DateTime.Now;
                string stringDtTime = dtTime.ToString("yyyyMMdd-HHmmss");
                string xmlFileName = xmlFilePath + "\\" + "DATA-" + stringDtTime + ".xml";
                XmlDataOperation.DataSetToXMLFile(dataSet, xmlFileName);
                //Assign a value to the global variable and print the data later
                GlobalVariableController.GlobalDataSet = dataSet.Copy();
                //Has the data been saved before the screen is closed
                CLOSE_FLAG++;
            }
            catch (Exception ex)
            {
                LogHelper.GetSingleObj().Error("修改数据画面保存", ex.Message);
                BtnToEnabled(true);
                MessageBox.Show("修改数据画面保存:" + ex.Message, "ERROR");
                return;
            }
            LogHelper.GetSingleObj().Info("修改数据画面保存", "数据保存成功");
            BtnToEnabled(true);
            MessageBox.Show("XML data file generation completed!");
        }
        /// <summary>
        /// Data Must Check
        /// </summary>
        /// <returns></returns>
        private string DataCheck()
        {
            String strError = "";
            //PPI Selected
            if (cbPPI.Checked)
            {
                //TYPE/泵型
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColType"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "TYPE/泵型" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANA/冲洗方案
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanA"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANA/冲洗方案A" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANB/冲洗方案B
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanB"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANA/冲洗方案B" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANC/冲洗方案C
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanC"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANC/冲洗方案C" + "不能为空" + "\n";
                        break;
                    }
                }
                //排残管路阀门
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColPumpDrainValve"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "排残管路阀门" + "不能为空" + "\n";
                        break;
                    }
                }
                //暖泵管路
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColWarmUpPiping"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "暖泵管路" + "不能为空" + "\n";
                        break;
                    }
                }
                //BEARING/轴承配置
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColBearing"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "BEARING/轴承配置" + "不能为空" + "\n";
                        break;
                    }
                }
                //COMPANION FLG. / 配对法兰要求
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColCompanionFlg"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "COMPANION FLG./配对法兰要求" + "不能为空" + "\n";
                        break;
                    }
                }
                //壳体是否带夹套
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColCaseBracket"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "壳体是否带夹套" + "不能为空" + "\n";
                        break;
                    }
                }
                //轴承室带排油管路
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColBearighousingDrain"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "轴承室带排油管路" + "不能为空" + "\n";
                        break;
                    }
                }
                //ITEM No/行号
                //数字CHECK
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColItemNo"].Value;
                    if (!CheckNumber(strChk))
                    {
                        strError = strError + "行号" + "请入力数字" + "\n";
                        break;
                    }
                }
            }
            //PUC Selected
            if (cbPUC.Checked)
            {
                //TYPE/泵型
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColType"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "TYPE/泵型" + "不能为空" + "\n";
                        break;
                    }
                }
                //SIZE/规格
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSize"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SIZE/规格" + "不能为空" + "\n";
                        break;
                    }
                }
                //MOTOR POWER/ 电机功率(kW) 
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColMotorPower"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "MOTOR POWER/电机功率(kW)" + "不能为空" + "\n";
                        break;
                    }
                }
                //VOLTAGE/ 电压(V) 
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColVoltage"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "VOLTAGE/电压(V)" + "不能为空" + "\n";
                        break;
                    }
                }
                //密封冲洗量
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealFlushFlow"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "密封冲洗量" + "不能为空" + "\n";
                        break;
                    }
                }
                //外接冷却水冲洗量
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColCoolingWaterFlow"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "外接冷却水冲洗量" + "不能为空" + "\n";
                        break;
                    }
                }
            }
            //PNT Selected
            if (cbPNT.Checked)
            {
                //Customer PO NO客户订单号
                string strChkCustomerPONO = txtCustomerPONO.Text;
                if (string.IsNullOrEmpty(strChkCustomerPONO))
                {
                    strError = strError + "客户订单号" + "不能为空" + "\n";
                }
                //Sulzer Order NO/苏尔寿订单号
                string strChkSulzerOrderNO = txtSulzerOrderNO.Text;
                if (string.IsNullOrEmpty(strChkSulzerOrderNO))
                {
                    strError = strError + "苏尔寿订单号" + "不能为空" + "\n";
                }
                //ITEM No/行号
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColItemNo"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "ITEM No/行号" + "不能为空" + "\n";
                        break;
                    }
                }
                //Series No/苏尔寿编号
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSeriesNo"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "Series No/苏尔寿编号" + "不能为空" + "\n";
                        break;
                    }
                }
                //TAG NO/客户位号
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColTagNo"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "TAG NO/客户位号" + "不能为空" + "\n";
                        break;
                    }
                }
                //TYPE/泵型
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColType"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "TYPE/泵型" + "不能为空" + "\n";
                        break;
                    }
                }
                //SIZE/规格
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSize"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SIZE/规格" + "不能为空" + "\n";
                        break;
                    }
                }
                //SERVICE / 泵名称
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColService"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SERVICE/泵名称" + "不能为空" + "\n";
                        break;
                    }
                }
                //FLOW / 流量 (m3/h)
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColFlow"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "FLOW/流量(m3/h)" + "不能为空" + "\n";
                        break;
                    }
                }
                //HEAD / 扬程 (m)
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColHead"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "HEAD/扬程(m)" + "不能为空" + "\n";
                        break;
                    }
                }
                //RATED POWER / 额定功率(kW)
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColRatedPower"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "RATED POWER/额定功率(kW)" + "不能为空" + "\n";
                        break;
                    }
                }
                //NPSHr  (m)
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColNpshr"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "NPSHr(m)" + "不能为空" + "\n";
                        break;
                    }
                }
                //MAWP (Mpa)
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColMawp"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "MAWP(Mpa)" + "不能为空" + "\n";
                        break;
                    }
                }
                //MAWT (℃)
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColMawt"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "MAWT(℃)" + "不能为空" + "\n";
                        break;
                    }
                }
                //Hydrostatic P (Mpa)
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColHydrostaticP"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "Hydrostatic P(Mpa)" + "不能为空" + "\n";
                        break;
                    }
                }
                //MATERIAL CLASS / 材料等级
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColMaterialClass"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "MATERIAL CLASS/材料等级" + "不能为空" + "\n";
                        break;
                    }
                }
                //BEARING / 轴承配置
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColBearing"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "BEARING/轴承配置" + "不能为空" + "\n";
                        break;
                    }
                }
                //轴承保持架材质
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColBearingCageMaterial"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "轴承保持架材质" + "不能为空" + "\n";
                        break;
                    }
                }
            }
            //PGA Selected
            if (cbPGA.Checked)
            {
                //TYPE/泵型
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColType"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "TYPE/泵型" + "不能为空" + "\n";
                        break;
                    }
                }
                //SIZE/规格
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSize"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SIZE/规格" + "不能为空" + "\n";
                        break;
                    }
                }
                //FRAME SIZE / 基座号
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColFrameSize"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "FRAME SIZE/基座号" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANA/冲洗方案
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanA"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANA/冲洗方案A" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANB/冲洗方案B
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanB"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANA/冲洗方案B" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANC/冲洗方案C
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanC"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANC/冲洗方案C" + "不能为空" + "\n";
                        break;
                    }
                }
                //FLANGE STANDARD / 法兰标准
                if (string.IsNullOrEmpty(cobFlangeStandard.SelectedValue.ToString()))
                {
                    strError = strError + "FLANGE STANDARD/法兰标准" + "不能为空" + "\n";
                }
                //排残管口口径
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColPumpDrainNozzleSize"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "排残管口口径" + "不能为空" + "\n";
                        break;
                    }
                }
                //排残管路阀门
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColPumpDrainValve"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "排残管路阀门" + "不能为空" + "\n";
                        break;
                    }
                }
                //暖泵管路
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColWarmUpPiping"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "暖泵管路" + "不能为空" + "\n";
                        break;
                    }
                }
                //COMPANION FLG. / 配对法兰要求
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColCompanionFlg"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "COMPANION FLG./配对法兰要求" + "不能为空" + "\n";
                        break;
                    }
                }
                //壳体是否带夹套
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColCaseBracket"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "壳体是否带夹套" + "不能为空" + "\n";
                        break;
                    }
                }
                //预留油雾润滑
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColReservedPureLubrication"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "预留油雾润滑" + "不能为空" + "\n";
                        break;
                    }
                }
                //轴承室带排油管路
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColBearighousingDrain"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "轴承室带排油管路" + "不能为空" + "\n";
                        break;
                    }
                }
                //PRE泵压力等级
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strType = (String)dgvSumList.Rows[i].Cells["ColType"].Value;
                    if (!strType.Contains("PR"))
                    {
                        break;
                    }
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColPREPumpPressureClass"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "PRE泵压力等级" + "不能为空" + "\n";
                        break;
                    }
                }
                //排残阀类型
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColDrainValveType"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "排残阀类型" + "不能为空" + "\n";
                        break;
                    }
                }

            }
            //PCS Selected
            if (cbPCS.Checked)
            {
                //TYPE/泵型
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColType"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "TYPE/泵型" + "不能为空" + "\n";
                        break;
                    }
                }
                //SIZE/规格
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSize"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SIZE/规格" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANA/冲洗方案
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanA"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANA/冲洗方案A" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANB/冲洗方案B
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanB"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANA/冲洗方案B" + "不能为空" + "\n";
                        break;
                    }
                }
                //SEAL PLANC/冲洗方案C
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColSealPlanC"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "SEAL PLANC/冲洗方案C" + "不能为空" + "\n";
                        break;
                    }
                }
                //MATERIAL CLASS / 材料等级
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColMaterialClass"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "MATERIAL CLASS/材料等级" + "不能为空" + "\n";
                        break;
                    }
                }
                //干湿轴
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColDryWetShaft"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "干湿轴" + "不能为空" + "\n";
                        break;
                    }
                    else
                    {
                        if(strChk.Equals("Wet"))
                        {
                            string strChkInpu = (String)dgvSumList.Rows[i].Cells["ColWetInput"].Value;
                            if(string.IsNullOrEmpty(strChkInpu))
                            {
                                strError = strError + "湿轴的情况下,Wet Shaft(自己填写)" + "不能为空" + "\n";
                                break;
                            }
                        }
                    }
                }
                //BEARING / 轴承配置
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColBearing"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "BEARING/轴承配置" + "不能为空" + "\n";
                        break;
                    }
                }
                //轴承保持架材质
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColBearingCageMaterial"].Value;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "轴承保持架材质" + "不能为空" + "\n";
                        break;
                    }
                }
                //壳体是否带夹套
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColCaseBracket"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "壳体是否带夹套" + "不能为空" + "\n";
                        break;
                    }
                }
                //预留油雾润滑
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColReservedPureLubrication"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "预留油雾润滑" + "不能为空" + "\n";
                        break;
                    }
                }
                //叶轮带平衡孔
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColImpellerBalanceHole"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "叶轮带平衡孔" + "不能为空" + "\n";
                        break;
                    }
                }
                //是否带诱导轮
                for (int i = 0; i < dgvSumList.Rows.Count - 1; i++)
                {
                    string strType = (String)dgvSumList.Rows[i].Cells["ColType"].Value;
                    if(!strType.Contains("PR"))
                    {
                        break;
                    }
                    string strChk = (String)dgvSumList.Rows[i].Cells["ColPREInducerYesNo"].FormattedValue;
                    if (string.IsNullOrEmpty(strChk))
                    {
                        strError = strError + "是否带诱导轮" + "不能为空" + "\n";
                        break;
                    }
                }
            }
            return strError;
        }
        /// <summary>
        /// number check
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public bool CheckNumber(string number)
        {
            bool bolFlag = false;
            if (!string.IsNullOrEmpty(number))
            {
                Regex regex = new Regex(@"^(-)?\d+(\.\d+)?$");
                if (regex.IsMatch(number))
                {
                    bolFlag = true;
                }
                else
                {
                    bolFlag = false;
                }
            }
            return bolFlag;
        }

        /// <summary>
        /// PPI Selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CbPPI_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPPI.Checked == true)
            {
                CbPPIChecked();
            }
            else
            {
                //ChangeColor White
                ChangeColorWhite();
                if(cbPUC.Checked == true)
                {
                    CbPUCChecked();
                }
                if(cbPNT.Checked == true)
                {
                    CbPNTChecked();
                }
                if (cbPDS.Checked == true)
                {
                    CbPDSChecked();
                }
                if (cbPSL.Checked == true)
                {
                    CbPSLChecked();
                }
                if (cbPGA.Checked == true)
                {
                    CbPGAChecked();
                }
                if (cbPCS.Checked == true)
                {
                    CbPCSChecked();
                }
            }
        }
        /// <summary>
        /// PUC Selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CbPUC_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPUC.Checked == true)
            {
                CbPUCChecked();
            }
            else
            {
                //ChangeColor White
                ChangeColorWhite();
                if (cbPPI.Checked == true)
                {
                    CbPPIChecked();
                }
                if (cbPNT.Checked == true)
                {
                    CbPNTChecked();
                }
                if (cbPDS.Checked == true)
                {
                    CbPDSChecked();
                }
                if (cbPSL.Checked == true)
                {
                    CbPSLChecked();
                }
                if (cbPGA.Checked == true)
                {
                    CbPGAChecked();
                }
                if (cbPCS.Checked == true)
                {
                    CbPCSChecked();
                }
            }
        }
        /// <summary>
        /// PNT Selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CbPNT_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPNT.Checked == true)
            {
                CbPNTChecked();
            }
            else
            {
                //ChangeColor White
                ChangeColorWhite();
                if (cbPPI.Checked == true)
                {
                    CbPPIChecked();
                }
                if (cbPUC.Checked == true)
                {
                    CbPUCChecked();
                }
                if (cbPDS.Checked == true)
                {
                    CbPDSChecked();
                }
                if (cbPSL.Checked == true)
                {
                    CbPSLChecked();
                }
                if (cbPGA.Checked == true)
                {
                    CbPGAChecked();
                }
                if (cbPCS.Checked == true)
                {
                    CbPCSChecked();
                }
            }
        }
        /// <summary>
        /// PDS Selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CbPDS_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPDS.Checked == true)
            {
                CbPDSChecked();
            }
            else
            {
                //ChangeColor White
                ChangeColorWhite();
                if (cbPPI.Checked == true)
                {
                    CbPPIChecked();
                }
                if (cbPUC.Checked == true)
                {
                    CbPUCChecked();
                }
                if (cbPNT.Checked == true)
                {
                    CbPNTChecked();
                }
                if (cbPSL.Checked == true)
                {
                    CbPSLChecked();
                }
                if (cbPGA.Checked == true)
                {
                    CbPGAChecked();
                }
                if (cbPCS.Checked == true)
                {
                    CbPCSChecked();
                }
            }
        }
        /// <summary>
        /// PSL Selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CbPSL_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPSL.Checked == true)
            {
                CbPSLChecked();
            }
            else
            {
                //ChangeColor White
                ChangeColorWhite();
                if (cbPPI.Checked == true)
                {
                    CbPPIChecked();
                }
                if (cbPUC.Checked == true)
                {
                    CbPUCChecked();
                }
                if (cbPNT.Checked == true)
                {
                    CbPNTChecked();
                }
                if (cbPDS.Checked == true)
                {
                    CbPDSChecked();
                }
                if (cbPGA.Checked == true)
                {
                    CbPGAChecked();
                }
                if (cbPCS.Checked == true)
                {
                    CbPCSChecked();
                }
            }
        }
        /// <summary>
        /// PGA Selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CbPGA_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPGA.Checked == true)
            {
                CbPGAChecked();
            }
            else
            {
                //ChangeColor White
                ChangeColorWhite();
                if (cbPPI.Checked == true)
                {
                    CbPPIChecked();
                }
                if (cbPUC.Checked == true)
                {
                    CbPUCChecked();
                }
                if (cbPNT.Checked == true)
                {
                    CbPNTChecked();
                }
                if (cbPDS.Checked == true)
                {
                    CbPDSChecked();
                }
                if (cbPSL.Checked == true)
                {
                    CbPSLChecked();
                }
                if (cbPCS.Checked == true)
                {
                    CbPCSChecked();
                }
            }
        }
        /// <summary>
        /// PCS Selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CbPCS_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPCS.Checked == true)
            {
                CbPCSChecked();
            }
            else
            {
                //ChangeColor White
                ChangeColorWhite();
                if (cbPPI.Checked == true)
                {
                    CbPPIChecked();
                }
                if (cbPUC.Checked == true)
                {
                    CbPUCChecked();
                }
                if (cbPNT.Checked == true)
                {
                    CbPNTChecked();
                }
                if (cbPDS.Checked == true)
                {
                    CbPDSChecked();
                }
                if (cbPSL.Checked == true)
                {
                    CbPSLChecked();
                }
                if (cbPGA.Checked == true)
                {
                    CbPGAChecked();
                }
            }
        }
        /// <summary>
        /// PPI change colour
        /// </summary>
        private void CbPPIChecked()
        {
            //TYPE / 泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案A
            dgvSumList.Columns["ColSealPlanA"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案B
            dgvSumList.Columns["ColSealPlanB"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案C
            dgvSumList.Columns["ColSealPlanC"].DefaultCellStyle.BackColor = Color.Gold;
            //排残管路 阀门
            dgvSumList.Columns["ColPumpDrainValve"].DefaultCellStyle.BackColor = Color.Gold;
            //暖泵管路
            dgvSumList.Columns["ColWarmUpPiping"].DefaultCellStyle.BackColor = Color.Gold;
            //BEARING / 轴承配置
            dgvSumList.Columns["ColBearing"].DefaultCellStyle.BackColor = Color.Gold;
            //COMPANION FLG. / 配对法兰要求
            dgvSumList.Columns["ColCompanionFlg"].DefaultCellStyle.BackColor = Color.Gold;
            //壳体是否带夹套
            dgvSumList.Columns["ColCaseBracket"].DefaultCellStyle.BackColor = Color.Gold;
            //轴承室带排油管路
            dgvSumList.Columns["ColBearighousingDrain"].DefaultCellStyle.BackColor = Color.Gold;
            //PRE泵压力等级
            dgvSumList.Columns["ColPREPumpPressureClass"].DefaultCellStyle.BackColor = Color.Gold;

            //Purchaser/买方
            txtPurchaser.BackColor = Color.Gold;
            //End User:/用户
            txtEndUser.BackColor = Color.Gold;
            //Project/项目
            txtProject.BackColor = Color.Gold;
            //Site/安装地点
            txtSite.BackColor = Color.Gold;
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.Gold;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.Gold;
            //Prepared by/编制
            txtPreparedBy.BackColor = Color.Gold;
            //Reviewed by/审核
            txtReviewedBy.BackColor = Color.Gold;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.Gold;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.Gold;
            //FLANGE STANDARD / 法兰标准
            cobFlangeStandard.BackColor = Color.Gold;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.Gold;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.Gold;
            //API 682 EDITION / API 682 版本
            dgvSumList.Columns["ColApiEdition"].DefaultCellStyle.BackColor = Color.Gold;
            //排残管口口径
            dgvSumList.Columns["ColPumpDrainNozzleSize"].DefaultCellStyle.BackColor = Color.Gold;
            //MATERIAL CLASS / 材料等级
            dgvSumList.Columns["ColMaterialClass"].DefaultCellStyle.BackColor = Color.Gold;
        }
        /// <summary>
        /// PUC change colour
        /// </summary>
        private void CbPUCChecked()
        {
            //TYPE / 泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE / 规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //MOTOR POWER/ 电机功率(kW) 
            dgvSumList.Columns["ColMotorPower"].DefaultCellStyle.BackColor = Color.Gold;
            //VOLTAGE/ 电压(V) 
            dgvSumList.Columns["ColVoltage"].DefaultCellStyle.BackColor = Color.Gold;
            //密封冲洗量
            dgvSumList.Columns["ColSealFlushFlow"].DefaultCellStyle.BackColor = Color.Gold;
            //外接冷却水冲洗量
            dgvSumList.Columns["ColCoolingWaterFlow"].DefaultCellStyle.BackColor = Color.Gold;

            //Purchaser/买方
            txtPurchaser.BackColor = Color.Gold;
            //End User:/用户
            txtEndUser.BackColor = Color.Gold;
            //Project/项目
            txtProject.BackColor = Color.Gold;
            //Site/安装地点
            txtSite.BackColor = Color.Gold;
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.Gold;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.Gold;
            //Prepared by/编制
            txtPreparedBy.BackColor = Color.Gold;
            //Reviewed by/审核
            txtReviewedBy.BackColor = Color.Gold;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.Gold;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.Gold;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.Gold;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.Gold;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.Gold;
        }
        /// <summary>
        /// PNT change colour
        /// </summary>
        private void CbPNTChecked()
        {
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.Gold;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.Gold;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.Gold;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TYPE/泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.Gold;
            //转数 r/min
            dgvSumList.Columns["ColSpeed"].DefaultCellStyle.BackColor = Color.Gold;
            //FLOW / 流量 (m3/h)
            dgvSumList.Columns["ColFlow"].DefaultCellStyle.BackColor = Color.Gold;
            //HEAD / 扬程 (m)
            dgvSumList.Columns["ColHead"].DefaultCellStyle.BackColor = Color.Gold;
            //RATED POWER / 额定功率(kW)
            dgvSumList.Columns["ColRatedPower"].DefaultCellStyle.BackColor = Color.Gold;
            //NPSHr  (m)
            dgvSumList.Columns["ColNpshr"].DefaultCellStyle.BackColor = Color.Gold;
            //MAWP (Mpa)
            dgvSumList.Columns["ColMawp"].DefaultCellStyle.BackColor = Color.Gold;
            //MAWT (℃)
            dgvSumList.Columns["ColMawt"].DefaultCellStyle.BackColor = Color.Gold;
            //Hydrostatic P (Mpa)
            dgvSumList.Columns["ColHydrostaticP"].DefaultCellStyle.BackColor = Color.Gold;
            //MATERIAL CLASS / 材料等级
            dgvSumList.Columns["ColMaterialClass"].DefaultCellStyle.BackColor = Color.Gold;
            //BEARING / 轴承配置
            dgvSumList.Columns["ColBearing"].DefaultCellStyle.BackColor = Color.Gold;
            //轴承保持架材质
            dgvSumList.Columns["ColBearingCageMaterial"].DefaultCellStyle.BackColor = Color.Gold;

            //Prepared by/编制
            txtPreparedBy.BackColor = Color.Gold;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.Gold;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.Gold;
        }
        /// <summary>
        /// PDS change colour
        /// </summary>
        private void CbPDSChecked()
        {
            //Purchaser/买方
            txtPurchaser.BackColor = Color.Gold;
            //End User:/用户
            txtEndUser.BackColor = Color.Gold;
            //Project/项目
            txtProject.BackColor = Color.Gold;
            //Site/安装地点
            txtSite.BackColor = Color.Gold;
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.Gold;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.Gold;
            //Prepared by/编制
            txtPreparedBy.BackColor = Color.Gold;
            //Reviewed by/审核
            txtReviewedBy.BackColor = Color.Gold;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.Gold;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.Gold;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.Gold;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TYPE/泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.Gold;
        }
        /// <summary>
        /// PSL change colour
        /// </summary>
        private void CbPSLChecked()
        {
            //Purchaser/买方
            txtPurchaser.BackColor = Color.Gold;
            //End User:/用户
            txtEndUser.BackColor = Color.Gold;
            //Project/项目
            txtProject.BackColor = Color.Gold;
            //Site/安装地点
            txtSite.BackColor = Color.Gold;
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.Gold;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.Gold;
            //Prepared by/编制
            txtPreparedBy.BackColor = Color.Gold;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.Gold;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.Gold;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.Gold;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TYPE/泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.Gold;
            //Qty / 数量
            dgvSumList.Columns["ColQty"].DefaultCellStyle.BackColor = Color.Gold;
        }
        /// <summary>
        /// PCS change colour
        /// </summary>
        private void CbPCSChecked()
        {
            //TYPE/泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案A
            dgvSumList.Columns["ColSealPlanA"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案B
            dgvSumList.Columns["ColSealPlanB"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案C
            dgvSumList.Columns["ColSealPlanC"].DefaultCellStyle.BackColor = Color.Gold;
            //MATERIAL CLASS / 材料等级
            dgvSumList.Columns["ColMaterialClass"].DefaultCellStyle.BackColor = Color.Gold;
            //干湿轴
            dgvSumList.Columns["ColDryWetShaft"].DefaultCellStyle.BackColor = Color.Gold;
            //轴承保持架材质
            dgvSumList.Columns["ColBearingCageMaterial"].DefaultCellStyle.BackColor = Color.Gold;
            //壳体是否带夹套
            dgvSumList.Columns["ColCaseBracket"].DefaultCellStyle.BackColor = Color.Gold;
            //预留油雾润滑
            dgvSumList.Columns["ColReservedPureLubrication"].DefaultCellStyle.BackColor = Color.Gold;
            //叶轮带平衡孔
            dgvSumList.Columns["ColImpellerBalanceHole"].DefaultCellStyle.BackColor = Color.Gold;
            //密封冲洗量
            //dgvSumList.Columns["ColSealFlushFlow"].DefaultCellStyle.BackColor = Color.Gold;
            //外接冷却水冲洗量
            //dgvSumList.Columns["ColCoolingWaterFlow"].DefaultCellStyle.BackColor = Color.Gold;
            //是否带诱导轮
            dgvSumList.Columns["ColPREInducerYesNo"].DefaultCellStyle.BackColor = Color.Gold;

            //Purchaser/买方
            txtPurchaser.BackColor = Color.Gold;
            //End User:/用户
            txtEndUser.BackColor = Color.Gold;
            //Project/项目
            txtProject.BackColor = Color.Gold;
            //Site/安装地点
            txtSite.BackColor = Color.Gold;
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.Gold;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.Gold;
            //Prepared by/编制
            txtPreparedBy.BackColor = Color.Gold;
            //Reviewed by/审核
            txtReviewedBy.BackColor = Color.Gold;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.Gold;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.Gold;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.Gold;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TYPE/泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.Gold;

        }
        /// <summary>
        /// PGA change colour
        /// </summary>
        private void CbPGAChecked()
        {
            //TYPE/泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.Gold;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.Gold;
            //电机型号
            dgvSumList.Columns["ColMotorFrame"].DefaultCellStyle.BackColor = Color.Gold;
            //FRAME SIZE / 基座号
            dgvSumList.Columns["ColFrameSize"].DefaultCellStyle.BackColor = Color.Gold;
            //MOTOR POWER/ 电机功率(kW) 
            dgvSumList.Columns["ColMotorPower"].DefaultCellStyle.BackColor = Color.Gold;
            //VOLTAGE/ 电压(V) 
            dgvSumList.Columns["ColVoltage"].DefaultCellStyle.BackColor = Color.Gold;
            //转数 r/min
            dgvSumList.Columns["ColSpeed"].DefaultCellStyle.BackColor = Color.Gold;
            //电机长度K
            dgvSumList.Columns["ColMotorLengthK"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案A
            dgvSumList.Columns["ColSealPlanA"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案B
            dgvSumList.Columns["ColSealPlanB"].DefaultCellStyle.BackColor = Color.Gold;
            //SEAL PLAN / 冲洗方案C
            dgvSumList.Columns["ColSealPlanC"].DefaultCellStyle.BackColor = Color.Gold;
            //FLANGE STANDARD / 法兰标准
            cobFlangeStandard.BackColor = Color.Gold;
            //排残管口口径
            dgvSumList.Columns["ColPumpDrainNozzleSize"].DefaultCellStyle.BackColor = Color.Gold;
            //排残管路 阀门
            dgvSumList.Columns["ColPumpDrainValve"].DefaultCellStyle.BackColor = Color.Gold;
            //暖泵管路
            dgvSumList.Columns["ColWarmUpPiping"].DefaultCellStyle.BackColor = Color.Gold;
            //MATERIAL CLASS / 材料等级
            dgvSumList.Columns["ColMaterialClass"].DefaultCellStyle.BackColor = Color.Gold;
            //COMPANION FLG. / 配对法兰要求
            dgvSumList.Columns["ColCompanionFlg"].DefaultCellStyle.BackColor = Color.Gold;
            //壳体是否带夹套
            dgvSumList.Columns["ColCaseBracket"].DefaultCellStyle.BackColor = Color.Gold;
            //管口载荷要求
            dgvSumList.Columns["ColNozzleLoad"].DefaultCellStyle.BackColor = Color.Gold;
            //预留油雾润滑
            dgvSumList.Columns["ColReservedPureLubrication"].DefaultCellStyle.BackColor = Color.Gold;
            //轴承室带排油管路
            dgvSumList.Columns["ColBearighousingDrain"].DefaultCellStyle.BackColor = Color.Gold;
            //PRE泵压力等级
            dgvSumList.Columns["ColPREPumpPressureClass"].DefaultCellStyle.BackColor = Color.Gold;
            //冷却水母管口径
            dgvSumList.Columns["ColCoolWaterMainPipeSize"].DefaultCellStyle.BackColor = Color.Gold;
            //排残阀类型
            dgvSumList.Columns["ColDrainValveType"].DefaultCellStyle.BackColor = Color.Gold;

            //Purchaser/买方
            txtPurchaser.BackColor = Color.Gold;
            //End User:/用户
            txtEndUser.BackColor = Color.Gold;
            //Project/项目
            txtProject.BackColor = Color.Gold;
            //Site/安装地点
            txtSite.BackColor = Color.Gold;
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.Gold;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.Gold;
            //Prepared by/编制
            txtPreparedBy.BackColor = Color.Gold;
            //Reviewed by/审核
            txtReviewedBy.BackColor = Color.Gold;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.Gold;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.Gold;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.Gold;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.Gold;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.Gold;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.Gold;
        }
        /// <summary>
        /// All controls change white
        /// </summary>
        private void ChangeColorWhite()
        {
            //Purchaser/买方
            txtPurchaser.BackColor = Color.White;
            //End User:/用户
            txtEndUser.BackColor = Color.White;
            //Project/项目
            txtProject.BackColor = Color.White;
            //Site/安装地点
            txtSite.BackColor = Color.White;
            //Customer PO NO客户订单号
            txtCustomerPONO.BackColor = Color.White;
            //Sulzer Order NO/苏尔寿订单号
            txtSulzerOrderNO.BackColor = Color.White;
            //Prepared by/编制
            txtPreparedBy.BackColor = Color.White;
            //Reviewed by/审核
            txtReviewedBy.BackColor = Color.White;
            //Approved by/批准
            txtApprovedBy.BackColor = Color.White;
            //Date/日期
            dateTimeWorkDate.BackColor = Color.White;
            //FLANGE STANDARD / 法兰标准
            cobFlangeStandard.BackColor = Color.White;
            //ITEM No/行号
            dgvSumList.Columns["ColItemNo"].DefaultCellStyle.BackColor = Color.White;
            //Series No/苏尔寿编号
            dgvSumList.Columns["ColSeriesNo"].DefaultCellStyle.BackColor = Color.White;
            //TAG NO/客户位号
            dgvSumList.Columns["ColTagNo"].DefaultCellStyle.BackColor = Color.White;
            //TYPE/泵型
            dgvSumList.Columns["ColType"].DefaultCellStyle.BackColor = Color.White;
            //SIZE/规格
            dgvSumList.Columns["ColSize"].DefaultCellStyle.BackColor = Color.White;
            //SERVICE / 泵名称
            dgvSumList.Columns["ColService"].DefaultCellStyle.BackColor = Color.White;
            //Qty / 数量
            dgvSumList.Columns["ColQty"].DefaultCellStyle.BackColor = Color.White;
            //MOTOR MANUFACTURER / 电机供应商
            dgvSumList.Columns["ColMotorManufacturer"].DefaultCellStyle.BackColor = Color.White;
            //电机型号
            dgvSumList.Columns["ColMotorFrame"].DefaultCellStyle.BackColor = Color.White;
            //FRAME SIZE/基座号
            dgvSumList.Columns["ColFrameSize"].DefaultCellStyle.BackColor = Color.White;
            //MOTOR POWER/ 电机功率(kW) 
            dgvSumList.Columns["ColMotorPower"].DefaultCellStyle.BackColor = Color.White;
            //VOLTAGE/ 电压(V) 
            dgvSumList.Columns["ColVoltage"].DefaultCellStyle.BackColor = Color.White;
            //转数 r/min
            dgvSumList.Columns["ColSpeed"].DefaultCellStyle.BackColor = Color.White;
            //电机长度K
            dgvSumList.Columns["ColMotorLengthK"].DefaultCellStyle.BackColor = Color.White;
            //API 682 EDITION / API 682 版本
            dgvSumList.Columns["ColApiEdition"].DefaultCellStyle.BackColor = Color.White;
            //SEAL PLAN A/ 冲洗方案
            dgvSumList.Columns["ColSealPlanA"].DefaultCellStyle.BackColor = Color.White;
            //SEAL PLAN B/ 冲洗方案
            dgvSumList.Columns["ColSealPlanB"].DefaultCellStyle.BackColor = Color.White;
            //SEAL PLAN C/ 冲洗方案
            dgvSumList.Columns["ColSealPlanC"].DefaultCellStyle.BackColor = Color.White;
            //FLOW / 流量 (m3/h)
            dgvSumList.Columns["ColFlow"].DefaultCellStyle.BackColor = Color.White;
            //HEAD / 扬程 (m)
            dgvSumList.Columns["ColHead"].DefaultCellStyle.BackColor = Color.White;
            //EFFICIENCY / 效率(%)
            dgvSumList.Columns["ColEfficiency"].DefaultCellStyle.BackColor = Color.White;
            //RATED POWER / 额定功率(kW)
            dgvSumList.Columns["ColRatedPower"].DefaultCellStyle.BackColor = Color.White;
            //NPSHr  (m)
            dgvSumList.Columns["ColNpshr"].DefaultCellStyle.BackColor = Color.White;
            //MAWP (Mpa)
            dgvSumList.Columns["ColMawp"].DefaultCellStyle.BackColor = Color.White;
            //MAWT (℃)
            dgvSumList.Columns["ColMawt"].DefaultCellStyle.BackColor = Color.White;
            //Hydrostatic P (Mpa)
            dgvSumList.Columns["ColHydrostaticP"].DefaultCellStyle.BackColor = Color.White;
            //排残管口口径
            dgvSumList.Columns["ColPumpDrainNozzleSize"].DefaultCellStyle.BackColor = Color.White;
            //排残管路 阀门
            dgvSumList.Columns["ColPumpDrainValve"].DefaultCellStyle.BackColor = Color.White;
            //暖泵管路
            dgvSumList.Columns["ColWarmUpPiping"].DefaultCellStyle.BackColor = Color.White;
            //MATERIAL CLASS / 材料等级
            dgvSumList.Columns["ColMaterialClass"].DefaultCellStyle.BackColor = Color.White;
            //干湿轴
            dgvSumList.Columns["ColDryWetShaft"].DefaultCellStyle.BackColor = Color.White;
            //BEARING / 轴承配置
            dgvSumList.Columns["ColBearing"].DefaultCellStyle.BackColor = Color.White;
            //轴承保持架材质
            dgvSumList.Columns["ColBearingCageMaterial"].DefaultCellStyle.BackColor = Color.White;
            //COMPANION FLG. / 配对法兰要求
            dgvSumList.Columns["ColCompanionFlg"].DefaultCellStyle.BackColor = Color.White;
            //壳体是否带夹套
            dgvSumList.Columns["ColCaseBracket"].DefaultCellStyle.BackColor = Color.White;
            //管口载荷要求
            dgvSumList.Columns["ColNozzleLoad"].DefaultCellStyle.BackColor = Color.White;
            //预留油雾润滑
            dgvSumList.Columns["ColReservedPureLubrication"].DefaultCellStyle.BackColor = Color.White;
            //轴承室带排油管路
            dgvSumList.Columns["ColBearighousingDrain"].DefaultCellStyle.BackColor = Color.White;
            //叶轮带平衡孔
            dgvSumList.Columns["ColImpellerBalanceHole"].DefaultCellStyle.BackColor = Color.White;
            //密封冲洗量
            dgvSumList.Columns["ColSealFlushFlow"].DefaultCellStyle.BackColor = Color.White;
            //外接冷却水冲洗量
            dgvSumList.Columns["ColCoolingWaterFlow"].DefaultCellStyle.BackColor = Color.White;
            //是否带诱导轮
            dgvSumList.Columns["ColPREInducerYesNo"].DefaultCellStyle.BackColor = Color.White;
            //PRE泵压力等级
            dgvSumList.Columns["ColPREPumpPressureClass"].DefaultCellStyle.BackColor = Color.White;
            //冷却水母管口径
            dgvSumList.Columns["ColCoolWaterMainPipeSize"].DefaultCellStyle.BackColor = Color.White;
            //排残阀类型
            dgvSumList.Columns["ColDrainValveType"].DefaultCellStyle.BackColor = Color.White;
        }
        /// <summary>
        /// Clear按钮处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClear_Click(object sender, EventArgs e)
        {
            ClearAllControls();
            ChangeColorWhite();
        }
        /// <summary>
        /// Initialize, empty all controls
        /// </summary>
        private void ClearAllControls()
        {
            //Clear old data
            txtSumListFilePath.Clear();
            textXMLPath.Clear();
            cbPGA.Checked = false;
            cbPPI.Checked = false;
            cbPCS.Checked = false;
            cbPUC.Checked = false;
            cbPNT.Checked = false;
            cbPDS.Checked = false;
            cbPSL.Checked = false;
            txtPurchaser.Clear();
            txtEndUser.Clear();
            txtProject.Clear();
            txtSite.Clear();
            txtCustomerPONO.Clear();
            txtSulzerOrderNO.Clear();
            txtPreparedBy.Clear();
            txtReviewedBy.Clear();
            txtApprovedBy.Clear();
            dateTimeWorkDate.Value = System.DateTime.Now;
            cobFlangeStandard.SelectedValue = "";
            dgvSumList.Rows.Clear();
            txtXMLFilePath.Clear();
            dgvSumList.Columns["ColVoluteCasing"].Visible = false;
            dgvSumList.Columns["ColImpeller"].Visible = false;
            dgvSumList.Columns["ColStationaryWearRing"].Visible = false;
            dgvSumList.Columns["ColCoverCoolingPlate"].Visible = false;
            dgvSumList.Columns["ColThroatBush"].Visible = false;
            dgvSumList.Columns["ColImpellerKey"].Visible = false;
            dgvSumList.Columns["ColGasket"].Visible = false;
            dgvSumList.Columns["ColSpacerSleeve"].Visible = false;
            dgvSumList.Columns["ColImpellerNut"].Visible = false;
            dgvSumList.Columns["ColInducer"].Visible = false;
            dgvSumList.Columns["ColInducerSleeve"].Visible = false;
            dgvSumList.Columns["ColWearPlate"].Visible = false;
        }
        /// <summary>
        /// close button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        /// <summary>
        /// XML file save path open
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnXMLPathOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "Please select a XML file";
            dialog.Filter = "XML File|*.xml*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = dialog.FileName;
                textXMLPath.Text = savePath;
            }
        }
        /// <summary>
        /// XMLDataImport Button down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnXMLDataImport_Click(object sender, EventArgs e)
        {
            if (textXMLPath.Text == "")
            {
                LogHelper.GetSingleObj().Error("XML文件导入", "Please select the XML file!");
                MessageBox.Show("Please select the XML file!", "ERROR");
                return;
            }
            try
            {
                BtnToEnabled(false);
                string strXMLPath = textXMLPath.Text;
                DataSet dataSet = XmlDataOperation.XMLFileToDataSet(strXMLPath);
                DataTable dtHead = dataSet.Tables["DataTableHead"];
                DataTable dtBody = dataSet.Tables["DataTableBody"];
                txtSumListFilePath.Clear();
                txtXMLFilePath.Clear();
                //all data clear
                ClearAllControls();
                //read XML file to dgvSumList
                DataBindingFromXML(dtHead, dtBody);
            }
            catch(Exception ex)
            {
                LogHelper.GetSingleObj().Error("XML文件导入", ex.Message);
                BtnToEnabled(true);
                MessageBox.Show("XML文件导入:" + ex.Message, "ERROR");
                return;
            }
            //Cursor release
            dgvSumList.CurrentCell = null;
            LogHelper.GetSingleObj().Info("XML文件导入", "XML文件导入成功");
            BtnToEnabled(true);
            MessageBox.Show("Read the XML file successfully!");
        }
        /// <summary>
        /// Paste data from clipboard to DataGridView
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //paste data
            PasteData(dgvSumList);
        }
        /// <summary>
        /// Removes the specified row
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dgr in dgvSumList.SelectedRows)
            {
                //Removes the specified row from the collection
                this.dgvSumList.Rows.RemoveAt(dgr.Index);
            }
        }
        /// <summary>
        /// insert new line
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ArrayList index = new ArrayList();
            int i = 0;
            int j = 1;
            foreach (DataGridViewRow dgr in dgvSumList.SelectedRows)
            {
                index.Add(dgr.Index);
            }
            index.Sort();
            foreach (object ite in index)
            {
                if (i != 0)
                {
                    if (Convert.ToInt32(index[i]) - Convert.ToInt32(index[i - 1]) == 1)
                    {
                        this.dgvSumList.Rows.Insert(Convert.ToInt32(index[i - j]) + i);
                        j++;
                    }
                    else
                    {
                        this.dgvSumList.Rows.Insert(Convert.ToInt32(index[i]) + i);
                        j = 1;
                    }
                }
                else
                {
                    this.dgvSumList.Rows.Insert(Convert.ToInt32(index[i]) + i);
                }
                i++;
            }
        }
        /// <summary>
        /// Press the right mouse button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DgvSumList_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    //this.dgvSumList.Rows[e.RowIndex].Selected = true;//是否选中当前行
                    //this.dgvSumList.CurrentCell = this.dgvSumList.Rows[e.RowIndex].Cells[0];
                    //this.contextMenuStripGDV.Show(this.dgvSumList, e.Location);
                    //this.contextMenuStripGDV.Show(Cursor.Position);//锁定右键列表出现的位置
                    dgvSumList.CurrentCell = dgvSumList.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    contextMenuStripGDV.Show(MousePosition.X, MousePosition.Y);

                }
            }
        }
        /// <summary>
        /// Event in form close
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SumListForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //No data saved
            if (CLOSE_FLAG == 0)
            {
                if (MessageBox.Show("The data has not been saved. Are you sure to close it?", 
                    "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, 
                    MessageBoxDefaultButton.Button2) == DialogResult.OK)
                {
                    e.Cancel = false;
                }
                else
                {
                    e.Cancel = true;
                }
            }
            //data saved
            else
            {
                if (cbPGA.Checked == true)
                {
                    GlobalVariableController.GlobalButtonFlag[1] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[1] = false;
                }
                if (cbPPI.Checked == true)
                {
                    GlobalVariableController.GlobalButtonFlag[2] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[2] = false;
                }
                if (cbPCS.Checked == true)
                {
                    GlobalVariableController.GlobalButtonFlag[3] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[3] = false;
                }
                if (cbPUC.Checked == true)
                {
                    GlobalVariableController.GlobalButtonFlag[4] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[4] = false;
                }
                if (cbPNT.Checked == true)
                {
                    GlobalVariableController.GlobalButtonFlag[5] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[5] = false;
                }
                if (cbPDS.Checked == true)
                {
                    GlobalVariableController.GlobalButtonFlag[6] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[6] = false;
                }
                if (cbPSL.Checked == true)
                {
                    GlobalVariableController.GlobalButtonFlag[7] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[7] = false;
                }
                e.Cancel = false;

                return;
            }
        }
        /// <summary>
        /// SEAL PLAN A/ linkage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBoxPlanA_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.dgvSumList.CurrentCell.ColumnIndex == this.dgvSumList.Columns["ColSealPlanA"].Index)
            {
                SortedDictionary<double, string> DicCombboxPlanB = new SortedDictionary<double, string>();
                ComboBox comboBox = (ComboBox)sender;
                //SEAL PLAN A
                string strPlanA = comboBox.Text;

                foreach (DataRow drPlan in COMBOBOX_DATASAT.Tables["API682PLANTABLE"].Rows)
                {
                    String drPlanA = drPlan.Field<string>("PLANA");
                    String drPlanB = drPlan.Field<string>("PLANB");
                    double keyNoB = 0;
                    if (drPlanB == "空白")
                    {
                        keyNoB = 99;
                    }
                    else if (drPlanB.Contains("53B"))
                    {
                        keyNoB = 53.5;
                    }
                    else
                    {
                        keyNoB = Int32.Parse(System.Text.RegularExpressions.Regex.Replace(drPlanB, @"[^0-9]+", ""));
                    }
                    if (drPlanA == strPlanA)
                    {
                        if (!(DicCombboxPlanB.ContainsKey(keyNoB)))
                        {
                            DicCombboxPlanB.Add(keyNoB, drPlanB);
                        }
                    }
                }

                DataTable dtPlanB = new DataTable();
                string strColumnName = "SEALPLANB";
                DataColumn columnName = new DataColumn(strColumnName);
                dtPlanB.Columns.Add(columnName);

                if (DicCombboxPlanB.Count > 0)
                {
                    foreach (string strValue in DicCombboxPlanB.Values)
                    {
                        DataRow drOne = dtPlanB.NewRow();
                        drOne[0] = strValue;
                        dtPlanB.Rows.Add(drOne);
                    }
                }

                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSealPlanB"];
                combo.DataSource = dtPlanB;
                combo.DisplayMember = "SEALPLANB";
                combo.ValueMember = "SEALPLANB";
            }
        }
        /// <summary>
        /// SEAL PLAN B/ linkage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBoxPlanB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.dgvSumList.CurrentCell.ColumnIndex == this.dgvSumList.Columns["ColSealPlanB"].Index)
            {
                SortedDictionary<double, string> DicCombboxPlanC = new SortedDictionary<double, string>();
                ComboBox comboBox = (ComboBox)sender;
                //SEAL PLAN B
                string strPlanB = comboBox.Text;

                foreach (DataRow drPlan in COMBOBOX_DATASAT.Tables["API682PLANTABLE"].Rows)
                {
                    String drPlanB = drPlan.Field<string>("PLANB");
                    String drPlanC = drPlan.Field<string>("PLANC");
                    double keyNoC = 0;
                    if (drPlanC == "空白")
                    {
                        keyNoC = 99;
                    }
                    else if (drPlanC.Contains("53B"))
                    {
                        keyNoC = 53.5;
                    }
                    else
                    {
                        keyNoC = Int32.Parse(System.Text.RegularExpressions.Regex.Replace(drPlanC, @"[^0-9]+", ""));
                    }
                    if (drPlanB == strPlanB)
                    {
                        if (!(DicCombboxPlanC.ContainsKey(keyNoC)))
                        {
                            DicCombboxPlanC.Add(keyNoC, drPlanC);
                        }
                    }
                }

                DataTable dtPlanC = new DataTable();
                string strColumnName = "SEALPLANC";
                DataColumn columnName = new DataColumn(strColumnName);
                dtPlanC.Columns.Add(columnName);

                if (DicCombboxPlanC.Count > 0)
                {
                    foreach (string strValue in DicCombboxPlanC.Values)
                    {
                        DataRow drOne = dtPlanC.NewRow();
                        drOne[0] = strValue;
                        dtPlanC.Rows.Add(drOne);
                    }
                }

                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSealPlanC"];
                combo.DataSource = dtPlanC;
                combo.DisplayMember = "SEALPLANC";
                combo.ValueMember = "SEALPLANC";
            }
        }
        /// <summary>
        /// datagridview cell cursor away
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DgvSumList_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //Drop down menu fuzzy query
            if (dgvSumList.CurrentCellAddress.X == dgvSumList.Columns["ColSize"].Index)
            {
                //If you enter the value of valuemember, it's OK.
                //But if you enter the value of displaymember, it won't be selected
                dgvSumList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = e.FormattedValue;
            }
            if (dgvSumList.CurrentCellAddress.X == dgvSumList.Columns["ColMaterialClass"].Index)
            {
                //If you enter the value of valuemember, it's OK.
                //But if you enter the value of displaymember, it won't be selected
                dgvSumList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = e.FormattedValue;
            }
        }
        /// <summary>
        /// datagridview cell edit event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DgvSumList_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //Block the right-click message.
            //When editing a cell, the right-click menu of the system will not pop up
            e.Control.ContextMenuStrip = new ContextMenuStrip();
            //Drop down menu fuzzy query
            if (dgvSumList.CurrentCellAddress.X == dgvSumList.Columns["ColSize"].Index)
            {
                ComboBox cbo = e.Control as ComboBox;
                if (cbo != null)
                {
                    cbo.DropDownStyle = ComboBoxStyle.DropDown;
                    cbo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                }
            }
            if (dgvSumList.CurrentCellAddress.X == dgvSumList.Columns["ColMaterialClass"].Index)
            {
                ComboBox cbo = e.Control as ComboBox;
                if (cbo != null)
                {
                    cbo.DropDownStyle = ComboBoxStyle.DropDown;
                    cbo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                }
            }
            //TYPE / 泵型
            if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColType") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
                ComboBox comboBox = (ComboBox)e.Control;
                comboBox.SelectedValueChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            }
            //SIZE/规格
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColSize") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //FRAME SIZE/基座号
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColFrameSize") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //API 682 EDITION / API 682 版本
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColApiEdition") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //SEAL PLAN A/ 冲洗方案
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColSealPlanA") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
                ComboBox comboBoxPlanA = (ComboBox)e.Control;
                comboBoxPlanA.SelectedValueChanged += new EventHandler(ComboBoxPlanA_SelectedIndexChanged);
            }
            //SEAL PLAN B/ 冲洗方案
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColSealPlanB") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
                ComboBox comboBoxPlanB = (ComboBox)e.Control;
                comboBoxPlanB.SelectedValueChanged += new EventHandler(ComboBoxPlanB_SelectedIndexChanged);
            }
            //SEAL PLAN C/ 冲洗方案
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColSealPlanC") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //WI/WO PUMP DRAIN VALVE/排残管路阀门
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColPumpDrainValve") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //WI/WO WARM-UP PIPING/暖泵管路
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColWarmUpPiping") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //MATERIAL CLASS/材料等级
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColMaterialClass") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
                ComboBox comboBox = (ComboBox)e.Control;
                comboBox.SelectedValueChanged += new EventHandler(ComboBoxMaterialClass_SelectedIndexChanged);
            }
            //DRY/WET SHAFT/干湿轴
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColDryWetShaft") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
                ComboBox comboBox = (ComboBox)e.Control;
                comboBox.SelectedValueChanged += new EventHandler(ComboBoxDryWet_SelectedIndexChanged);
            }
            //BEARING/轴承配置
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColBearing") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //BEARING CAGE MATERIAL/轴承保持架材质
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColBearingCageMaterial") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //COMPANION FLG./配对法兰要求
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColCompanionFlg") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //CASE  WI/WO BRACKET/壳体是否带夹套
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColCaseBracket") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //WI/WO RESERVED PURE OIL-MIST LUBRICATION/预留油雾润滑
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColReservedPureLubrication") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //BEARING HOUSING WI/WO OIL DRAIN/轴承室带排油管路
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColBearighousingDrain") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //IMPELLER WI/WO BALANCE HOLE/叶轮带平衡孔
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColImpellerBalanceHole") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //PRE泵压力等级
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColPREPumpPressureClass") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //是否带诱导轮
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColPREInducerYesNo") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //冷却水母管口径
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColCoolWaterMainPipeSize") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
            //排残阀类型
            else if (dgvSumList.CurrentCell.OwningColumn.Name.Equals("ColDrainValveType") && e.Control is ComboBox)
            {
                e.CellStyle.BackColor = Color.FromName("window");
            }
        }
        /// <summary>
        /// MATERIAL CLASS/材料等级 linkage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBoxMaterialClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.dgvSumList.CurrentCell.ColumnIndex == this.dgvSumList.Columns["ColMaterialClass"].Index)
            {
                ComboBox comboBox = (ComboBox)sender;
                if (comboBox.Text.Contains("MOD"))
                {
                    //泵体/泵盖
                    dgvSumList.Columns["ColVoluteCasing"].Visible = true;
                    //dgvSumList.Columns["ColVoluteCasing"].ReadOnly = true;
                    //叶轮
                    dgvSumList.Columns["ColImpeller"].Visible = true;
                    //口环
                    dgvSumList.Columns["ColStationaryWearRing"].Visible = true;
                    //泵盖水冷密封板
                    dgvSumList.Columns["ColCoverCoolingPlate"].Visible = true;
                    //喉部衬套
                    dgvSumList.Columns["ColThroatBush"].Visible = true;
                    //叶轮键
                    dgvSumList.Columns["ColImpellerKey"].Visible = true;
                    //泵体泵盖之间的密封垫
                    dgvSumList.Columns["ColGasket"].Visible = true;
                    //轴套
                    dgvSumList.Columns["ColSpacerSleeve"].Visible = true;
                    //叶轮螺母
                    dgvSumList.Columns["ColImpellerNut"].Visible = true;
                    //诱导轮
                    dgvSumList.Columns["ColInducer"].Visible = true;
                    //诱导轮衬套
                    dgvSumList.Columns["ColInducerSleeve"].Visible = true;
                    //耐磨板
                    dgvSumList.Columns["ColWearPlate"].Visible = true;
                    //TYPE/泵型
                    string strType = (String)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColType"].Value;
                    if (!strType.Contains("PR"))
                    {
                        //泵体/泵盖
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColVoluteCasing"].ReadOnly = false;
                        //叶轮
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColImpeller"].ReadOnly = false;
                        //口环
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColStationaryWearRing"].ReadOnly = false;
                        //喉部衬套
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColThroatBush"].ReadOnly = false;
                        //叶轮键
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColImpellerKey"].ReadOnly = false;
                        //泵体泵盖之间的密封垫
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColGasket"].ReadOnly = false;
                        //叶轮螺母
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColImpellerNut"].ReadOnly = false;
                        if (strType.Equals("ZAO") || strType.Equals("ZEO"))
                        {
                            //耐磨板
                            dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = false;
                        }
                        else
                        {
                            //耐磨板
                            dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                        }
                    }
                    else
                    {
                        //泵体/泵盖
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColVoluteCasing"].ReadOnly = false;
                        //叶轮
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColImpeller"].ReadOnly = false;
                        //口环
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColStationaryWearRing"].ReadOnly = false;
                        //泵盖水冷密封板
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCoverCoolingPlate"].ReadOnly = false;
                        //喉部衬套
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColThroatBush"].ReadOnly = false;
                        //泵体泵盖之间的密封垫
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColGasket"].ReadOnly = false;
                        //轴套
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSpacerSleeve"].ReadOnly = false;
                        //叶轮螺母
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColImpellerNut"].ReadOnly = false;
                        //诱导轮
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColInducer"].ReadOnly = false;
                        //诱导轮衬套
                        dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColInducerSleeve"].ReadOnly = false;
                        if(strType.Equals("PRESF"))
                        {
                            //耐磨板
                            dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = false;
                        }
                        else
                        {
                            //耐磨板
                            dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                        }
                    }
                }
                else
                {
                    int iCount = 0;
                    for (int i = 0; i < dgvSumList.RowCount; i++)
                    {
                        if (dgvSumList.CurrentCell.RowIndex == i)
                        {
                            continue;
                        }
                        if (dgvSumList.Rows[i].Cells["ColMaterialClass"].Value.ToString().Contains("MOD"))
                        {
                            iCount++;
                        }
                    }
                    if (iCount == 0)
                    {
                        //泵体/泵盖
                        dgvSumList.Columns["ColVoluteCasing"].Visible = false;
                        //叶轮
                        dgvSumList.Columns["ColImpeller"].Visible = false;
                        //口环
                        dgvSumList.Columns["ColStationaryWearRing"].Visible = false;
                        //泵盖水冷密封板
                        dgvSumList.Columns["ColCoverCoolingPlate"].Visible = false;
                        //喉部衬套
                        dgvSumList.Columns["ColThroatBush"].Visible = false;
                        //叶轮键
                        dgvSumList.Columns["ColImpellerKey"].Visible = false;
                        //泵体泵盖之间的密封垫
                        dgvSumList.Columns["ColGasket"].Visible = false;
                        //轴套
                        dgvSumList.Columns["ColSpacerSleeve"].Visible = false;
                        //叶轮螺母
                        dgvSumList.Columns["ColImpellerNut"].Visible = false;
                        //诱导轮
                        dgvSumList.Columns["ColInducer"].Visible = false;
                        //诱导轮衬套
                        dgvSumList.Columns["ColInducerSleeve"].Visible = false;
                        //耐磨板
                        dgvSumList.Columns["ColWearPlate"].Visible = false;
                    }
                }
            }
        }
        /// <summary>
        /// DRY/WET SHAFT/干湿轴 linkage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBoxDryWet_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            if (comboBox.Text.Contains("Wet"))
            {
                //Wet Shaft(自己填写)
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWetInput"].ReadOnly = false;
            }
            else
            {
                //Wet Shaft(自己填写)
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWetInput"].ReadOnly = true;
            }
        }
        /// <summary>
        /// TYPE/泵型列和SIZE/规格列 linkage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            //Get the current cell, get its value,
            //and bind different DataTables to adjacent drop-down box cells according to different values
            if (comboBox.Text.Equals("ZA"))
            {
                //SIZE/规格
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                //datatable binding
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZATABLE"];
                combo.DisplayMember = "ZA";
                combo.ValueMember = "ZA";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "1";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("ZAO"))
            {
                //SIZE / 规格
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZAOTABLE"];
                combo.DisplayMember = "ZAO";
                combo.ValueMember = "ZAO";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "1";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = false;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if(comboBox.Text.Equals("ZAG"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZAGTABLE"];
                combo.DisplayMember = "ZAG";
                combo.ValueMember = "ZAG";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "1";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = true;
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("ZE(8th)"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZE(8th)TABLE"];
                combo.DisplayMember = "ZE(8th)";
                combo.ValueMember = "ZE(8th)";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("ZE(10th)"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZE(10th)TABLE"];
                combo.DisplayMember = "ZE(10th)";
                combo.ValueMember = "ZE(10th)";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("ZEO"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZEOTABLE"];
                combo.DisplayMember = "ZEO";
                combo.ValueMember = "ZEO";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = false;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("ZEG"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZEGTABLE"];
                combo.DisplayMember = "ZEG";
                combo.ValueMember = "ZEG";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = true;
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("ZF"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZFTABLE"];
                combo.DisplayMember = "ZF";
                combo.ValueMember = "ZF";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].Value = "";
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].Value = "";
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("PRE"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPRETABLE"];
                combo.DisplayMember = "PRE";
                combo.ValueMember = "PRE";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = false;
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = false;
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("PREZ"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPREZTABLE"];
                combo.DisplayMember = "PREZ";
                combo.ValueMember = "PREZ";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = false;
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = false;
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("PRESF"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPRESFTABLE"];
                combo.DisplayMember = "PRESF";
                combo.ValueMember = "PRESF";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = false;
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = false;
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = false;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("PRELF"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPRELFTABLE"];
                combo.DisplayMember = "PRELF";
                combo.ValueMember = "PRELF";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = false;
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = false;
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
            else if (comboBox.Text.Equals("PREK"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPREKTABLE"];
                combo.DisplayMember = "PREK";
                combo.ValueMember = "PREK";
                //PRE泵压力等级
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREPumpPressureClass"].ReadOnly = false;
                //是否带诱导轮
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColPREInducerYesNo"].ReadOnly = false;
                //管口载荷要求
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColNozzleLoad"].Value = "2";
                //壳体是否带夹套
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColCaseBracket"].Value = "";
                //耐磨板
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].ReadOnly = true;
                dgvSumList.Rows[dgvSumList.CurrentCell.RowIndex].Cells["ColWearPlate"].Value = "";
            }
        }
        /// <summary>
        /// SIZE/规格
        /// </summary>
        /// <param name="strType">TYPE</param>
        /// <param name="i">line no</param>
        private void SelecteSizeComboBox(string strType, int i)
        {
            if (strType.Equals("ZA"))
            {
                //SIZE/规格
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                //datatable binding
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZATABLE"];
                combo.DisplayMember = "ZA";
                combo.ValueMember = "ZA";
            }
            else if (strType.Equals("ZAO"))
            {
                //SIZE/规格
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZAOTABLE"];
                combo.DisplayMember = "ZAO";
                combo.ValueMember = "ZAO";
            }
            else if (strType.Equals("ZAG"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZAGTABLE"];
                combo.DisplayMember = "ZAG";
                combo.ValueMember = "ZAG";
            }
            else if (strType.Equals("ZE(8th)"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZE(8th)TABLE"];
                combo.DisplayMember = "ZE(8th)";
                combo.ValueMember = "ZE(8th)";
            }
            else if (strType.Equals("ZE(10th)"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZE(10th)TABLE"];
                combo.DisplayMember = "ZE(10th)";
                combo.ValueMember = "ZE(10th)";
            }
            else if (strType.Equals("ZEO"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZEOTABLE"];
                combo.DisplayMember = "ZEO";
                combo.ValueMember = "ZEO";
            }
            else if (strType.Equals("ZEG"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZEGTABLE"];
                combo.DisplayMember = "ZEG";
                combo.ValueMember = "ZEG";
            }
            else if (strType.Equals("ZF"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEZFTABLE"];
                combo.DisplayMember = "ZF";
                combo.ValueMember = "ZF";
            }
            else if (strType.Equals("PRE"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPRETABLE"];
                combo.DisplayMember = "PRE";
                combo.ValueMember = "PRE";
            }
            else if (strType.Equals("PREZ"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPREZTABLE"];
                combo.DisplayMember = "PREZ";
                combo.ValueMember = "PREZ";
            }
            else if (strType.Equals("PRESF"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPRESFTABLE"];
                combo.DisplayMember = "PRESF";
                combo.ValueMember = "PRESF";
            }
            else if (strType.Equals("PRELF"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPRELFTABLE"];
                combo.DisplayMember = "PRELF";
                combo.ValueMember = "PRELF";
            }
            else if (strType.Equals("PREK"))
            {
                DataGridViewComboBoxCell combo = (DataGridViewComboBoxCell)dgvSumList.Rows[i].Cells["ColSize"];
                combo.Style.BackColor = Color.FromName("window");
                combo.DataSource = COMBOBOX_DATASAT.Tables["TYPESIZEPREKTABLE"];
                combo.DisplayMember = "PREK";
                combo.ValueMember = "PREK";
            }
        }
        /// <summary>
        /// Button available settings
        /// </summary>
        /// <param name="bolFlag"></param>
        private void BtnToEnabled(bool bolFlag)
        {
            btnSumListDataImport.Enabled = bolFlag;
            btnXMLDataImport.Enabled = bolFlag;
            btnClear.Enabled = bolFlag;
            btnSave.Enabled = bolFlag;
            btnClose.Enabled = bolFlag;
        }
        /// <summary>
        /// paste data
        /// </summary>
        /// <param name="dgvSumList"></param> 
        private void PasteData(DataGridView dgvSumList)
        {
            try
            {
                //Gets the contents of the clipboard
                string clipboardText = Clipboard.GetText();
                if (string.IsNullOrEmpty(clipboardText))
                {
                    return;
                }
                int colnum = 0;
                int rownum = 0;
                for (int i = 0; i < clipboardText.Length; i++)
                {
                    if (clipboardText.Substring(i, 1) == "\t")
                    {
                        colnum++;
                    }
                    if (clipboardText.Substring(i, 1) == "\n")
                    {
                        rownum++;
                    }
                }
                //When the data on the pasteboard comes from excel,
                //there is \ n at the end of each row. If it comes from DataGridView,
                //there is no \ n at the end of the last row
                if (clipboardText.Substring(clipboardText.Length - 1, 1) == "\n")
                {
                    rownum--;
                }
                colnum = colnum / (rownum + 1);
                //Defines a two-dimensional array of object types
                object[,] data;
                //Instantiate the array based on the number of rows and columns of the clipboard
                data = new object[rownum + 1, colnum + 1];
                string rowStr = "";
                //Assign values to the elements of the array
                for (int i = 0; i <= rownum; i++)
                {
                    for (int j = 0; j <= colnum; j++)
                    {
                        //Other columns in a row
                        if (j != colnum)
                        {
                            rowStr = clipboardText.Substring(0, clipboardText.IndexOf("\t"));
                            clipboardText = clipboardText.Substring(clipboardText.IndexOf("\t") + 1);
                        }
                        //The last column in a row
                        if (j == colnum && clipboardText.IndexOf("\r") != -1)
                        {
                            rowStr = clipboardText.Substring(0, clipboardText.IndexOf("\r"));
                        }
                        //Last column of last row
                        if (j == colnum && clipboardText.IndexOf("\r") == -1)
                        {
                            rowStr = clipboardText.Substring(0);
                        }
                        data[i, j] = rowStr;
                    }
                    //Intercept the next row and subsequent data
                    clipboardText = clipboardText.Substring(clipboardText.IndexOf("\n") + 1);
                }
                //Gets the column sequence number of the currently selected cell
                int colIndex = dgvSumList.CurrentRow.Cells.IndexOf(dgvSumList.CurrentCell);
                //Gets the row sequence number of the currently selected cell
                int rowIndex = dgvSumList.CurrentRow.Index;
                for (int i = 0; i <= rownum; i++)
                {
                    for (int j = 0; j <= colnum; j++)
                    {
                        dgvSumList.Rows[i + rowIndex].Cells[j + colIndex].Value = data[i, j];
                    }
                }
            }
            catch
            {
                MessageBox.Show("粘贴区域大小不一致");
                return;
            }
        }

    }
    /// <summary>
    /// File Name       ：Choice
    /// Description     ：Construction class
    /// Author          ：LIU
    /// Create Date     ：2021/09/15 10:43
    /// Modifier        ：
    /// Modifier Date   ：
    /// Revision History：
    /// Version         ：V1.0
    /// </summary>
    public class Choice
    {
        public string Name { get; private set; }
        public Boolean Value { get; private set; }
        public Choice(string name, Boolean  value)
        {
            Name = name;
            Value = value;
        }
        private static readonly List<Choice> possibleChoices = new List<Choice>
        {
            { new Choice("Yes/是", true ) },
            { new Choice("No/否", false) }
        };
        public static List<Choice> GetChoices()
        {
            return possibleChoices;
        }
    }
    /// <summary>
    /// File Name       ：ChoicePDValve
    /// Description     ：Construction class
    /// Author          ：LIU
    /// Create Date     ：2021/09/15 10:43
    /// Modifier        ：
    /// Modifier Date   ：
    /// Revision History：
    /// Version         ：V1.0
    /// </summary>
    public class ChoicePDValve
    {
        public string PDName { get; private set; }
        public string PDValue { get; private set; }
        public ChoicePDValve(string pdname, string pdvalue)
        {
            PDName = pdname;
            PDValue = pdvalue;
        }
        private static readonly List<ChoicePDValve> possibleChoicePDValve = new List<ChoicePDValve>
        {
            { new ChoicePDValve("Without Valve/无阀", "0" ) },
            { new ChoicePDValve("Single Valve/单阀", "1") },
            { new ChoicePDValve("Double Valves/双阀", "2" ) },
        };
        public static List<ChoicePDValve> GetChoicePDValve()
        {
            return possibleChoicePDValve;
        }
        private static readonly List<ChoicePDValve> possibleChoiceCageValve = new List<ChoicePDValve>
        {
            { new ChoicePDValve("Nylon/P尼龙(ZA默认)", "P" ) },
            { new ChoicePDValve("Steel/J钢(ZE/ZF默认)", "J") },
            { new ChoicePDValve("Brass/M铜", "M" ) },
        };
        public static List<ChoicePDValve> GetChoiceCageValve()
        {
            return possibleChoiceCageValve;
        }
        private static readonly List<ChoicePDValve> possiblewetDryChoices = new List<ChoicePDValve>
        {
            { new ChoicePDValve("Wet Shaft(自己填写)", "Wet" ) },
            { new ChoicePDValve("Dry Shaft(30Cr13)", "Dry") }
        };
        public static List<ChoicePDValve> GetWetDryChoices()
        {
            return possiblewetDryChoices;
        }
    }
}
