using AcAdAmComm;
using System;
using System.Data;
using System.Windows.Forms;

namespace SESDTools
{
    /// <summary>
    /// 类    名：AdministratorForm
    /// 功能描述：管理画面类
    /// 创 建 人：SY
    /// 创建时间：2021/07/23 10:43
    /// 描    述：基本数据维护
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    /// </summary>
    public partial class AdministratorForm : Form
    {
        public AdministratorForm()
        {
            //画面居中表示
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            //画面初始化赋值
            InitializeDataSettings();
        }
        /// <summary>
        /// 画面初始化赋值
        /// </summary>
        private void InitializeDataSettings()
        {
            string nodeNamePassword = "password";
            string nodeNamePasswordConfirmation = "passwordConfirmation";
            string nodeNameAdminiPhone = "adminiPhone";
            string nodeNameAdminiMail = "adminiMail";
            string nodeNameVersionNumber = "versionNumber";
            string nodeNameVersionUpgradeExplanation = "versionUpgradeExplanation";
            string nodeNameDeveloperContactInformation = "developerContactInformation";
            string nodeNameComputerName = "computerName";
            //Password
            txtPassword.Text = XMLHelper.ReadAdminiPasswordSettings(nodeNamePassword);
            //PasswordConfirmation
            txtPasswordConfirmation.Text = XMLHelper.ReadAdminiPasswordSettings(nodeNamePasswordConfirmation);
            //Phone
            txtPhone.Text = XMLHelper.ReadAdminiContactInformationSettings(nodeNameAdminiPhone);
            //Mail
            txtMail.Text = XMLHelper.ReadAdminiContactInformationSettings(nodeNameAdminiMail);
            //VersionNumber
            txtVersionNumber.Text = XMLHelper.ReadVersionInformationSettings(nodeNameVersionNumber);
            //VersionInstructions
            txtVersionInstructions.Text = XMLHelper.ReadVersionInformationSettings(nodeNameVersionUpgradeExplanation);
            //DeveloperPhone
            txtDeveloperPhone.Text = XMLHelper.ReadVersionInformationSettings(nodeNameDeveloperContactInformation);
            //ComputerNameList
            string strDgvComputerName = XMLHelper.ReadSecureInterfaceSettings(nodeNameComputerName);
            string[] rowDataComputerName = strDgvComputerName.Split(';');
            for(int i = 0;i< rowDataComputerName.Length;i++)
            {
                string strRow = rowDataComputerName[i];
                if(strRow.Length == 0)
                {
                    break;
                }
                string[] cellDataComputerName = strRow.Split(',');
                this.dgvComputerName.Rows.Add();
                //UserName
                dgvComputerName.Rows[i].Cells[0].Value = cellDataComputerName[0];
                //ComputerName
                dgvComputerName.Rows[i].Cells[1].Value = cellDataComputerName[1];
                
            }
        }
        /// <summary>
        /// select the excel folder path
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExcelFilePathOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "Please select a excel file";
            dialog.Filter = "Excel File|*.xlsx*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = dialog.FileName;
                txtExcelFliePath.Text = savePath;
            }
        }
        /// <summary>
        /// select the path to generate XML folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnXmlFilePathOpen_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "Please select the path to generate XML folder";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = dialog.SelectedPath;
                txtXmlFilePath.Text = savePath;
            }
        }
        /// <summary>
        /// To Xml File Button down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDoneToXmlFile_Click(object sender, EventArgs e)
        {
            if (txtExcelFliePath.Text == "")
            {
                //Log写入
                LogHelper.GetSingleObj().Error("XML文件生成", "Please select the excel folder path");
                MessageBox.Show("Please select the excel folder path!", "ERROR");
                return;
            }
            if (txtXmlFilePath.Text == "")
            {
                //Log写入
                LogHelper.GetSingleObj().Error("XML文件生成", "Please select the path to generate XML folder");
                MessageBox.Show("Please select the path to generate XML folder!", "ERROR");
                return;
            }
            //按钮不可用
            BtnToEnabled(false);
            string excelFilePath = txtExcelFliePath.Text;
            string xmlFilePath = txtXmlFilePath.Text;
            string fileName = System.IO.Path.GetFileName(excelFilePath);//文件名
            string xmlFileName = fileName.Remove(fileName.Length - 5, 5);
            try
            {
                //读取一个EXCEL文件里所有数据到dataSet里
                DataSet dataSet = ExcelDataOperation.ExcelToDataSet(excelFilePath, xmlFileName, true);
                //All ComboBox Data
                DataSet dataSetTemp = new DataSet();
                dataSetTemp.Tables.Add(dataSet.Tables["TYPETABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZATABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZAOTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZAGTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZE(8th)TABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZE(10th)TABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZEOTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZEGTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEZFTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEPRETABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEPREZTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEPRESFTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEPRELFTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["TYPESIZEPREKTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["FRAMESIZETABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["APIEDITIONTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["SEALPLANATABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["SEALPLANBTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["SEALPLANCTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["FLANGESTANDARDTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["MATERIALCLASSTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["DRYWETSHAFTTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["BEARINGTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["BEARINGCAGEMATERIALTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["PUMPDRAINVALVETABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["BEARINGWATERCOOLINGTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["API682PLANTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["PREPUMPPRESSURECLASSTABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["COOLWATERMAINPIPESIZETABLE"].Copy());
                dataSetTemp.Tables.Add(dataSet.Tables["DRAINVALVETYPETABLE"].Copy());
                string xmlComboBoxName = xmlFilePath + "\\" + "COMBOBOXDATATABLE.xml";
                //生成一个ComboBoxXML文件
                XmlDataOperation.DataSetToXMLFile(dataSetTemp, xmlComboBoxName);
                //ComboBox数据以外的数据生成一个XML文件
                dataSet.Tables.Remove(dataSet.Tables["TYPETABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZATABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZAOTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZAGTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZE(8th)TABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZE(10th)TABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZEOTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZEGTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEZFTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEPRETABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEPREZTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEPRESFTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEPRELFTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["TYPESIZEPREKTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["FRAMESIZETABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["APIEDITIONTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["SEALPLANATABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["SEALPLANBTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["SEALPLANCTABLE"]);
                //dataSet.Tables.Remove(dataSet.Tables["FLANGESTANDARDTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["MATERIALCLASSTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["DRYWETSHAFTTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["BEARINGTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["BEARINGCAGEMATERIALTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["PUMPDRAINVALVETABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["BEARINGWATERCOOLINGTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["PREPUMPPRESSURECLASSTABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["COOLWATERMAINPIPESIZETABLE"]);
                dataSet.Tables.Remove(dataSet.Tables["DRAINVALVETYPETABLE"]);
                string xmlBasicDataName = xmlFilePath + "\\" + "BASICDATATABLE.xml";
                //生成一个BasicDataXML文件
                XmlDataOperation.DataSetToXMLFile(dataSet, xmlBasicDataName);
                //for (int i = 0; i < dataSet.Tables.Count; i++)
                //{
                //    DataTable dataTable = dataSet.Tables[i];
                //    string strDataTableName = dataTable.TableName;
                //    string xmlFullName = xmlFilePath + "\\" + strDataTableName + ".xml";
                //    XmlDataOperation.DataTableToXMLFile(dataTable, xmlFullName);
                //}
            }
            catch(Exception ex)
            {
                //Log写入
                LogHelper.GetSingleObj().Error("XML文件生成", ex.Message);
                MessageBox.Show("Please select the correct file template!", "ERROR");
                //按钮可用
                BtnToEnabled(true);
                return;
            }
            //基础数据文件生成路径保存
            string nodeNameXmlDataFilePath = "xmlDataFilePath";
            XMLHelper.AddXmlDataFilePathSettings(nodeNameXmlDataFilePath, xmlFilePath);
            //按钮可用
            BtnToEnabled(true);
            MessageBox.Show("XML data file generation completed!");

        }
        /// <summary>
        /// Administrator Password Update Button down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDoneToPasswordUpdate_Click(object sender, EventArgs e)
        {
            string nodeNamePassword = "password";
            string nodeNamePasswordConfirmation = "passwordConfirmation";
            if (txtPassword.Text == "")
            {
                //Log写入
                LogHelper.GetSingleObj().Error("密码更新", "Please enter a password");
                MessageBox.Show("Please enter a password!", "ERROR");
                return;
            }
            if (txtPasswordConfirmation.Text == "")
            {
                //Log写入
                LogHelper.GetSingleObj().Error("密码更新", "Please enter the confirmation password");
                MessageBox.Show("Please enter the confirmation password!", "ERROR");
                return;
            }
            if (txtPassword.Text != txtPasswordConfirmation.Text)
            {
                //Log写入
                LogHelper.GetSingleObj().Error("密码更新", "The password does not match the confirmation password");
                MessageBox.Show("The password does not match the confirmation password!", "ERROR");
                return;
            }
            //按钮不可用
            BtnToEnabled(false);
            XMLHelper.AddAdminiPasswordSettings(nodeNamePassword, txtPassword.Text);
            XMLHelper.AddAdminiPasswordSettings(nodeNamePasswordConfirmation, txtPasswordConfirmation.Text);
            //Log写入
            LogHelper.GetSingleObj().Info("密码更新", "密码更新成功");
            //按钮可用
            BtnToEnabled(true);
            MessageBox.Show("The password has been updated!");
        }
        /// <summary>
        /// Administrator Information Update Button down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDoneToAdminiInforUpdate_Click(object sender, EventArgs e)
        {
            string nodeNameAdminiPhone = "adminiPhone";
            string nodeNameAdminiMail = "adminiMail";
            //按钮不可用
            BtnToEnabled(false);
            XMLHelper.AddAdminiContactInformationSettings(nodeNameAdminiPhone, txtPhone.Text);
            XMLHelper.AddAdminiContactInformationSettings(nodeNameAdminiMail, txtMail.Text);
            //Log写入
            LogHelper.GetSingleObj().Info("管理者信息更新", "管理者信息更新成功");
            //按钮可用
            BtnToEnabled(true);
            MessageBox.Show("The administrator contact information has been updated!");
        }
        /// <summary>
        /// Version Information Update Button down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDoneToVersionInforUpdate_Click(object sender, EventArgs e)
        {
            string nodeNameVersionNumber = "versionNumber";
            string nodeNameVersionUpgradeExplanation = "versionUpgradeExplanation";
            string nodeNameDeveloperContactInformation = "developerContactInformation";
            //按钮不可用
            BtnToEnabled(false);
            XMLHelper.AddVersionInformationSettings(nodeNameVersionNumber, txtVersionNumber.Text);
            XMLHelper.AddVersionInformationSettings(nodeNameVersionUpgradeExplanation, txtVersionInstructions.Text);
            XMLHelper.AddVersionInformationSettings(nodeNameDeveloperContactInformation, txtDeveloperPhone.Text);
            //Log写入
            LogHelper.GetSingleObj().Info("版本信息更新", "版本信息更新成功");
            //按钮可用
            BtnToEnabled(true);
            MessageBox.Show("The version information has been updated!");
        }
        /// <summary>
        /// DataGridView行号自动添加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DgvComputerName_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }
        /// <summary>
        /// Secure Interface Update Button down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDoneToSecureInterfaceUpdate_Click(object sender, EventArgs e)
        {
            string nodeNameComputerName = "computerName";
            bool boolFlag = false;
            string secureInterface = "";
            int row = dgvComputerName.Rows.Count;//得到总行数
            //得到总行数并在之内循环  
            for (int i = 0; i < row; i++) 
            {
                string strRowCellOne = "";
                string strRowCellTwo = "";
                if (dgvComputerName.Rows[i].Cells[0].Value != null)
                {
                    strRowCellOne = dgvComputerName.Rows[i].Cells[0].Value.ToString();
                }
                if (dgvComputerName.Rows[i].Cells[1].Value != null)
                {
                    strRowCellTwo = dgvComputerName.Rows[i].Cells[1].Value.ToString();
                }
                if(string.IsNullOrEmpty(strRowCellOne) && !string.IsNullOrEmpty(strRowCellTwo))
                {
                    boolFlag = true;
                    break;
                }
                if (!string.IsNullOrEmpty(strRowCellOne) && string.IsNullOrEmpty(strRowCellTwo))
                {
                    boolFlag = true;
                    break;
                }
                if (string.IsNullOrEmpty(strRowCellOne) && string.IsNullOrEmpty(strRowCellTwo))
                {
                    continue;
                }

                secureInterface = secureInterface + strRowCellOne + "," + strRowCellTwo + ";";
            }

            if(boolFlag)
            {
                //Log写入
                LogHelper.GetSingleObj().Error("用户名计算机名更新", "User name and computer name cannot be empty");
                MessageBox.Show("User name and computer name cannot be empty!", "ERROR");
                return;
            }
            else
            {
                //按钮不可用
                BtnToEnabled(false);
                XMLHelper.AddSecureInterfaceSettings(nodeNameComputerName, secureInterface);
                //Log写入
                LogHelper.GetSingleObj().Info("用户名计算机名更新", "用户名计算机名更新成功");
                //按钮可用
                BtnToEnabled(true);
                MessageBox.Show("The user name and computer name have been updated!");
            }
        }

        /// <summary>
        /// Close Button down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        /// <summary>
        /// 按钮可用不可用设定
        /// </summary>
        /// <param name="bolFlag"></param>
        private void BtnToEnabled(bool bolFlag)
        {
            btnDoneToXmlFile.Enabled = bolFlag;
            btnDoneToPasswordUpdate.Enabled = bolFlag;
            btnDoneToAdminiInforUpdate.Enabled = bolFlag;
            btnDoneToSecureInterfaceUpdate.Enabled = bolFlag;
            btnDoneToVersionInforUpdate.Enabled = bolFlag;
        }
    }
}
