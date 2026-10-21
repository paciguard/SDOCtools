using AcAdAmComm;
using System;
using System.Data;
using System.Windows.Forms;

namespace SESDTools
{
    /// <summary>
    /// 类    名：SelectPathForm
    /// 功能描述：路径选择画面类
    /// 创 建 人：SY
    /// 创建时间：2021/07/23 10:43
    /// 描    述：共通画面，各个打印功能调用
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    public partial class SelectPathForm : Form
    {
        /// <summary>
        /// 画面数据
        /// </summary>
        public DataSet SUMLISTDATASAT;
        /// <summary>
        /// 基础数据
        /// </summary>
        public DataSet BASICDATASAT;
        public string SelectedPath;
        public SelectPathForm(String functionNameParameter, DataSet ds)
        {
            //画面居中表示
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            InitializeBasicData();
            this.labelRecieve.Text = functionNameParameter;
            this.SUMLISTDATASAT = ds.Copy();
        }
        /// <summary>
        /// BasicData初始化赋值
        /// </summary>
        private void InitializeBasicData()
        {
            //BasicData数据文件路径
            string nodeNamePassword = "xmlDataFilePath";
            string fileName = "BASICDATATABLE.xml";
            DataSet dataSetBasicData = new DataSet();
            try
            {
                string xmlDataFilePath = XMLHelper.ReadXmlDataFilePathSettings(nodeNamePassword) + "\\" + fileName;
                //BasicData数据取得
                dataSetBasicData = XmlDataOperation.XMLFileToDataSet(xmlDataFilePath);
            }
            catch (Exception ex)
            {
                //Log写入
                LogHelper.GetSingleObj().Error("基础数据初始化", ex.Message);
                MessageBox.Show("Please generate the database XML file!", "ERROR");
                this.Close();
            }
            //BasicData全局变量赋值
            BASICDATASAT = dataSetBasicData;
        }
        /// <summary>
        /// 文件保存路径选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnPrintFilePathOpen_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "Please select the print file folder path";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = dialog.SelectedPath;
                txtPrintFilePath.Text = savePath;
                SelectedPath = savePath;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        /// <summary>
        /// 出力按钮处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnOutput_Click(object sender, EventArgs e)
        {
            if (txtPrintFilePath.Text == "")
            {
                //Log写入
                LogHelper.GetSingleObj().Error("PDF文件出力", "Please select the print file save path");
                MessageBox.Show("Please select the print file save path!", "ERROR");
                return;
            }
            string strFunctionName = labelRecieve.Text;
            //PPI画面迁移
            if(strFunctionName.Equals("PPI"))
            {
                DataSet dsPPISumList = SUMLISTDATASAT.Copy();
                DataSet dsPPIBasicData = BASICDATASAT;
                PPIController ppiController = new PPIController();
                ppiController.XXXXXX(dsPPISumList, dsPPIBasicData);
            }

            MessageBox.Show("The document has been printed!");
        }

    }
}
