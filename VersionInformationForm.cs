using AcAdAmComm;
using System.Windows.Forms;

namespace SESDTools
{
    /// <summary>
    /// 类    名：VersionInformationForm
    /// 功能描述：版本信息确认画面类
    /// 创 建 人：SY
    /// 创建时间：2021/07/23 10:43
    /// 描    述：版本信息确认
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    public partial class VersionInformationForm : Form
    {
        public VersionInformationForm()
        {
            //画面居中表示
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();

            string nodeNameVersionNumber = "versionNumber";
            string nodeNameVersionUpgradeExplanation = "versionUpgradeExplanation";
            string nodeNameDeveloperContactInformation = "developerContactInformation";
            //版本号
            txtVersionNumber.Text = XMLHelper.ReadVersionInformationSettings(nodeNameVersionNumber);
            //版本升级信息
            txtVersionInstructions.Text = XMLHelper.ReadVersionInformationSettings(nodeNameVersionUpgradeExplanation);
            //开发者电话
            txtDeveloperPhone.Text = XMLHelper.ReadVersionInformationSettings(nodeNameDeveloperContactInformation);

        }
    }
}
