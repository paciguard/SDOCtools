using AcAdAmComm;
using System.Windows.Forms;

namespace SESDTools
{
    /// <summary>
    /// 类    名：AdministratorInformationForm
    /// 功能描述：管理者信息确认画面类
    /// 创 建 人：SY
    /// 创建时间：2021/07/23 10:43
    /// 描    述：管理者信息确认
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    public partial class AdministratorInformationForm : Form
    {
        public AdministratorInformationForm()
        {
            //画面居中表示
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();

            string nodeNameAdminiPhone = "adminiPhone";
            string nodeNameAdminiMail = "adminiMail";
            //Phone
            txtPhone.Text = XMLHelper.ReadAdminiContactInformationSettings(nodeNameAdminiPhone);
            //Mail
            txtMail.Text = XMLHelper.ReadAdminiContactInformationSettings(nodeNameAdminiMail);
        }
    }
}
