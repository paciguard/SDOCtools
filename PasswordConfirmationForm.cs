using AcAdAmComm;
using System;
using System.Windows.Forms;

namespace SESDTools
{
    /// <summary>
    /// 类    名：PasswordConfirmationForm
    /// 功能描述：密码确认画面类
    /// 创 建 人：SY
    /// 创建时间：2021/07/23 10:43
    /// 描    述：密码是否正确确认
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    public partial class PasswordConfirmationForm : Form
    {
        public PasswordConfirmationForm()
        {
            //画面居中表示
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDone_Click(object sender, EventArgs e)
        {
            string nodeNamePassword = "password";
            string strPassword = XMLHelper.ReadAdminiPasswordSettings(nodeNamePassword);
            if(string.IsNullOrEmpty(txtPassword.Text))
            {
                //Log写入
                LogHelper.GetSingleObj().Error("密码确认", "Please enter a password");
                MessageBox.Show("Please enter a password!", "ERROR");
                return;
            }
            if(strPassword.Equals(txtPassword.Text))
            {
                this.Hide();
                AdministratorForm fm = new AdministratorForm();
                fm.ShowDialog();
            }
            else
            {
                //Log写入
                LogHelper.GetSingleObj().Error("密码确认", "The password you entered is incorrect. Please enter it again");
                MessageBox.Show("The password you entered is incorrect. Please enter it again!", "ERROR");
                return;
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
    }
}
