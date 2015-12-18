using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace 傑偲工程管理系統_途程簡易新增程式
{
    public partial class Login : Form
    {
        string SQLCon = Properties.Settings.Default["OpenSQL"].ToString();
        string SQLComm = "";
        public string User = "";
        string UnitId = "EE";
        string UserId = "EM041";
        public Login()
        {
            InitializeComponent();
        }

        private void LoginForm_FormClosing(object sender,FormClosingEventArgs e)
        {
            if (User == "")
            {
                Application.Exit();
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if(txtUserName.Text=="" | txtPassword.Text=="")
            {
                MessageBox.Show("帳號和密碼不可為空白！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                SQLComm = "select UserId,UserName,UnitId from CURdUsers where UserId='" + txtUserName.Text +
                    "' and UserPassword collate Chinese_Taiwan_Stroke_CS_AS='" + txtPassword.Text + "'";
                using (SqlConnection sqlcon = new SqlConnection(SQLCon))
                {
                    sqlcon.Open();
                    using (SqlCommand sqlcomm = new SqlCommand(SQLComm, sqlcon))
                    {
                        SqlDataReader reader = sqlcomm.ExecuteReader();
                        if(reader.HasRows)
                        {
                            reader.Read();
                            if (reader.GetValue(0).ToString() == UserId | reader.GetValue(2).ToString() == UnitId)
                            {
                                DialogResult = DialogResult.OK;
                                User = reader.GetValue(0).ToString() + reader.GetValue(1).ToString();
                            }
                            else
                            {
                                MessageBox.Show("您沒有此程式操作權限！" + Environment.NewLine + "程式即將結束。", "提示訊息",
                                    MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                Application.Exit();
                            }
                        }
                        else
                        {
                            MessageBox.Show("帳號或密碼錯誤！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }
        
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
