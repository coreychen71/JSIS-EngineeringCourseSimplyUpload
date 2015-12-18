using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace 傑偲工程管理系統_途程簡易新增程式
{
    public partial class MainProgram : Form
    {
        string SQLCon = Properties.Settings.Default["OpenSQL"].ToString();
        string FilePath = "";
        public MainProgram()
        {
            InitializeComponent();
            this.Hide();
            Login log = new Login();
            if(log.ShowDialog()==DialogResult.OK)
            {
                lblUser.Text = "（" + log.User + "）";
                this.Show();
                log.Dispose();
            }
            lblExplain.ForeColor = Color.Red;
        }

        /// <summary>
        /// 檢查Excel版本，回傳對應版本的OleDb連線字串
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <returns></returns>
        private string CheckExcelVersion(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            string extension = file.Extension;
            string strConn = "";
            //透過上傳的Excel副檔名來判斷Excel版本，以對應符合版本的連線字串
            switch (extension)
            {
                //Excel2003
                case ".xls":
                    strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath +
                        ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;Readonly=0'";
                    break;

                //Excel2007 Up
                case ".xlsx":
                    strConn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath +
                        ";Extended Properties='Excel 12.0 Xml;HDR=YES'";
                    break;

                //副檔名都不符合，就套用Excel2003
                default:
                    strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath +
                        ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;Readonly=0'";
                    break;
            }
            return strConn;
        }

        /// <summary>
        /// 取得Excel Sheet名稱然後以DataTable回傳
        /// </summary>
        /// <param name="filepath">檔案路徑</param>
        /// <returns></returns>
        private DataTable GetExcelSheetNames(string filepath)
        {
            OleDbConnection Conn = new OleDbConnection(CheckExcelVersion(filepath));
            DataTable dt = new DataTable();
            Conn.Open();
            dt = Conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            Conn.Close();
            return dt;
        }

        /// <summary>
        /// 使用SQL語法取得Excel Sheet資料內容，以DataTable回傳
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="sql">SQL語法</param>
        /// <returns></returns>
        private DataTable GetExcelDataTable(string filepath, string sql)
        {
            OleDbConnection Conn = new OleDbConnection(CheckExcelVersion(filepath));
            OleDbDataAdapter Load = new OleDbDataAdapter(sql, Conn);
            DataTable dt = new DataTable();
            Exception error = new Exception();
            Load.Fill(dt);
            Conn.Close();
            return dt;
        }

        //選擇要上傳的Excel檔
        private void btnDot_Click(object sender, EventArgs e)
        {
            OpenFileDialog Path = new OpenFileDialog();
            if(Path.ShowDialog()==DialogResult.OK)
            {
                FilePath = Path.FileName;
            }
            txtFilePath.Text = FilePath;
        }

        //讀取Excel內容，並秀在DataGridView
        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (txtFilePath.Text == "")
            {
                MessageBox.Show("尚未選擇要上傳的Excel檔案！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            string tableMas = "Mas$";
            string tableDtl = "Dtl$";
            int SheetNum = 0;
            DataTable dtSheet = GetExcelSheetNames(FilePath);
            //檢查Sheet名稱是否符合規定的名稱格式，若不符就中止程式
            for (int i = 0; i < dtSheet.Rows.Count; i++)
            {
                if (dtSheet.Rows[i]["TABLE_NAME"].ToString() == tableMas | dtSheet.Rows[i]["TABLE_NAME"].ToString()
                    == tableDtl)
                {
                    SheetNum++;
                }
            }
            if(SheetNum!=2)
            {
                MessageBox.Show("Excel檔的Sheet名稱錯誤，請檢查！", "提示訊息", MessageBoxButtons.OK,
                    MessageBoxIcon.Asterisk);
                return;
            }
            //透過SQL語法去撈Excel Sheet資料時，需在Sheet名稱前後補上[]
            string sql = "select * from [" + tableMas + "]";
            DataTable dtMas = GetExcelDataTable(FilePath, sql);
            sql = "select * from [" + tableDtl + "]";
            DataTable dtDtl = GetExcelDataTable(FilePath, sql);
            dgvMasShow.DataSource = dtMas;
            dgvDtlShow.DataSource = dtDtl;
            dgvMasShow.Columns[0].Width = 180;
            dgvMasShow.Columns[1].Width = 70;
            dgvMasShow.Columns[2].Width = 500;
            dgvDtlShow.Columns[0].Width = 180;
            dgvDtlShow.Columns[1].Width = 60;
            dgvDtlShow.Columns[2].Width = 80;
            dgvDtlShow.Columns[3].Width = 380;
            dgvMasShow.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDtlShow.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDtlShow.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        //將DataGridView資料寫入SQL資料庫
        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (dgvMasShow.Rows.Count == 1 & dgvDtlShow.Rows.Count == 1)
            {
                MessageBox.Show("尚未載入資料！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            string SQLComm = "select * from EMOdTmpRouteMas where TmpId='" + dgvMasShow.Rows[0].Cells[0].Value.
                ToString().Trim() + "'";
            int Mas = 0;
            int Dtl = 0;
            int user = 0;
            //先檢查途程編碼是否重複
            using (SqlConnection sqlcon = new SqlConnection(SQLCon))
            {
                sqlcon.Open();
                using (SqlCommand sqlcomm = new SqlCommand(SQLComm, sqlcon))
                {
                    SqlDataReader Read = sqlcomm.ExecuteReader();
                    if(Read.HasRows)
                    {
                        MessageBox.Show("途程編碼和資料庫重複！" + Environment.NewLine + "途程編號為唯一索引值。",
                            "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        return;
                    }
                }
            }
            //開始寫入資料
            using (SqlConnection sqlcon = new SqlConnection(SQLCon))
            {
                sqlcon.Open();
                using (SqlCommand sqlcomm = new SqlCommand())
                {
                    //寫入Mas
                    for (int i = 0; i < dgvMasShow.Rows.Count - 1; i++)
                    {
                        SQLComm = "insert into EMOdTmpRouteMas(TmpId,Notes,Status) values('" + dgvMasShow.Rows[i].
                            Cells[0].Value.ToString().Trim() + "','" + dgvMasShow.Rows[i].Cells[2].Value.ToString()
                            .Trim() + "','" + dgvMasShow.Rows[i].Cells[1].Value.ToString().Trim() + "')";
                        sqlcomm.Connection = sqlcon;
                        sqlcomm.CommandText = SQLComm;
                        if (sqlcomm.ExecuteNonQuery() == 1)
                        {
                            Mas++;
                        }
                    }
                    //寫入Dtl
                    for (int i = 0; i < dgvDtlShow.Rows.Count - 1; i++)
                    {
                        SQLComm = "insert into EMOdTmpRouteDtl(TmpId,SerialNum,ProcCode) values('" + dgvDtlShow.Rows[i].
                            Cells[0].Value.ToString().Trim() + "','" + dgvDtlShow.Rows[i].Cells[1].Value.ToString()
                            .Trim() + "','" + dgvDtlShow.Rows[i].Cells[2].Value.ToString().Trim() + "')";
                        sqlcomm.Connection = sqlcon;
                        sqlcomm.CommandText = SQLComm;
                        if (sqlcomm.ExecuteNonQuery() == 1)
                        {
                            Dtl++;
                        }
                    }
                    //寫入上傳人員與時間
                    SQLComm = "insert into EMOdTmpRouteUploadPeople(TmpId,UserId,UserName,UploadTime) values('" +
                        dgvMasShow.Rows[0].Cells[0].Value.ToString().Trim() + "','" + lblUser.Text.Substring(1, 5) +
                        "','" + lblUser.Text.Substring(6, 3) + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                        "')";
                    sqlcomm.CommandText = SQLComm;
                    if (sqlcomm.ExecuteNonQuery() == 1)
                    {
                        user++;
                    }
                }
            }
            if (Mas == dgvMasShow.Rows.Count - 1 & Dtl == dgvDtlShow.Rows.Count - 1 & user == 1)
            {
                MessageBox.Show("已上傳完成！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
    }
}
