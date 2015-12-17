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
        string FilePath = "";
        public MainProgram()
        {
            InitializeComponent();
            lblExplain.ForeColor = Color.Red;
        }

        private DataTable GetExcelSheetNames(string filePath)
        {
            //Office 2003
            //OleDbConnection Conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                //filePath + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'");

            //Office 2007
            OleDbConnection Conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'");

            DataSet ds = new DataSet();
            Conn.Open();
            DataTable dt = Conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            Conn.Close();
            return dt;
        }

        private DataTable GetExcelDataTable(string filePath, string sql)
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

            OleDbConnection Conn = new OleDbConnection(strConn);

            /*
            //Office 2003
            OleDbConnection Conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath +
                ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;Readonly=0'");

            //Office 2007
            OleDbConnection Conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                filePath + ";Extended Properties='Excel 12.0 Xml;HDR=YES'");
            */

            OleDbDataAdapter Load = new OleDbDataAdapter(sql, Conn);
            DataTable dt = new DataTable();
            Load.Fill(dt);
            //dt.TableName = "tmp";
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
            string tableMas = "[Mas$]";
            string tableDtl = "[Dtl$]";
            string sql = "select * from " + tableMas;
            DataTable dtMas = GetExcelDataTable(FilePath, sql);
            sql = "select * from " + tableDtl;
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
            string SQLConn = "server=ERP;database=EW;uid=JSIS;pwd=JSIS";
            string SQLComm = "select * from EMOdTmpRouteMas where TmpId='" + dgvMasShow.Rows[0].Cells[0].Value.
                ToString().Trim() + "'";
            int Mas = 0;
            int Dtl = 0;
            //先檢查途程編碼是否重複
            using (SqlConnection sqlcon = new SqlConnection(SQLConn))
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
            using (SqlConnection sqlcon = new SqlConnection(SQLConn))
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
                }
            }
            if (Mas == dgvMasShow.Rows.Count - 1 & Dtl == dgvDtlShow.Rows.Count - 1)
            {
                MessageBox.Show("已上傳完成！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
    }
}
