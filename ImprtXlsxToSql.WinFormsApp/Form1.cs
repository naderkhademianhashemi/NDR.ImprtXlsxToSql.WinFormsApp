using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Data.OleDb;

namespace ImprtXlsxToSql.WinFormsApp
{
    public partial class FormImport : Form
    {
        public FormImport()
        {
            InitializeComponent();
        }

        private void btnBrws_Click(object sender, EventArgs e)
        {
            txtPath.Text = GetPath(txtPath.Text);
        }

        private string GetPath(string startPath)
        {
            string tmpPath = startPath;
            OpenFileDialog fdlg = new OpenFileDialog()
            {
                Title = RSRC.Farsi.SelectFile,
                InitialDirectory = startPath,
                FilterIndex = 1,
                RestoreDirectory = true,
                FileName = startPath
            };

            tmpPath = (fdlg.ShowDialog() == DialogResult.OK) ? fdlg.FileName : startPath;

            return tmpPath;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            ImportXlsx2Sql(txtPath.Text);
        }

        private void ImportXlsx2Sql(string excelFilePath)
        {
            //throw new NotImplementedException();
            string ssqltable = "Categories";

            try
            {
                string ssqlconnectionstring = RSRC.Setting.NWSqlConStr;

                String sheetName = "Sheet1";
                
                String constr = RSRC.Setting.ExcelConStr +
                    excelFilePath +
                    RSRC.Setting.ExcelConStr2;

                var oledbconn = new OleDbConnection(constr);
                var oledbcmd = new OleDbCommand("Select * From [" + sheetName + "$]", oledbconn);
                oledbconn.Open();

                var sda = new OleDbDataAdapter(oledbcmd);
                DataTable data = new DataTable();
                sda.Fill(data);

                OleDbDataReader dr = oledbcmd.ExecuteReader();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(ssqlconnectionstring);
                bulkcopy.DestinationTableName = ssqltable;
                while (dr.Read())
                {
                    bulkcopy.WriteToServer(dr);
                }
                dr.Close();
                oledbconn.Close();
                MessageBox.Show("File imported into sql server successfully.");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}