using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace NewExcelApp
{
    static class Program
    {
        /// <summary>
        /// Ponto de entrada principal para o aplicativo.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

    }

    public class ReadFile
    {
        public DataTable ReadExcel(string fileName, string fileExt, string sheetName)
        {
            string conn;

            DataTable dtExcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source =" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//excel below 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007

            using(OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + sheetName + "$A:I]", con);
                    oleAdpt.Fill(dtExcel); //fill excel data into dataTable 
                }
                catch(Exception ex) {
                    throw ex;
                }
            }

            return dtExcel;
        }
    }
}
