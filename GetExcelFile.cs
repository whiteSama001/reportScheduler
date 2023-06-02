using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mailScheduler
{
    public class GetExcelFile
    {
        public void GetDataAndExportExcel()
        {
            DataSet ds = null;
            using (SqlConnection con = new SqlConnection("Data Source=./;Initial Catalog=Loans;Integrated Security=True"))
            {
                try
                {
                    con.Open();
                    Console.WriteLine("About fetch from database...");
                    SqlCommand cmd = new SqlCommand("SELECT * FROM [Loans].[dbo].[CreditApproval]", con);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter da = new SqlDataAdapter();
                    da.SelectCommand = cmd;
                    ds = new DataSet();
                    da.Fill(ds);
                    Console.WriteLine("Database spool successful");

                    ExportDataSetToExcel(ds);
                    Console.WriteLine("Excel file successfully generated");
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    ds?.Dispose();
                    con.Close();
                }
            }
        }

        private void ExportDataSetToExcel(DataSet ds)
        {
            string AppLocation = "";
            AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            AppLocation = AppLocation.Replace("file:\\", "");
            string file = AppLocation + "\\ExcelFiles\\Bet9jaReport.xlsx";
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ds.Tables[0]);
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;
                wb.SaveAs(file);
            }
        }
    }
}
