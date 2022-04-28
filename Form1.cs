using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.Sql;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Diagnostics;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat; // TO AUTOFORMAT THE SHEET.
using System.IO;

namespace Auto_Report_Tool_to_Office
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string cs = ConfigurationManager.ConnectionStrings["SCPM"].ConnectionString;

        string FileName = @"C:\AutoMail\AutoTool.xls";


        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(cs);
            con.Close();

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandText = @"[Proc_2022_04_28_AutoTool_ReportToOffice_1]";
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            con.Open();
            sda.SelectCommand = cmd;
            DataSet ds = new DataSet("myDataset");
            sda.Fill(ds);

            Excel.Application ExcelApp = new Excel.Application();

            Excel.Workbook ExcelWorkBook = null;

            Excel.Worksheet ExcelWorkSheet = null;

            ExcelApp.Visible = true;

            object misValue = System.Reflection.Missing.Value;
            ExcelWorkBook = ExcelApp.Workbooks.Add(misValue);

            List<string> SheetNames = new List<string>();

            SheetNames.Add("Tool");

            try

            {

                for (int i = 1; i < ds.Tables.Count; i++)

                    ExcelWorkBook.Worksheets.Add(); //Adding New sheet in Excel Workbook



                for (int i = 0; i < ds.Tables.Count; i++)

                {

                    int r = 3; // Initialize Excel Row Start Position  = 3


                    ExcelWorkSheet = ExcelWorkBook.Worksheets[i + 1];

                    
                    //Writing Columns Name in Excel Sheet

                    for (int col = 1; col <= ds.Tables[i].Columns.Count; col++)

                        ExcelWorkSheet.Cells[r, col] = ds.Tables[i].Columns[col - 1].ColumnName;

                    r++;

                    //ExcelWorkSheet.Columns.AutoFit();
                    //ExcelWorkSheet.UsedRange.Columns.AutoFit();
                    //aRange.Columns.AutoFit();


                    //Writing Rows into Excel Sheet

                    for (int row = 0; row < ds.Tables[i].Rows.Count; row++) //r stands for ExcelRow and col for ExcelColumn

                    {
                        // Excel row and column start positions for writing Row=1 and Col=1

                        for (int col = 1; col <= ds.Tables[i].Columns.Count; col++)

                            ExcelWorkSheet.Cells[r, col] = ds.Tables[i].Rows[row][col - 1].ToString();
                        r++;
                    }

                    ExcelWorkSheet.Name = SheetNames[i];//Renaming the ExcelSheets

                    ExcelWorkSheet.Columns.AutoFit();
                    ExcelWorkSheet.Rows.AutoFit();
                    ExcelWorkSheet.Columns.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    ExcelWorkSheet.Range["A3:O3"].EntireRow.Font.Bold = true;
                    ExcelWorkSheet.Range["I:O"].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                //ExcelWorkBook.SaveAs(FileName);

                //ExcelWorkSheet.Columns.AutoFit();
                //ExcelWorkSheet.Rows.AutoFit();
                //***************************************************************TITLE**************************************************************
                string datetime_ = DateTime.Now.ToString("MMMM dd, yyyy");

                if (ExcelWorkSheet.Name == "Tool")
                {
                    //var tit = ExcelWorkSheet.Cells[1, 1];
                    var tit = ExcelWorkSheet.Range["A1:O2"];
                    tit.Merge();
                    tit.Value = "Southern California Precision Machining" +
                        "\r\n" + "Auto generate Tool was created on " + datetime_;
                    tit.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    tit.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    //var po_1 = ExcelWorkSheet.Range["H1:I1"];
                    //po_1.Merge();
                    //po_1.Value = "PO# " + gage_Form;
                    //po_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    //var po_2 = ExcelWorkSheet.Range["H2:I2"];
                    //po_2.Merge();
                    ////po_2.Value = "ORDER# ";
                    //po_2.Formula = "=" + '"' + "ORDER# " + '"' + "&" + "TEXT(TODAY()," + '"' + "mmddyy" + '"' + ")";
                    //po_2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    //var po_3 = ExcelWorkSheet.Range["H3:I3"];
                    //po_3.Merge();
                    //po_3.Value = "RECEIVED BY";
                    //po_3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                }



                //*****************************************************************************************************************************

                ExcelApp.DisplayAlerts = false;
                ExcelWorkBook.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, misValue, misValue);


                ExcelWorkBook.Close();

                ExcelApp.Quit();

                Marshal.ReleaseComObject(ExcelWorkSheet);

                Marshal.ReleaseComObject(ExcelWorkBook);

                Marshal.ReleaseComObject(ExcelApp);


                //Chep file vao Server
                string sourceDir = @"C:\AutoMail\AutoTool.xls";
                string desDir = @"\\scpfs01\Shared\AutoMail\AutoTool.xls";

                File.Copy(sourceDir, desDir, true);

                Application.Exit();
            }

            catch (Exception exHandle)
            {
                Console.WriteLine("Exception: " + exHandle.Message);
                Console.ReadLine();
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))

                    process.Kill();

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button1_Click(this, new EventArgs());
        }
    }
}
