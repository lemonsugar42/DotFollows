using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new Form1());
        }
    }
}

namespace ExcelApp
{
    public class Excel
    {
        private static Microsoft.Office.Interop.Excel.Application excelApp;
        public static Microsoft.Office.Interop.Excel.Application ExcelApp()
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null) throw new Exception("Excel is not installed");
            else
            {
                //excelApp.DisplayAlerts = false;
                return excelApp;
            }
        }
        private static bool Launch()
        {
            bool first;
            string curdir = Environment.CurrentDirectory;
            try
            {
                excelApp.Workbooks.Open($@"{curdir}\..\..\Database\Records.xlsx", Editable: true);
                first = false;
            }
            catch
            {
                Directory.CreateDirectory($@"{curdir}\..\..\Database");
                excelApp.Workbooks.Add();
                excelApp.ActiveWorkbook.SaveAs($@"{curdir}\..\..\Database\Records.xlsx");
                first = true;
            }
            return first;
        }
        private static void Sort(Worksheet excelSheet, int rows)
        {
            Range records = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rows, 3]];
            records.Sort(records.Columns[3], XlSortOrder.xlDescending, records.Columns[1], Type.Missing, XlSortOrder.xlAscending);
        }
        public static void Update(string name, string score)
        {
            bool first = Launch();
            Worksheet excelSheet = excelApp.ActiveWorkbook.Sheets[1];
            int id = 1;
            if (!first)
            {
                id = excelSheet.UsedRange.Rows.Count + 1;
            }
            excelSheet.Cells[id, 1] = id;
            excelSheet.Cells[id, 2] = name;
            excelSheet.Cells[id, 3] = score;
            Sort(excelSheet, id);
            excelApp.ActiveWorkbook.Save();
            excelApp.Visible = true;
            //excelApp.ActiveWorkbook.Close();
        }
    }
}
