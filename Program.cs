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
    class ExcelApp
    {
        public static Microsoft.Office.Interop.Excel.Application NewApp()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null) throw new Exception("Excel is not installed");
            return excelApp;
        }
        public static void Update(Microsoft.Office.Interop.Excel.Application excelApp, string text)
        {
            bool first; // по-хорошему весь блок трай-кэтч надо вынести в bool-метод opening
            try // а еще место сейва другое, поближе к проекту
            {
                excelApp.Workbooks.Open(@"C:\Users\79080\Documents\Книга1.xlsx", Editable: true);
                first = false;
            }
            catch
            {
                excelApp.Workbooks.Add();
                excelApp.ActiveWorkbook.SaveAs(@"C:\Users\79080\Documents\Книга1.xlsx");
                first = true;
            }
            Worksheet excelSheet = excelApp.ActiveWorkbook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int id;
            if (first)
            {
                id = 1;
            }
            else
            {
                for (id = 1; id < rows; id++)
                {
                    if (excelSheet.Cells[id, 1] == null) break;
                }
                id++;
            }
            excelSheet.Cells[id, 1] = id;
            excelSheet.Cells[id, 2] = text;
            excelApp.ActiveWorkbook.Save();
            excelApp.Visible = true;
            //excelApp.ActiveWorkbook.Close();
        }
    }
}
