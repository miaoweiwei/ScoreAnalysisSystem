using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreAnalysisSystem.Common
{
    public static class MyMessageBox
    {
        public static void Show(string info, Excel.Application app)
        {
            Show(info, app.ActiveWorkbook);
        }

        public static void Show(string info, Excel.Workbook workbook)
        {
            Show(info, workbook.ActiveSheet);
        }

        public static void Show(string info, Excel.Worksheet worksheet)
        {
            Excel.Application app = worksheet.Application;
            Excel.Workbook workbook = app.ActiveWorkbook;
            info = $@"{app.Name} {workbook.Name} {worksheet.Name}{Environment.NewLine}"+ info;
            Show(info);
        }

        public static void Show(string info, string title="提示")
        {
            MessageBox.Show(info, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}