using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
//using Spire.Xls;
namespace ExcelSubjectAddIn
{
    class rendering_diagram
    {
        /*
        public void renderClassSheet(string WorkbookName,string SheetName)
        {
            //加载Excel文档
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(WorkbookName);
            Worksheet ws = workbook.Worksheets[SheetName];
            Chart[] charts = new Chart[10];
            charts[0] = ws.Charts.Add();

            //2.指定生成图表的区域
            charts[0].DataRange = ws.Range["A1:C6"];
            charts[0].SeriesDataFromRange = false;

            //将饼形图所有系列都分离20%

            for (int i = 0; i < charts[0].Series.Count; i++)

            {

                charts[0].Series[i].DataFormat.Percent = 20;

            }

        }
        */
        public void renderClassSheet(Excel.Worksheet ClassSheet, string SheetName)
        {
  
        }
        public void renderIndividualSheet(Excel.Worksheet IndividualSheet, string SheetName)
        {
            
        }
        public void renderLessonSheet(Excel.Worksheet LessonSheet, string SheetName)
        {
            // share.excelEdit.InsertActiveChart(Microsoft.Office.Interop.Excel.XlChartType.xlPie, SheetName, 3, 1, 3, 6, Microsoft.Office.Interop.Excel.XlRowCol.xlRows);
            //Excel.Range data = LessonSheet.Range[LessonSheet.Cells[2, 2], LessonSheet.Cells[2, 6]];
            //Excel.Range data2 = LessonSheet.Range[LessonSheet.Cells[3, 2], LessonSheet.Cells[3, 6]];
            //Excel.Range data = LessonSheet.Range["B2:F2,B3:F3"];
            //share.excelEdit.InsertActiveChart(Microsoft.Office.Interop.Excel.XlChartType.xlPie, SheetName, data, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns);
            ExcelGraph p = new ExcelGraph();
            Excel.Range data = LessonSheet.Range["B2:F2,B3:F3"];
            p.CreateChart(share.ExcelApp.ActiveWorkbook, LessonSheet, data);
            
        }
    }
}
