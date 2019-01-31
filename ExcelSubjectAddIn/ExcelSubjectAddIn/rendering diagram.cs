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
            ExcelGraph p = new ExcelGraph();
            Excel.Range data = LessonSheet.Range["B2:F2,B3:F3"];
            //p.CreateChart(share.ExcelApp.ActiveWorkbook, LessonSheet, data, );
        }
        public void addChart_LessonSheet(Excel.Worksheet LessonSheet, string ClassName, int Chart_index)
        {
            string Range_string = "B2:F2";
            int mark_row = 0;
            for(int i= share.lessonMenu_row + 1; i<=share.subject_num; i++ )
            {
                if(LessonSheet.Cells[i,1].value == ClassName)
                {
                    mark_row = i;
                    Range_string = Range_string + "," + "B" + mark_row + ":" + "F" + mark_row;
                    break;
                }
            }
            
            ExcelGraph p = new ExcelGraph();
            Excel.Range data = LessonSheet.Range[Range_string];
            p.CreateChart(share.ExcelApp.ActiveWorkbook, LessonSheet, data, ClassName, Chart_index);
            
            /// <param name="m_Book">_Workbook</param>
            /// <param name="m_Sheet">_Worksheet</param>
            /// <param name="CharTop">距页面顶部位置（按格数算）</param>
            /// <param name="CharLeft">距页面左侧位置（按格数算）</param>
            /// <param name="Width">图表外框宽度</param>
            /// <param name="Height">图表外框高度</param>
            /// <param name="Title">图表标题名称</param>
            /// <param name="range">要插入图表的范围值</param>
            /// <param name="CategoryLabels">类别标签值</param>
            /// <param name="SeriesLabels">系列标签值</param>
            /// <param name="MinimumScale">x轴最小值</param>
            /// <param name="MaximumScale">x轴最大值</param>
            /// <param name="CharName">图表名称(为了区份操作的不是一个图，无其他用处)</param>


            //Excel.Range data2 = LessonSheet.Range[Range_string];
            //share.excelEdit.CreateRadarChart(share.ExcelApp.ActiveWorkbook, LessonSheet, 1, 11, 288, 200, ClassName, data2, ClassName);


        }
        private string NunToChar(int number)
        {
            if (65 <= number && 90 >= number)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] btNumber = new byte[] { (byte)number };
                return asciiEncoding.GetString(btNumber);
            }
            return "数字不在转换范围内";
        }

        public void addChart_IndividualSheet(Excel.Worksheet IndividualSheet, string  StudentNumber, int Chart_index)
        {
            string Range_string = "C3:" + NunToChar(64 + 2 + share.subject_num -3) + "3";
            int mark_row = 0;
    

            for (int i = 0 ; i < share.subject_num; i++)
            {
                if (Convert.ToString(IndividualSheet.Cells[i + share.individualMenu_row + 2, 1].value) == StudentNumber)
                {
                    mark_row = i + share.individualMenu_row + 2;
                    Range_string = Range_string + "," + "C" + mark_row + ":" + NunToChar(64 + 2 + share.subject_num - 3) + mark_row;
                    break;
                }
            }  
            Excel.Range data = IndividualSheet.Range[Range_string];

            share.excelEdit.CreateRadarChart(share.ExcelApp.ActiveWorkbook, IndividualSheet, 1, 11, 288, 200, StudentNumber, data, StudentNumber, Chart_index);

        }

        public void delChart_LessonSheet(Excel.Worksheet LessonSheet, string ClassName)
        {
            string Range_string = "B2:F2";
            int mark_row = 0;
            for (int i = share.lessonMenu_row + 1; i <= share.subject_num; i++)
            {
                if (LessonSheet.Cells[i, 1].value == ClassName)
                {
                    mark_row = i;
                    Range_string = Range_string + "," + "B" + mark_row + ":" + "F" + mark_row;
                    break;
                }
            }

            ExcelGraph p = new ExcelGraph();
            Excel.Range data = LessonSheet.Range[Range_string];
            //p.CreateChart(share.ExcelApp.ActiveWorkbook, LessonSheet, data, ClassName);
            
        }


    }
}
