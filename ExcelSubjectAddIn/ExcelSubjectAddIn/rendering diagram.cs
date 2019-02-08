﻿using System;
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
            //第一行文字渲染
            ClassSheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ClassSheet.Cells[1, 1].Font.Name = "华文中宋";
            ClassSheet.Cells[1, 1].Font.Size = 14;
            //班级情况总结处单元格添加背景颜色
            Excel.Range Summary_range = ClassSheet.get_Range("A2", "B6");
            Summary_range.Interior.ColorIndex = 44;
            Summary_range.Borders.LineStyle = 1;
            Summary_range.EntireColumn.AutoFit();
            //添加单元格边框线
            Excel.Range Data_range = ClassSheet.get_Range("A"+share.classMenu_row, "" + NumToChar(64 + share.subject_num - 3 + 8) + (share.classMenu_row + share.student_num));
            Data_range.Borders.LineStyle = 1;
        }
        public void renderIndividualSheet(Excel.Worksheet IndividualSheet, string SheetName)
        {
            //第一行文字居中
            IndividualSheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            IndividualSheet.Cells[1, 1].Font.Name = "华文中宋";
            IndividualSheet.Cells[1, 1].Font.Size = 14;

            Excel.Range Menu_range = IndividualSheet.get_Range("A2", NumToChar(64 + 9 + share.subject_num - 3) + "3");
            Menu_range.Font.Name = "宋体";
            Menu_range.Font.Size = 11;
            Menu_range.Font.Bold = true;
            Menu_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Menu_range.WrapText = true;

            Excel.Range Menu2_range = IndividualSheet.get_Range("A3", NumToChar(64 + 9 + share.subject_num - 3) + "3");
            Menu2_range.RowHeight = 27;

            Excel.Range Column1_range = IndividualSheet.get_Range("A1", "A" + share.student_num + share.individualMenu_row +1);
            Column1_range.ColumnWidth = 15;
            Column1_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //添加单元格边框线
            Excel.Range Data_range = IndividualSheet.get_Range("A1", "" + NumToChar(64 + share.subject_num - 3 + 9) + (share.student_num + 3));
            Data_range.Borders.LineStyle = 1;
            //将值为“否”的单元格渲染成黄色
            for (int i=0; i< share.student_num; i++)
            {
                string Grade4 = IndividualSheet.Cells[i + share.individualMenu_row + 2, share.subject_num - 3 + 8].value;
                string Grade6 = IndividualSheet.Cells[i + share.individualMenu_row + 2, share.subject_num - 3 + 9].value;
                if (Grade4 == "否")
                {
                    IndividualSheet.Cells[i + share.individualMenu_row + 2, share.subject_num - 3 + 8].Interior.ColorIndex =6;
                }
                if (Grade6 == "否")
                {
                    IndividualSheet.Cells[i + share.individualMenu_row + 2, share.subject_num - 3 + 9].Interior.ColorIndex = 6;
                }


            }
        }
        public void renderLessonSheet(Excel.Worksheet LessonSheet, string SheetName)
        {
            Excel.Range Title_range = LessonSheet.get_Range("A1", "" + NumToChar(64 + 10) + 1);
            Title_range.Font.Name = "华文中宋";
            Title_range.Font.Size = 14;
            Title_range.Borders.LineStyle = 1;
            Title_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Excel.Range Menu_range = LessonSheet.get_Range("A2", "" + NumToChar(64 + 10) + 2);
            Menu_range.Font.Name = "宋体";
            Menu_range.Font.Bold = true;
            Menu_range.Font.Size = 11;
            Menu_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Excel.Range Data_range = LessonSheet.get_Range("A2", "" + NumToChar(64 + 10) + (share.subject_num-3 + 2));
            Data_range.Borders.LineStyle = 1;
            Data_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Data_range.ColumnWidth = 13;
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
        private string NumToChar(int number)
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
            string Range_string = "C3:" + NumToChar(64 + 2 + share.subject_num -3) + "3";
            int mark_row = 0;

            for (int i = 0 ; i < share.student_num; i++)
            {
                //查找对应学号所在行mark_row
                string str = Convert.ToString(IndividualSheet.Cells[i + share.individualMenu_row + 2, 1].value);
                if (Convert.ToString(IndividualSheet.Cells[i + share.individualMenu_row + 2, 1].value) == StudentNumber)
                {
                    mark_row = i + share.individualMenu_row + 2;
                    Range_string = Range_string + "," + "C" + mark_row + ":" + NumToChar(64 + 2 + share.subject_num - 3) + mark_row;
                    break;
                }
            }  
            Excel.Range data = IndividualSheet.Range[Range_string];
            string StudentName = IndividualSheet.Cells[mark_row, 2].value;
            string Grade4 = IndividualSheet.Cells[mark_row, share.subject_num -3 +8].value;
            string Grade6 = IndividualSheet.Cells[mark_row, share.subject_num -3 +9].value;
            string subtitleText = "四级: " + Grade4 + "\n" + "六级: " + Grade6;
            share.excelEdit.CreateRadarChart(share.ExcelApp.ActiveWorkbook, IndividualSheet, 1, 11, 288, 200, StudentName, data, StudentNumber, Chart_index, subtitleText);

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
