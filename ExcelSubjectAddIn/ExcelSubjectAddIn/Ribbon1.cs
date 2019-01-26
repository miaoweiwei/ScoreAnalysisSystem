using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
namespace ExcelSubjectAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            share.excelEdit.wb = share.ExcelApp.ActiveWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet importWorkSheet = share.excelEdit.AddSheet("数据导入工作表");
            importWorkSheet.Cells[1, 1] = "数据导入工作表";
            share.excelEdit.UniteCells(importWorkSheet, 1, 1, 1, 10);   //合并单元格
            importWorkSheet.Cells[2, 1] = "学号";
            importWorkSheet.Cells[2, 2] = "姓名";
            importWorkSheet.Cells[2, 3] = "课程1";
            importWorkSheet.Cells[2, 4] = "课程2";
            importWorkSheet.Cells[2, 5] = "课程...";
         
      
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            share.excelEdit.wb = share.ExcelApp.ActiveWorkbook; //指定工作薄
            string WorkbookName = share.ExcelApp.ActiveWorkbook.Path + "\\"+share.ExcelApp.ActiveWorkbook.Name;
            Excel.Worksheet ClassSheet = share.excelEdit.AddSheet("班级总体学习情况");
            Excel.Worksheet IndividualSheet = share.excelEdit.AddSheet("个人学习情况分析");
            Excel.Worksheet LessonSheet = share.excelEdit.AddSheet("课程学习情况分析");
            Excel.Worksheet importWorkSheet = share.excelEdit.GetSheet("数据导入工作表");
            share.dataAnalysis.wb = share.ExcelApp.ActiveWorkbook;
            share.dataAnalysis.analyClassStudyStatus(importWorkSheet,ClassSheet);
            share.dataAnalysis.analyIndividualStatus(importWorkSheet,IndividualSheet);
            share.dataAnalysis.analyLessonStatus(importWorkSheet,LessonSheet);


            share.rendering_diagram.renderClassSheet(ClassSheet,"班级总体学习情况");
            //share.rendering_diagram.renderIndividualStatus(IndividualSheet,"个人学习情况分析");
            share.rendering_diagram.renderLessonSheet(LessonSheet, "课程学习情况分析");
        }
    }
}
