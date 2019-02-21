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
            //创建导入工作表基本样式
            share.excelEdit.wb = share.ExcelApp.ActiveWorkbook; //指定工作薄
            if (null == share.excelEdit.GetSheet("数据导入工作表"))
            {
                create_importWorksheet();
            }
        }
        private Excel.Worksheet create_importWorksheet()
        {
            share.excelEdit.wb = share.ExcelApp.ActiveWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet importWorkSheet = share.excelEdit.AddSheet("数据导入工作表");
            importWorkSheet.Cells[1, 1] = "数据导入工作表";
            share.excelEdit.UniteCells(importWorkSheet, 1, 1, 1, 10);   //合并单元格
            importWorkSheet.Cells[2, 1] = "学号";
            importWorkSheet.Cells[2, 2] = "姓名";
            importWorkSheet.Cells[2, 3] = "课程1";
            importWorkSheet.Cells[2, 4] = "课程2";
            importWorkSheet.Cells[2, 5] = "..课程n";
            importWorkSheet.Cells[2, 6] = "四级";
            importWorkSheet.Cells[2, 7] = "六级";
            importWorkSheet.Cells[2, 8] = "目前绩点";
            return importWorkSheet;
        }
        

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

            //分配变量
            share.excelEdit.wb = share.ExcelApp.ActiveWorkbook; //指定工作薄
            //string WorkbookName = share.ExcelApp.ActiveWorkbook.Path + "\\"+share.ExcelApp.ActiveWorkbook.Name;
            Excel.Worksheet ClassSheet = share.excelEdit.AddSheet("班级总体学习情况");
            Excel.Worksheet IndividualSheet = share.excelEdit.AddSheet("个人学习情况分析");
            Excel.Worksheet LessonSheet = share.excelEdit.AddSheet("课程学习情况分析");
            Excel.Worksheet importWorkSheet = null;
            if (null != share.excelEdit.GetSheet("数据导入工作表"))
            {
                importWorkSheet = share.excelEdit.GetSheet("数据导入工作表");
            }
            else
            {
                importWorkSheet = create_importWorksheet();
            }
            share.ClassSheet = ClassSheet;
            share.IndividualSheet = IndividualSheet;
            share.LessonSheet = LessonSheet;
            share.importWorkSheet = importWorkSheet;
            share.dataAnalysis.wb = share.ExcelApp.ActiveWorkbook;
            //清除任务窗格中的记录
            share.myUserControl_individual.cleanListBox();
            share.myUserControl_Lesson.cleanListBox();




            share.dataAnalysis.analyClassStudyStatus(importWorkSheet,ClassSheet);
            share.dataAnalysis.analyIndividualStatus(importWorkSheet,IndividualSheet);
            share.dataAnalysis.analyLessonStatus(importWorkSheet,LessonSheet);


            share.rendering_diagram.renderClassSheet(ClassSheet,"班级总体学习情况");
            share.rendering_diagram.renderIndividualSheet(IndividualSheet,"个人学习情况分析");
            share.rendering_diagram.renderLessonSheet(LessonSheet, "课程学习情况分析");


            //补丁V1.0 
            //防止提前点击任务窗口
            share.TaskPane_Ready = 1;
            share.myUserControl_Lesson.AllButtonEnable();
            share.myUserControl_individual.AllButtonEnable();
            //防止第一次运行任务窗口不显示            
            if (share.ExcelApp.ActiveSheet != null && share.TaskPane_Ready== 1)
            {
                if (share.ExcelApp.ActiveSheet.Name == "课程学习情况分析")
                {
                    share.myCustomTaskPane_individual.Visible = false;
                    share.myCustomTaskPane_Lesson.Visible = true;
                }
                else if (share.ExcelApp.ActiveSheet.Name == "个人学习情况分析")
                {
                    share.myCustomTaskPane_Lesson.Visible = false;
                    share.myCustomTaskPane_individual.Visible = true;

                }
                else
                {
                    share.myCustomTaskPane_individual.Visible = false;
                    share.myCustomTaskPane_Lesson.Visible = false;
                }
            }

        }
        /*
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
           if( share.myCustomTaskPane_Lesson.Visible == true)
            {
                share.myCustomTaskPane_Lesson.Visible = false;
            }
            else
            {
                share.myCustomTaskPane_Lesson.Visible = true;
            }

        }
        */
    }
}
