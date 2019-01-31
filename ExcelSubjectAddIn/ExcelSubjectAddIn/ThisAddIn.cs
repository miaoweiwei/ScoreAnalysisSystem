using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
namespace ExcelSubjectAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            share.ExcelApp = Globals.ThisAddIn.Application;   //获取加载项所在的Excel应用程序
            share.myForm = new Form1();
            share.excelEdit = new ExcelEdit();
            share.dataAnalysis = new dataAnalysis();
            share.rendering_diagram = new rendering_diagram();
            share.myUserControl_Lesson = new UserControl1();
            share.myCustomTaskPane_Lesson = Globals.ThisAddIn.CustomTaskPanes.Add(share.myUserControl_Lesson, "课程情况分析");//添加任务窗
            share.myCustomTaskPane_Lesson.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            share.myUserControl_individual = new UserControl2();
            share.myCustomTaskPane_individual = Globals.ThisAddIn.CustomTaskPanes.Add(share.myUserControl_individual, "个人情况分析");
            share.myCustomTaskPane_individual.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            share.ExcelApp.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(Workbook_SheetActivate);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        void Workbook_SheetActivate(object sheet)
        {
            
            ReportEventWithSheetParameter("Workbook.SheetActivate", sheet);
        }

        void ReportEventWithSheetParameter(string eventName, object sheet)
        {
            Excel.Worksheet worksheet = sheet as Excel.Worksheet;

            if (worksheet != null)
            {
                if(worksheet.Name == "课程学习情况分析")
                {
                    share.myCustomTaskPane_individual.Visible = false;
                    share.myCustomTaskPane_Lesson.Visible = true;
                }
                else if(worksheet.Name == "个人学习情况分析")
                {
                    share.myCustomTaskPane_Lesson.Visible = false;
                    share.myCustomTaskPane_individual.Visible = true;

                }
                else
                {
                    share.myCustomTaskPane_individual.Visible = false;
                    share.myCustomTaskPane_Lesson.Visible = false;
                }
                //MessageBox.Show(String.Format("{0} ({1})", eventName, worksheet.Name));
            }

            Excel.Chart chart = sheet as Excel.Chart;

            if (chart != null)
            {
                //MessageBox.Show(String.Format("{0} ({1})", eventName, chart.Name));
            }
        }


        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
