using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using ScoreAnalysisSystem.Common;
using ScoreAnalysisSystem.Properties;
using ScoreAnalysisSystem.Services;
using ScoreAnalysisSystem.Services.Transcript;
using ScoreAnalysisSystem.View;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreAnalysisSystem
{
    [ComVisible(true)]
    public class RibbonMenu : ExcelRibbon, IExcelAddIn
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger("RibbonMenu");
        private IRibbonUI _ribbonUi;
        //保存打开的CTP窗体
        private static readonly List<CustomTaskPane> CustomTaskPaneList = new List<CustomTaskPane>();

        private bool _BtnClassAnalysisEnabled = false;
        private bool _BtnStudentAnalysisEnabled = false;
        private bool _BtnCourseAnalysisEnabled = false;
        private bool _BtnCourseChartVisible = false;
        private bool _BtnStudentChartVisible = false;

        public void RibbonMenu_Load(IRibbonUI ribbonUi)
        {
            Log.Debug("系统初始化成功");
            _ribbonUi = ribbonUi;
        }

        #region 成绩单

        public void BtnDataSortOut_Click(IRibbonControl control)
        {
            ScoreTable scoreTable = new ScoreTable();
            bool dataSort = scoreTable.ScorePrintToExcel();
            if (dataSort)
                MessageBox.Show("Success!");
        }

        public void BtnExamAbsent_Click(IRibbonControl control)
        {
            if (true) //分析成功
            {
                _BtnClassAnalysisEnabled = true;
                _BtnStudentAnalysisEnabled = true;
                _BtnCourseAnalysisEnabled = true;
            }
            else
            {
                _BtnClassAnalysisEnabled = false;
                _BtnStudentAnalysisEnabled = false;
                _BtnCourseAnalysisEnabled = false;
            }

            RefreshControl();
        }

        /// <summary>
        /// 成绩计算公式填写
        /// </summary>
        /// <param name="control"></param>
        public void BtnScoreFormula_Click(IRibbonControl control)
        {
            CloseVisibleCtp();
            //ScoreFormulaUserControl scoreFormula = new ScoreFormulaUserControl();
            CourseCreditUserControl scoreFormula = new CourseCreditUserControl(new string[]{"语文", "数学","英语" });
            CustomTaskPane scoreFormulaPane = CustomTaskPaneFactory.CreateCustomTaskPane(scoreFormula, "成绩计算公式");
            scoreFormulaPane.Width = 230;
            scoreFormulaPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight; //在右边弹出

            CustomTaskPaneList.Add(scoreFormulaPane);
            scoreFormulaPane.Visible = true;
        }

        public void BtnScoreCalculate_Click(IRibbonControl control)
        {
        }

        #endregion

        #region 班级情况

        public void BtnClassAnalysis_Click(IRibbonControl control)
        {
            RefreshControl();
        }

        public bool GetBtnClassAnalysis_Enabled(IRibbonControl control)
        {
            return _BtnClassAnalysisEnabled;
        }

        #endregion

        #region 个人情况

        public void BtnStudentAnalysis_Click(IRibbonControl control)
        {
            if (true) //个人成绩分析成功
            {
                _BtnStudentChartVisible = true;
            }

            RefreshControl();
        }

        public void BtnStudentChart_Click(IRibbonControl control)
        {
        }

        private bool _btnAverageScorePressed = false;

        private bool _btnGradePointPressed = true;

        public bool GetBtnAverageScore_Pressed(IRibbonControl control)
        {
            return _btnAverageScorePressed;
        }

        public bool GetBtnGradePoint_Pressed(IRibbonControl control)
        {
            return _btnGradePointPressed;
        }

        public void BtnAverageScore_Click(IRibbonControl control, bool pressed)
        {
            _btnAverageScorePressed = pressed;
            _btnGradePointPressed = !pressed;

            RefreshControl();
        }

        public void BtnGradePoint_Click(IRibbonControl control, bool pressed)
        {
            _btnGradePointPressed = pressed;
            _btnAverageScorePressed = !pressed;

            RefreshControl();
        }


        public bool GetStudentChart_Visible(IRibbonControl control)
        {
            return _BtnStudentChartVisible;
        }


        public bool GetBtnStudentAnalysis_Enabled(IRibbonControl control)
        {
            return _BtnStudentAnalysisEnabled;
        }

        #endregion

        #region 课程情况

        public void BtnCourseAnalysis_Click(IRibbonControl control)
        {
            if (true) //课程成绩分析成功
            {
                _BtnCourseChartVisible = true;
            }

            RefreshControl();
        }

        public void BtnCourseChart_Click(IRibbonControl control)
        {
        }


        public bool GetBtnCourseAnalysis_Enabled(IRibbonControl control)
        {
            return _BtnCourseAnalysisEnabled;
        }

        public bool GetCourseChart_Visible(IRibbonControl control)
        {
            return _BtnCourseChartVisible;
        }

        #endregion

        public void Start_Click(IRibbonControl control)
        {
            //for (int i = 0; i < 100; i++)
            //{
            //    Thread.Sleep(200);
            //    ExcelHelper.SetStatusBar(i+1,100);
            //}

            Excel.Workbook workbook = ExcelHelper.GetCurrentWorkbook();
            Excel.Worksheet worksheet = ExcelHelper.GetCurrentWorksheet();

            MessageBox.Show(worksheet.Name + " " + worksheet.CodeName, workbook.Name + " " + workbook.CodeName);

        }

        /// <summary>
        /// 刷新全部控件
        /// </summary>
        private void RefreshControl()
        {
            _ribbonUi.Invalidate();
        }

        /// <summary>
        /// 刷新指定Id的控件
        /// </summary>
        private void RefreshControl(string controlId)
        {
            _ribbonUi.InvalidateControl(controlId);
        }

        /// <summary> 关闭当前活动的workbook已经打开的窗体 </summary>
        private static void CloseVisibleCtp()
        {
            if (CustomTaskPaneList.Count > 0)
            {
                for (var i = 0; i < CustomTaskPaneList.Count; i++)
                {
                    CustomTaskPaneList[i].Delete();
                    CustomTaskPaneList[i] = null;
                }
                CustomTaskPaneList.Clear();
            }
        }

        /// <summary>
        /// XLL加载时调用
        /// </summary>
        public void AutoOpen()
        {

        }

        /// <summary>
        /// XLL卸载时调用
        /// </summary>
        public void AutoClose()
        {

        }
    }
}