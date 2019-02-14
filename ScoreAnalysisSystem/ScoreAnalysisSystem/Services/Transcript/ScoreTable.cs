using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ScoreAnalysisSystem.Common;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreAnalysisSystem.Services.Transcript
{
    /// <summary>
    /// 成绩单
    /// </summary>
    public class ScoreTable
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger("ScoreTable");

        public ScoreTable()
        {
        }

        /// <summary>
        /// 数据验证
        /// </summary>
        /// <returns></returns>
        public string[,] DataValidation(Excel.Worksheet worksheet)
        {
            string[,] data = ExcelHelper.GetDataForExcel(worksheet, out int startRow, out int startCol);
            if (data == null)
            {
                MyMessageBox.Show("数据填写有误或没有填入数据!", worksheet);
                return null;
            }

            if (startCol != 1 || startRow != 1) //输入数据必须从第一行第一列开始
            {
                MyMessageBox.Show(@"请从“A1”开始填写数据！", worksheet);
                return null;
            }
            return data;
        }

        public void ScoreAnalysis()
        {
        }
        /// <summary>
        /// 计算绩点并且打印到Excel
        /// </summary>
        public bool ScorePrintToExcel()
        {
            Excel.Application excelApp = ExcelHelper.GetXlApplication();
            Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;
            Excel.Worksheet currentWorksheet = currentWorkbook.ActiveSheet;
            string[,] data = DataValidation(currentWorksheet);
            //数据的前面三列必须是班级、学号、姓名 后面为课程名称、四六级。
            //绩点有系统自动计算
            DataTable dt = ExcelHelper.ConvertToDataTable(data, "成绩单");
            if(dt.Rows.Count>0)
                ExcelHelper.PrintToExcel(dt, currentWorksheet, 20, 1);
            return true;
        }

        /// <summary>
        /// 计算绩点
        /// </summary>
        public string[,] CalculatedGrade(string[,] data)
        {
            string[,] score = null;

            return score;
        }
    }
}