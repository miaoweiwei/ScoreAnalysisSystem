using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Logging;

namespace ScoreAnalysisSystem.Common
{
    public static class ExcelHelper
    {
        #region Excel相关对象的获取 application，workbook，worksheet等

        /// <summary>
        /// 获取当前excel application对象
        /// </summary>
        /// <returns></returns>
        public static Excel.Application GetXlApplication()
        {
            return (Excel.Application)ExcelDnaUtil.Application;
        }

        /// <summary>
        /// 获取当前活动的Workbook
        /// </summary>
        /// <returns></returns>
        public static Excel.Workbook GetCurrentWorkbook()
        {
            Excel.Application app = GetXlApplication();
            return app.ActiveWorkbook;
        }

        /// <summary>
        /// 获取当前活动的WorkSheet
        /// </summary>
        /// <returns></returns>
        public static Excel.Worksheet GetCurrentWorksheet()
        {
            //Excel.Workbook workbook = GetCurrentWorkbook();
            //return workbook.ActiveSheet;
            Excel.Application app = GetXlApplication();
            return app.ActiveSheet;
        }

        #endregion

        #region 设置操作进度

        /// <summary>
        /// 状态图标（未完成）
        /// </summary>
        //private const char ImgUnfinishedStartBar = '□';
        private const char ImgUnfinishedStartBar = '○';

        /// <summary>
        /// 状态图标（已经完成）
        /// </summary>
        //private const char ImgFinishStartBar = '■';
        private const char ImgFinishStartBar = '●';

        /// <summary>
        /// 状态栏显示小图标的总个数
        /// </summary>
        private const int BarShowCount = 20;

        internal delegate List<string> SendRequestSyncCallBack(Dictionary<string, dynamic> date);

        /// <summary>
        /// 设置进度条
        /// </summary>
        /// <param name="callBack"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        internal static object SetStatusBar(SendRequestSyncCallBack callBack, Dictionary<string, dynamic> param)
        {
            bool runStart = true;
            object result = null;
            Thread thread = new Thread(() =>
            {
                result = callBack(param);
                runStart = false;
            });
            thread.Start();

            while (runStart)
            {
                Thread.Sleep(100);
                SetStatusBar(0, 0, ChangeBar(20));
            }

            SetStatusBar(0, 0, "");
            return result;
        }

        private static int index = 0;
        private static string bar = "";

        public static string ChangeBar(int barLength)
        {
            if (barLength == 0)
            {
                return bar;
            }

            if (bar.Length == 0)
            {
                bar = bar.PadRight(barLength, ImgUnfinishedStartBar);
            }

            if (index >= bar.Length)
            {
                index = bar.Length;
            }

            char[] point = bar.ToCharArray();
            point[index] = (point[index] == ImgUnfinishedStartBar ? ImgFinishStartBar : ImgUnfinishedStartBar);
            bar = new string(point);

            if (index == (bar.Length - 1))
            {
                index = 0;
            }
            else
            {
                index++;
            }

            return bar;
        }

        static readonly Excel.Application XlApp = GetXlApplication();

        /// <summary>
        /// 设置进度条 
        /// </summary>
        /// <param name="curValue"></param>
        /// <param name="maxValue"></param>
        /// <param name="desText"></param>
        internal static void SetStatusBar(int curValue, int maxValue, string desText = "当前计算进度：")
        {
            if (XlApp == null)
            {
                return;
            }

            if (curValue == 0 && maxValue == 0)
            {
                if (String.IsNullOrEmpty(desText))
                    XlApp.StatusBar = false;
                else
                    XlApp.StatusBar = desText;
                return;
            }

            string s = "";
            if (curValue >= maxValue)
            {
                s = s.PadRight(BarShowCount, ImgFinishStartBar);
                XlApp.StatusBar = desText + s + "  100%";
            }

            int j = curValue / maxValue * BarShowCount;
            s = s.PadRight(j, ImgFinishStartBar);
            s = s.PadRight(BarShowCount, ImgUnfinishedStartBar);
            var progress = curValue * 1.0 / maxValue * 100;
            XlApp.StatusBar = desText + s + "  " + progress.ToString("F") + "%";
            if (progress.Equals(100))
            {
                XlApp.StatusBar = false;
            }
        }

        #endregion

        #region 获取Excel里的数据

        /// <summary>
        /// 在指定Sheet上获取指定起始Cell和指定结束Cell的Range
        /// </summary>
        /// <param name="sheet">指定Sheet</param>
        /// <param name="startRow">指定起始Cell的行号</param>
        /// <param name="startColumn">指定起始Cell的列号</param>
        /// <param name="endRow">指定结束Cell的行号</param>
        /// <param name="endColumn">指定结束Cell的列号</param>
        /// <returns></returns>
        public static Excel.Range GetRange(Excel.Worksheet sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Excel.Range rng = sheet.Range[(Excel.Range)sheet.Cells[startRow, startColumn], (Excel.Range)sheet.Cells[endRow, endColumn]];
            return rng;
        }
        /// <summary>
        /// 获取指定Sheet里所有已经使用的Range里的数据
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static string[,] GetDataForExcel(Excel.Worksheet worksheet)
        {
            string[,] data = GetDataForExcel(worksheet, out var row, out var col);
            return data;
        }

        /// <summary>
        /// 获取指定Sheet里所有已经使用的Range里的数据
        /// </summary>
        /// <param name="worksheet">指定的WorkSheet</param>
        /// <param name="startRow">Range的开始行号</param>
        /// <param name="startCol">Range的开始列号</param>
        /// <returns></returns>
        public static string[,] GetDataForExcel(Excel.Worksheet worksheet, out int startRow, out int startCol)
        {
            Excel.Range useRange = worksheet.UsedRange;
            object obj = useRange.Value2;
            startCol = useRange.Column;
            startRow = useRange.Row;
            if (obj == null)
                return null;
            if (obj.GetType().Name == "Object[,]")//数据转换为字符串
            {
                object[,] objArr = (object[,])obj;
                string[,] data = ObjectToStrings(objArr);
                return data;
            }
            else//只有一个单元格被使用了
            {
                string[,] data = { { obj.ToString() }, };
                return data;
            }

        }

        /// <summary>
        /// Object的二维数组转成string的二维数组
        /// </summary>
        /// <param name="objArr"></param>
        /// <returns></returns>
        public static string[,] ObjectToStrings(object[,] objArr)
        {
            //数据转换为字符串
            string[,] data = new string[objArr.GetLength(0), objArr.GetLength(1)];
            for (int i = 0; i < objArr.GetLength(0); i++) //行
                for (int j = 0; j < objArr.GetLength(1); j++) //列
                    data[i, j] = objArr[i + 1, j + 1] == null ? "" : objArr[i + 1, j + 1].ToString();
            return data;
        }

        #endregion

        #region 数据转换

        /// <summary>
        /// 将一个二维数组转换成DataTable
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="tableTitle"></param>
        /// <returns></returns>
        public static DataTable ConvertToDataTable(string[,] arr, string tableTitle = "")
        {
            DataTable dt = new DataTable(tableTitle);
            if (arr == null)
                return dt;
            for (int i = 0; i < arr.GetLength(1); i++) //根据数据的列数来设置列头
                dt.Columns.Add(new DataColumn(arr[0, i], arr[0, 0].GetType()));
            for (int i = 1; i < arr.GetLength(0); i++)//去掉第一行的列头
            {
                DataRow newRow = dt.NewRow();
                for (int j = 0; j < arr.GetLength(1); j++)
                    newRow[j] = arr[i, j];
                dt.Rows.Add(newRow);
            }
            return dt;
        }

        #endregion

        #region 数据填充

        /// <summary>
        /// 填充数据到指定的Range中
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endRow"></param>
        /// <param name="endCol"></param>
        /// <param name="objData"></param>
        public static void SetData(object objData, Excel.Worksheet sheet, int startRow, int startCol, int endRow, int endCol)
        {
            try
            {
                var dataRange = GetRange(sheet, startRow, startCol, endRow, endCol);
                dataRange.Value = objData;
            }
            catch (Exception ee)
            {
                LogUtil.Error("数据写入错误" + ee);
            }
        }

        /// <summary>
        /// 用excel表格展示数据
        /// </summary>
        /// <param name="startRow"> 开始单元格的行</param>
        /// <param name="startCol">开始单元格的列</param>
        /// <param name="objData">展示的内容</param>
        /// <param name="worksheet"></param>
        public static void PrintToExcel(object[,] objData, Excel.Worksheet worksheet, int startRow, int startCol)
        {
            Excel.Range xlRange = worksheet.Range[worksheet.Cells[startRow, startCol], worksheet.Cells[startRow + objData.GetLength(0) - 1, startCol + objData.GetLength(1) - 1]];
            xlRange.Value2 = objData;
        }

        /// <summary>
        /// 打印数据到Excel指定的Sheet，从指定的Cell开始填充数据
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="worksheet"></param>
        /// <param name="startRow"></param>
        /// <param name="startCol"></param>
        public static void PrintToExcel(DataTable dt, Excel.Worksheet worksheet, int startRow, int startCol)
        {
            int rowCount = dt.Rows.Count + 1;//加上列头
            int colCount = dt.Columns.Count;
            object[,] objData = new object[rowCount, colCount];

            for (int i = 0; i < dt.Columns.Count; i++)
                objData[0, i] = dt.Columns[i].ColumnName;

            for (int i = 0; i < dt.Rows.Count; i++)
                for (int j = 0; j < dt.Rows[i].ItemArray.Length; j++)
                    objData[i + 1, j] = dt.Rows[i].ItemArray[j];
            PrintToExcel(objData, worksheet, startRow, startCol);
        }

        /// <summary>
        /// 把object[,]从当前选中的cell开始填充到excel
        /// </summary>
        /// <param name="objectArr"></param>
        public static void PrintToExcelAtCurrentCellStart(object[,] objectArr)
        {
            if (objectArr == null || objectArr.GetLength(0) <= 0 || objectArr.GetLength(1) <= 0)
                objectArr = new object[,] { { "无数据", "" }, };
            Excel.Application xlApp = GetXlApplication();
            Excel.Worksheet xlSheet = xlApp.ActiveSheet;
            var xlRange = (Excel.Range)xlApp.Selection;
            var startRow = xlRange.Row;
            var startCol = xlRange.Column;

            int endRow = objectArr.GetLength(0) + startRow - 1;
            int endCol = objectArr.GetLength(1) + startCol - 1;

            SetData(objectArr, xlSheet, startRow, startCol, endRow, endCol);
        }

        #endregion
    }
}