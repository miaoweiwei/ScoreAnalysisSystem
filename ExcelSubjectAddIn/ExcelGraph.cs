using System;

using System.Collections.Generic;

using System.Text;

using System.IO;

using System.Data;

using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using  System.Windows.Forms;
namespace ExcelSubjectAddIn
{

    class ExcelGraph
    {

        private static string strCurrentPath = @"D:\";

        private static string title = "testGraph";

        public void CreateExcel(string title, string fileName, string sheetNames)
        {
            //待生成的文件名称
            string FileName = fileName;

            string FilePath = strCurrentPath + FileName;

            FileInfo fi = new FileInfo(FilePath);

            if (fi.Exists)     //判断文件是否已经存在,如果存在就删除!
            {

                fi.Delete();

            }

            if (sheetNames != null && sheetNames != "")
            {

                Microsoft.Office.Interop.Excel.Application m_Excel = new Microsoft.Office.Interop.Excel.Application();//创建一个Excel对象(同时启动EXCEL.EXE进程) 

                m_Excel.SheetsInNewWorkbook = 1;//工作表的个数 

                Microsoft.Office.Interop.Excel._Workbook m_Book = (Microsoft.Office.Interop.Excel._Workbook)(m_Excel.Workbooks.Add(Missing.Value));//添加新工作簿 

                Microsoft.Office.Interop.Excel._Worksheet m_Sheet = (Microsoft.Office.Interop.Excel._Worksheet)(m_Excel.Worksheets.Add(Missing.Value));

                #region 处理

                DataTable auto = new DataTable();

                auto.Columns.Add("LaunchName");

                auto.Columns.Add("Usage");

                auto.Rows.Add(new Object[] { "win8 apac", "100" });
                auto.Rows.Add(new Object[] { "win8 china", "200" });
                auto.Rows.Add(new Object[] { "win8 india", "300" });
                // DataSet ds = ScData.ListData("exec Vote_2008.dbo.P_VoteResult_Update " + int.Parse(fdate));
                DataTableToSheet(title, auto, m_Sheet, m_Book, 1);

                #endregion

                #region 保存Excel,清除进程

                m_Book.SaveAs(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                //m_Excel.ActiveWorkbook._SaveAs(FilePath, Excel.XlFileFormat.xlExcel9795, null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);

                m_Book.Close(false, Missing.Value, Missing.Value);

                m_Excel.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(m_Book);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(m_Excel);

                m_Book = null;

                m_Sheet = null;

                m_Excel = null;

                GC.Collect();

                //this.Close();//关闭窗体

                #endregion

            }

        }


        /// <summary>

        /// 将DataTable中的数据写到Excel的指定Sheet中

        /// </summary>

        /// <param name="dt"></param>

        /// <param name="m_Sheet"></param>

        public void DataTableToSheet(string title, DataTable dt, Microsoft.Office.Interop.Excel._Worksheet m_Sheet,

                                    Microsoft.Office.Interop.Excel._Workbook m_Book, int startrow)
        {
            //以下是填写EXCEL中数据

            Microsoft.Office.Interop.Excel.Range range = m_Sheet.Range[m_Sheet.Cells[1, 1], m_Sheet.Cells[1, 2]];
            // range.MergeCells = true;  //合并单元格

            range.Font.Bold = true;   //加粗单元格内字符

            //写入题目

            m_Sheet.Cells[startrow, startrow] = title;

            int rownum = dt.Rows.Count;//行数

            int columnnum = dt.Columns.Count;//列数

            int num = rownum + 2;   //得到数据中的最大行数

            //写入列标题

            for (int j = 0; j < columnnum; j++)
            {

                int bt_startrow = startrow + 1;

                //将字段名写入文档

                m_Sheet.Cells[bt_startrow, 1 + j] = dt.Columns[j].ColumnName;

                //单元格内背景色
                m_Sheet.Range[m_Sheet.Cells[bt_startrow, 1 + j], m_Sheet.Cells[bt_startrow, 1 + j]].Interior.ColorIndex = 15;

            }

            //逐行写入数据 

            for (int i = 0; i < rownum; i++)
            {

                for (int j = 0; j < columnnum; j++)
                {

                    m_Sheet.Cells[startrow + 2 + i, 1 + j] = dt.Rows[i][j].ToString();

                }

            }

            m_Sheet.Columns.AutoFit();

            //在当前工作表中根据数据生成图表

           // CreateChart(m_Book, m_Sheet, num);

        }



        public void CreateChart(Microsoft.Office.Interop.Excel._Workbook m_Book, Microsoft.Office.Interop.Excel._Worksheet m_Sheet, Excel.Range data, string ChartName, int Chart_index, int Chart_LocationToLeft) 
        {

            Microsoft.Office.Interop.Excel.Range oResizeRange;

            Microsoft.Office.Interop.Excel.Series oSeries;

            m_Book.Charts.Add(Missing.Value, Missing.Value, 1, Missing.Value);
       
            m_Book.ActiveChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;//设置图形
            m_Book.ActiveChart.ChartStyle = 253;
            
            //设置数据取值范围

            m_Book.ActiveChart.SetSourceData(data, Microsoft.Office.Interop.Excel.XlRowCol.xlRows);

            //m_Book.ActiveChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAutomatic, ChartName);
            
            //以下是给图表放在指定位置

            m_Book.ActiveChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAutomatic, m_Sheet.Name);
            
            oResizeRange = (Microsoft.Office.Interop.Excel.Range)m_Sheet.Rows.get_Item(10, Missing.Value);

            m_Sheet.Shapes.Item("Chart 1").Name = ChartName;
            //调图表的位置上边距
            m_Sheet.Shapes.Item(ChartName).Top = Chart_index*200;  
            //调图表的位置左边距
            oResizeRange = (Microsoft.Office.Interop.Excel.Range)m_Sheet.Columns.get_Item(Chart_LocationToLeft, Missing.Value);  
            m_Sheet.Shapes.Item(ChartName).Left = (float)(double)oResizeRange.Left;
            m_Sheet.Shapes.Item(ChartName).Width = 288;   //调图表的宽度
            m_Sheet.Shapes.Item(ChartName).Height = 200;  //调图表的高度
            //m_Book.ActiveChart.PlotArea.Interior.Color = "blue";  //设置绘图区的背景色 
            m_Book.ActiveChart.PlotArea.Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;//设置绘图区边框线条
            m_Book.ActiveChart.PlotArea.Width = 160;
            m_Book.ActiveChart.PlotArea.Height = 120;
            m_Book.ActiveChart.PlotArea.Top = 30;
            m_Book.ActiveChart.PlotArea.Left = 0;
            // m_Book.ActiveChart.ChartArea.Interior.ColorIndex = 10; //设置整个图表的背影颜色
            // m_Book.ActiveChart.ChartArea.Border.ColorIndex = 8;// 设置整个图表的边框颜色
            m_Book.ActiveChart.ChartArea.Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//设置边框线条
            m_Book.ActiveChart.HasDataTable = false;
            m_Book.ActiveChart.HasTitle = true;
            m_Book.ActiveChart.HasLegend = true;
            m_Book.ActiveChart.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 50, 50);

            //设置Legend图例的位置和格式
            //m_Book.ActiveChart.Legend.Top = 50; //具体设置图例的上边距
            m_Book.ActiveChart.Legend.Left = 410;//具体设置图例的左边距
            m_Book.ActiveChart.Legend.Interior.ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone;
            m_Book.ActiveChart.Legend.Width = 100;
            m_Book.ActiveChart.Legend.Font.Size = 12;
            m_Book.ActiveChart.Legend.Font.Bold = true;
            m_Book.ActiveChart.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;//设置图例的位置
            m_Book.ActiveChart.Legend.Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;//设置图例边框线条

            oSeries = (Microsoft.Office.Interop.Excel.Series)m_Book.ActiveChart.SeriesCollection(1);

            oSeries.Border.ColorIndex = 45;
            oSeries.Border.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            //m_Book.ActiveChart.SaveAs("a.pic");

            m_Book.ActiveChart.ChartTitle.Text = ChartName;
        }

    }

}