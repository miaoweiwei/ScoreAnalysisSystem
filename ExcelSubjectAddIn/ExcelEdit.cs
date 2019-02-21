using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
namespace ExcelSubjectAddIn
{
    public class ExcelEdit
    {
        public string mFilename;
        public Microsoft.Office.Interop.Excel.Application app;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;
        public ExcelEdit()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
            //Microsoft.Office.Interop.Excel.Workbook wb = new Microsoft.Office.Interop.Excel.Application.
        }

        public void Create()//创建一个Microsoft.Office.Interop.Excel对象
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(true);
        }
        public void Open(string FileName)//打开一个Microsoft.Office.Interop.Excel文件
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            //wb = wbs.Open(FileName, 0, true, 5,"", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "t", false, false, 0, true,Type.Missing,Type.Missing);
            //wb = wbs.Open(FileName,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
            mFilename = FileName;
        }
        public Microsoft.Office.Interop.Excel.Worksheet GetSheet(string SheetName)
        //获取一个工作表
        {
            int flag = 0;
            Microsoft.Office.Interop.Excel.Worksheet s;
            for (int i = 1; i <= wb.Sheets.Count; i++)
            {
                if (string.Compare(wb.Sheets[i].Name, SheetName) == 0)
                {
                    flag = 1;
                }
            }
            if (flag == 1)
            {
                s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[SheetName];
            }
            else
            {
                //MessageBox.Show("请先创建:" + SheetName);
                //s= share.excelEdit.AddSheet(SheetName);
                s = null;
            }
            return s;
        }
        public Microsoft.Office.Interop.Excel.Worksheet AddSheet(string SheetName)
        //添加一个工作表,若已存在则返回已存在的工作表
        {

            int flag = 0;

            for (int i = 1; i <= wb.Sheets.Count; i++)
            {
                if (string.Compare(wb.Sheets[i].Name, SheetName) == 0)
                {
                    flag = 1;
                }
            }
            if (flag == 1)
            {
                //MessageBox.Show("已经存在"+SheetName+"工作表，无需新建");
            }
            else
            {
                Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                s.Name = SheetName;
                return s;
            }
            return wb.Sheets[SheetName];

        }

        public void DelSheet(string SheetName)//删除一个工作表
        {
            ((Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[SheetName]).Delete();
        }
        public Microsoft.Office.Interop.Excel.Worksheet ReNameSheet(string OldSheetName, string NewSheetName)//重命名一个工作表一
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[OldSheetName];
            s.Name = NewSheetName;
            return s;
        }

        public Microsoft.Office.Interop.Excel.Worksheet ReNameSheet(Microsoft.Office.Interop.Excel.Worksheet Sheet, string NewSheetName)//重命名一个工作表二
        {

            Sheet.Name = NewSheetName;

            return Sheet;
        }

        public void SetCellValue(Microsoft.Office.Interop.Excel.Worksheet ws, int x, int y, object value)
        //ws：要设值的工作表     X行Y列     value   值
        {
            ws.Cells[x, y] = value;
        }
        public void SetCellValue(string ws, int x, int y, object value)
        //ws：要设值的工作表的名称 X行Y列 value 值
        {

            GetSheet(ws).Cells[x, y] = value;
        }

        public void SetCellProperty(Microsoft.Office.Interop.Excel.Worksheet ws, int Startx, int Starty, int Endx, int Endy, int size, string name, Microsoft.Office.Interop.Excel.Constants color, Microsoft.Office.Interop.Excel.Constants HorizontalAlignment)
        //设置一个单元格的属性   字体，   大小，颜色   ，对齐方式
        {
            name = "宋体";
            size = 12;
            color = Microsoft.Office.Interop.Excel.Constants.xlAutomatic;
            HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = name;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }

        public void SetCellProperty(string wsn, int Startx, int Starty, int Endx, int Endy, int size, string name, Microsoft.Office.Interop.Excel.Constants color, Microsoft.Office.Interop.Excel.Constants HorizontalAlignment)
        {
            //name = "宋体";
            //size = 12;
            //color = Microsoft.Office.Interop.Excel.Constants.xlAutomatic;
            //HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;

            Microsoft.Office.Interop.Excel.Worksheet ws = GetSheet(wsn);
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = name;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;

            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }


        public void UniteCells(Microsoft.Office.Interop.Excel.Worksheet ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            ws.Range[ws.Cells[x1, y1], ws.Cells[x2, y2]].Merge(Type.Missing);
        }

        public void UniteCells(string ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            GetSheet(ws).Range[GetSheet(ws).Cells[x1, y1], GetSheet(ws).Cells[x2, y2]].Merge(Type.Missing);

        }


        public void InsertTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格插入到Microsoft.Office.Interop.Excel指定工作表的指定位置 为在使用模板时控制格式时使用一
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    GetSheet(ws).Cells[startX + i, j + startY] = dt.Rows[i][j].ToString();

                }

            }

        }
        public void InsertTable(System.Data.DataTable dt, Microsoft.Office.Interop.Excel.Worksheet ws, int startX, int startY)
        //将内存中数据表格插入到Microsoft.Office.Interop.Excel指定工作表的指定位置二
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    ws.Cells[startX + i, j + startY] = dt.Rows[i][j];

                }

            }

        }


        public void AddTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格添加到Microsoft.Office.Interop.Excel指定工作表的指定位置一
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    GetSheet(ws).Cells[i + startX, j + startY] = dt.Rows[i][j];

                }

            }

        }
        public void AddTable(System.Data.DataTable dt, Microsoft.Office.Interop.Excel.Worksheet ws, int startX, int startY)
        //将内存中数据表格添加到Microsoft.Office.Interop.Excel指定工作表的指定位置二
        {


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    ws.Cells[i + startX, j + startY] = dt.Rows[i][j];

                }
            }

        }
        /*
        public void InsertPictures(string Filename, string ws)
        //插入图片操作一
        {
            GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
            //后面的数字表示位置
        }
        */
        //public void InsertPictures(string Filename, string ws, int Height, int Width)
        //插入图片操作二
        //{
        //    GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Height = Height;
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Width = Width;
        //}
        //public void InsertPictures(string Filename, string ws, int left, int top, int Height, int Width)
        //插入图片操作三
        //{

        //    GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).IncrementLeft(left);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).IncrementTop(top);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Height = Height;
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Width = Width;
        //}

        public void InsertActiveChart(Microsoft.Office.Interop.Excel.XlChartType ChartType, string ws, int DataSourcesX1, int DataSourcesY1, int DataSourcesX2, int DataSourcesY2, Microsoft.Office.Interop.Excel.XlRowCol ChartDataType)
        //插入图表操作
        {
            ChartDataType = Microsoft.Office.Interop.Excel.XlRowCol.xlColumns;
            wb.Charts.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            {
                wb.ActiveChart.ChartType = ChartType;
                wb.ActiveChart.SetSourceData(GetSheet(ws).Range[GetSheet(ws).Cells[DataSourcesX1, DataSourcesY1], GetSheet(ws).Cells[DataSourcesX2, DataSourcesY2]], ChartDataType);
                wb.ActiveChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsObject, ws);

            }
        }
        public void InsertActiveChart(Microsoft.Office.Interop.Excel.XlChartType ChartType, string ws, Excel.Range data, Microsoft.Office.Interop.Excel.XlRowCol ChartDataType)
        //插入图表操作
        {
            ChartDataType = Microsoft.Office.Interop.Excel.XlRowCol.xlColumns;
            wb.Charts.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            {
                wb.ActiveChart.ChartType = ChartType;
                wb.ActiveChart.SetSourceData(data, ChartDataType);
                wb.ActiveChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsObject, ws);

            }
        }
        public bool Save()
        //保存文档
        {
            if (mFilename == "")
            {
                return false;
            }
            else
            {
                try
                {
                    wb.Save();
                    return true;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return false;
                }
            }
        }
        public bool SaveAs(object FileName)
        //文档另存为
        {
            try
            {
                wb.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;

            }

            catch (Exception ex)
            {
                return false;

            }
        }
        public void Close()
        //关闭一个Microsoft.Office.Interop.Excel对象，销毁对象
        {
            //wb.Save();
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }


        /// <summary>
        /// 创建图表
        /// </summary>
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
        public void CreateChart(Microsoft.Office.Interop.Excel._Workbook m_Book, Microsoft.Office.Interop.Excel._Worksheet m_Sheet, int CharTop, int CharLeft, float Width, float Height, string Title, Excel.Range range, object CategoryLabels, object SeriesLabels, double MinimumScale, double MaximumScale, string CharName)
        {
            Microsoft.Office.Interop.Excel.Range oResizeRange;
            Microsoft.Office.Interop.Excel.Series oSeries;
            m_Book.Charts.Add(Type.Missing, Type.Missing, 1, Type.Missing);
            m_Book.ActiveChart.ChartWizard(range, Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered, Type.Missing, Microsoft.Office.Interop.Excel.XlRowCol.xlColumns, CategoryLabels, SeriesLabels, true, Title, "各市", "百分比(%)", Type.Missing);
            //以下是给图表放在指定位置
            m_Book.ActiveChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsObject, m_Sheet.Name);
            oResizeRange = (Microsoft.Office.Interop.Excel.Range)m_Sheet.Rows.get_Item(CharTop, Type.Missing);
            m_Sheet.Shapes.Item(CharName).Top = (float)(double)oResizeRange.Top;  //调图表的位置上边距
            oResizeRange = (Microsoft.Office.Interop.Excel.Range)m_Sheet.Columns.get_Item(CharLeft, Type.Missing);
            m_Sheet.Shapes.Item(CharName).Left = (float)(double)oResizeRange.Left;//调图表的位置左边距
            m_Sheet.Shapes.Item(CharName).Width = Width;   //调图表的宽度
            m_Sheet.Shapes.Item(CharName).Height = Height;  //调图表的高度
            m_Book.ActiveChart._ApplyDataLabels();//数据标签
            m_Book.ActiveChart.PlotArea.Interior.ColorIndex = 19;  //设置绘图区的背景色
            m_Book.ActiveChart.PlotArea.Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;//设置绘图区边框线条
            m_Book.ActiveChart.ChartArea.Border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;//设置边框线条
            m_Book.ActiveChart.HasDataTable = false;
            //设置Legend图例的位置和格式
            m_Book.ActiveChart.Legend.Interior.ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone;
            m_Book.ActiveChart.Legend.Font.Name = "宋体";
            //设置X轴的显示
            Microsoft.Office.Interop.Excel.Axis xAxis = (Microsoft.Office.Interop.Excel.Axis)m_Book.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary);
            xAxis.MajorGridlines.Border.ColorIndex = 1;//gridLine横向线条的颜色
            xAxis.HasTitle = true;
            xAxis.MinimumScale = MinimumScale;
            xAxis.MaximumScale = MaximumScale;
            xAxis.TickLabels.Font.Name = "宋体";
            xAxis.TickLabels.Font.Size = 8;
            //设置Y轴的显示
            Microsoft.Office.Interop.Excel.Axis yAxis = (Microsoft.Office.Interop.Excel.Axis)m_Book.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary);
            yAxis.TickLabels.Orientation = Microsoft.Office.Interop.Excel.XlTickLabelOrientation.xlTickLabelOrientationHorizontal;//Y轴显示的方向,是水平还是垂直等
            yAxis.TickLabels.Font.Size = 8;
            yAxis.TickLabels.Font.Name = "宋体";
            oSeries = (Microsoft.Office.Interop.Excel.Series)m_Book.ActiveChart.SeriesCollection(1);
            oSeries.Border.ColorIndex = 45;
        }

        public void CreateRadarChart(Microsoft.Office.Interop.Excel._Workbook m_Book, Microsoft.Office.Interop.Excel._Worksheet m_Sheet, int CharTop, int CharLeft, float Width, float Height, string Title, Excel.Range range,string CharName , double Chart_index , int Chart_LocationToLeft,string subtitleText)
        {

            Microsoft.Office.Interop.Excel.Range oResizeRange;

            Microsoft.Office.Interop.Excel.Series oSeries;

            m_Book.Charts.Add(Type.Missing, Type.Missing, 1, Type.Missing);


            m_Book.ActiveChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlRadarFilled;
            //设置图形
            //m_Book.ActiveChart.ChartStyle = 253;

            //设置数据取值范围
            m_Book.ActiveChart.SetSourceData(range, Microsoft.Office.Interop.Excel.XlRowCol.xlRows);
            //以下是给图表放在指定位置
            m_Book.ActiveChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAutomatic, m_Sheet.Name);
            oResizeRange = (Microsoft.Office.Interop.Excel.Range)m_Sheet.Rows.get_Item(10, Type.Missing);
            m_Sheet.Shapes.Item("Chart 1").Name = CharName;
            //调图表的位置上边距
            m_Sheet.Shapes.Item(CharName).Top = (float)(Chart_index) * 200;
            //调图表的位置左边距
            oResizeRange = (Microsoft.Office.Interop.Excel.Range)m_Sheet.Columns.get_Item(Chart_LocationToLeft, Type.Missing);  //Chart_LocationToLeft 为1~65535
            m_Sheet.Shapes.Item(CharName).Left = (float)(double)oResizeRange.Left;
            m_Sheet.Shapes.Item(CharName).Width = 288;   //调图表的宽度
            m_Sheet.Shapes.Item(CharName).Height = 200;  //调图表的高度
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

            var subframe = m_Book.ActiveChart.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 500, 400, 100, 100).TextFrame2;
            //(Excel.Series)m_Book.ActiveChart.SeriesCollection(1).DataLabels(1) as Excel.DataLabel).Text ="a";
            // Add title
            var subtitle = subframe.TextRange;
            //subtitle.Text = "六级 是\n四级 是";
            subtitle.Text = subtitleText;
            subtitle.Font.NameFarEast = "微软雅黑";
            subtitle.Font.Size = 12;
            

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
            oSeries.Name = CharName;
            oSeries.Border.ColorIndex = 45;
            oSeries.Border.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            //oSeries.Values = 1;
            m_Book.ActiveChart.ChartTitle.Text = Title;
            //m_Book.ActiveChart.SaveAs("C:\\Users\\Public\\Pictures\\" + Title +".jpg");


            m_Book.ActiveChart.Axes(Excel.XlAxisType.xlValue).MaximumScaleIsAuto = false;
            m_Book.ActiveChart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 100;
            m_Book.ActiveChart.Axes(Excel.XlAxisType.xlValue).MinimumScaleIsAuto = false;
            m_Book.ActiveChart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 50;
            m_Book.ActiveChart.Axes(Excel.XlAxisType.xlValue).MajorUnitIsAuto = false;
            m_Book.ActiveChart.Axes(Excel.XlAxisType.xlValue).MajorUnit = 10;

        }


    }
}

