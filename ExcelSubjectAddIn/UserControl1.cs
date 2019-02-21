using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelSubjectAddIn
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        public void addCheckItem(string ItemsName)
        {
            checkedListBox_Lesson.Items.Add(ItemsName);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //图表清零
            int shapes_count = share.LessonSheet.Shapes.Count;
            for (int i = 0; i < shapes_count; i++)
            {
                share.LessonSheet.Shapes.Item(1).Delete();
            }
            //添加饼图
            int Chart_index = -1;
            for (int i=0; i < checkedListBox_Lesson.Items.Count; i++)
            {
                if (checkedListBox_Lesson.GetItemChecked(i) == true)
                {
                    Chart_index += 1;    
                    share.rendering_diagram.addChart_LessonSheet(share.LessonSheet, checkedListBox_Lesson.GetItemText(checkedListBox_Lesson.Items[i]), Chart_index);
                }
            }
        }
        public void cleanListBox()
        {
            for(int i=0; i< checkedListBox_Lesson.Items.Count;)
            {
                checkedListBox_Lesson.Items.RemoveAt(0);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for(int i=0; i< checkedListBox_Lesson.Items.Count; i++)
            {
                checkedListBox_Lesson.SetItemChecked(i, true);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox_Lesson.Items.Count; i++)
            {
                checkedListBox_Lesson.SetItemChecked(i, false);
            }
 

        }

        private void checkedListBox_Lesson_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //导出饼图
        private void button4_Click(object sender, EventArgs e)
        {

            share.excelEdit.wb = share.ExcelApp.ActiveWorkbook; //指定工作薄
            Excel.Worksheet PieCollectSheet = null;
            if (null != share.excelEdit.GetSheet("课程情况饼图汇总表"))
            {
                PieCollectSheet = share.excelEdit.GetSheet("课程情况饼图汇总表");
            }
            else
            {
                share.excelEdit.wb = share.ExcelApp.ActiveWorkbook;
                PieCollectSheet = share.excelEdit.AddSheet("课程情况饼图汇总表");
            }
            //PieCollectSheet图表清零
            int shapes_count = PieCollectSheet.Shapes.Count;
            for (int i = 0; i < shapes_count; i++)
            {
                PieCollectSheet.Shapes.Item(1).Delete();
            }
            //根据lessonSheet数据 导出饼图到 PieCollectSheet
            int Chart_index = -1;
            int Chart_indexToLeft = 0;
            int Chart_LocationToLeft = 0;
            int Chart_LocationToTop = 0;
            int Chart_Width = 6; //一个雷达图6个单元格宽度
            int NumofChartinaLine = 3;
            for (int i = 0; i < checkedListBox_Lesson.Items.Count; i++)
            {
                if (checkedListBox_Lesson.GetItemChecked(i) == true)
                {
                    Chart_index += 1;
                    Chart_indexToLeft += 1;
                    if (Chart_indexToLeft == 4)
                    {
                        Chart_indexToLeft = 1;
                    }
                    Chart_LocationToTop = Chart_index / NumofChartinaLine;
                    Chart_LocationToLeft = (Chart_indexToLeft - 1) * Chart_Width + 1;

                    share.rendering_diagram.addChart_PieCollectSheet(share.LessonSheet, PieCollectSheet, checkedListBox_Lesson.GetItemText(checkedListBox_Lesson.Items[i]), Chart_LocationToTop, Chart_LocationToLeft);
                }
            }
        }
        public void AllButtonDisable()
        {
            foreach (Control control in this.Controls)
            {
                //遍历所有Button...
                if (control is Button)
                {
                    Button t = (Button)control;
                    t.Enabled = false;
                }
            }

        }
        public void AllButtonEnable()
        {
            foreach (Control control in this.Controls)
            {
                //遍历所有Button...
                if (control is Button)
                {
                    Button t = (Button)control;
                    t.Enabled = true;
                }
            }
        }
        private void UserControl1_Load(object sender, EventArgs e)
        {

        }
    }
}
