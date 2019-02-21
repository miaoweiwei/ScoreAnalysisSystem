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
    public partial class UserControl2 : UserControl
    {
        public UserControl2()
        {
            InitializeComponent();
        }
        public void addCheckItem(string ItemsName)
        {
            checkedListBox_individual.Items.Add(ItemsName);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox_individual.Items.Count; i++)
            {
                checkedListBox_individual.SetItemChecked(i, true);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox_individual.Items.Count; i++)
            {
                checkedListBox_individual.SetItemChecked(i, false);
            }
        }
        public void cleanListBox()
        {
            for (int i = 0; checkedListBox_individual.Items.Count >i;)
            {
                checkedListBox_individual.Items.RemoveAt(0);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //图表清零
            int shapes_count = share.IndividualSheet.Shapes.Count;
            for (int i = 0; i < shapes_count; i++)
            {
                share.IndividualSheet.Shapes.Item(1).Delete();
            }
            //添加
            int Chart_index = -1;
            string studentNumber = "";
            for (int i = 0; i < checkedListBox_individual.Items.Count; i++)
            {
                if (checkedListBox_individual.GetItemChecked(i) == true)
                {
                    Chart_index += 1;
                    studentNumber = Convert.ToString(share.IndividualSheet.Cells[share.individualMenu_row + 2+i,1].value);
                    share.rendering_diagram.addChart_IndividualSheet(share.IndividualSheet, studentNumber, Chart_index);
                }
            }
        }
        //导出雷达图
        private void button4_Click(object sender, EventArgs e)
        {

            share.excelEdit.wb = share.ExcelApp.ActiveWorkbook; //指定工作薄
            Excel.Worksheet RadarCollectSheet = null;
            if (null != share.excelEdit.GetSheet("学生能力雷达图汇总表"))
            {
                RadarCollectSheet = share.excelEdit.GetSheet("学生能力雷达图汇总表");
            }
            else
            {
                share.excelEdit.wb = share.ExcelApp.ActiveWorkbook;
                RadarCollectSheet = share.excelEdit.AddSheet("学生能力雷达图汇总表");
            }
            //RadarCollectSheet 图表清零
            int shapes_count = RadarCollectSheet.Shapes.Count;
            for (int i = 0; i < shapes_count; i++)
            {
                RadarCollectSheet.Shapes.Item(1).Delete();
            }
            //根据IndividualSheet数据 导出雷达图到 RadarCollectSheet
            int Chart_index = -1;
            int Chart_indexToLeft = 0;
            int Chart_LocationToLeft = 0;
            int Chart_LocationToTop = 0;
            int Chart_Width = 6; //一个雷达图6个单元格宽度
            int NumofChartinaLine = 3;
            string studentNumber = "";
            for (int i = 0; i < checkedListBox_individual.Items.Count; i++)
            {
                if (checkedListBox_individual.GetItemChecked(i) == true)
                {
                    Chart_index += 1;
                    Chart_indexToLeft += 1;
                    if (Chart_indexToLeft == NumofChartinaLine + 1)
                    {
                        Chart_indexToLeft = 1;
                    }
                    studentNumber = Convert.ToString(share.IndividualSheet.Cells[share.individualMenu_row + 2 + i, 1].value);
                    Chart_LocationToTop = Chart_index / NumofChartinaLine;
                    Chart_LocationToLeft = (Chart_indexToLeft - 1) * Chart_Width + 1;
                    share.rendering_diagram.addChart_RadarCollectSheet(share.IndividualSheet, RadarCollectSheet, studentNumber, Chart_LocationToTop, Chart_LocationToLeft);
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
        private void UserControl2_Load(object sender, EventArgs e)
        {

        }
    }
}
