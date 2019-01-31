using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            //添加
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
    }
}
