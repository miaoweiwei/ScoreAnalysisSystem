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
    }
}
