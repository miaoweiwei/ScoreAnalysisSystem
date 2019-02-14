using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ScoreAnalysisSystem.View
{
    public partial class CourseCreditUserControl : UserControl
    {
        private string[] courseNames;
        public CourseCreditUserControl(string[]courseNames)
        {
            this.courseNames = courseNames;
            InitializeComponent();
            foreach (string s in courseNames)
            {
                DataGridViewRow row=new DataGridViewRow();
                DataGridViewTextBoxCell nameBoxCell = new DataGridViewTextBoxCell();
                DataGridViewTextBoxCell courseCreditBoxCell = new DataGridViewTextBoxCell();
                nameBoxCell.Value = s;
                row.Cells.Add(nameBoxCell);
                row.Cells.Add(courseCreditBoxCell);
                this.courseDgv.Rows.Add(row);
            }
        }

        public Dictionary<string, float> GetCourseCredit()
        {
            Dictionary<string,float>courseCreditDic=new Dictionary<string, float>();



            return courseCreditDic;
        }
    }
}
