using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelSubjectAddIn
{
    class share
    {
        public static Excel.Application ExcelApp;
        public static Office.CommandBar cmb;
        public static Microsoft.Office.Tools.ActionsPane acp;
        public static ExcelEdit excelEdit;
        public static dataAnalysis dataAnalysis;
        public static rendering_diagram rendering_diagram;
        public static Form1 myForm;
        public static UserControl1 myUserControl_Lesson;
        public static UserControl2 myUserControl_individual;
        public static Excel.Workbook CurrentWorkbook;
        public static Excel.Worksheet ClassSheet;
        public static Excel.Worksheet IndividualSheet;
        public static Excel.Worksheet LessonSheet;
        public static Excel.Worksheet importWorkSheet;
        public static Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane_Lesson;
        public static Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane_individual;
        public static int subject_num;
        public static int student_num;
        public static int importMenu_row;
        public static int classMenu_row;
        public static int individualMenu_row;
        public static int lessonMenu_row;

    }
}
