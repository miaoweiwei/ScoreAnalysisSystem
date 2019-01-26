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
        public static int subject_num;
        public static int student_num;
    }
}
