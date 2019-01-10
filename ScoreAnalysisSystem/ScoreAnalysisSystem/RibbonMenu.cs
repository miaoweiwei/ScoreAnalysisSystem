using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration.CustomUI;

namespace ScoreAnalysisSystem
{
    [ComVisible(true)]
    public class RibbonMenu : ExcelRibbon
    {
        public void RibbonMenu_Load(IRibbonUI ribbonUi)
        {

        }

        public void btnLogin_Click(IRibbonControl control)
        {

        }

        private string _loginLable = "登录";
        public string GetLoginLabel(IRibbonControl control)
        {
            return _loginLable;
        }
    }
}
