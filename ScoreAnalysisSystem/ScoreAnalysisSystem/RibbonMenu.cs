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
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger("RibbonMenu");
        private IRibbonUI _ribbonUi;

        public void RibbonMenu_Load(IRibbonUI ribbonUi)
        {
            Log.Debug("初始化成功");
            _ribbonUi = ribbonUi;
        }

        public void btnAnalysis_Click(IRibbonControl control)
        {
            //_ribbonUi.InvalidateControl("btnAnalysis");
            _ribbonUi.Invalidate(); //刷新显示
        }

        private string _analysisLabel = "分析成绩";

        public string GetAnalysisLabel(IRibbonControl control)
        {
            return _analysisLabel;
        }

        /// <summary>
        /// 刷新全部控件
        /// </summary>
        private void RefreshControl()
        {
            _ribbonUi.Invalidate();
        }

        /// <summary>
        /// 刷新指定Id的控件
        /// </summary>
        private void RefreshControl(string controlId)
        {
            _ribbonUi.InvalidateControl(controlId);
        }
    }
}