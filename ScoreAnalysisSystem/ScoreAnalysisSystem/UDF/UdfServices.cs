using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace ScoreAnalysisSystem.UDF
{
    public class UdfServices
    {
        [ExcelFunction(Name = "CalculatedGrade", Description = "课程绩点", Category = "成绩分析系统")]
        public static object CalculatedGrade(
            [ExcelArgument(Name = "courseCredit", Description = "课程学分")]string courseCredit,
            [ExcelArgument(Name = "courseScore", Description = "课程成绩")]string courseScore)
        {
            if (string.IsNullOrEmpty(courseCredit) || !float.TryParse(courseCredit, out var credit))
                return 0;
            if (string.IsNullOrEmpty(courseScore) || !float.TryParse(courseScore, out var score))
                return 0;

            score = (score / 10 - 5) * credit;
            if (score < 0)
                score = 0;
            return score;
        }
    }
}
