using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
namespace ExcelSubjectAddIn
{
    class dataAnalysis
    {
        public string mFilename;
        public Microsoft.Office.Interop.Excel.Application app;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;

        private int importMenu_row;
        private int classMenu_row;
        private int individualMenu_row;
        private int lessonMenu_row;
        
        public dataAnalysis()
        {
            importMenu_row = 2; //导入数据从第2行开始
            classMenu_row = 7;  //班级情况明细表从第7行开始
            lessonMenu_row = 2; //课程情况明细从第2行开始
            individualMenu_row = 2;
            share.individualMenu_row = individualMenu_row;
            share.importMenu_row = importMenu_row;
            share.classMenu_row = classMenu_row;
            share.lessonMenu_row = lessonMenu_row;
        }
        public void analyClassStudyStatus(Excel.Worksheet importWorkSheet, Excel.Worksheet ClassSheet)
        {
            ClassSheet.Cells[1, 1] = "班级学习情况";
            share.excelEdit.UniteCells(ClassSheet, 1, 1, 1, share.subject_num -3 +8);    //合并单元格
            //统计科目数量
            int count = 0;
            while (importWorkSheet.Cells[2, 3 + count].value != null)
            {
                count++;
            }
            share.subject_num = count;
            //统计学生人数
            count = 0;
            while (importWorkSheet.Cells[3 +count , 1].value != null)
            {
                count++;
            }
            share.student_num = count;
            //填写项目
                      
            ClassSheet.Cells[classMenu_row, 1].value = "学号";
            ClassSheet.Cells[classMenu_row, 2].value = "姓名";
            ClassSheet.Cells[classMenu_row, 3].value = "绩点";
            ClassSheet.Cells[classMenu_row, 4].value = "不及格科目数";
            ClassSheet.Cells[classMenu_row, 5].value = "总分";
            ClassSheet.Cells[classMenu_row, 6].value = "平均分";
            for (int i=0;i<share.subject_num-1;i++) //去除最后绩点
            {
                ClassSheet.Cells[classMenu_row, 7 + i].value = importWorkSheet.Cells[importMenu_row, 3 + i].value ;
            }
            //填写班级情况明细表   
            for (int i= 1; i<= share.student_num; i++)    
            //外循环遍历importWorkSheet中学生所在行
            {
                //填写学号
                ClassSheet.Cells[classMenu_row + i, 1].value = importWorkSheet.Cells[importMenu_row + i, 1].value;
                //填写姓名
                ClassSheet.Cells[classMenu_row + i, 2].value = importWorkSheet.Cells[importMenu_row + i, 2].value;
                //填写绩点
                ClassSheet.Cells[classMenu_row + i, 3].value = importWorkSheet.Cells[importMenu_row + i, 2+share.subject_num].value;
                //填写该学生不及格科目数
                ClassSheet.Cells[classMenu_row + i, 4].value = noPass_num(importWorkSheet, importMenu_row + i);
                //填写该学生总分
                ClassSheet.Cells[classMenu_row + i, 5].value = totalScore(importWorkSheet, importMenu_row + i);
                //填写该学生平均分
                ClassSheet.Cells[classMenu_row + i, 6].value = ClassSheet.Cells[classMenu_row + i, 5].value / (share.subject_num -3);
                //填写该学生所有科目考试成绩
                for(int j=0; j< share.subject_num -1 ; j++) //去除绩点
                {
                    if (importWorkSheet.Cells[importMenu_row + i, 3 + j].value == null)
                    {
                        ClassSheet.Cells[classMenu_row + i, 7 + j].value = 0;
                    }
                    else if (importWorkSheet.Cells[importMenu_row + i, 3 + j].value.GetType() == typeof(string))
                    {
                        if (importWorkSheet.Cells[importMenu_row + i, 3 + j].value == "")
                        {
                            ClassSheet.Cells[classMenu_row + i, 7 + j].value = 0;
                        } else if (importWorkSheet.Cells[importMenu_row + i, 3 + j].value == "是" || importWorkSheet.Cells[importMenu_row + i, 3 + j].value == "否")
                        {
                            ClassSheet.Cells[classMenu_row + i, 7 + j].value = importWorkSheet.Cells[importMenu_row + i, 3 + j].value;
                        }
                        else
                        {
                            MessageBox.Show("data wrong");
                        }

                    }else
                    {
                        ClassSheet.Cells[classMenu_row + i, 7 + j].value = importWorkSheet.Cells[importMenu_row + i, 3 + j].value;
                    }                                                          
                }
            }
            
            //最后算
            ClassSheet.Cells[2, 1] = "班级平均分";
            ClassSheet.Cells[2, 2] = calClassAverageScore(ClassSheet,5);    //统计第5列的均值
            ClassSheet.Cells[3, 1] = "班级平均绩点";
            ClassSheet.Cells[3, 2] = calClassAverageJD(ClassSheet,3);   //统计第3列的均值
            ClassSheet.Cells[4, 1] = "不及格率";
            ClassSheet.Cells[4, 2] = calClassUnpassRate(ClassSheet,4);  //统计第4列的和，再除以总考试次数
            ClassSheet.Cells[5, 1] = "四级通过率";
            ClassSheet.Cells[5, 2] = calClassG4passRate(importWorkSheet);
            ClassSheet.Cells[6, 1] = "六级通过率";
            ClassSheet.Cells[6, 2] = calClassG6passRate(importWorkSheet);
        }
        //统计不及格科目数
        private double noPass_num(Excel.Worksheet importWorkSheet,int row)
        {
            int count = 0;
            for(int i=0; i<share.subject_num-1; i++)  //最后绩点不要
            {
                if(importWorkSheet.Cells[row, 3 + i].value == null )
                {
                    count++;
                }
                else if (importWorkSheet.Cells[row, 3 + i].value.GetType() == typeof(string))
                {
                    //if(importWorkSheet.Cells[row, 3 + i].value == "否")  count++;  //统计四六级
                    if (importWorkSheet.Cells[row, 3 + i].value == "")  count++;
                }
                else
                {
                    if (importWorkSheet.Cells[row, 3 + i].value < 60) count++;
                }

            }
            return count;
        }
        //统计学生总分
        private double totalScore(Excel.Worksheet importWorkSheet,int row)
        {
            int sum = 0;
            for (int i = 0; i < share.subject_num-1; i++) //最后绩点不要
            {
                if (importWorkSheet.Cells[row, 3 + i].value.GetType() == typeof(string))
                {
                    sum += 0;
                }
                else
                {
                    sum += importWorkSheet.Cells[row, 3 + i].value;
                }
            }
            return sum;
        }
        //统计班级均分
        private double calClassAverageScore(Excel.Worksheet ClassSheet, int column)
        {
            double sum = 0;
            for(int i=1; i<=share.student_num; i++)
            {
                sum += ClassSheet.Cells[classMenu_row + i, column].value;
            }
            double average = sum / share.student_num;
            return average;
        }
        //统计班级平均绩点
        private double calClassAverageJD(Excel.Worksheet ClassSheet ,int column)
        {
            double sum = 0;
            for (int i = 1; i <= share.student_num; i++)
            {
       
               sum += ClassSheet.Cells[classMenu_row + i, column].value;
            }
            double average = sum / share.student_num;
            return average;
        }
        //统计班级及格率，不包括四六级
        private double calClassUnpassRate(Excel.Worksheet ClassSheet, int column)
        {
            double sum = 0;
            for (int i = 1; i <= share.student_num; i++)
            {
                sum += ClassSheet.Cells[classMenu_row + i, column].value;
            }
            double totalTest_num = share.student_num * (share.subject_num - 3);  //不包括四六级和绩点
            double average = sum / totalTest_num;
            return average;
        }
        private double calClassG4passRate(Excel.Worksheet importWorkSheet)
        {
            double passCount = 0;
            for (int i = 1; i <= share.student_num; i++)
            {
                string Grade4 = importWorkSheet.Cells[classMenu_row + i, 2 + share.subject_num - 2].value;
                if (Grade4 == "是")
                {
                    passCount += 1;
                }
            }
            double rate = passCount / share.student_num;
            return rate;
        }
        private double calClassG6passRate(Excel.Worksheet importWorkSheet)
        {
            double passCount = 0;
            for (int i = 1; i <= share.student_num; i++)
            {
                string Grade6 = importWorkSheet.Cells[classMenu_row + i, 2 + share.subject_num-1].value;
                if (Grade6 == "是")
                {
                    passCount += 1;
                }
            }
            double rate = passCount / share.student_num;
            return rate;
        }
        public void analyIndividualStatus(Excel.Worksheet importWorkSheet, Excel.Worksheet IndividualSheet)
        {
            IndividualSheet.Cells[1, 1] = "2015级计算机科学与技术（师范）";
            share.excelEdit.UniteCells(IndividualSheet, 1, 1, 1, share.subject_num -3 +9);    //合并单元格
            IndividualSheet.Cells[individualMenu_row, 1] = "学号";
            share.excelEdit.UniteCells(IndividualSheet, individualMenu_row, 1, individualMenu_row + 1, 1);    //合并单元格
            IndividualSheet.Cells[individualMenu_row, 2] = "姓名";
            share.excelEdit.UniteCells(IndividualSheet, individualMenu_row, 2, individualMenu_row+1, 2);    //合并单元格
            IndividualSheet.Cells[individualMenu_row, 3] = "课程";
            share.excelEdit.UniteCells(IndividualSheet, individualMenu_row, 3, individualMenu_row, 2 + share.subject_num - 3);    //合并单元格
            for (int i = 0; i < share.subject_num - 3; i++) //除去绩点、四六级
            {
                //打印课程名
                IndividualSheet.Cells[individualMenu_row + 1,3+i].value = importWorkSheet.Cells[importMenu_row, 3 + i].value;
            }
            share.IndividualSheet.Cells[individualMenu_row, 2 + share.subject_num - 3 + 1] = "不及格科目数";
            share.IndividualSheet.Cells[individualMenu_row, 2 + share.subject_num - 3 + 2] = "平均分";
            share.IndividualSheet.Cells[individualMenu_row, 2 + share.subject_num - 3 + 3] = "平均分排名";
            share.IndividualSheet.Cells[individualMenu_row, 2 + share.subject_num - 3 + 4] = "绩点";
            share.IndividualSheet.Cells[individualMenu_row, 2 + share.subject_num - 3 + 5] = "绩点排名";
            share.IndividualSheet.Cells[individualMenu_row, 2 + share.subject_num - 3 + 6] = "四级";
            share.IndividualSheet.Cells[individualMenu_row, 2 + share.subject_num - 3 + 7] = "六级";
            for (int i=1; i<=7; i++)
            {
                share.excelEdit.UniteCells(IndividualSheet, 2, 2 + share.subject_num - 3 + i, 3, 2 + share.subject_num - 3 + i);    //合并单元格
            }

           
            int[] sort_index = sort(3);
            for (int i=0; i< share.student_num; i++)
            {
                //IndividualSheet添加学号
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 1].value = share.ClassSheet.Cells[classMenu_row + 1 + sort_index[i], 1];
                //IndividualSheet添加姓名
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 2].value = share.ClassSheet.Cells[classMenu_row +1 +sort_index[i] ,2];

                //会耦合
                //添加姓名复选框
                share.myUserControl_individual.addCheckItem(IndividualSheet.Cells[individualMenu_row + i + 2, 2].value);

                //IndividualSheet添加课程
                for (int j=0; j<share.subject_num -3;j++)
                {
                    share.IndividualSheet.Cells[individualMenu_row + 2 + i, 3 +j].value = share.ClassSheet.Cells[classMenu_row + 1 + sort_index[i], 7+j];
                }
                //IndividualSheet添加不及格科目数
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 3 + share.subject_num - 3 ].value = share.ClassSheet.Cells[classMenu_row + 1 + sort_index[i], 4];
                //平均分
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 4 + share.subject_num - 3].value = share.ClassSheet.Cells[classMenu_row + 1 + sort_index[i], 6];

                //平均分排名
                double individual_average = share.IndividualSheet.Cells[individualMenu_row + 2 + i, 4 + share.subject_num - 3].value;
                int individual_average_index = 1;
                for (int j=0; j<share.student_num;j++)
                {
                    if (share.ClassSheet.Cells[classMenu_row + 1 + j, 6].value > individual_average)
                    {
                        individual_average_index += 1;
                    }
                }
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 5 + share.subject_num - 3].value = individual_average_index;
                //绩点
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 6 + share.subject_num - 3].value = share.ClassSheet.Cells[classMenu_row + 1 + sort_index[i], 3];
                //绩点排名
                double individual_JD = share.IndividualSheet.Cells[individualMenu_row + 2 + i, 6 + share.subject_num - 3].value;
                int individual_JD_index = 1;
                for (int j = 0; j < share.student_num; j++)
                {
                    if (share.ClassSheet.Cells[classMenu_row + 1 + j, 3].value > individual_JD)
                    {
                        individual_JD_index += 1;
                    }
                }
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 7 + share.subject_num - 3].value = individual_JD_index;
                //四级
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 8 + share.subject_num - 3].value = share.ClassSheet.Cells[classMenu_row + 1 + sort_index[i], 7 + share.subject_num - 3].value;
                //六级
                share.IndividualSheet.Cells[individualMenu_row + 2 + i, 9 + share.subject_num - 3].value = share.ClassSheet.Cells[classMenu_row + 1 + sort_index[i], 8 + share.subject_num - 3].value;

            }

        }
        //返回绩点排序sort_index
        private int[] sort(int column)
        {
            //int column = 3;
            double[] record = new double[share.student_num];
            int[] sort_index = new int[share.student_num];
            for (int i = 0; i < share.student_num; i++)
            {
                sort_index[i] = i;
            }
            for (int i=0; i<share.student_num; i++)
            {
                record[i] = share.ClassSheet.Cells[classMenu_row + 1 + i, column].value;//绩点数据在第三列
            }
            double maxValue =0;
            double temp = 0;
            int temp2 = 0;
            int mark_j = 0;
            for(int i=share.student_num; i>0; i--)
            {
                mark_j = share.student_num - i;
                maxValue = record[share.student_num - i];
                
                for (int j = share.student_num - i +1; j < share.student_num; j++)
                {
                    if (record[j] > maxValue)
                    {
                        mark_j = j;
                        maxValue = record[j];
                    }
                }
                temp = record[share.student_num - i];
                record[share.student_num - i] = record[mark_j];
                record[mark_j] = temp;

                temp2 = sort_index[share.student_num - i];
                sort_index[share.student_num - i] = sort_index[mark_j];
                sort_index[mark_j] = temp2;
            }
            return sort_index;
        }

 
        public void analyLessonStatus(Excel.Worksheet importWorkSheet, Excel.Worksheet LessonSheet)
        {
            LessonSheet.Cells[1, 1] = "课程学习情况";
            share.excelEdit.UniteCells(LessonSheet, 1, 1, 1, 10);    //合并单元格
            //60分以下   60 - 69   70 - 79   80 - 89   90 - 100  不及格率 平均分 最高分 最低分
            double[] frequence = new double[5];   //五个成绩段
            LessonSheet.Cells[lessonMenu_row, 1] = "课程";
            LessonSheet.Cells[lessonMenu_row, 2] = "60分以下";            
            LessonSheet.Cells[lessonMenu_row, 3] = "60~69分";
            LessonSheet.Cells[lessonMenu_row, 4] = "70~79分";
            LessonSheet.Cells[lessonMenu_row, 5] = "80~89分";
            LessonSheet.Cells[lessonMenu_row, 6] = "90~100分";
            LessonSheet.Cells[lessonMenu_row, 7] = "不及格率";
            LessonSheet.Cells[lessonMenu_row, 8] = "平均分";
            LessonSheet.Cells[lessonMenu_row, 9] = "最高分";

            LessonSheet.Cells[lessonMenu_row, 10] = "最低分";


            for (int i=0; i<share.subject_num - 3; i++) //除去绩点、四六级
            {
                //打印课程名
                LessonSheet.Cells[lessonMenu_row +i +1, 1].value = importWorkSheet.Cells[importMenu_row, 3 + i];
                //添加复选框
                share.myUserControl_Lesson.addCheckItem(LessonSheet.Cells[lessonMenu_row + i + 1, 1].value);
                
                //计算成绩分布情况
                frequence = calFrequence(importWorkSheet, 3 + i);   //importWorkSheet第3列开始存储成绩
                //打印成绩分布情况
                LessonSheet.Cells[lessonMenu_row + i +1, 2].value = frequence[0];
                LessonSheet.Cells[lessonMenu_row + i +1, 3].value = frequence[1];
                LessonSheet.Cells[lessonMenu_row + i +1, 4].value = frequence[2];
                LessonSheet.Cells[lessonMenu_row + i +1, 5].value = frequence[3];
                LessonSheet.Cells[lessonMenu_row + i +1, 6].value = frequence[4];
                LessonSheet.Cells[lessonMenu_row + i + 1, 7].value = frequence[0] / share.student_num;
                LessonSheet.Cells[lessonMenu_row + i + 1, 8].value = calThisLessonClassAverage(importWorkSheet,3+i);
                LessonSheet.Cells[lessonMenu_row + i + 1, 9].value = calThisLessonClassHighestScore(importWorkSheet, 3 + i);
                LessonSheet.Cells[lessonMenu_row + i + 1, 10].value = calThisLessonClassLowestScore(importWorkSheet, 3 + i);
            }
            
        }
        //统计课程班级成绩分布
        private double[] calFrequence(Excel.Worksheet importWorkSeet, int column)
        {
            double[] frequence = new double[5] { 0, 0, 0, 0, 0 };
            for ( int i=1; i<=share.student_num; i++)
            {
                if(importWorkSeet.Cells[importMenu_row +i,column].value ==null) frequence[0] += 1;

                if (importWorkSeet.Cells[importMenu_row + i,column].value.GetType() == typeof(string))    
                 //防止出现字符
                {
                    if (importWorkSeet.Cells[importMenu_row + i, column].value == "")
                    {
                        frequence[0] += 1;
                    }
                    else
                    {
                        MessageBox.Show("出现非法字符：" + importWorkSeet.Cells[importMenu_row + i, column].value);
                    }
                }
                else
                //确保纯数值比较
                {
                    if (importWorkSeet.Cells[importMenu_row + i, column].value < 60)
                    {
                        frequence[0] += 1;
                    }else if(importWorkSeet.Cells[importMenu_row + i, column].value < 70)
                    {
                        frequence[1] += 1;
                    }
                    else if (importWorkSeet.Cells[importMenu_row + i, column].value < 80)
                    {
                        frequence[2] += 1;
                    }
                    else if (importWorkSeet.Cells[importMenu_row + i, column].value < 90)
                    {
                        frequence[3] += 1;
                    }
                    else if (importWorkSeet.Cells[importMenu_row + i, column].value <= 100)
                    {
                        frequence[4] += 1;
                    }
                }
            }
            return frequence;
        }
        //统计课程班级成绩平均分
        private double calThisLessonClassAverage(Excel.Worksheet importWorkSeet, int column)
        {
            double sum = 0;
            for (int i = 1; i <= share.student_num; i++)
            {
                if (importWorkSeet.Cells[importMenu_row + i, column].value.GetType() == typeof(string))
                {
                    sum = sum + 0;
                }
                else
                {
                    sum += importWorkSeet.Cells[importMenu_row + i, column].value;
                }
              
            }
            double average = sum / share.student_num;
            return average;
        }
        private double calThisLessonClassHighestScore(Excel.Worksheet importWorkSeet, int column)
        {
            double HighestScore = 0;
            double tempScore = 0;
            for (int i = 1; i <= share.student_num; i++)
            {
                if (importWorkSeet.Cells[importMenu_row + i, column].value.GetType() == typeof(string))
                {
                    tempScore = 0;
                }
                else
                {
                    tempScore = importWorkSeet.Cells[importMenu_row + i, column].value;
                }

                if (tempScore > HighestScore)
                {
                    HighestScore = tempScore;
                }
            }        
            return HighestScore;
        }
        private double calThisLessonClassLowestScore(Excel.Worksheet importWorkSeet, int column)
        {
            double LowestScore = 100;
            double tempScore = 0;
            for (int i = 1; i <= share.student_num; i++)
            {
                if (importWorkSeet.Cells[importMenu_row + i, column].value.GetType() == typeof(string))
                {
                    tempScore = 0;
                }
                else
                {
                    tempScore = importWorkSeet.Cells[importMenu_row + i, column].value;
                }
                if (tempScore < LowestScore)
                {
                    LowestScore = tempScore;
                }
            }
            return LowestScore;
        }
    }
}
