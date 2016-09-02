using Aspose.Words;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.IO;
using K12.Data;
using 個人學期成績單.UDT;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;


namespace 個人學期成績單
{

    //2016/7/15 穎驊製作測試，主要測試之前上傳的固定排名名單時做成報表時會不會有甚麼狀況。

    public class Program
    {
        [FISCA.MainMethod]

        //public static void Main() 的public 超級重要，沒有系統就load 不進去啦QQ
       
        public static void Main()
        {
            var btn = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["期末成績通知單(穎驊測試版)"];

            btn.Enable = true;
           
            //btn.Click += new EventHandler(Program_Click);

            btn.Click += delegate
            {
                PersonalSemesterScoreReportForm PSSR = new PersonalSemesterScoreReportForm();

                PSSR.ShowDialog();

            };

        }


        static void Program_Click(object sender_, EventArgs e_)
        {
          
            List<String> Error_List = new List<string>();

            String StudentID = "61943";

            int school_year = 104;

            int semester = 1;

            int rank_sequence = 4;
                       
                //foreach (DataRow dr in dt.Rows)
                //{

                //    retVal.Add(dr["uid"].ToString(), new Rank());

                //    retVal[dr["uid"].ToString()].SchoolYear = (int)int.Parse(dr["school_year"].ToString());

                //    retVal[dr["UID"].ToString()].Semester = (int)int.Parse(dr["semester"].ToString());

                //    retVal[dr["UID"].ToString()].GradeYear = (int)int.Parse(dr["grade_year"].ToString());

                //    retVal[dr["UID"].ToString()].RankType = (string)dr["rank_type"];

                //    retVal[dr["UID"].ToString()].RankSequence = (int)int.Parse(dr["rank_sequence"].ToString());

                //    retVal[dr["UID"].ToString()].DisplayName = (string)dr["display_name"];

                //    retVal[dr["UID"].ToString()].Memo = (string)dr["memo"];

                //    retVal[dr["UID"].ToString()].CreateTime = (DateTime)DateTime.Parse(dr["create_time"].ToString());

                //    retVal[dr["UID"].ToString()].Active = (bool)bool.Parse(dr["active"].ToString());

                //}

            System.ComponentModel.BackgroundWorker BGW = new BackgroundWorker();

            BGW.WorkerReportsProgress = true;
            
            BGW.DoWork += delegate(object sender, DoWorkEventArgs e)
                {
                    decimal progress = 0m;

                    BGW.ReportProgress(5);

                    Aspose.Words.Document Template;
                    Template = new Aspose.Words.Document(new MemoryStream(Properties.Resources.個人學期成績單樣板));
                  
                    // 取得選取學生
                    List<K12.Data.StudentRecord> StudentList = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);

                    DataTable table = new DataTable();

                    table.Columns.Add("學校名稱");
                    table.Columns.Add("學年度");
                    table.Columns.Add("學期");
                    table.Columns.Add("班級");
                    table.Columns.Add("姓名");
                    table.Columns.Add("定期評量");

                    table.Columns.Add("類別排名1");
                    table.Columns.Add("類別排名2");

                    table.Columns.Add("總分");
                    table.Columns.Add("平均");
                    table.Columns.Add("加權總分");
                    table.Columns.Add("加權平均");

                    table.Columns.Add("類別1總分");
                    table.Columns.Add("類別1平均");
                    table.Columns.Add("類別1加權總分");
                    table.Columns.Add("類別1加權平均");

                    table.Columns.Add("類別2總分");
                    table.Columns.Add("類別2平均");
                    table.Columns.Add("類別2加權總分");
                    table.Columns.Add("類別2加權平均");

                    table.Columns.Add("總分班排名");
                    table.Columns.Add("總分班排名母數");

                    table.Columns.Add("平均班排名");
                    table.Columns.Add("平均班排名母數");

                    table.Columns.Add("加權總分班排名");
                    table.Columns.Add("加權總分班排名母數");

                    table.Columns.Add("加權平均班排名");
                    table.Columns.Add("加權平均班排名母數");


                    table.Columns.Add("類別1總分排名");
                    table.Columns.Add("類別1總分排名母數");

                    table.Columns.Add("類別1平均排名");
                    table.Columns.Add("類別1平均排名母數");

                    table.Columns.Add("類別1加權總分排名");
                    table.Columns.Add("類別1加權總分排名母數");

                    table.Columns.Add("類別1加權平均排名");
                    table.Columns.Add("類別1加權平均排名母數");

                    table.Columns.Add("類別2總分排名");
                    table.Columns.Add("類別2總分排名母數");

                    table.Columns.Add("類別2平均排名");
                    table.Columns.Add("類別2平均排名母數");

                    table.Columns.Add("類別2加權總分排名");
                    table.Columns.Add("類別2加權總分排名母數");

                    table.Columns.Add("類別2加權平均排名");
                    table.Columns.Add("類別2加權平均排名母數");

                    foreach (var stuRec in StudentList)
                    {                        
                        StudentID = stuRec.ID;

                        DataRow row = table.NewRow();

                        FISCA.Data.QueryHelper qh = new FISCA.Data.QueryHelper();

                        //string strSQL = "SELECT * FROM $ischool.sunflower.rank";


                        //  襪靠 !! 學到一招啦!!  String 前面加@ 會自動把跨行的進去，省掉很多麻煩阿
                        string strSQL = @"
                SELECT $ischool.sunflower.rank.school_year, $ischool.sunflower.rank.semester,$ischool.sunflower.rank.rank_sequence ,group_hash_key,subject_type,subject_name,group_type,group_name,member_count, $ischool.sunflower.rankstudent.name as stu_name, $ischool.sunflower.rankstudent.class_name as stu_class_name, $ischool.sunflower.rankstudent.rank_include,$ischool.sunflower.rankdetail.* 
                FROM $ischool.sunflower.rank
                LEFT OUTER JOIN $ischool.sunflower.rankgroup ON $ischool.sunflower.rankgroup.ref_rank_id = $ischool.sunflower.rank.uid
                LEFT OUTER JOIN $ischool.sunflower.rankdetail ON $ischool.sunflower.rankgroup.uid = $ischool.sunflower.rankdetail.ref_rank_group_id
                LEFT OUTER JOIN $ischool.sunflower.rankstudent ON $ischool.sunflower.rankstudent.ref_rank_id = $ischool.sunflower.rank.uid AND $ischool.sunflower.rankstudent.ref_student_id = $ischool.sunflower.rankdetail.ref_student_id
                WHERE                 
                $ischool.sunflower.rank.active = true
                AND school_year =" + school_year +
                        @"
                AND semester =" + semester +
                        @"
                AND rank_sequence=" + rank_sequence +
                        @"                
                AND $ischool.sunflower.rankdetail.ref_student_id=" + StudentID +
                        @"
                AND (subject_type='科目' OR subject_type='總計')";

                        //AND $ischool.sunflower.rankstudent.class_name ='2017LIPa'

                        //2016/7/18 穎驊發現一件很重要的事，拿下來的DataTable dt，如果後續程式碼沒有用的話(因為是第一次寫邊寫邊檢查，後面要用的部分先註解所致)，系統會自動把它給掃除，
                        //造成DataTable dt 為null 的狀況，解決方法就是寫一些短程式先去用他，感謝恩正大大的幫忙ㄎㄎ

                        System.Data.DataTable dt = qh.Select(strSQL);


                        // 先前恩正所寫，為了讓dt不被回收而意思意思使用一下的CODE，後來已經完全不需要了，但此例太經典了，故留下註解以資紀念。
                        //var dt2 = qh.Select(strSQL);
                        //if (dt.Rows.Count != dt2.Rows.Count)
                        //    MsgBox.Show("Test");
                        //else
                        //    MsgBox.Show("Test2");


                        if (dt.Rows.Count != 0)
                        {

                            row["學校名稱"] = K12.Data.School.ChineseName;
                            row["學年度"] = dt.Rows[0]["school_year"];
                            row["學期"] = dt.Rows[0]["semester"];
                            row["班級"] = dt.Rows[0]["stu_class_name"];
                            row["姓名"] = dt.Rows[0]["stu_name"];
                            row["定期評量"] = "第" + dt.Rows[0]["rank_sequence"] + "次評量考試";
                                           
                            //科目
                            string col_subject = "";

                            //學分數
                            string col_credit = "";

                            //科目成績
                            string col_score = "";

                            //班排名
                            string col_rank = "";

                            //班排名母數
                            string col_rank_member_count = "";

                            //類別1排名
                            string col_rank_Cat1 = "";

                            //類別1排名母數
                            string col_rank_Cat1_member_count = "";


                            //類別2排名
                            string col_rank_Cat2 = "";

                            //類別2排名母數
                            string col_rank_Cat2_member_count = "";

                            int Rows_Counter = 1;

                            foreach (var dr in dt.Select("subject_type='科目' AND group_type= '班排名'"))
                            {

                                col_subject = string.Format("科目名稱{0}", Rows_Counter);

                                col_credit = string.Format("學分數{0}", Rows_Counter);

                                col_score = string.Format("科目成績{0}", Rows_Counter);

                                col_rank = string.Format("班排名{0}", Rows_Counter);

                                col_rank_member_count = string.Format("班排名母數{0}", Rows_Counter);


                                if (!table.Columns.Contains(col_subject))
                                {
                                    table.Columns.Add(col_subject);
                                }

                                if (!table.Columns.Contains(col_credit))
                                {
                                    table.Columns.Add(col_credit);
                                }

                                if (!table.Columns.Contains(col_score))
                                {
                                    table.Columns.Add(col_score);
                                }

                                if (!table.Columns.Contains(col_rank))
                                {
                                    table.Columns.Add(col_rank);
                                }


                                if (!table.Columns.Contains(col_rank_member_count))
                                {
                                    table.Columns.Add(col_rank_member_count);
                                }





                                row[col_subject] = dr["subject_name"];

                                row[col_credit] = dr["credit"];

                                row[col_score] = dr["score"];

                                row[col_rank] = dr["rank"];

                                row[col_rank_member_count] = dr["member_count"];



                                foreach (var dr_Cat1 in dt.Select("subject_type='科目' AND group_type= '類別1'"))
                                {

                                    row["類別排名1"] = dr_Cat1["group_name"];

                                    //  這邊的 ""+ 很好用，可以自動幫忙 object  轉型 成 string

                                    if ("" + dr_Cat1["subject_name"] == "" + dr["subject_name"])
                                    {

                                        col_rank_Cat1 = string.Format("類別1排名{0}", Rows_Counter);

                                        col_rank_Cat1_member_count = string.Format("類別1排名母數{0}", Rows_Counter);

                                        if (!table.Columns.Contains(col_rank_Cat1))
                                        {
                                            table.Columns.Add(col_rank_Cat1);
                                        }

                                        if (!table.Columns.Contains(col_rank_Cat1_member_count))
                                        {
                                            table.Columns.Add(col_rank_Cat1_member_count);
                                        }
                                        row[col_rank_Cat1] = dr_Cat1["rank"];

                                        row[col_rank_Cat1_member_count] = dr_Cat1["member_count"];


                                    }
                                }


                                foreach (var dr_Cat2 in dt.Select("subject_type='科目' AND group_type= '類別2'"))
                                {

                                    row["類別排名2"] = dr_Cat2["group_name"];

                                    //  這邊的 ""+ 很好用，可以自動幫忙 object  轉型 成 string

                                    if ("" + dr_Cat2["subject_name"] == "" + dr["subject_name"])
                                    {

                                        col_rank_Cat2 = string.Format("類別2排名{0}", Rows_Counter);

                                        col_rank_Cat2_member_count = string.Format("類別2排名母數{0}", Rows_Counter);

                                        if (!table.Columns.Contains(col_rank_Cat2))
                                        {
                                            table.Columns.Add(col_rank_Cat2);
                                        }

                                        if (!table.Columns.Contains(col_rank_Cat2_member_count))
                                        {
                                            table.Columns.Add(col_rank_Cat2_member_count);
                                        }
                                        row[col_rank_Cat2] = dr_Cat2["rank"];

                                        row[col_rank_Cat2_member_count] = dr_Cat2["member_count"];


                                    }
                                }




                                Rows_Counter++;

                            }

                            
                            
                            
                            

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='總分' AND group_type= '班排名'"))
                            {
                                row["總分"] = item["score"];
                                row["總分班排名"] = item["rank"];
                                row["總分班排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='平均' AND group_type= '班排名'"))
                            {
                                row["平均"] = item["score"];
                                row["平均班排名"] = item["rank"];
                                row["平均班排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='加權總分' AND group_type= '班排名'"))
                            {
                                row["加權總分"] = item["score"];
                                row["加權總分班排名"] = item["rank"];
                                row["加權總分班排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='加權平均' AND group_type= '班排名'"))
                            {
                                row["加權平均"] = item["score"];
                                row["加權平均班排名"] = item["rank"];
                                row["加權平均班排名母數"] = item["member_count"];

                            }

                           
                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='總分' AND group_type= '類別1'"))
                            {
                                row["類別1總分"] = item["score"];
                                row["類別1總分排名"] = item["rank"];
                                row["類別1總分排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='平均' AND group_type= '類別1'"))
                            {
                                row["類別1平均"] = item["score"];
                                row["類別1平均排名"] = item["rank"];
                                row["類別1平均排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='加權總分' AND group_type= '類別1'"))
                            {
                                row["類別1加權總分"] = item["score"];
                                row["類別1加權總分排名"] = item["rank"];
                                row["類別1加權總分排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='加權平均' AND group_type= '類別1'"))
                            {
                                row["類別1加權平均"] = item["score"];
                                row["類別1加權平均排名"] = item["rank"];
                                row["類別1加權平均排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='總分' AND group_type= '類別2'"))
                            {
                                row["類別2總分"] = item["score"];
                                row["類別2總分排名"] = item["rank"];
                                row["類別2總分排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='平均' AND group_type= '類別2'"))
                            {
                                row["類別2平均"] = item["score"];
                                row["類別2平均排名"] = item["rank"];
                                row["類別2平均排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='加權總分' AND group_type= '類別2'"))
                            {
                                row["類別2加權總分"] = item["score"];
                                row["類別2加權總分排名"] = item["rank"];
                                row["類別2加權總分排名母數"] = item["member_count"];

                            }

                            foreach (var item in dt.Select("subject_type='總計' AND subject_name='加權平均' AND group_type= '類別2'"))
                            {
                                row["類別2加權平均"] = item["score"];
                                row["類別2加權平均排名"] = item["rank"];
                                row["類別2加權平均排名母數"] = item["member_count"];

                            }

                            table.Rows.Add(row);

                        }
                        else {

                            Error_List.Add("班級:"+stuRec.Class.Name + "學生:" + stuRec.Name + "沒有固定排名資料，請檢查是否為非排名學生");
                                                
                        }

                        // 下行 1 之前必須要加(decimal) 幫忙轉型，因為StudentList.Count 是int 型別， 1/ (int) 結果還是int，如果數字太小，會造成progress永遠都是0 的窘境

                        progress += (((decimal)1 /StudentList.Count)*89);

                        int progress_int = (int)Math.Round(progress,0,MidpointRounding.AwayFromZero);

                        BGW.ReportProgress(10 + progress_int );                       
                    }
                    

                    Document PageOne = (Document)Template.Clone(true);
                    PageOne.MailMerge.Execute(table);
                    PageOne.MailMerge.DeleteFields();
                    BGW.ReportProgress(100,"作業完成");


                    e.Result = PageOne;
                };

            BGW.RunWorkerAsync();

            #region 計算DoWork完成百分比

            BGW.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("期末成績單產生中...", e.ProgressPercentage);
            };

            #endregion


            
            BGW.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                // 顯示錯誤訊息
                if (Error_List.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var errorMsg in Error_List)
                    {
                        sb.AppendLine(errorMsg);
                    }
                    MsgBox.Show(sb.ToString());
                }



                #region RunWorkerCompleted
                if (e.Cancelled)
                {
                    MsgBox.Show("作業已被中止!!");
                }
                else
                {
                    if (e.Error == null)
                    {
                        Document inResult = (Document)e.Result;

                        try
                        {
                            SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                            SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                            SaveFileDialog1.FileName = "個人學期成績單";

                            if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                            {
                                inResult.Save(SaveFileDialog1.FileName);
                                Process.Start(SaveFileDialog1.FileName);
                            }
                            else
                            {
                                FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                                return;
                            }
                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                            return;
                        }

                        FISCA.Presentation.MotherForm.SetStatusBarMessage("個人學期成績單產生完成", 100);
                    }
                    else
                    {
                        MsgBox.Show("列印資料發生錯誤\n" + e.Error.Message);
                    }
                }
                #endregion
            };


        }

    }
}
