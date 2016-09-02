using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace 固定排名中繼Excel整理
{
    //[FISCA.UDT.TableName("ischool.StudentSH_SemesterScoreReport.Configure")]


    // 2016/6/27 穎驊新增Config，避免另外一個功能使用到同一個設定
    [FISCA.UDT.TableName("ischool.FixedScoreRankReport.Configure")]
    public class Configure : FISCA.UDT.ActiveRecord
    {
        public Configure()
        {
            PrintSubjectList = new List<string>();
            TagRank1TagList = new List<string>();
            TagRank1SubjectList = new List<string>();
            TagRank2TagList = new List<string>();
            TagRank2SubjectList = new List<string>();
            RankFilterTagList = new List<string>();



            SubjectTypeList = new List<string>();

            SubjectNameList = new List<string>();

        }
        /// <summary>
        /// 設定檔名稱
        /// </summary>
        [FISCA.UDT.Field]
        public string Name { get; set; }

        ///2016/6/23 穎驊新增
        /// <summary>
        /// 年級
        /// </summary>
        [FISCA.UDT.Field]
        public string GradeYear { get; set; }


        ///2016/7/4 穎驊新增
        /// <summary>
        /// 類別
        /// </summary>
        [FISCA.UDT.Field]

        private String SubjectTypeString { get; set; }

        public List<string> SubjectTypeList { get; set; }


        ///2016/7/4 穎驊新增
        /// <summary>
        /// 科目名稱
        /// </summary>
        [FISCA.UDT.Field]

        private String SubjectNameString { get; set; }
        public List<string> SubjectNameList { get; set; }



        /// <summary>
        /// 學年度
        /// </summary>
        [FISCA.UDT.Field]
        public string SchoolYear { get; set; }
        /// <summary>
        /// 學期
        /// </summary>
        [FISCA.UDT.Field]
        public string Semester { get; set; }
        /// <summary>
        /// 樣板包含學年成績
        /// </summary>
        [FISCA.UDT.Field]
        public bool WithSchoolYearScore { get; set; }
        /// <summary>
        /// 樣板包含上學期成績
        /// </summary>
        [FISCA.UDT.Field]
        public bool WithPrevSemesterScore { get; set; }



        /// 樣板，已經不需要了，穎驊來準備把他給拆了吧... 
        /// <summary>
        /// 列印樣板
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream { get; set; }
        public Aspose.Words.Document Template { get; set; }



        /// <summary>
        /// 樣板中支援列印科目的最大數
        /// </summary>
        [FISCA.UDT.Field]
        public int SubjectLimit { get; set; }

        /// <summary>
        /// 列印試別
        /// </summary>
        [FISCA.UDT.Field]
        private string ExamRecordID { get; set; }
        public K12.Data.ExamRecord ExamRecord { get; set; }
        
        /// <summary>
        /// 參考成績試別
        /// </summary>
        [FISCA.UDT.Field]
        private string RefenceExamRecordID { get; set; }
        public K12.Data.ExamRecord RefenceExamRecord { get; set; }
        
        
        /// <summary>
        /// 列印科別
        /// </summary>
        [FISCA.UDT.Field]
        private string PrintSubjectListString { get; set; }
        public List<string> PrintSubjectList { get; private set; }
        /// <summary>
        /// 類別排名1
        /// </summary>
        [FISCA.UDT.Field]
        public string TagRank1TagName { get; set; }
        public List<string> TagRank1TagList { get; private set; }
        /// <summary>
        /// 類別排名1，排名科目
        /// </summary>
        [FISCA.UDT.Field]
        private string TagRank1SubjectListString { get; set; }
        public List<string> TagRank1SubjectList { get; private set; }
        /// <summary>
        /// 類別排名2
        /// </summary>
        [FISCA.UDT.Field]
        public string TagRank2TagName { get; set; }
        public List<string> TagRank2TagList { get; private set; }
        /// <summary>
        /// 類別排名2，排名科目
        /// </summary>
        [FISCA.UDT.Field]
        private string TagRank2SubjectListString { get; set; }
        public List<string> TagRank2SubjectList { get; private set; }
        /// <summary>
        /// 不參與排名學生類別
        /// </summary>
        [FISCA.UDT.Field]
        public string RankFilterTagName { get; set; }
        public List<string> RankFilterTagList { get; private set; }


        [FISCA.UDT.Field]
        public string FileName { get; set;}

        /// <summary>
        /// 在儲存前，把資料填入儲存欄位中
        /// </summary>
        public void Encode()
        {
            this.ExamRecordID = (this.ExamRecord == null ? "" : this.ExamRecord.ID);

            this.PrintSubjectListString = "";
            foreach (var item in this.PrintSubjectList)
            {
                this.PrintSubjectListString += (this.PrintSubjectListString == "" ? "" : "^^^") + item;
            }

            this.RefenceExamRecordID = (this.RefenceExamRecord == null ? "" : this.RefenceExamRecord.ID);

            this.TagRank1SubjectListString = "";
            foreach (var item in this.TagRank1SubjectList)
            {
                this.TagRank1SubjectListString += (this.TagRank1SubjectListString == "" ? "" : "^^^") + item;
            }


            this.TagRank2SubjectListString = "";
            foreach (var item in this.TagRank2SubjectList)
            {
                this.TagRank2SubjectListString += (this.TagRank2SubjectListString == "" ? "" : "^^^") + item;
            }

            
            // 2016/7/4 穎驊依樣畫葫蘆做類別、科目成績項目，看起來資料欄位只能儲存字串，卻不能存List，所以要拆

            if (this.SubjectTypeList != null)
            {


                this.SubjectTypeString = "";
                foreach (var item in this.SubjectTypeList)
                {
                    this.SubjectTypeString += (this.SubjectTypeString == "" ? "" : "^^^") + item;
                }
            }

            if (this.SubjectNameList != null)
            {


                this.SubjectNameString = "";
                foreach (var item in this.SubjectNameList)
                {
                    this.SubjectNameString += (this.SubjectNameString == "" ? "" : "^^^") + item;
                }
            }


            
            //System.IO.MemoryStream stream = new System.IO.MemoryStream();
            //this.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
            //this.TemplateStream = Convert.ToBase64String(stream.ToArray());





        }
        /// <summary>
        /// 在資料取出後，把資料從儲存欄位轉換至資料欄位
        /// </summary>
        public void Decode()
        {
            this.ExamRecord = K12.Data.Exam.SelectByID(this.ExamRecordID);

            //2016/8/29 防弊
            if (this.PrintSubjectListString != null) {

                this.PrintSubjectList = new List<string>(this.PrintSubjectListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            }
            
            
            //this.RefenceExamRecord = K12.Data.Exam.SelectByID(this.RefenceExamRecordID);


            if (this.TagRank1SubjectListString != null) {
                this.TagRank1SubjectList = new List<string>(this.TagRank1SubjectListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            }


            if (this.TagRank2SubjectListString != null) {
                this.TagRank2SubjectList = new List<string>(this.TagRank2SubjectListString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            }
            


            // 2016/7/4 穎驊依樣畫葫蘆做類別、科目成績項目，看起來資料欄位只能儲存字串，卻不能存List，所以要拆完變成一個字串儲存後，在這邊在打散成List

            if (this.SubjectTypeString != null) {

                this.SubjectTypeList = new List<string>(this.SubjectTypeString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            }


            if (this.SubjectNameString != null) {
                this.SubjectNameList = new List<string>(this.SubjectNameString.Split(new string[] { "^^^" }, StringSplitOptions.RemoveEmptyEntries));
            }
            

            //this.Template = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream)));

        }
    }
}
