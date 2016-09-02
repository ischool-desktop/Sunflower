using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.IO;
//using K12.Data;
//using SHSchool.Data;
using SmartSchool.Customization.Data;
using System.Threading;
using SmartSchool.Customization.Data.StudentExtension;
using FISCA.Permission;
using SmartSchool;
using FISCA.Presentation;
using Aspose.Cells;
using FISCA;
using FISCA.Presentation.Controls;
using K12.Presentation;

using SmartSchool.Customization;
using 固定排名中繼Excel整理.UDT;

namespace 固定排名中繼Excel整理
{
    // 2016/6/16(馬叔歸國)，穎驊紀錄，此Code原本為期末成績單，但由於最近公司要產出"固定排名"的功能，恩正說拿原本算學生的期末成績單Code的部分功能來修改，
    //主要目標要生出一張記錄詳情報告的Excel表單，可供正是上傳Sever、使用者驗證資料，估計這一個禮拜都要做這個，這是我進公司已來看過最複雜的報表(CODE我已經刪掉一堆了)


    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            var btn = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["固定排名中繼Excel整理(測試版)"]["固定排名中繼Excel整理(測試版666666666666666)"];

            btn.Enable = Permissions.固定排名中繼Excel整理test權限;

            btn.Click += new EventHandler(Program_Click);

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"];
            permission.Add(new RibbonFeature(Permissions.固定排名中繼Excel整理test, "固定排名中繼Excel整理(測試版)"));
        }

        private static string GetNumber(decimal? p)
        {
            if (p == null) return "";
            string levelNumber;
            switch (((int)p.Value))
            {
                #region 對應levelNumber
                case 0:
                    levelNumber = "";
                    break;
                case 1:
                    levelNumber = "Ⅰ";
                    break;
                case 2:
                    levelNumber = "Ⅱ";
                    break;
                case 3:
                    levelNumber = "Ⅲ";
                    break;
                case 4:
                    levelNumber = "Ⅳ";
                    break;
                case 5:
                    levelNumber = "Ⅴ";
                    break;
                case 6:
                    levelNumber = "Ⅵ";
                    break;
                case 7:
                    levelNumber = "Ⅶ";
                    break;
                case 8:
                    levelNumber = "Ⅷ";
                    break;
                case 9:
                    levelNumber = "Ⅸ";
                    break;
                case 10:
                    levelNumber = "Ⅹ";
                    break;
                default:
                    levelNumber = "" + (p);
                    break;
                #endregion
            }
            return levelNumber;
        }

        static Dictionary<string, decimal> _studPassSumCreditDict1 = new Dictionary<string, decimal>();
        static Dictionary<string, decimal> _studPassSumCreditDictAll = new Dictionary<string, decimal>();

        // 累計取得必修學分
        static Dictionary<string, decimal> _studPassSumCreditDictC1 = new Dictionary<string, decimal>();
        // 累計取得選修學分
        static Dictionary<string, decimal> _studPassSumCreditDictC2 = new Dictionary<string, decimal>();

        static void Program_Click(object sender_, EventArgs e_)
        {
            AccessHelper helper = new AccessHelper();
            List<StudentRecord> lista = helper.StudentHelper.GetSelectedStudent();

            // 取得學生及格與補考標準
            Dictionary<string, Dictionary<string, decimal>> StudentApplyLimitDict = Utility.GetStudentApplyLimitDict(lista);







            ConfigForm form = new ConfigForm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                AccessHelper accessHelper = new AccessHelper();
                //return;
                List<StudentRecord> overflowRecords = new List<StudentRecord>();
                //取得列印設定
                Configure conf = form.Configure;
                //建立測試的選取學生(先期不管怎麼選就是印這些人)
                //List<string> selectedStudents = K12.Presentation.NLDPanels.Student.SelectedSource;
                List<string> selectedStudents = new List<string>();
                #region 以年級為單位


                //2016/6/23，穎驊新增 看你項目選幾年級 就加入該全年級，資料來源為在ConfigForm時，你所選的欄位

                string targetGrade = conf.GradeYear;
                foreach (var stuRec in accessHelper.StudentHelper.GetAllStudent())
                {
                    if (stuRec.RefClass != null && stuRec.RefClass.GradeYear == targetGrade)
                        selectedStudents.Add(stuRec.StudentID);
                }



                #region 整理回歸科目

                //2016/7/5 穎驊新增，整理回歸科目使用

                //整理中過渡使用
                Dictionary<string, string> SubjectTypeArrange = new Dictionary<string, string>();

                //整理完成，每一個科目會對應一個回歸科目
                Dictionary<string, string> SubjectTypeArrange_complete = new Dictionary<string, string>();

                //整理中使用，來接拆完的String
                List<string> SubjectNameList_Split = new List<string>();


                for (int i = 0; i < conf.SubjectTypeList.Count; i++)
                {

                    //  加這一行的用意是，有的時候使用者會在GridView 點一下下面新增Row的空白行後，又不打東西，然後又儲存了，就會產生有null的List，造成新增Dictionary的錯誤
                    if (conf.SubjectTypeList[i] != null)
                    {
                        SubjectTypeArrange.Add(conf.SubjectTypeList[i], conf.SubjectNameList[i]);
                    }
                }

                foreach (var item in SubjectTypeArrange)
                {


                    SubjectNameList_Split = new List<string>(item.Value.Split(new string[] { "，" }, StringSplitOptions.RemoveEmptyEntries));

                    foreach (var SubjectNameList_SplitName in SubjectNameList_Split)
                    {

                        SubjectTypeArrange_complete.Add(SubjectNameList_SplitName, item.Key);

                    }
                }
                #endregion







                #endregion
                //建立合併欄位總表
                DataTable table = new DataTable();
                #region 所有的合併欄位


                table.Columns.Add("科別名稱");
                table.Columns.Add("試別");

                table.Columns.Add("收件人");
                table.Columns.Add("學年度");
                table.Columns.Add("學期");
                table.Columns.Add("班級科別名稱");
                table.Columns.Add("班級");
                table.Columns.Add("班導師");
                table.Columns.Add("座號");
                table.Columns.Add("學號");
                table.Columns.Add("姓名");
                table.Columns.Add("定期評量");
                table.Columns.Add("本學期取得學分數");
                table.Columns.Add("累計取得學分數");
                table.Columns.Add("累計取得必修學分");
                table.Columns.Add("累計取得選修學分");
                table.Columns.Add("系統學年度");
                table.Columns.Add("系統學期");

                for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                {
                    table.Columns.Add("科目名稱" + subjectIndex);
                    table.Columns.Add("學分數" + subjectIndex);
                    table.Columns.Add("前次成績" + subjectIndex);
                    table.Columns.Add("科目成績" + subjectIndex);
                    // 新增學期科目相關成績--
                    table.Columns.Add("科目必選修" + subjectIndex);
                    table.Columns.Add("科目校部定" + subjectIndex);
                    table.Columns.Add("科目註記" + subjectIndex);
                    table.Columns.Add("科目取得學分" + subjectIndex);
                    table.Columns.Add("科目未取得學分註記" + subjectIndex);
                    table.Columns.Add("學期科目原始成績" + subjectIndex);
                    table.Columns.Add("學期科目補考成績" + subjectIndex);
                    table.Columns.Add("學期科目重修成績" + subjectIndex);
                    table.Columns.Add("學期科目手動調整成績" + subjectIndex);
                    table.Columns.Add("學期科目學年調整成績" + subjectIndex);
                    table.Columns.Add("學期科目成績" + subjectIndex);
                    table.Columns.Add("學期科目原始成績註記" + subjectIndex);
                    table.Columns.Add("學期科目補考成績註記" + subjectIndex);
                    table.Columns.Add("學期科目重修成績註記" + subjectIndex);
                    table.Columns.Add("學期科目手動成績註記" + subjectIndex);
                    table.Columns.Add("學期科目學年成績註記" + subjectIndex);
                    table.Columns.Add("學期科目需要補考註記" + subjectIndex);
                    table.Columns.Add("學期科目需要重修註記" + subjectIndex);
                    // 新增學期科目排名
                    table.Columns.Add("學期科目排名成績" + subjectIndex);
                    table.Columns.Add("學期科目班排名" + subjectIndex);
                    table.Columns.Add("學期科目班排名母數" + subjectIndex);
                    table.Columns.Add("學期科目科排名" + subjectIndex);
                    table.Columns.Add("學期科目科排名母數" + subjectIndex);
                    table.Columns.Add("學期科目類別1排名" + subjectIndex);
                    table.Columns.Add("學期科目類別1排名母數" + subjectIndex);
                    table.Columns.Add("學期科目類別2排名" + subjectIndex);
                    table.Columns.Add("學期科目類別2排名母數" + subjectIndex);
                    table.Columns.Add("學期科目全校排名" + subjectIndex);
                    table.Columns.Add("學期科目全校排名母數" + subjectIndex);
                    // 新增上學期科目相關成績--
                    table.Columns.Add("上學期科目原始成績" + subjectIndex);
                    table.Columns.Add("上學期科目補考成績" + subjectIndex);
                    table.Columns.Add("上學期科目重修成績" + subjectIndex);
                    table.Columns.Add("上學期科目手動調整成績" + subjectIndex);
                    table.Columns.Add("上學期科目學年調整成績" + subjectIndex);
                    table.Columns.Add("上學期科目成績" + subjectIndex);
                    table.Columns.Add("上學期科目原始成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目補考成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目重修成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目手動成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目學年成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目取得學分" + subjectIndex);
                    table.Columns.Add("上學期科目未取得學分註記" + subjectIndex);
                    table.Columns.Add("上學期科目需要補考註記" + subjectIndex);
                    table.Columns.Add("上學期科目需要重修註記" + subjectIndex);

                    // 新增學年科目成績--
                    table.Columns.Add("學年科目成績" + subjectIndex);
                    //定期評量成績個項欄位--
                    table.Columns.Add("班排名" + subjectIndex);
                    table.Columns.Add("班排名母數" + subjectIndex);
                    table.Columns.Add("科排名" + subjectIndex);
                    table.Columns.Add("科排名母數" + subjectIndex);
                    table.Columns.Add("類別1排名" + subjectIndex);
                    table.Columns.Add("類別1排名母數" + subjectIndex);
                    table.Columns.Add("類別2排名" + subjectIndex);
                    table.Columns.Add("類別2排名母數" + subjectIndex);
                    table.Columns.Add("全校排名" + subjectIndex);
                    table.Columns.Add("全校排名母數" + subjectIndex);
                    #region 瘋狂的組距及分析
                    table.Columns.Add("班高標" + subjectIndex); table.Columns.Add("科高標" + subjectIndex); table.Columns.Add("校高標" + subjectIndex); table.Columns.Add("類1高標" + subjectIndex); table.Columns.Add("類2高標" + subjectIndex);
                    table.Columns.Add("班均標" + subjectIndex); table.Columns.Add("科均標" + subjectIndex); table.Columns.Add("校均標" + subjectIndex); table.Columns.Add("類1均標" + subjectIndex); table.Columns.Add("類2均標" + subjectIndex);
                    table.Columns.Add("班低標" + subjectIndex); table.Columns.Add("科低標" + subjectIndex); table.Columns.Add("校低標" + subjectIndex); table.Columns.Add("類1低標" + subjectIndex); table.Columns.Add("類2低標" + subjectIndex);
                    table.Columns.Add("班標準差" + subjectIndex); table.Columns.Add("科標準差" + subjectIndex); table.Columns.Add("校標準差" + subjectIndex); table.Columns.Add("類1標準差" + subjectIndex); table.Columns.Add("類2標準差" + subjectIndex);
                    table.Columns.Add("班組距" + subjectIndex + "count90"); table.Columns.Add("科組距" + subjectIndex + "count90"); table.Columns.Add("校組距" + subjectIndex + "count90"); table.Columns.Add("類1組距" + subjectIndex + "count90"); table.Columns.Add("類2組距" + subjectIndex + "count90");
                    table.Columns.Add("班組距" + subjectIndex + "count80"); table.Columns.Add("科組距" + subjectIndex + "count80"); table.Columns.Add("校組距" + subjectIndex + "count80"); table.Columns.Add("類1組距" + subjectIndex + "count80"); table.Columns.Add("類2組距" + subjectIndex + "count80");
                    table.Columns.Add("班組距" + subjectIndex + "count70"); table.Columns.Add("科組距" + subjectIndex + "count70"); table.Columns.Add("校組距" + subjectIndex + "count70"); table.Columns.Add("類1組距" + subjectIndex + "count70"); table.Columns.Add("類2組距" + subjectIndex + "count70");
                    table.Columns.Add("班組距" + subjectIndex + "count60"); table.Columns.Add("科組距" + subjectIndex + "count60"); table.Columns.Add("校組距" + subjectIndex + "count60"); table.Columns.Add("類1組距" + subjectIndex + "count60"); table.Columns.Add("類2組距" + subjectIndex + "count60");
                    table.Columns.Add("班組距" + subjectIndex + "count50"); table.Columns.Add("科組距" + subjectIndex + "count50"); table.Columns.Add("校組距" + subjectIndex + "count50"); table.Columns.Add("類1組距" + subjectIndex + "count50"); table.Columns.Add("類2組距" + subjectIndex + "count50");
                    table.Columns.Add("班組距" + subjectIndex + "count40"); table.Columns.Add("科組距" + subjectIndex + "count40"); table.Columns.Add("校組距" + subjectIndex + "count40"); table.Columns.Add("類1組距" + subjectIndex + "count40"); table.Columns.Add("類2組距" + subjectIndex + "count40");
                    table.Columns.Add("班組距" + subjectIndex + "count30"); table.Columns.Add("科組距" + subjectIndex + "count30"); table.Columns.Add("校組距" + subjectIndex + "count30"); table.Columns.Add("類1組距" + subjectIndex + "count30"); table.Columns.Add("類2組距" + subjectIndex + "count30");
                    table.Columns.Add("班組距" + subjectIndex + "count20"); table.Columns.Add("科組距" + subjectIndex + "count20"); table.Columns.Add("校組距" + subjectIndex + "count20"); table.Columns.Add("類1組距" + subjectIndex + "count20"); table.Columns.Add("類2組距" + subjectIndex + "count20");
                    table.Columns.Add("班組距" + subjectIndex + "count10"); table.Columns.Add("科組距" + subjectIndex + "count10"); table.Columns.Add("校組距" + subjectIndex + "count10"); table.Columns.Add("類1組距" + subjectIndex + "count10"); table.Columns.Add("類2組距" + subjectIndex + "count10");
                    table.Columns.Add("班組距" + subjectIndex + "count100Up"); table.Columns.Add("科組距" + subjectIndex + "count100Up"); table.Columns.Add("校組距" + subjectIndex + "count100Up"); table.Columns.Add("類1組距" + subjectIndex + "count100Up"); table.Columns.Add("類2組距" + subjectIndex + "count100Up");
                    table.Columns.Add("班組距" + subjectIndex + "count90Up"); table.Columns.Add("科組距" + subjectIndex + "count90Up"); table.Columns.Add("校組距" + subjectIndex + "count90Up"); table.Columns.Add("類1組距" + subjectIndex + "count90Up"); table.Columns.Add("類2組距" + subjectIndex + "count90Up");
                    table.Columns.Add("班組距" + subjectIndex + "count80Up"); table.Columns.Add("科組距" + subjectIndex + "count80Up"); table.Columns.Add("校組距" + subjectIndex + "count80Up"); table.Columns.Add("類1組距" + subjectIndex + "count80Up"); table.Columns.Add("類2組距" + subjectIndex + "count80Up");
                    table.Columns.Add("班組距" + subjectIndex + "count70Up"); table.Columns.Add("科組距" + subjectIndex + "count70Up"); table.Columns.Add("校組距" + subjectIndex + "count70Up"); table.Columns.Add("類1組距" + subjectIndex + "count70Up"); table.Columns.Add("類2組距" + subjectIndex + "count70Up");
                    table.Columns.Add("班組距" + subjectIndex + "count60Up"); table.Columns.Add("科組距" + subjectIndex + "count60Up"); table.Columns.Add("校組距" + subjectIndex + "count60Up"); table.Columns.Add("類1組距" + subjectIndex + "count60Up"); table.Columns.Add("類2組距" + subjectIndex + "count60Up");
                    table.Columns.Add("班組距" + subjectIndex + "count50Up"); table.Columns.Add("科組距" + subjectIndex + "count50Up"); table.Columns.Add("校組距" + subjectIndex + "count50Up"); table.Columns.Add("類1組距" + subjectIndex + "count50Up"); table.Columns.Add("類2組距" + subjectIndex + "count50Up");
                    table.Columns.Add("班組距" + subjectIndex + "count40Up"); table.Columns.Add("科組距" + subjectIndex + "count40Up"); table.Columns.Add("校組距" + subjectIndex + "count40Up"); table.Columns.Add("類1組距" + subjectIndex + "count40Up"); table.Columns.Add("類2組距" + subjectIndex + "count40Up");
                    table.Columns.Add("班組距" + subjectIndex + "count30Up"); table.Columns.Add("科組距" + subjectIndex + "count30Up"); table.Columns.Add("校組距" + subjectIndex + "count30Up"); table.Columns.Add("類1組距" + subjectIndex + "count30Up"); table.Columns.Add("類2組距" + subjectIndex + "count30Up");
                    table.Columns.Add("班組距" + subjectIndex + "count20Up"); table.Columns.Add("科組距" + subjectIndex + "count20Up"); table.Columns.Add("校組距" + subjectIndex + "count20Up"); table.Columns.Add("類1組距" + subjectIndex + "count20Up"); table.Columns.Add("類2組距" + subjectIndex + "count20Up");
                    table.Columns.Add("班組距" + subjectIndex + "count10Up"); table.Columns.Add("科組距" + subjectIndex + "count10Up"); table.Columns.Add("校組距" + subjectIndex + "count10Up"); table.Columns.Add("類1組距" + subjectIndex + "count10Up"); table.Columns.Add("類2組距" + subjectIndex + "count10Up");
                    table.Columns.Add("班組距" + subjectIndex + "count90Down"); table.Columns.Add("科組距" + subjectIndex + "count90Down"); table.Columns.Add("校組距" + subjectIndex + "count90Down"); table.Columns.Add("類1組距" + subjectIndex + "count90Down"); table.Columns.Add("類2組距" + subjectIndex + "count90Down");
                    table.Columns.Add("班組距" + subjectIndex + "count80Down"); table.Columns.Add("科組距" + subjectIndex + "count80Down"); table.Columns.Add("校組距" + subjectIndex + "count80Down"); table.Columns.Add("類1組距" + subjectIndex + "count80Down"); table.Columns.Add("類2組距" + subjectIndex + "count80Down");
                    table.Columns.Add("班組距" + subjectIndex + "count70Down"); table.Columns.Add("科組距" + subjectIndex + "count70Down"); table.Columns.Add("校組距" + subjectIndex + "count70Down"); table.Columns.Add("類1組距" + subjectIndex + "count70Down"); table.Columns.Add("類2組距" + subjectIndex + "count70Down");
                    table.Columns.Add("班組距" + subjectIndex + "count60Down"); table.Columns.Add("科組距" + subjectIndex + "count60Down"); table.Columns.Add("校組距" + subjectIndex + "count60Down"); table.Columns.Add("類1組距" + subjectIndex + "count60Down"); table.Columns.Add("類2組距" + subjectIndex + "count60Down");
                    table.Columns.Add("班組距" + subjectIndex + "count50Down"); table.Columns.Add("科組距" + subjectIndex + "count50Down"); table.Columns.Add("校組距" + subjectIndex + "count50Down"); table.Columns.Add("類1組距" + subjectIndex + "count50Down"); table.Columns.Add("類2組距" + subjectIndex + "count50Down");
                    table.Columns.Add("班組距" + subjectIndex + "count40Down"); table.Columns.Add("科組距" + subjectIndex + "count40Down"); table.Columns.Add("校組距" + subjectIndex + "count40Down"); table.Columns.Add("類1組距" + subjectIndex + "count40Down"); table.Columns.Add("類2組距" + subjectIndex + "count40Down");
                    table.Columns.Add("班組距" + subjectIndex + "count30Down"); table.Columns.Add("科組距" + subjectIndex + "count30Down"); table.Columns.Add("校組距" + subjectIndex + "count30Down"); table.Columns.Add("類1組距" + subjectIndex + "count30Down"); table.Columns.Add("類2組距" + subjectIndex + "count30Down");
                    table.Columns.Add("班組距" + subjectIndex + "count20Down"); table.Columns.Add("科組距" + subjectIndex + "count20Down"); table.Columns.Add("校組距" + subjectIndex + "count20Down"); table.Columns.Add("類1組距" + subjectIndex + "count20Down"); table.Columns.Add("類2組距" + subjectIndex + "count20Down");
                    table.Columns.Add("班組距" + subjectIndex + "count10Down"); table.Columns.Add("科組距" + subjectIndex + "count10Down"); table.Columns.Add("校組距" + subjectIndex + "count10Down"); table.Columns.Add("類1組距" + subjectIndex + "count10Down"); table.Columns.Add("類2組距" + subjectIndex + "count10Down");
                    #endregion
                }
                table.Columns.Add("總分");
                table.Columns.Add("總分班排名");
                table.Columns.Add("總分班排名母數");
                table.Columns.Add("總分科排名");
                table.Columns.Add("總分科排名母數");
                table.Columns.Add("總分全校排名");
                table.Columns.Add("總分全校排名母數");
                table.Columns.Add("平均");
                table.Columns.Add("平均班排名");
                table.Columns.Add("平均班排名母數");
                table.Columns.Add("平均科排名");
                table.Columns.Add("平均科排名母數");
                table.Columns.Add("平均全校排名");
                table.Columns.Add("平均全校排名母數");
                // 學期分項成績 --
                table.Columns.Add("學期學業成績");
                table.Columns.Add("學期體育成績");
                table.Columns.Add("學期國防通識成績");
                table.Columns.Add("學期健康與護理成績");
                table.Columns.Add("學期實習科目成績");
                table.Columns.Add("學期學業(原始)成績");
                table.Columns.Add("學期體育(原始)成績");
                table.Columns.Add("學期國防通識(原始)成績");
                table.Columns.Add("學期健康與護理(原始)成績");
                table.Columns.Add("學期實習科目(原始)成績");
                table.Columns.Add("學期專業科目成績");
                table.Columns.Add("學期專業科目(原始)成績");

                table.Columns.Add("學期德行成績");
                // 學期學業成績排名
                table.Columns.Add("學期學業成績班排名");
                table.Columns.Add("學期學業成績科排名");
                table.Columns.Add("學期學業成績類別1排名");
                table.Columns.Add("學期學業成績類別2排名");
                table.Columns.Add("學期學業成績校排名");
                table.Columns.Add("學期學業成績班排名母數");
                table.Columns.Add("學期學業成績科排名母數");
                table.Columns.Add("學期學業成績類別1排名母數");
                table.Columns.Add("學期學業成績類別2排名母數");
                table.Columns.Add("學期學業成績校排名母數");
                // 導師評語 --
                table.Columns.Add("導師評語");
                // 獎懲統計 --
                table.Columns.Add("大功統計");
                table.Columns.Add("小功統計");
                table.Columns.Add("嘉獎統計");
                table.Columns.Add("大過統計");
                table.Columns.Add("小過統計");
                table.Columns.Add("警告統計");
                table.Columns.Add("留校察看");
                // 上學期分項成績 --
                table.Columns.Add("上學期學業成績");
                table.Columns.Add("上學期體育成績");
                table.Columns.Add("上學期國防通識成績");
                table.Columns.Add("上學期健康與護理成績");
                table.Columns.Add("上學期實習科目成績");
                table.Columns.Add("上學期德行成績");
                // 學年分項成績 --
                table.Columns.Add("學年學業成績");
                table.Columns.Add("學年體育成績");
                table.Columns.Add("學年國防通識成績");
                table.Columns.Add("學年健康與護理成績");
                table.Columns.Add("學年實習科目成績");
                table.Columns.Add("學年德行成績");
                table.Columns.Add("學年學業成績班排名");

                // 服務學習時數
                table.Columns.Add("前學期服務學習時數");
                table.Columns.Add("本學期服務學習時數");
                table.Columns.Add("學年服務學習時數");

                // 缺曠統計
                // 動態新增缺曠統計，使用模式一般_曠課、一般_事假..
                foreach (string name in Utility.GetATMappingKey())
                {
                    table.Columns.Add("前學期" + name);
                    table.Columns.Add("本學期" + name);
                    table.Columns.Add("學年" + name);
                }
                // --
                table.Columns.Add("加權總分");
                table.Columns.Add("加權總分班排名");
                table.Columns.Add("加權總分班排名母數");
                table.Columns.Add("加權總分科排名");
                table.Columns.Add("加權總分科排名母數");
                table.Columns.Add("加權總分全校排名");
                table.Columns.Add("加權總分全校排名母數");
                table.Columns.Add("加權平均");
                table.Columns.Add("加權平均班排名");
                table.Columns.Add("加權平均班排名母數");
                table.Columns.Add("加權平均科排名");
                table.Columns.Add("加權平均科排名母數");
                table.Columns.Add("加權平均全校排名");
                table.Columns.Add("加權平均全校排名母數");

                table.Columns.Add("類別排名1");
                table.Columns.Add("類別1總分");
                table.Columns.Add("類別1總分排名");
                table.Columns.Add("類別1總分排名母數");
                table.Columns.Add("類別1平均");
                table.Columns.Add("類別1平均排名");
                table.Columns.Add("類別1平均排名母數");
                table.Columns.Add("類別1加權總分");
                table.Columns.Add("類別1加權總分排名");
                table.Columns.Add("類別1加權總分排名母數");
                table.Columns.Add("類別1加權平均");
                table.Columns.Add("類別1加權平均排名");
                table.Columns.Add("類別1加權平均排名母數");

                table.Columns.Add("類別排名2");
                table.Columns.Add("類別2總分");
                table.Columns.Add("類別2總分排名");
                table.Columns.Add("類別2總分排名母數");
                table.Columns.Add("類別2平均");
                table.Columns.Add("類別2平均排名");
                table.Columns.Add("類別2平均排名母數");
                table.Columns.Add("類別2加權總分");
                table.Columns.Add("類別2加權總分排名");
                table.Columns.Add("類別2加權總分排名母數");
                table.Columns.Add("類別2加權平均");
                table.Columns.Add("類別2加權平均排名");
                table.Columns.Add("類別2加權平均排名母數");

                #region 瘋狂的組距及分析
                table.Columns.Add("總分班高標"); table.Columns.Add("總分科高標"); table.Columns.Add("總分校高標"); table.Columns.Add("平均班高標"); table.Columns.Add("平均科高標"); table.Columns.Add("平均校高標"); table.Columns.Add("加權總分班高標"); table.Columns.Add("加權總分科高標"); table.Columns.Add("加權總分校高標"); table.Columns.Add("加權平均班高標"); table.Columns.Add("加權平均科高標"); table.Columns.Add("加權平均校高標"); table.Columns.Add("類1總分高標"); table.Columns.Add("類1平均高標"); table.Columns.Add("類1加權總分高標"); table.Columns.Add("類1加權平均高標"); table.Columns.Add("類2總分高標"); table.Columns.Add("類2平均高標"); table.Columns.Add("類2加權總分高標"); table.Columns.Add("類2加權平均高標");
                table.Columns.Add("總分班均標"); table.Columns.Add("總分科均標"); table.Columns.Add("總分校均標"); table.Columns.Add("平均班均標"); table.Columns.Add("平均科均標"); table.Columns.Add("平均校均標"); table.Columns.Add("加權總分班均標"); table.Columns.Add("加權總分科均標"); table.Columns.Add("加權總分校均標"); table.Columns.Add("加權平均班均標"); table.Columns.Add("加權平均科均標"); table.Columns.Add("加權平均校均標"); table.Columns.Add("類1總分均標"); table.Columns.Add("類1平均均標"); table.Columns.Add("類1加權總分均標"); table.Columns.Add("類1加權平均均標"); table.Columns.Add("類2總分均標"); table.Columns.Add("類2平均均標"); table.Columns.Add("類2加權總分均標"); table.Columns.Add("類2加權平均均標");
                table.Columns.Add("總分班低標"); table.Columns.Add("總分科低標"); table.Columns.Add("總分校低標"); table.Columns.Add("平均班低標"); table.Columns.Add("平均科低標"); table.Columns.Add("平均校低標"); table.Columns.Add("加權總分班低標"); table.Columns.Add("加權總分科低標"); table.Columns.Add("加權總分校低標"); table.Columns.Add("加權平均班低標"); table.Columns.Add("加權平均科低標"); table.Columns.Add("加權平均校低標"); table.Columns.Add("類1總分低標"); table.Columns.Add("類1平均低標"); table.Columns.Add("類1加權總分低標"); table.Columns.Add("類1加權平均低標"); table.Columns.Add("類2總分低標"); table.Columns.Add("類2平均低標"); table.Columns.Add("類2加權總分低標"); table.Columns.Add("類2加權平均低標");
                table.Columns.Add("總分班標準差"); table.Columns.Add("總分科標準差"); table.Columns.Add("總分校標準差"); table.Columns.Add("平均班標準差"); table.Columns.Add("平均科標準差"); table.Columns.Add("平均校標準差"); table.Columns.Add("加權總分班標準差"); table.Columns.Add("加權總分科標準差"); table.Columns.Add("加權總分校標準差"); table.Columns.Add("加權平均班標準差"); table.Columns.Add("加權平均科標準差"); table.Columns.Add("加權平均校標準差"); table.Columns.Add("類1總分標準差"); table.Columns.Add("類1平均標準差"); table.Columns.Add("類1加權總分標準差"); table.Columns.Add("類1加權平均標準差"); table.Columns.Add("類2總分標準差"); table.Columns.Add("類2平均標準差"); table.Columns.Add("類2加權總分標準差"); table.Columns.Add("類2加權平均標準差");
                table.Columns.Add("總分班組距count90"); table.Columns.Add("總分科組距count90"); table.Columns.Add("總分校組距count90"); table.Columns.Add("平均班組距count90"); table.Columns.Add("平均科組距count90"); table.Columns.Add("平均校組距count90"); table.Columns.Add("加權總分班組距count90"); table.Columns.Add("加權總分科組距count90"); table.Columns.Add("加權總分校組距count90"); table.Columns.Add("加權平均班組距count90"); table.Columns.Add("加權平均科組距count90"); table.Columns.Add("加權平均校組距count90"); table.Columns.Add("類1總分組距count90"); table.Columns.Add("類1平均組距count90"); table.Columns.Add("類1加權總分組距count90"); table.Columns.Add("類1加權平均組距count90"); table.Columns.Add("類2總分組距count90"); table.Columns.Add("類2平均組距count90"); table.Columns.Add("類2加權總分組距count90"); table.Columns.Add("類2加權平均組距count90");
                table.Columns.Add("總分班組距count80"); table.Columns.Add("總分科組距count80"); table.Columns.Add("總分校組距count80"); table.Columns.Add("平均班組距count80"); table.Columns.Add("平均科組距count80"); table.Columns.Add("平均校組距count80"); table.Columns.Add("加權總分班組距count80"); table.Columns.Add("加權總分科組距count80"); table.Columns.Add("加權總分校組距count80"); table.Columns.Add("加權平均班組距count80"); table.Columns.Add("加權平均科組距count80"); table.Columns.Add("加權平均校組距count80"); table.Columns.Add("類1總分組距count80"); table.Columns.Add("類1平均組距count80"); table.Columns.Add("類1加權總分組距count80"); table.Columns.Add("類1加權平均組距count80"); table.Columns.Add("類2總分組距count80"); table.Columns.Add("類2平均組距count80"); table.Columns.Add("類2加權總分組距count80"); table.Columns.Add("類2加權平均組距count80");
                table.Columns.Add("總分班組距count70"); table.Columns.Add("總分科組距count70"); table.Columns.Add("總分校組距count70"); table.Columns.Add("平均班組距count70"); table.Columns.Add("平均科組距count70"); table.Columns.Add("平均校組距count70"); table.Columns.Add("加權總分班組距count70"); table.Columns.Add("加權總分科組距count70"); table.Columns.Add("加權總分校組距count70"); table.Columns.Add("加權平均班組距count70"); table.Columns.Add("加權平均科組距count70"); table.Columns.Add("加權平均校組距count70"); table.Columns.Add("類1總分組距count70"); table.Columns.Add("類1平均組距count70"); table.Columns.Add("類1加權總分組距count70"); table.Columns.Add("類1加權平均組距count70"); table.Columns.Add("類2總分組距count70"); table.Columns.Add("類2平均組距count70"); table.Columns.Add("類2加權總分組距count70"); table.Columns.Add("類2加權平均組距count70");
                table.Columns.Add("總分班組距count60"); table.Columns.Add("總分科組距count60"); table.Columns.Add("總分校組距count60"); table.Columns.Add("平均班組距count60"); table.Columns.Add("平均科組距count60"); table.Columns.Add("平均校組距count60"); table.Columns.Add("加權總分班組距count60"); table.Columns.Add("加權總分科組距count60"); table.Columns.Add("加權總分校組距count60"); table.Columns.Add("加權平均班組距count60"); table.Columns.Add("加權平均科組距count60"); table.Columns.Add("加權平均校組距count60"); table.Columns.Add("類1總分組距count60"); table.Columns.Add("類1平均組距count60"); table.Columns.Add("類1加權總分組距count60"); table.Columns.Add("類1加權平均組距count60"); table.Columns.Add("類2總分組距count60"); table.Columns.Add("類2平均組距count60"); table.Columns.Add("類2加權總分組距count60"); table.Columns.Add("類2加權平均組距count60");
                table.Columns.Add("總分班組距count50"); table.Columns.Add("總分科組距count50"); table.Columns.Add("總分校組距count50"); table.Columns.Add("平均班組距count50"); table.Columns.Add("平均科組距count50"); table.Columns.Add("平均校組距count50"); table.Columns.Add("加權總分班組距count50"); table.Columns.Add("加權總分科組距count50"); table.Columns.Add("加權總分校組距count50"); table.Columns.Add("加權平均班組距count50"); table.Columns.Add("加權平均科組距count50"); table.Columns.Add("加權平均校組距count50"); table.Columns.Add("類1總分組距count50"); table.Columns.Add("類1平均組距count50"); table.Columns.Add("類1加權總分組距count50"); table.Columns.Add("類1加權平均組距count50"); table.Columns.Add("類2總分組距count50"); table.Columns.Add("類2平均組距count50"); table.Columns.Add("類2加權總分組距count50"); table.Columns.Add("類2加權平均組距count50");
                table.Columns.Add("總分班組距count40"); table.Columns.Add("總分科組距count40"); table.Columns.Add("總分校組距count40"); table.Columns.Add("平均班組距count40"); table.Columns.Add("平均科組距count40"); table.Columns.Add("平均校組距count40"); table.Columns.Add("加權總分班組距count40"); table.Columns.Add("加權總分科組距count40"); table.Columns.Add("加權總分校組距count40"); table.Columns.Add("加權平均班組距count40"); table.Columns.Add("加權平均科組距count40"); table.Columns.Add("加權平均校組距count40"); table.Columns.Add("類1總分組距count40"); table.Columns.Add("類1平均組距count40"); table.Columns.Add("類1加權總分組距count40"); table.Columns.Add("類1加權平均組距count40"); table.Columns.Add("類2總分組距count40"); table.Columns.Add("類2平均組距count40"); table.Columns.Add("類2加權總分組距count40"); table.Columns.Add("類2加權平均組距count40");
                table.Columns.Add("總分班組距count30"); table.Columns.Add("總分科組距count30"); table.Columns.Add("總分校組距count30"); table.Columns.Add("平均班組距count30"); table.Columns.Add("平均科組距count30"); table.Columns.Add("平均校組距count30"); table.Columns.Add("加權總分班組距count30"); table.Columns.Add("加權總分科組距count30"); table.Columns.Add("加權總分校組距count30"); table.Columns.Add("加權平均班組距count30"); table.Columns.Add("加權平均科組距count30"); table.Columns.Add("加權平均校組距count30"); table.Columns.Add("類1總分組距count30"); table.Columns.Add("類1平均組距count30"); table.Columns.Add("類1加權總分組距count30"); table.Columns.Add("類1加權平均組距count30"); table.Columns.Add("類2總分組距count30"); table.Columns.Add("類2平均組距count30"); table.Columns.Add("類2加權總分組距count30"); table.Columns.Add("類2加權平均組距count30");
                table.Columns.Add("總分班組距count20"); table.Columns.Add("總分科組距count20"); table.Columns.Add("總分校組距count20"); table.Columns.Add("平均班組距count20"); table.Columns.Add("平均科組距count20"); table.Columns.Add("平均校組距count20"); table.Columns.Add("加權總分班組距count20"); table.Columns.Add("加權總分科組距count20"); table.Columns.Add("加權總分校組距count20"); table.Columns.Add("加權平均班組距count20"); table.Columns.Add("加權平均科組距count20"); table.Columns.Add("加權平均校組距count20"); table.Columns.Add("類1總分組距count20"); table.Columns.Add("類1平均組距count20"); table.Columns.Add("類1加權總分組距count20"); table.Columns.Add("類1加權平均組距count20"); table.Columns.Add("類2總分組距count20"); table.Columns.Add("類2平均組距count20"); table.Columns.Add("類2加權總分組距count20"); table.Columns.Add("類2加權平均組距count20");
                table.Columns.Add("總分班組距count10"); table.Columns.Add("總分科組距count10"); table.Columns.Add("總分校組距count10"); table.Columns.Add("平均班組距count10"); table.Columns.Add("平均科組距count10"); table.Columns.Add("平均校組距count10"); table.Columns.Add("加權總分班組距count10"); table.Columns.Add("加權總分科組距count10"); table.Columns.Add("加權總分校組距count10"); table.Columns.Add("加權平均班組距count10"); table.Columns.Add("加權平均科組距count10"); table.Columns.Add("加權平均校組距count10"); table.Columns.Add("類1總分組距count10"); table.Columns.Add("類1平均組距count10"); table.Columns.Add("類1加權總分組距count10"); table.Columns.Add("類1加權平均組距count10"); table.Columns.Add("類2總分組距count10"); table.Columns.Add("類2平均組距count10"); table.Columns.Add("類2加權總分組距count10"); table.Columns.Add("類2加權平均組距count10");
                table.Columns.Add("總分班組距count100Up"); table.Columns.Add("總分科組距count100Up"); table.Columns.Add("總分校組距count100Up"); table.Columns.Add("平均班組距count100Up"); table.Columns.Add("平均科組距count100Up"); table.Columns.Add("平均校組距count100Up"); table.Columns.Add("加權總分班組距count100Up"); table.Columns.Add("加權總分科組距count100Up"); table.Columns.Add("加權總分校組距count100Up"); table.Columns.Add("加權平均班組距count100Up"); table.Columns.Add("加權平均科組距count100Up"); table.Columns.Add("加權平均校組距count100Up"); table.Columns.Add("類1總分組距count100Up"); table.Columns.Add("類1平均組距count100Up"); table.Columns.Add("類1加權總分組距count100Up"); table.Columns.Add("類1加權平均組距count100Up"); table.Columns.Add("類2總分組距count100Up"); table.Columns.Add("類2平均組距count100Up"); table.Columns.Add("類2加權總分組距count100Up"); table.Columns.Add("類2加權平均組距count100Up");
                table.Columns.Add("總分班組距count90Up"); table.Columns.Add("總分科組距count90Up"); table.Columns.Add("總分校組距count90Up"); table.Columns.Add("平均班組距count90Up"); table.Columns.Add("平均科組距count90Up"); table.Columns.Add("平均校組距count90Up"); table.Columns.Add("加權總分班組距count90Up"); table.Columns.Add("加權總分科組距count90Up"); table.Columns.Add("加權總分校組距count90Up"); table.Columns.Add("加權平均班組距count90Up"); table.Columns.Add("加權平均科組距count90Up"); table.Columns.Add("加權平均校組距count90Up"); table.Columns.Add("類1總分組距count90Up"); table.Columns.Add("類1平均組距count90Up"); table.Columns.Add("類1加權總分組距count90Up"); table.Columns.Add("類1加權平均組距count90Up"); table.Columns.Add("類2總分組距count90Up"); table.Columns.Add("類2平均組距count90Up"); table.Columns.Add("類2加權總分組距count90Up"); table.Columns.Add("類2加權平均組距count90Up");
                table.Columns.Add("總分班組距count80Up"); table.Columns.Add("總分科組距count80Up"); table.Columns.Add("總分校組距count80Up"); table.Columns.Add("平均班組距count80Up"); table.Columns.Add("平均科組距count80Up"); table.Columns.Add("平均校組距count80Up"); table.Columns.Add("加權總分班組距count80Up"); table.Columns.Add("加權總分科組距count80Up"); table.Columns.Add("加權總分校組距count80Up"); table.Columns.Add("加權平均班組距count80Up"); table.Columns.Add("加權平均科組距count80Up"); table.Columns.Add("加權平均校組距count80Up"); table.Columns.Add("類1總分組距count80Up"); table.Columns.Add("類1平均組距count80Up"); table.Columns.Add("類1加權總分組距count80Up"); table.Columns.Add("類1加權平均組距count80Up"); table.Columns.Add("類2總分組距count80Up"); table.Columns.Add("類2平均組距count80Up"); table.Columns.Add("類2加權總分組距count80Up"); table.Columns.Add("類2加權平均組距count80Up");
                table.Columns.Add("總分班組距count70Up"); table.Columns.Add("總分科組距count70Up"); table.Columns.Add("總分校組距count70Up"); table.Columns.Add("平均班組距count70Up"); table.Columns.Add("平均科組距count70Up"); table.Columns.Add("平均校組距count70Up"); table.Columns.Add("加權總分班組距count70Up"); table.Columns.Add("加權總分科組距count70Up"); table.Columns.Add("加權總分校組距count70Up"); table.Columns.Add("加權平均班組距count70Up"); table.Columns.Add("加權平均科組距count70Up"); table.Columns.Add("加權平均校組距count70Up"); table.Columns.Add("類1總分組距count70Up"); table.Columns.Add("類1平均組距count70Up"); table.Columns.Add("類1加權總分組距count70Up"); table.Columns.Add("類1加權平均組距count70Up"); table.Columns.Add("類2總分組距count70Up"); table.Columns.Add("類2平均組距count70Up"); table.Columns.Add("類2加權總分組距count70Up"); table.Columns.Add("類2加權平均組距count70Up");
                table.Columns.Add("總分班組距count60Up"); table.Columns.Add("總分科組距count60Up"); table.Columns.Add("總分校組距count60Up"); table.Columns.Add("平均班組距count60Up"); table.Columns.Add("平均科組距count60Up"); table.Columns.Add("平均校組距count60Up"); table.Columns.Add("加權總分班組距count60Up"); table.Columns.Add("加權總分科組距count60Up"); table.Columns.Add("加權總分校組距count60Up"); table.Columns.Add("加權平均班組距count60Up"); table.Columns.Add("加權平均科組距count60Up"); table.Columns.Add("加權平均校組距count60Up"); table.Columns.Add("類1總分組距count60Up"); table.Columns.Add("類1平均組距count60Up"); table.Columns.Add("類1加權總分組距count60Up"); table.Columns.Add("類1加權平均組距count60Up"); table.Columns.Add("類2總分組距count60Up"); table.Columns.Add("類2平均組距count60Up"); table.Columns.Add("類2加權總分組距count60Up"); table.Columns.Add("類2加權平均組距count60Up");
                table.Columns.Add("總分班組距count50Up"); table.Columns.Add("總分科組距count50Up"); table.Columns.Add("總分校組距count50Up"); table.Columns.Add("平均班組距count50Up"); table.Columns.Add("平均科組距count50Up"); table.Columns.Add("平均校組距count50Up"); table.Columns.Add("加權總分班組距count50Up"); table.Columns.Add("加權總分科組距count50Up"); table.Columns.Add("加權總分校組距count50Up"); table.Columns.Add("加權平均班組距count50Up"); table.Columns.Add("加權平均科組距count50Up"); table.Columns.Add("加權平均校組距count50Up"); table.Columns.Add("類1總分組距count50Up"); table.Columns.Add("類1平均組距count50Up"); table.Columns.Add("類1加權總分組距count50Up"); table.Columns.Add("類1加權平均組距count50Up"); table.Columns.Add("類2總分組距count50Up"); table.Columns.Add("類2平均組距count50Up"); table.Columns.Add("類2加權總分組距count50Up"); table.Columns.Add("類2加權平均組距count50Up");
                table.Columns.Add("總分班組距count40Up"); table.Columns.Add("總分科組距count40Up"); table.Columns.Add("總分校組距count40Up"); table.Columns.Add("平均班組距count40Up"); table.Columns.Add("平均科組距count40Up"); table.Columns.Add("平均校組距count40Up"); table.Columns.Add("加權總分班組距count40Up"); table.Columns.Add("加權總分科組距count40Up"); table.Columns.Add("加權總分校組距count40Up"); table.Columns.Add("加權平均班組距count40Up"); table.Columns.Add("加權平均科組距count40Up"); table.Columns.Add("加權平均校組距count40Up"); table.Columns.Add("類1總分組距count40Up"); table.Columns.Add("類1平均組距count40Up"); table.Columns.Add("類1加權總分組距count40Up"); table.Columns.Add("類1加權平均組距count40Up"); table.Columns.Add("類2總分組距count40Up"); table.Columns.Add("類2平均組距count40Up"); table.Columns.Add("類2加權總分組距count40Up"); table.Columns.Add("類2加權平均組距count40Up");
                table.Columns.Add("總分班組距count30Up"); table.Columns.Add("總分科組距count30Up"); table.Columns.Add("總分校組距count30Up"); table.Columns.Add("平均班組距count30Up"); table.Columns.Add("平均科組距count30Up"); table.Columns.Add("平均校組距count30Up"); table.Columns.Add("加權總分班組距count30Up"); table.Columns.Add("加權總分科組距count30Up"); table.Columns.Add("加權總分校組距count30Up"); table.Columns.Add("加權平均班組距count30Up"); table.Columns.Add("加權平均科組距count30Up"); table.Columns.Add("加權平均校組距count30Up"); table.Columns.Add("類1總分組距count30Up"); table.Columns.Add("類1平均組距count30Up"); table.Columns.Add("類1加權總分組距count30Up"); table.Columns.Add("類1加權平均組距count30Up"); table.Columns.Add("類2總分組距count30Up"); table.Columns.Add("類2平均組距count30Up"); table.Columns.Add("類2加權總分組距count30Up"); table.Columns.Add("類2加權平均組距count30Up");
                table.Columns.Add("總分班組距count20Up"); table.Columns.Add("總分科組距count20Up"); table.Columns.Add("總分校組距count20Up"); table.Columns.Add("平均班組距count20Up"); table.Columns.Add("平均科組距count20Up"); table.Columns.Add("平均校組距count20Up"); table.Columns.Add("加權總分班組距count20Up"); table.Columns.Add("加權總分科組距count20Up"); table.Columns.Add("加權總分校組距count20Up"); table.Columns.Add("加權平均班組距count20Up"); table.Columns.Add("加權平均科組距count20Up"); table.Columns.Add("加權平均校組距count20Up"); table.Columns.Add("類1總分組距count20Up"); table.Columns.Add("類1平均組距count20Up"); table.Columns.Add("類1加權總分組距count20Up"); table.Columns.Add("類1加權平均組距count20Up"); table.Columns.Add("類2總分組距count20Up"); table.Columns.Add("類2平均組距count20Up"); table.Columns.Add("類2加權總分組距count20Up"); table.Columns.Add("類2加權平均組距count20Up");
                table.Columns.Add("總分班組距count10Up"); table.Columns.Add("總分科組距count10Up"); table.Columns.Add("總分校組距count10Up"); table.Columns.Add("平均班組距count10Up"); table.Columns.Add("平均科組距count10Up"); table.Columns.Add("平均校組距count10Up"); table.Columns.Add("加權總分班組距count10Up"); table.Columns.Add("加權總分科組距count10Up"); table.Columns.Add("加權總分校組距count10Up"); table.Columns.Add("加權平均班組距count10Up"); table.Columns.Add("加權平均科組距count10Up"); table.Columns.Add("加權平均校組距count10Up"); table.Columns.Add("類1總分組距count10Up"); table.Columns.Add("類1平均組距count10Up"); table.Columns.Add("類1加權總分組距count10Up"); table.Columns.Add("類1加權平均組距count10Up"); table.Columns.Add("類2總分組距count10Up"); table.Columns.Add("類2平均組距count10Up"); table.Columns.Add("類2加權總分組距count10Up"); table.Columns.Add("類2加權平均組距count10Up");
                table.Columns.Add("總分班組距count90Down"); table.Columns.Add("總分科組距count90Down"); table.Columns.Add("總分校組距count90Down"); table.Columns.Add("平均班組距count90Down"); table.Columns.Add("平均科組距count90Down"); table.Columns.Add("平均校組距count90Down"); table.Columns.Add("加權總分班組距count90Down"); table.Columns.Add("加權總分科組距count90Down"); table.Columns.Add("加權總分校組距count90Down"); table.Columns.Add("加權平均班組距count90Down"); table.Columns.Add("加權平均科組距count90Down"); table.Columns.Add("加權平均校組距count90Down"); table.Columns.Add("類1總分組距count90Down"); table.Columns.Add("類1平均組距count90Down"); table.Columns.Add("類1加權總分組距count90Down"); table.Columns.Add("類1加權平均組距count90Down"); table.Columns.Add("類2總分組距count90Down"); table.Columns.Add("類2平均組距count90Down"); table.Columns.Add("類2加權總分組距count90Down"); table.Columns.Add("類2加權平均組距count90Down");
                table.Columns.Add("總分班組距count80Down"); table.Columns.Add("總分科組距count80Down"); table.Columns.Add("總分校組距count80Down"); table.Columns.Add("平均班組距count80Down"); table.Columns.Add("平均科組距count80Down"); table.Columns.Add("平均校組距count80Down"); table.Columns.Add("加權總分班組距count80Down"); table.Columns.Add("加權總分科組距count80Down"); table.Columns.Add("加權總分校組距count80Down"); table.Columns.Add("加權平均班組距count80Down"); table.Columns.Add("加權平均科組距count80Down"); table.Columns.Add("加權平均校組距count80Down"); table.Columns.Add("類1總分組距count80Down"); table.Columns.Add("類1平均組距count80Down"); table.Columns.Add("類1加權總分組距count80Down"); table.Columns.Add("類1加權平均組距count80Down"); table.Columns.Add("類2總分組距count80Down"); table.Columns.Add("類2平均組距count80Down"); table.Columns.Add("類2加權總分組距count80Down"); table.Columns.Add("類2加權平均組距count80Down");
                table.Columns.Add("總分班組距count70Down"); table.Columns.Add("總分科組距count70Down"); table.Columns.Add("總分校組距count70Down"); table.Columns.Add("平均班組距count70Down"); table.Columns.Add("平均科組距count70Down"); table.Columns.Add("平均校組距count70Down"); table.Columns.Add("加權總分班組距count70Down"); table.Columns.Add("加權總分科組距count70Down"); table.Columns.Add("加權總分校組距count70Down"); table.Columns.Add("加權平均班組距count70Down"); table.Columns.Add("加權平均科組距count70Down"); table.Columns.Add("加權平均校組距count70Down"); table.Columns.Add("類1總分組距count70Down"); table.Columns.Add("類1平均組距count70Down"); table.Columns.Add("類1加權總分組距count70Down"); table.Columns.Add("類1加權平均組距count70Down"); table.Columns.Add("類2總分組距count70Down"); table.Columns.Add("類2平均組距count70Down"); table.Columns.Add("類2加權總分組距count70Down"); table.Columns.Add("類2加權平均組距count70Down");
                table.Columns.Add("總分班組距count60Down"); table.Columns.Add("總分科組距count60Down"); table.Columns.Add("總分校組距count60Down"); table.Columns.Add("平均班組距count60Down"); table.Columns.Add("平均科組距count60Down"); table.Columns.Add("平均校組距count60Down"); table.Columns.Add("加權總分班組距count60Down"); table.Columns.Add("加權總分科組距count60Down"); table.Columns.Add("加權總分校組距count60Down"); table.Columns.Add("加權平均班組距count60Down"); table.Columns.Add("加權平均科組距count60Down"); table.Columns.Add("加權平均校組距count60Down"); table.Columns.Add("類1總分組距count60Down"); table.Columns.Add("類1平均組距count60Down"); table.Columns.Add("類1加權總分組距count60Down"); table.Columns.Add("類1加權平均組距count60Down"); table.Columns.Add("類2總分組距count60Down"); table.Columns.Add("類2平均組距count60Down"); table.Columns.Add("類2加權總分組距count60Down"); table.Columns.Add("類2加權平均組距count60Down");
                table.Columns.Add("總分班組距count50Down"); table.Columns.Add("總分科組距count50Down"); table.Columns.Add("總分校組距count50Down"); table.Columns.Add("平均班組距count50Down"); table.Columns.Add("平均科組距count50Down"); table.Columns.Add("平均校組距count50Down"); table.Columns.Add("加權總分班組距count50Down"); table.Columns.Add("加權總分科組距count50Down"); table.Columns.Add("加權總分校組距count50Down"); table.Columns.Add("加權平均班組距count50Down"); table.Columns.Add("加權平均科組距count50Down"); table.Columns.Add("加權平均校組距count50Down"); table.Columns.Add("類1總分組距count50Down"); table.Columns.Add("類1平均組距count50Down"); table.Columns.Add("類1加權總分組距count50Down"); table.Columns.Add("類1加權平均組距count50Down"); table.Columns.Add("類2總分組距count50Down"); table.Columns.Add("類2平均組距count50Down"); table.Columns.Add("類2加權總分組距count50Down"); table.Columns.Add("類2加權平均組距count50Down");
                table.Columns.Add("總分班組距count40Down"); table.Columns.Add("總分科組距count40Down"); table.Columns.Add("總分校組距count40Down"); table.Columns.Add("平均班組距count40Down"); table.Columns.Add("平均科組距count40Down"); table.Columns.Add("平均校組距count40Down"); table.Columns.Add("加權總分班組距count40Down"); table.Columns.Add("加權總分科組距count40Down"); table.Columns.Add("加權總分校組距count40Down"); table.Columns.Add("加權平均班組距count40Down"); table.Columns.Add("加權平均科組距count40Down"); table.Columns.Add("加權平均校組距count40Down"); table.Columns.Add("類1總分組距count40Down"); table.Columns.Add("類1平均組距count40Down"); table.Columns.Add("類1加權總分組距count40Down"); table.Columns.Add("類1加權平均組距count40Down"); table.Columns.Add("類2總分組距count40Down"); table.Columns.Add("類2平均組距count40Down"); table.Columns.Add("類2加權總分組距count40Down"); table.Columns.Add("類2加權平均組距count40Down");
                table.Columns.Add("總分班組距count30Down"); table.Columns.Add("總分科組距count30Down"); table.Columns.Add("總分校組距count30Down"); table.Columns.Add("平均班組距count30Down"); table.Columns.Add("平均科組距count30Down"); table.Columns.Add("平均校組距count30Down"); table.Columns.Add("加權總分班組距count30Down"); table.Columns.Add("加權總分科組距count30Down"); table.Columns.Add("加權總分校組距count30Down"); table.Columns.Add("加權平均班組距count30Down"); table.Columns.Add("加權平均科組距count30Down"); table.Columns.Add("加權平均校組距count30Down"); table.Columns.Add("類1總分組距count30Down"); table.Columns.Add("類1平均組距count30Down"); table.Columns.Add("類1加權總分組距count30Down"); table.Columns.Add("類1加權平均組距count30Down"); table.Columns.Add("類2總分組距count30Down"); table.Columns.Add("類2平均組距count30Down"); table.Columns.Add("類2加權總分組距count30Down"); table.Columns.Add("類2加權平均組距count30Down");
                table.Columns.Add("總分班組距count20Down"); table.Columns.Add("總分科組距count20Down"); table.Columns.Add("總分校組距count20Down"); table.Columns.Add("平均班組距count20Down"); table.Columns.Add("平均科組距count20Down"); table.Columns.Add("平均校組距count20Down"); table.Columns.Add("加權總分班組距count20Down"); table.Columns.Add("加權總分科組距count20Down"); table.Columns.Add("加權總分校組距count20Down"); table.Columns.Add("加權平均班組距count20Down"); table.Columns.Add("加權平均科組距count20Down"); table.Columns.Add("加權平均校組距count20Down"); table.Columns.Add("類1總分組距count20Down"); table.Columns.Add("類1平均組距count20Down"); table.Columns.Add("類1加權總分組距count20Down"); table.Columns.Add("類1加權平均組距count20Down"); table.Columns.Add("類2總分組距count20Down"); table.Columns.Add("類2平均組距count20Down"); table.Columns.Add("類2加權總分組距count20Down"); table.Columns.Add("類2加權平均組距count20Down");
                table.Columns.Add("總分班組距count10Down"); table.Columns.Add("總分科組距count10Down"); table.Columns.Add("總分校組距count10Down"); table.Columns.Add("平均班組距count10Down"); table.Columns.Add("平均科組距count10Down"); table.Columns.Add("平均校組距count10Down"); table.Columns.Add("加權總分班組距count10Down"); table.Columns.Add("加權總分科組距count10Down"); table.Columns.Add("加權總分校組距count10Down"); table.Columns.Add("加權平均班組距count10Down"); table.Columns.Add("加權平均科組距count10Down"); table.Columns.Add("加權平均校組距count10Down"); table.Columns.Add("類1總分組距count10Down"); table.Columns.Add("類1平均組距count10Down"); table.Columns.Add("類1加權總分組距count10Down"); table.Columns.Add("類1加權平均組距count10Down"); table.Columns.Add("類2總分組距count10Down"); table.Columns.Add("類2平均組距count10Down"); table.Columns.Add("類2加權總分組距count10Down"); table.Columns.Add("類2加權平均組距count10Down");
                #endregion
                #endregion
                //宣告產生的報表
                Aspose.Words.Document document = new Aspose.Words.Document();
                //用一個BackgroundWorker包起來
                System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                bkw.WorkerReportsProgress = true;
                System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 期末成績單產生 S");
                bkw.ProgressChanged += delegate(object sender, System.ComponentModel.ProgressChangedEventArgs e)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("期末成績單產生中", e.ProgressPercentage);
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 期末成績單產生 " + e.ProgressPercentage);
                };
                Exception exc = null;


                bkw.RunWorkerCompleted += delegate(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
                {
                    if (exc != null)
                    {
                        throw exc;
                    }

                    Workbook workbook = e.Result as Workbook;

                    if (workbook == null)
                        return;

                    ImportWorkbook(workbook);

                    // 以後記得存Excel 都用新版的Xlsx，可以避免ㄧ些不必要的問題(EX: sheet 只能到1023張)
                    SaveFileDialog save = new SaveFileDialog();
                    save.Title = "另存新檔";
                    save.FileName = conf.FileName;
                    save.Filter = "Excel檔案 (*.Xlsx)|*.Xlsx|所有檔案 (*.*)|*.*";

                    if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            workbook.Save(save.FileName, Aspose.Cells.SaveFormat.Xlsx);
                            System.Diagnostics.Process.Start(save.FileName);


                        }
                        catch
                        {
                            MessageBox.Show("檔案儲存失敗");


                        }
                    }
                };


                bkw.DoWork += delegate(object sender, System.ComponentModel.DoWorkEventArgs e)
                {

                    // 為列印Excel 先New 物件
                    Workbook wb = new Workbook(new MemoryStream(Properties.Resources.固定排名資料明細));

                    Cells cs0 = wb.Worksheets[0].Cells;
                    Cells cs1 = wb.Worksheets[1].Cells;
                    Cells cs2 = wb.Worksheets[2].Cells;


                    var studentRecords = accessHelper.StudentHelper.GetStudents(selectedStudents);
                    Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>> studentExamSores = new Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>>();
                    Dictionary<string, Dictionary<string, ExamScoreInfo>> studentRefExamSores = new Dictionary<string, Dictionary<string, ExamScoreInfo>>();
                    ManualResetEvent scoreReady = new ManualResetEvent(false);
                    ManualResetEvent elseReady = new ManualResetEvent(false);
                    #region 偷跑取得考試成績
                    // 有成績科目名稱對照
                    new Thread(new ThreadStart(delegate
                    {
                        // 取得學生學期科目成績
                        int sSchoolYear, sSemester;
                        int.TryParse(conf.SchoolYear, out sSchoolYear);
                        int.TryParse(conf.Semester, out sSemester);
                        #region 整理學生定期評量成績
                        #region 篩選課程學年度、學期、科目取得有可能有需要的資料
                        List<CourseRecord> targetCourseList = new List<CourseRecord>();
                        try
                        {
                            foreach (var courseRecord in accessHelper.CourseHelper.GetAllCourse(sSchoolYear, sSemester))
                            {
                                //用科目濾出可能有用到的課程
                                if (conf.PrintSubjectList.Contains(courseRecord.Subject)
                                    || conf.TagRank1SubjectList.Contains(courseRecord.Subject)
                                    || conf.TagRank2SubjectList.Contains(courseRecord.Subject))
                                    targetCourseList.Add(courseRecord);
                            }
                        }
                        catch (Exception exception)
                        {
                            exc = exception;
                        }
                        #endregion
                        try
                        {
                            if (conf.ExamRecord != null || conf.RefenceExamRecord != null)
                            {
                                accessHelper.CourseHelper.FillExam(targetCourseList);
                                var tcList = new List<CourseRecord>();
                                var totalList = new List<CourseRecord>();
                                foreach (var courseRec in targetCourseList)
                                {
                                    if (conf.ExamRecord != null && courseRec.ExamList.Contains(conf.ExamRecord.Name))
                                    {
                                        tcList.Add(courseRec);
                                        totalList.Add(courseRec);
                                    }
                                    if (tcList.Count == 180)
                                    {
                                        accessHelper.CourseHelper.FillStudentAttend(tcList);
                                        accessHelper.CourseHelper.FillExamScore(tcList);
                                        tcList.Clear();
                                    }
                                }
                                accessHelper.CourseHelper.FillStudentAttend(tcList);
                                accessHelper.CourseHelper.FillExamScore(tcList);
                                foreach (var courseRecord in totalList)
                                {
                                    #region 整理本次定期評量成績
                                    if (conf.ExamRecord != null && courseRecord.ExamList.Contains(conf.ExamRecord.Name))
                                    {
                                        foreach (var attendStudent in courseRecord.StudentAttendList)
                                        {
                                            if (!studentExamSores.ContainsKey(attendStudent.StudentID)) studentExamSores.Add(attendStudent.StudentID, new Dictionary<string, Dictionary<string, ExamScoreInfo>>());
                                            if (!studentExamSores[attendStudent.StudentID].ContainsKey(courseRecord.Subject)) studentExamSores[attendStudent.StudentID].Add(courseRecord.Subject, new Dictionary<string, ExamScoreInfo>());
                                            studentExamSores[attendStudent.StudentID][courseRecord.Subject].Add("" + attendStudent.CourseID, null);
                                        }
                                        foreach (var examScoreRec in courseRecord.ExamScoreList)
                                        {
                                            if (examScoreRec.ExamName == conf.ExamRecord.Name)
                                            {
                                                studentExamSores[examScoreRec.StudentID][courseRecord.Subject]["" + examScoreRec.CourseID] = examScoreRec;
                                            }
                                        }
                                    }
                                    #endregion


                                    // 2016/6/27 穎驊:不再需要參考試別，把他給註解掉
                                    #region 整理前次定期評量成績
                                    //if (conf.RefenceExamRecord != null && courseRecord.ExamList.Contains(conf.RefenceExamRecord.Name))
                                    //{
                                    //    foreach (var examScoreRec in courseRecord.ExamScoreList)
                                    //    {
                                    //        if (examScoreRec.ExamName == conf.RefenceExamRecord.Name)
                                    //        {
                                    //            if (!studentRefExamSores.ContainsKey(examScoreRec.StudentID))
                                    //                studentRefExamSores.Add(examScoreRec.StudentID, new Dictionary<string, ExamScoreInfo>());
                                    //            studentRefExamSores[examScoreRec.StudentID].Add("" + examScoreRec.CourseID, examScoreRec);
                                    //        }
                                    //    }
                                    //}
                                    #endregion


                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            exc = exception;
                        }
                        finally
                        {
                            scoreReady.Set();
                        }
                        #endregion
                        #region 整理學生學期、學年成績
                        try
                        {
                            if (sSemester == 2 && conf.WithSchoolYearScore)
                            {
                                accessHelper.StudentHelper.FillSchoolYearEntryScore(true, studentRecords);
                                accessHelper.StudentHelper.FillSchoolYearSubjectScore(true, studentRecords);
                            }
                            accessHelper.StudentHelper.FillSemesterEntryScore(true, studentRecords);
                            accessHelper.StudentHelper.FillSemesterSubjectScore(true, studentRecords);
                            accessHelper.StudentHelper.FillSemesterMoralScore(true, studentRecords);
                            //accessHelper.StudentHelper.FillField("SemesterEntryClassRating", studentRecords);
                            accessHelper.StudentHelper.FillField("SchoolYearEntryClassRating", studentRecords);

                            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                            string sidList = "";
                            Dictionary<string, StudentRecord> stuDictionary = new Dictionary<string, StudentRecord>();
                            foreach (var stuRec in studentRecords)
                            {
                                sidList += (sidList == "" ? "" : ",") + stuRec.StudentID;
                                stuDictionary.Add(stuRec.StudentID, stuRec);
                            }
                            FISCA.Data.QueryHelper qh = new FISCA.Data.QueryHelper();
                            #region 學期學業成績排名
                            string strSQL = "select * from sems_entry_score where ref_student_id in (" + sidList + ") and school_year=" + sSchoolYear + " and semester=" + sSemester + "";
                            System.Data.DataTable dt = qh.Select(strSQL);
                            foreach (System.Data.DataRow dr in dt.Rows)
                            {
                                if ("" + dr["entry_group"] != "1") continue;
                                StudentRecord rec = stuDictionary["" + dr["ref_student_id"]];
                                if ("" + dr["class_rating"] != "")
                                {
                                    //學期學業成績班排名
                                    doc.LoadXml("" + dr["class_rating"]);
                                    System.Xml.XmlElement ele = (System.Xml.XmlElement)doc.SelectSingleNode("Rating/Item[@分項='學業']");
                                    if (ele != null)
                                    {
                                        //<Item 分項="學業" 成績="90.6" 成績人數="35" 排名="2"/>
                                        rec.Fields.Add("學期學業成績班排名", ele.GetAttribute("排名"));
                                        rec.Fields.Add("學期學業成績班排名母數", ele.GetAttribute("成績人數"));
                                    }
                                }
                                //學期學業成績科排名

                                if ("" + dr["dept_rating"] != "")
                                {
                                    doc.LoadXml("" + dr["dept_rating"]);
                                    System.Xml.XmlElement ele = (System.Xml.XmlElement)doc.SelectSingleNode("Rating/Item[@分項='學業']");
                                    if (ele != null)
                                    {
                                        //<Item 分項="學業" 成績="90.6" 成績人數="35" 排名="2"/>
                                        rec.Fields.Add("學期學業成績科排名", ele.GetAttribute("排名"));
                                        rec.Fields.Add("學期學業成績科排名母數", ele.GetAttribute("成績人數"));
                                    }
                                }
                                //學期學業成績類別1排名
                                //學期學業成績類別2排名

                                if ("" + dr["group_rating"] != "")
                                {
                                    doc.LoadXml("" + dr["group_rating"]);
                                    foreach (System.Xml.XmlElement element in doc.SelectNodes("Ratings/Rating"))
                                    {
                                        System.Xml.XmlElement ele = (System.Xml.XmlElement)element.SelectSingleNode("Item[@分項='學業']");
                                        if (ele != null)
                                        {
                                            if (!rec.Fields.ContainsKey("學期學業成績類別1"))
                                            {
                                                rec.Fields.Add("學期學業成績類別1", element.GetAttribute("類別"));
                                                if (!rec.Fields.ContainsKey("學期學業成績" + element.GetAttribute("類別") + "排名"))
                                                {
                                                    rec.Fields.Add("學期學業成績" + element.GetAttribute("類別") + "排名", ele.GetAttribute("排名"));
                                                    rec.Fields.Add("學期學業成績" + element.GetAttribute("類別") + "排名母數", ele.GetAttribute("成績人數"));
                                                }
                                            }
                                            else
                                            {
                                                rec.Fields.Add("學期學業成績類別2", element.GetAttribute("類別"));
                                                if (!rec.Fields.ContainsKey("學期學業成績" + element.GetAttribute("類別") + "排名"))
                                                {
                                                    rec.Fields.Add("學期學業成績" + element.GetAttribute("類別") + "排名", ele.GetAttribute("排名"));
                                                    rec.Fields.Add("學期學業成績" + element.GetAttribute("類別") + "排名母數", ele.GetAttribute("成績人數"));
                                                }
                                            }
                                        }
                                    }
                                }
                                //學期學業成績校排名
                                if ("" + dr["year_rating"] != "")
                                {
                                    doc.LoadXml("" + dr["year_rating"]);
                                    System.Xml.XmlElement ele = (System.Xml.XmlElement)doc.SelectSingleNode("Rating/Item[@分項='學業']");
                                    if (ele != null)
                                    {
                                        //<Item 分項="學業" 成績="90.6" 成績人數="35" 排名="2"/>
                                        rec.Fields.Add("學期學業成績校排名", ele.GetAttribute("排名"));
                                        rec.Fields.Add("學期學業成績校排名母數", ele.GetAttribute("成績人數"));
                                    }
                                }
                            }
                            #endregion
                            #region 學期科目成績排名
                            strSQL = "select * from sems_subj_score where ref_student_id in (" + sidList + ") and school_year=" + sSchoolYear + " and semester=" + sSemester + "";
                            dt = qh.Select(strSQL);
                            foreach (System.Data.DataRow dr in dt.Rows)
                            {
                                StudentRecord rec = stuDictionary["" + dr["ref_student_id"]];
                                //學期學業成績班排名
                                if ("" + dr["class_rating"] != "")
                                {
                                    doc.LoadXml("" + dr["class_rating"]);
                                    foreach (System.Xml.XmlElement ele in doc.SelectNodes("Rating/Item"))
                                    {
                                        //<Item 成績="83" 成績人數="50" 排名="33" 科目="公民與社會" 科目級別="1"/>
                                        rec.Fields.Add("學期科目排名成績" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("成績"));
                                        rec.Fields.Add("學期科目班排名" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("排名"));
                                        rec.Fields.Add("學期科目班排名母數" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("成績人數"));
                                    }
                                }
                                //學期學業成績科排名
                                if ("" + dr["dept_rating"] != "")
                                {
                                    doc.LoadXml("" + dr["dept_rating"]);
                                    foreach (System.Xml.XmlElement ele in doc.SelectNodes("Rating/Item"))
                                    {
                                        //<Item 分項="學業" 成績="90.6" 成績人數="35" 排名="2"/>
                                        rec.Fields.Add("學期科目科排名" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("排名"));
                                        rec.Fields.Add("學期科目科排名母數" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("成績人數"));
                                    }
                                }
                                //學期學業成績類別1排名
                                //學期學業成績類別2排名
                                if ("" + dr["group_rating"] != "")
                                {
                                    doc.LoadXml("" + dr["group_rating"]);
                                    foreach (System.Xml.XmlElement element in doc.SelectNodes("Ratings/Rating"))
                                    {
                                        string cat = element.GetAttribute("類別");
                                        if (!rec.Fields.ContainsKey("學期科目成績類別1"))
                                        {
                                            rec.Fields.Add("學期科目成績類別1", cat);
                                            foreach (System.Xml.XmlElement ele in element.SelectNodes("Item"))
                                            {
                                                if (!rec.Fields.ContainsKey("學期科目成績" + cat + "排名" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別")))
                                                {
                                                    rec.Fields.Add("學期科目成績" + cat + "排名" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("排名"));
                                                    rec.Fields.Add("學期科目成績" + cat + "排名母數" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("成績人數"));
                                                }
                                            }
                                        }
                                        else
                                        {
                                            rec.Fields.Add("學期科目成績類別2", cat);
                                            foreach (System.Xml.XmlElement ele in element.SelectNodes("Item"))
                                            {
                                                if (!rec.Fields.ContainsKey("學期科目成績" + cat + "排名" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別")))
                                                {
                                                    rec.Fields.Add("學期科目成績" + cat + "排名" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("排名"));
                                                    rec.Fields.Add("學期科目成績" + cat + "排名母數" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("成績人數"));
                                                }
                                            }
                                        }
                                    }
                                }
                                //學期學業成績校排名
                                if ("" + dr["year_rating"] != "")
                                {
                                    doc.LoadXml("" + dr["year_rating"]);
                                    foreach (System.Xml.XmlElement ele in doc.SelectNodes("Rating/Item"))
                                    {
                                        //<Item 分項="學業" 成績="90.6" 成績人數="35" 排名="2"/>
                                        rec.Fields.Add("學期科目校排名" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("排名"));
                                        rec.Fields.Add("學期科目校排名母數" + ele.GetAttribute("科目") + "^^^" + ele.GetAttribute("科目級別"), ele.GetAttribute("成績人數"));
                                    }
                                }
                            }
                            #endregion
                            accessHelper.StudentHelper.FillAttendance(studentRecords);
                            accessHelper.StudentHelper.FillReward(studentRecords);
                        }
                        catch (Exception exception)
                        {
                            exc = exception;
                        }
                        finally
                        {
                            elseReady.Set();
                        }
                        #endregion
                    })).Start();
                    #endregion
                    try
                    {
                        string key = "";
                        bkw.ReportProgress(0);

                        #region 整理同年級學生
                        //整理選取學生的年級
                        Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();
                        foreach (var studentRec in studentRecords)
                        {
                            string grade = "";
                            if (studentRec.RefClass != null)
                                grade = "" + studentRec.RefClass.GradeYear;
                            if (!gradeyearStudents.ContainsKey(grade))
                                gradeyearStudents.Add(grade, new List<StudentRecord>());
                            gradeyearStudents[grade].Add(studentRec);
                        }
                        foreach (var classRec in accessHelper.ClassHelper.GetAllClass())
                        {
                            if (gradeyearStudents.ContainsKey("" + classRec.GradeYear))
                            {
                                //用班級去取出可能有相關的學生
                                foreach (var studentRec in classRec.Students)
                                {
                                    string grade = "";
                                    if (studentRec.RefClass != null)
                                        grade = "" + studentRec.RefClass.GradeYear;
                                    if (!gradeyearStudents[grade].Contains(studentRec))
                                        gradeyearStudents[grade].Add(studentRec);
                                }
                            }
                        }
                        //List<string> gradeyearClasses = new List<string>();
                        //foreach (ClassRecord classRec in K12.Data.Class.SelectAll())
                        //{
                        //    if (gradeyearStudents.ContainsKey("" + classRec.GradeYear))
                        //    {
                        //        gradeyearClasses.Add(classRec.ID);
                        //    }
                        //}
                        //用班級去取出可能有相關的學生
                        //foreach (SHStudentRecord studentRec in SHStudent.SelectByClassIDs(gradeyearClasses))
                        //{
                        //    string grade = "";
                        //    if (studentRec.Class != null)
                        //        grade = "" + studentRec.Class.GradeYear;
                        //    if (!gradeyearStudents[grade].Contains(studentRec.ID))
                        //        gradeyearStudents[grade].Add(studentRec.ID);
                        //    if (!studentRecords.ContainsKey(studentRec.ID))
                        //        studentRecords.Add(studentRec.ID, studentRec);
                        //}
                        #endregion
                        bkw.ReportProgress(15);
                        #region 取得學生類別


                        Dictionary<string, List<K12.Data.StudentTagRecord>> studentTags = new Dictionary<string, List<K12.Data.StudentTagRecord>>();
                        List<string> list = new List<string>();
                        foreach (var sRecs in gradeyearStudents.Values)
                        {
                            foreach (var stuRec in sRecs)
                            {
                                list.Add(stuRec.StudentID);
                            }
                        }
                        foreach (var tag in K12.Data.StudentTag.SelectByStudentIDs(list))
                        {
                            if (!studentTags.ContainsKey(tag.RefStudentID))
                                studentTags.Add(tag.RefStudentID, new List<K12.Data.StudentTagRecord>());
                            studentTags[tag.RefStudentID].Add(tag);
                        }
                        #endregion
                        bkw.ReportProgress(20);
                        //等到成績載完
                        scoreReady.WaitOne();
                        bkw.ReportProgress(35);
                        int progressCount = 0;
                        #region 計算總分及各項目排名
                        Dictionary<string, string> studentTag1Group = new Dictionary<string, string>();
                        Dictionary<string, string> studentTag2Group = new Dictionary<string, string>();
                        Dictionary<string, bool> joinRank = new Dictionary<string, bool>();
                        Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                        Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
                        Dictionary<string, decimal> studentPrintSubjectSum = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectSum = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectSum = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentPrintSubjectAvg = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectAvg = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectAvg = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentPrintSubjectSumW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectSumW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectSumW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentPrintSubjectAvgW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectAvgW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectAvgW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> analytics = new Dictionary<string, decimal>();




                        // 2016/7/5 穎驊新增，回歸科目的計算，其為每一科子科目加權平均後的成績。


                        // 一般列印項目的回歸科目使用
                        Dictionary<string, decimal> studentRegressSubjectTotalW = new Dictionary<string, decimal>();

                        Dictionary<string, decimal> studentRegressSubjectTotalCredit = new Dictionary<string, decimal>();

                        Dictionary<string, decimal> RegressSubjectAvgW = new Dictionary<string, decimal>(); //這個算總分得後來發現暫時用不到

                        Dictionary<string, Dictionary<string, decimal>> studentIDRegressSubjectAvgW = new Dictionary<string, Dictionary<string, decimal>>();


                        // 類別一Tag1的回歸科目使用
                        Dictionary<string, decimal> studentRegressSubjectTotalW_Tag1 = new Dictionary<string, decimal>();

                        Dictionary<string, decimal> studentRegressSubjectTotalCredit_Tag1 = new Dictionary<string, decimal>();

                        Dictionary<string, Dictionary<string, decimal>> studentIDRegressSubjectAvgW_Tag1 = new Dictionary<string, Dictionary<string, decimal>>();


                        // 類別二Tag2的回歸科目使用
                        Dictionary<string, decimal> studentRegressSubjectTotalW_Tag2 = new Dictionary<string, decimal>();

                        Dictionary<string, decimal> studentRegressSubjectTotalCredit_Tag2 = new Dictionary<string, decimal>();

                        Dictionary<string, Dictionary<string, decimal>> studentIDRegressSubjectAvgW_Tag2 = new Dictionary<string, Dictionary<string, decimal>>();


                        int total = 0;
                        foreach (var gss in gradeyearStudents.Values)
                        {
                            total += gss.Count;
                        }
                        bkw.ReportProgress(40);
                        foreach (string gradeyear in gradeyearStudents.Keys)
                        {
                            //找出全年級學生
                            foreach (var studentRec in gradeyearStudents[gradeyear])
                            {
                                string studentID = studentRec.StudentID;
                                bool rank = true;
                                string tag1ID = "";
                                string tag2ID = "";
                                #region 分析學生所屬類別
                                if (studentTags.ContainsKey(studentID))
                                {
                                    foreach (var tag in studentTags[studentID])
                                    {
                                        #region 判斷學生是否屬於不排名類別
                                        if (conf.RankFilterTagList.Contains(tag.RefTagID))
                                        {
                                            rank = false;
                                        }
                                        #endregion
                                        #region 判斷學生在類別排名1中所屬的類別
                                        if (tag1ID == "" && conf.TagRank1TagList.Contains(tag.RefTagID))
                                        {
                                            tag1ID = tag.RefTagID;
                                            studentTag1Group.Add(studentID, tag1ID);
                                        }
                                        #endregion
                                        #region 判斷學生在類別排名2中所屬的類別
                                        if (tag2ID == "" && conf.TagRank2TagList.Contains(tag.RefTagID))
                                        {
                                            tag2ID = tag.RefTagID;
                                            studentTag2Group.Add(studentID, tag2ID);
                                        }
                                        #endregion
                                    }
                                }
                                #endregion
                                joinRank.Add(studentRec.StudentID, (rank && studentRec.Status == "一般"));
                                bool summaryRank = true;
                                bool tag1SummaryRank = true;
                                bool tag2SummaryRank = true;


                                //2016/7/5 穎驊新增，作為回歸科目計算使用
                                bool RegressRank = true;

                                if (studentExamSores.ContainsKey(studentID))
                                {
                                    decimal printSubjectSum = 0;
                                    int printSubjectCount = 0;
                                    decimal tag1SubjectSum = 0;
                                    int tag1SubjectCount = 0;
                                    decimal tag2SubjectSum = 0;
                                    int tag2SubjectCount = 0;
                                    decimal printSubjectSumW = 0;
                                    decimal printSubjectCreditSum = 0;
                                    decimal tag1SubjectSumW = 0;
                                    decimal tag1SubjectCreditSum = 0;
                                    decimal tag2SubjectSumW = 0;
                                    decimal tag2SubjectCreditSum = 0;


                                    //2016/7/5 穎驊新增，作為回歸科目計算使用

                                    decimal RegressSubjectSumW = 0;
                                    decimal RegressSubjectCreditSum = 0;



                                    // 穎驊筆記，每算完一個學生都要把計算用Dictionary清空，才不會有B學生的成績影響到A學生
                                    studentRegressSubjectTotalW.Clear();
                                    studentRegressSubjectTotalCredit.Clear();

                                    studentRegressSubjectTotalW_Tag1.Clear();
                                    studentRegressSubjectTotalCredit_Tag1.Clear();

                                    studentRegressSubjectTotalW_Tag2.Clear();
                                    studentRegressSubjectTotalCredit_Tag2.Clear();

                                    RegressSubjectAvgW.Clear();


                                    foreach (var subjectName in studentExamSores[studentID].Keys)
                                    {
                                        if (conf.PrintSubjectList.Contains(subjectName))
                                        {


                                            #region 是列印科目
                                            foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                            {
                                                if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                {
                                                    printSubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                    printSubjectCount++;
                                                    //計算加權總分
                                                    printSubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                    printSubjectCreditSum += sceTakeRecord.CreditDec();
                                                    if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                    {
                                                        if (sceTakeRecord.RefClass != null)
                                                        {
                                                            //各科目班排名
                                                            key = "班排名" + sceTakeRecord.RefClass.ClassID + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(sceTakeRecord.ExamScore);
                                                            rankStudents[key].Add(studentID);
                                                        }
                                                        if (sceTakeRecord.Department != "")
                                                        {
                                                            //各科目科排名
                                                            key = "科排名" + sceTakeRecord.Department + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(sceTakeRecord.ExamScore);
                                                            rankStudents[key].Add(studentID);
                                                        }
                                                        //各科目全校排名
                                                        key = "全校排名" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                        if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                        if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                        ranks[key].Add(sceTakeRecord.ExamScore);
                                                        rankStudents[key].Add(studentID);
                                                    }
                                                }
                                                else
                                                {
                                                    summaryRank = false;
                                                }
                                            }
                                            #endregion



                                            #region 計算回歸科目

                                            // 2016/7/5 穎驊新增，用來計算有回歸科目的子科目成績加權起來

                                            foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                            {

                                                if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "" && SubjectTypeArrange_complete.ContainsKey(sceTakeRecord.Subject))
                                                {

                                                    // 新增回歸科目的成績欄位，從0開始
                                                    if (!studentRegressSubjectTotalW.ContainsKey(SubjectTypeArrange_complete[sceTakeRecord.Subject]))
                                                    {

                                                        studentRegressSubjectTotalW.Add(SubjectTypeArrange_complete[sceTakeRecord.Subject], 0);

                                                    }

                                                    // 新增回歸科目的學分欄位，從0開始
                                                    if (!studentRegressSubjectTotalCredit.ContainsKey(SubjectTypeArrange_complete[sceTakeRecord.Subject]))
                                                    {

                                                        studentRegressSubjectTotalCredit.Add(SubjectTypeArrange_complete[sceTakeRecord.Subject], 0);

                                                    }

                                                    //開始針對該領域進行加權總分計算
                                                    studentRegressSubjectTotalW[SubjectTypeArrange_complete[sceTakeRecord.Subject]] += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();

                                                    //開始針對該領域進行加權總學分計算
                                                    studentRegressSubjectTotalCredit[SubjectTypeArrange_complete[sceTakeRecord.Subject]] += sceTakeRecord.CreditDec();

                                                    //可進行總分的Rank排名，但我們目前"回歸科目"只需要加權平均後的結果，故先註解起來
                                                    #region 加權總分
                                                    ////計算加權總分
                                                    //RegressSubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                    //RegressSubjectCreditSum += sceTakeRecord.CreditDec();

                                                    //if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                    //{
                                                    //    if (sceTakeRecord.RefClass != null)
                                                    //    {
                                                    //        //各科目班排名
                                                    //        key = "班排名" + sceTakeRecord.RefClass.ClassID + "^^^" + "回歸科目" + "^^^" + SubjectTypeArrange_complete[sceTakeRecord.Subject] + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                    //        if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    //        if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    //        ranks[key].Add(sceTakeRecord.ExamScore);
                                                    //        rankStudents[key].Add(studentID);
                                                    //    }
                                                    //    if (sceTakeRecord.Department != "")
                                                    //    {
                                                    //        //各科目科排名
                                                    //        key = "科排名" + sceTakeRecord.Department + "^^^" + gradeyear + "^^^" + "回歸科目" + "^^^" + SubjectTypeArrange_complete[sceTakeRecord.Subject] + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                    //        if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    //        if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    //        ranks[key].Add(sceTakeRecord.ExamScore);
                                                    //        rankStudents[key].Add(studentID);
                                                    //    }
                                                    //    //各科目全校排名
                                                    //    key = "全校排名" + gradeyear + "^^^" + "回歸科目" + "^^^" + SubjectTypeArrange_complete[sceTakeRecord.Subject] + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                    //    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    //    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    //    ranks[key].Add(sceTakeRecord.ExamScore);
                                                    //    rankStudents[key].Add(studentID);
                                                    //} 
                                                    #endregion
                                                }
                                                else
                                                {
                                                    RegressRank = false;
                                                }


                                            }
                                            #endregion

                                        }



                                        if (tag1ID != "" && conf.TagRank1SubjectList.Contains(subjectName))
                                        {
                                            #region 有Tag1且是排名科目
                                            foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                            {
                                                if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                {
                                                    tag1SubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                    tag1SubjectCount++;
                                                    //計算加權總分
                                                    tag1SubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                    tag1SubjectCreditSum += sceTakeRecord.CreditDec();
                                                    //各科目類別1排名
                                                    if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                    {
                                                        if (conf.PrintSubjectList.Contains(subjectName))//是列印科目才算科目排名                                                
                                                        {
                                                            key = "類別1排名" + tag1ID + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(sceTakeRecord.ExamScore);
                                                            rankStudents[key].Add(studentID);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    tag1SummaryRank = false;
                                                }
                                            }
                                            #endregion
                                        }



                                        if (tag2ID != "" && conf.TagRank2SubjectList.Contains(subjectName))
                                        {
                                            #region 有Tag2且是排名科目
                                            foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                            {
                                                if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                {
                                                    tag2SubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                    tag2SubjectCount++;
                                                    //計算加權總分
                                                    tag2SubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                    tag2SubjectCreditSum += sceTakeRecord.CreditDec();
                                                    //各科目類別2排名
                                                    if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                    {
                                                        if (conf.PrintSubjectList.Contains(subjectName))//是列印科目才算科目排名                                                
                                                        {
                                                            key = "類別2排名" + tag2ID + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(sceTakeRecord.ExamScore);
                                                            rankStudents[key].Add(studentID);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    tag2SummaryRank = false;
                                                }
                                            }
                                            #endregion
                                        }



                                        // 以下為計算類別1 Tag1 的回歸科目項目

                                        foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                        {
                                            // 具有Tag1，且該科目在Tag1的科目列表，還有該科目也在回歸科目列表中
                                            if (tag1ID != "" && conf.TagRank1SubjectList.Contains(subjectName) && SubjectTypeArrange_complete.ContainsKey(sceTakeRecord.Subject))
                                            {
                                                #region 有Tag1且是排名科目

                                                if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                {

                                                    // 新增類別1 Tag1回歸科目的成績欄位，從0開始
                                                    if (!studentRegressSubjectTotalW_Tag1.ContainsKey(SubjectTypeArrange_complete[sceTakeRecord.Subject]))
                                                    {

                                                        studentRegressSubjectTotalW_Tag1.Add(SubjectTypeArrange_complete[sceTakeRecord.Subject], 0);

                                                    }

                                                    // 新增類別1 Tag1回歸科目的學分欄位，從0開始
                                                    if (!studentRegressSubjectTotalCredit_Tag1.ContainsKey(SubjectTypeArrange_complete[sceTakeRecord.Subject]))
                                                    {

                                                        studentRegressSubjectTotalCredit_Tag1.Add(SubjectTypeArrange_complete[sceTakeRecord.Subject], 0);

                                                    }

                                                    //開始針對該領域進行加權總分計算
                                                    studentRegressSubjectTotalW_Tag1[SubjectTypeArrange_complete[sceTakeRecord.Subject]] += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();

                                                    //開始針對該領域進行加權總學分計算
                                                    studentRegressSubjectTotalCredit_Tag1[SubjectTypeArrange_complete[sceTakeRecord.Subject]] += sceTakeRecord.CreditDec();


                                                }

                                            }
                                                #endregion
                                        }

                                        // 以下為計算類別2 Tag2 的回歸科目項目

                                        foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                        {
                                            // 具有Tag2，且該科目在Tag2的科目列表，還有該科目也在回歸科目列表中
                                            if (tag2ID != "" && conf.TagRank2SubjectList.Contains(subjectName) && SubjectTypeArrange_complete.ContainsKey(sceTakeRecord.Subject))
                                            {
                                                #region 有Tag2且是排名科目

                                                if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                {


                                                    // 新增類別2 Tag2回歸科目的成績欄位，從0開始
                                                    if (!studentRegressSubjectTotalW_Tag2.ContainsKey(SubjectTypeArrange_complete[sceTakeRecord.Subject]))
                                                    {

                                                        studentRegressSubjectTotalW_Tag2.Add(SubjectTypeArrange_complete[sceTakeRecord.Subject], 0);

                                                    }

                                                    // 新增類別2 Tag2回歸科目的學分欄位，從0開始
                                                    if (!studentRegressSubjectTotalCredit_Tag2.ContainsKey(SubjectTypeArrange_complete[sceTakeRecord.Subject]))
                                                    {

                                                        studentRegressSubjectTotalCredit_Tag2.Add(SubjectTypeArrange_complete[sceTakeRecord.Subject], 0);

                                                    }

                                                    //開始針對該領域進行加權總分計算
                                                    studentRegressSubjectTotalW_Tag2[SubjectTypeArrange_complete[sceTakeRecord.Subject]] += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();

                                                    //開始針對該領域進行加權總學分計算
                                                    studentRegressSubjectTotalCredit_Tag2[SubjectTypeArrange_complete[sceTakeRecord.Subject]] += sceTakeRecord.CreditDec();


                                                }

                                            }
                                                #endregion
                                        }





                                    }
                                    if (printSubjectCount > 0)
                                    {
                                        #region 有列印科目處理加總成績
                                        //總分
                                        studentPrintSubjectSum.Add(studentID, printSubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentPrintSubjectAvg.Add(studentID, Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                        if (rank && studentRec.Status == "一般" && summaryRank == true)//不在過濾名單且沒有特殊成績狀況且為一般生才做排名
                                        {
                                            //總分班排名
                                            key = "總分班排名" + studentRec.RefClass.ClassID;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(printSubjectSum);
                                            rankStudents[key].Add(studentID);
                                            //總分科排名
                                            key = "總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(printSubjectSum);
                                            rankStudents[key].Add(studentID);
                                            //總分全校排名
                                            key = "總分全校排名" + gradeyear;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(printSubjectSum);
                                            rankStudents[key].Add(studentID);
                                            //平均班排名
                                            key = "平均班排名" + studentRec.RefClass.ClassID;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                            rankStudents[key].Add(studentID);
                                            //平均科排名
                                            key = "平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                            rankStudents[key].Add(studentID);
                                            //平均全校排名
                                            key = "平均全校排名" + gradeyear;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                            rankStudents[key].Add(studentID);
                                        }
                                        #endregion
                                        if (printSubjectCreditSum > 0)
                                        {
                                            #region 有總學分數處理加總
                                            //加權總分
                                            studentPrintSubjectSumW.Add(studentID, printSubjectSumW);
                                            //加權平均四捨五入至小數點第二位
                                            studentPrintSubjectAvgW.Add(studentID, Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                            if (rank && studentRec.Status == "一般" && summaryRank == true)//不在過濾名單且為一般生才做排名
                                            {
                                                //加權總分班排名
                                                key = "加權總分班排名" + studentRec.RefClass.ClassID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(printSubjectSumW);
                                                rankStudents[key].Add(studentID);
                                                //加權總分科排名
                                                key = "加權總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(printSubjectSumW);
                                                rankStudents[key].Add(studentID);
                                                //加權總分全校排名
                                                key = "加權總分全校排名" + gradeyear;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(printSubjectSumW);
                                                rankStudents[key].Add(studentID);
                                                //加權平均班排名
                                                key = "加權平均班排名" + studentRec.RefClass.ClassID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                rankStudents[key].Add(studentID);
                                                //加權平均科排名
                                                key = "加權平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                rankStudents[key].Add(studentID);
                                                //加權平均全校排名
                                                key = "加權平均全校排名" + gradeyear;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                        }


                                        //2016/7/5 穎驊新增，處理回歸科目

                                        // 假如前面所處理科目字典裏頭，有回歸科目被加入
                                        if (studentRegressSubjectTotalW.Count > 0)
                                        {
                                            #region 有總學分數處理加總


                                            if (!studentIDRegressSubjectAvgW.ContainsKey(studentID))
                                            {

                                                studentIDRegressSubjectAvgW.Add(studentID, new Dictionary<string, decimal>());
                                            }

                                            //加權平均四捨五入至小數點第二位

                                            foreach (String RegressSubject in studentRegressSubjectTotalW.Keys)
                                            {

                                                //if (!RegressSubjectAvgW.ContainsKey(RegressSubject))
                                                //{
                                                //    RegressSubjectAvgW.Add(RegressSubject, Math.Round(studentRegressSubjectTotalW[RegressSubject] / studentRegressSubjectTotalCredit[RegressSubject], 2, MidpointRounding.AwayFromZero));

                                                //}

                                                studentIDRegressSubjectAvgW[studentID].Add(RegressSubject, Math.Round(studentRegressSubjectTotalW[RegressSubject] / studentRegressSubjectTotalCredit[RegressSubject], 2, MidpointRounding.AwayFromZero));



                                                if (rank && studentRec.Status == "一般")//不在過濾名單且為一般生才做排名
                                                {


                                                    //回歸科目班排名
                                                    key = "回歸科目" + "_" + RegressSubject + "_" + "班排名" + studentRec.RefClass.ClassID;
                                                    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    ranks[key].Add(studentIDRegressSubjectAvgW[studentID][RegressSubject]);
                                                    rankStudents[key].Add(studentID);


                                                    //回歸科目科排名
                                                    key = "回歸科目" + "_" + RegressSubject + "_" + "科排名" + "_" + studentRec.Department + "^^^" + gradeyear;
                                                    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    ranks[key].Add(studentIDRegressSubjectAvgW[studentID][RegressSubject]);
                                                    rankStudents[key].Add(studentID);


                                                    //回歸科目全校排名
                                                    key = "回歸科目" + "_" + RegressSubject + "_" + "全校排名" + gradeyear;
                                                    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    ranks[key].Add(studentIDRegressSubjectAvgW[studentID][RegressSubject]);
                                                    rankStudents[key].Add(studentID);

                                                }

                                            }
                                            #endregion
                                        }
                                    }


                                    // 類別1 回歸科目分數計算、給Key填Rank
                                    if (studentRegressSubjectTotalW_Tag1.Count > 0)
                                    {


                                        if (!studentIDRegressSubjectAvgW_Tag1.ContainsKey(studentID))
                                        {
                                            studentIDRegressSubjectAvgW_Tag1.Add(studentID, new Dictionary<string, decimal>());
                                        }

                                        //加權平均四捨五入至小數點第二位

                                        foreach (String RegressSubject in studentRegressSubjectTotalW_Tag1.Keys)
                                        {

                                            studentIDRegressSubjectAvgW_Tag1[studentID].Add(RegressSubject, Math.Round(studentRegressSubjectTotalW_Tag1[RegressSubject] / studentRegressSubjectTotalCredit_Tag1[RegressSubject], 2, MidpointRounding.AwayFromZero));

                                            if (rank && studentRec.Status == "一般" && studentRegressSubjectTotalCredit_Tag1.Count > 0)//不在過濾名單且為一般生才做排名
                                            {
                                                key = "類別1回歸科目排名" + "_" + RegressSubject + "^^^" + gradeyear + "^^^" + tag1ID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(studentIDRegressSubjectAvgW_Tag1[studentID][RegressSubject]);
                                                rankStudents[key].Add(studentID);
                                            }
                                        }

                                    }

                                    // 類別2 回歸科目分數計算、給Key填Rank
                                    if (studentRegressSubjectTotalW_Tag2.Count > 0)
                                    {

                                        if (!studentIDRegressSubjectAvgW_Tag2.ContainsKey(studentID))
                                        {
                                            studentIDRegressSubjectAvgW_Tag2.Add(studentID, new Dictionary<string, decimal>());
                                        }

                                        //加權平均四捨五入至小數點第二位

                                        foreach (String RegressSubject in studentRegressSubjectTotalW_Tag2.Keys)
                                        {

                                            studentIDRegressSubjectAvgW_Tag2[studentID].Add(RegressSubject, Math.Round(studentRegressSubjectTotalW_Tag2[RegressSubject] / studentRegressSubjectTotalCredit_Tag2[RegressSubject], 2, MidpointRounding.AwayFromZero));


                                            if (rank && studentRec.Status == "一般" && studentRegressSubjectTotalCredit_Tag2.Count > 0)//不在過濾名單且為一般生才做排名
                                            {

                                                key = "類別2回歸科目排名" + "_" + RegressSubject + "^^^" + gradeyear + "^^^" + tag2ID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(studentIDRegressSubjectAvgW_Tag2[studentID][RegressSubject]);
                                                rankStudents[key].Add(studentID);

                                            }
                                        }

                                    }






                                    //類別1總分平均排名
                                    if (tag1SubjectCount > 0)
                                    {
                                        //總分
                                        studentTag1SubjectSum.Add(studentID, tag1SubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentTag1SubjectAvg.Add(studentID, Math.Round(tag1SubjectSum / tag1SubjectCount, 2, MidpointRounding.AwayFromZero));
                                        if (rank && studentRec.Status == "一般" && tag1SummaryRank == true)//不在過濾名單且為一般生才做排名
                                        {
                                            key = "類別1總分排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(tag1SubjectSum);
                                            rankStudents[key].Add(studentID);

                                            key = "類別1平均排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(Math.Round(tag1SubjectSum / tag1SubjectCount, 2, MidpointRounding.AwayFromZero));
                                            rankStudents[key].Add(studentID);
                                        }
                                        //類別1加權總分平均排名
                                        if (tag1SubjectCreditSum > 0)
                                        {
                                            studentTag1SubjectSumW.Add(studentID, tag1SubjectSumW);
                                            studentTag1SubjectAvgW.Add(studentID, Math.Round(tag1SubjectSumW / tag1SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                            if (rank && studentRec.Status == "一般" && tag1SummaryRank == true)//不在過濾名單且為一般生才做排名
                                            {
                                                key = "類別1加權總分排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(tag1SubjectSumW);
                                                rankStudents[key].Add(studentID);

                                                key = "類別1加權平均排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(Math.Round(tag1SubjectSumW / tag1SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                rankStudents[key].Add(studentID);
                                            }
                                        }
                                    }
                                    //類別2總分平均排名
                                    if (tag2SubjectCount > 0)
                                    {
                                        //總分
                                        studentTag2SubjectSum.Add(studentID, tag2SubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentTag2SubjectAvg.Add(studentID, Math.Round(tag2SubjectSum / tag2SubjectCount, 2, MidpointRounding.AwayFromZero));
                                        if (rank && studentRec.Status == "一般" && tag2SummaryRank == true)//不在過濾名單且為一般生才做排名
                                        {
                                            key = "類別2總分排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(tag2SubjectSum);
                                            rankStudents[key].Add(studentID);
                                            key = "類別2平均排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            ranks[key].Add(Math.Round(tag2SubjectSum / tag2SubjectCount, 2, MidpointRounding.AwayFromZero));
                                            rankStudents[key].Add(studentID);
                                        }
                                        //類別2加權總分平均排名
                                        if (tag2SubjectCreditSum > 0)
                                        {
                                            studentTag2SubjectSumW.Add(studentID, tag2SubjectSumW);
                                            studentTag2SubjectAvgW.Add(studentID, Math.Round(tag2SubjectSumW / tag2SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                            if (rank && studentRec.Status == "一般" && tag2SummaryRank == true)//不在過濾名單且為一般生才做排名
                                            {
                                                key = "類別2加權總分排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(tag2SubjectSumW);
                                                rankStudents[key].Add(studentID);

                                                key = "類別2加權平均排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(Math.Round(tag2SubjectSumW / tag2SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                rankStudents[key].Add(studentID);
                                            }
                                        }
                                    }
                                }
                                progressCount++;
                                bkw.ReportProgress(40 + progressCount * 30 / total);
                            }
                        }
                        foreach (var k in ranks.Keys)
                        {
                            var rankscores = ranks[k];
                            //排序
                            rankscores.Sort();
                            rankscores.Reverse();
                            //高均標、組距
                            if (rankscores.Count > 0)
                            {
                                #region 算高標的中點
                                int middleIndex = 0;
                                int count = 1;
                                var score = rankscores[0];
                                while (rankscores.Count > middleIndex)
                                {
                                    if (score != rankscores[middleIndex])
                                    {
                                        if (count * 2 >= rankscores.Count) break;
                                        score = rankscores[middleIndex];
                                    }
                                    middleIndex++;
                                    count++;
                                }
                                if (rankscores.Count == middleIndex)
                                {
                                    middleIndex--;
                                    count--;
                                }
                                #endregion
                                analytics.Add(k + "^^^高標", Math.Round(rankscores.GetRange(0, count).Average(), 2, MidpointRounding.AwayFromZero));
                                analytics.Add(k + "^^^均標", Math.Round(rankscores.Average(), 2, MidpointRounding.AwayFromZero));
                                #region 算低標的中點
                                middleIndex = rankscores.Count - 1;
                                count = 1;
                                score = rankscores[middleIndex];
                                while (middleIndex >= 0)
                                {
                                    if (score != rankscores[middleIndex])
                                    {
                                        if (count * 2 >= rankscores.Count) break;
                                        score = rankscores[middleIndex];
                                    }
                                    middleIndex--;
                                    count++;
                                }
                                if (middleIndex < 0)
                                {
                                    middleIndex++;
                                    count--;
                                }
                                #endregion
                                analytics.Add(k + "^^^低標", Math.Round(rankscores.GetRange(middleIndex, count).Average(), 2, MidpointRounding.AwayFromZero));
                                //Compute the Average      
                                var avg = (double)rankscores.Average();
                                //Perform the Sum of (value-avg)_2_2      
                                var sum = (double)rankscores.Sum(d => Math.Pow((double)d - avg, 2));
                                //Put it all together      
                                analytics.Add(k + "^^^標準差", Math.Round((decimal)Math.Sqrt((sum) / rankscores.Count()), 2, MidpointRounding.AwayFromZero));
                            }
                            #region 計算級距
                            int count90 = 0, count80 = 0, count70 = 0, count60 = 0, count50 = 0, count40 = 0, count30 = 0, count20 = 0, count10 = 0;
                            int count100Up = 0, count90Up = 0, count80Up = 0, count70Up = 0, count60Up = 0, count50Up = 0, count40Up = 0, count30Up = 0, count20Up = 0, count10Up = 0;
                            int count90Down = 0, count80Down = 0, count70Down = 0, count60Down = 0, count50Down = 0, count40Down = 0, count30Down = 0, count20Down = 0, count10Down = 0;
                            foreach (var score in rankscores)
                            {
                                if (score >= 100)
                                    count100Up++;
                                else if (score >= 90)
                                    count90++;
                                else if (score >= 80)
                                    count80++;
                                else if (score >= 70)
                                    count70++;
                                else if (score >= 60)
                                    count60++;
                                else if (score >= 50)
                                    count50++;
                                else if (score >= 40)
                                    count40++;
                                else if (score >= 30)
                                    count30++;
                                else if (score >= 20)
                                    count20++;
                                else if (score >= 10)
                                    count10++;
                                else
                                    count10Down++;
                            }
                            count90Up = count100Up + count90;
                            count80Up = count90Up + count80;
                            count70Up = count80Up + count70;
                            count60Up = count70Up + count60;
                            count50Up = count60Up + count50;
                            count40Up = count50Up + count40;
                            count30Up = count40Up + count30;
                            count20Up = count30Up + count20;
                            count10Up = count20Up + count10;

                            count20Down = count10Down + count10;
                            count30Down = count20Down + count20;
                            count40Down = count30Down + count30;
                            count50Down = count40Down + count40;
                            count60Down = count50Down + count50;
                            count70Down = count60Down + count60;
                            count80Down = count70Down + count70;
                            count90Down = count80Down + count80;

                            analytics.Add(k + "^^^count90", count90);
                            analytics.Add(k + "^^^count80", count80);
                            analytics.Add(k + "^^^count70", count70);
                            analytics.Add(k + "^^^count60", count60);
                            analytics.Add(k + "^^^count50", count50);
                            analytics.Add(k + "^^^count40", count40);
                            analytics.Add(k + "^^^count30", count30);
                            analytics.Add(k + "^^^count20", count20);
                            analytics.Add(k + "^^^count10", count10);
                            analytics.Add(k + "^^^count100Up", count100Up);
                            analytics.Add(k + "^^^count90Up", count90Up);
                            analytics.Add(k + "^^^count80Up", count80Up);
                            analytics.Add(k + "^^^count70Up", count70Up);
                            analytics.Add(k + "^^^count60Up", count60Up);
                            analytics.Add(k + "^^^count50Up", count50Up);
                            analytics.Add(k + "^^^count40Up", count40Up);
                            analytics.Add(k + "^^^count30Up", count30Up);
                            analytics.Add(k + "^^^count20Up", count20Up);
                            analytics.Add(k + "^^^count10Up", count10Up);
                            analytics.Add(k + "^^^count90Down", count90Down);
                            analytics.Add(k + "^^^count80Down", count80Down);
                            analytics.Add(k + "^^^count70Down", count70Down);
                            analytics.Add(k + "^^^count60Down", count60Down);
                            analytics.Add(k + "^^^count50Down", count50Down);
                            analytics.Add(k + "^^^count40Down", count40Down);
                            analytics.Add(k + "^^^count30Down", count30Down);
                            analytics.Add(k + "^^^count20Down", count20Down);
                            analytics.Add(k + "^^^count10Down", count10Down);
                            #endregion
                        }
                        #endregion

                        // 先取得 K12 StudentRec,因為後面透過 k12.data 取資料有的傳入ID,有的傳入 Record 有點亂
                        List<K12.Data.StudentRecord> StudRecList = new List<K12.Data.StudentRecord>();
                        List<string> StudIDList = (from data in studentRecords select data.StudentID).ToList();
                        StudRecList = K12.Data.Student.SelectByIDs(StudIDList);

                        int SchoolYear, Semester;
                        int.TryParse(conf.SchoolYear, out SchoolYear);
                        int.TryParse(conf.Semester, out Semester);


                        List<K12.Data.PeriodMappingInfo> PeriodMappingList = K12.Data.PeriodMapping.SelectAll();
                        // 節次>類別
                        Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
                        foreach (K12.Data.PeriodMappingInfo rec in PeriodMappingList)
                        {
                            if (!PeriodMappingDict.ContainsKey(rec.Name))
                                PeriodMappingDict.Add(rec.Name, rec.Type);
                        }

                        bkw.ReportProgress(70);
                        elseReady.WaitOne();

                        _studPassSumCreditDict1.Clear();
                        _studPassSumCreditDictAll.Clear();
                        _studPassSumCreditDictC1.Clear();
                        _studPassSumCreditDictC2.Clear();

                        progressCount = 0;
                        int cs2RowIndex = 2;

                        //學年度
                        cs0[2, 0].PutValue(conf.SchoolYear);
                        //學期
                        cs0[2, 1].PutValue(conf.Semester);
                        //年級
                        cs0[2, 2].PutValue(targetGrade);
                        //排名類型
                        cs0[2, 3].PutValue("定期評量成績排名");
                        // 排名次序
                        cs0[2, 4].PutValue(conf.ExamRecord.DisplayOrder);
                        //顯示名稱
                        cs0[2, 5].PutValue(conf.ExamRecord.Name);
                        //註記
                        cs0[2, 6].PutValue("");
                        //建立時間
                        cs0[2, 7].PutValue(DateTime.Now.ToString("MM-dd-yyyy"));

                        int cs1RowCount = 2;
                        #region 填入資料表
                        foreach (var stuRec in studentRecords)
                        {
                            // 本學期取得學分數
                            if (!_studPassSumCreditDict1.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDict1.Add(stuRec.StudentID, 0);

                            // 累計取得學分數
                            if (!_studPassSumCreditDictAll.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDictAll.Add(stuRec.StudentID, 0);

                            if (!_studPassSumCreditDictC1.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDictC1.Add(stuRec.StudentID, 0);

                            if (!_studPassSumCreditDictC2.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDictC2.Add(stuRec.StudentID, 0);

                            string studentID = stuRec.StudentID;
                            string gradeYear = (stuRec.RefClass == null ? "" : "" + stuRec.RefClass.GradeYear);
                            DataRow row = table.NewRow();


                            #region 基本資料


                            //學生系統編號
                            cs1[cs1RowCount, 0].PutValue(stuRec.StudentID);
                            //學號
                            cs1[cs1RowCount, 1].PutValue(stuRec.StudentNumber);
                            //姓名
                            cs1[cs1RowCount, 2].PutValue(stuRec.StudentName);
                            //班級
                            cs1[cs1RowCount, 3].PutValue(stuRec.RefClass == null ? "" : stuRec.RefClass.ClassName);
                            //座號
                            cs1[cs1RowCount, 4].PutValue(stuRec.SeatNo);
                            //班導師
                            cs1[cs1RowCount, 5].PutValue((stuRec.RefClass == null || stuRec.RefClass.RefTeacher == null) ? "" : stuRec.RefClass.RefTeacher.TeacherName);
                            //科別
                            cs1[cs1RowCount, 6].PutValue(stuRec.RefClass == null ? "" : stuRec.Department);

                            if (studentTag1Group.ContainsKey(studentID))
                            {
                                foreach (var tag in studentTags[studentID])
                                {
                                    if (tag.RefTagID == studentTag1Group[studentID])
                                    {
                                        //類別1
                                        cs1[cs1RowCount, 7].PutValue(tag.Name);
                                    }
                                }
                            }

                            if (studentTag2Group.ContainsKey(studentID))
                            {
                                foreach (var tag in studentTags[studentID])
                                {
                                    if (tag.RefTagID == studentTag2Group[studentID])
                                    {
                                        //類別2
                                        cs1[cs1RowCount, 8].PutValue(tag.Name);
                                    }

                                }
                            }
                            //參與排名
                            cs1[cs1RowCount, 9].PutValue(joinRank[stuRec.StudentID] ? "是" : "否");
                            cs1RowCount++;
                            #endregion
                            #region 成績資料
                            #region 各科成績資料
                            #region 分項成績
                            int currentGradeYear = -1;
                            foreach (var semesterEntryScore in stuRec.SemesterEntryScoreList)
                            {
                                if (("" + semesterEntryScore.SchoolYear) == conf.SchoolYear && ("" + semesterEntryScore.Semester) == conf.Semester)
                                {
                                    row["學期" + semesterEntryScore.Entry + "成績"] = semesterEntryScore.Score;
                                    currentGradeYear = semesterEntryScore.GradeYear;
                                }
                            }
                            #region 學期學業成績排名

                            foreach (var k in new string[] { "班", "科", "校" })
                            {
                                if (stuRec.Fields.ContainsKey("學期學業成績" + k + "排名")) row["學期學業成績" + k + "排名"] = "" + stuRec.Fields["學期學業成績" + k + "排名"];
                                if (stuRec.Fields.ContainsKey("學期學業成績" + k + "排名母數")) row["學期學業成績" + k + "排名母數"] = "" + stuRec.Fields["學期學業成績" + k + "排名母數"];
                            }
                            //類別1
                            if (studentTag1Group.ContainsKey(studentID))
                            {
                                foreach (var tag in studentTags[studentID])
                                {
                                    if (tag.RefTagID == studentTag1Group[studentID])
                                    {
                                        key = "學期學業成績" + tag.Name + "排名";
                                        if (stuRec.Fields.ContainsKey(key))
                                            row["學期學業成績類別1排名"] = "" + stuRec.Fields[key];
                                        key = "學期學業成績" + tag.Name + "排名母數";
                                        if (stuRec.Fields.ContainsKey(key))
                                            row["學期學業成績類別1排名母數"] = "" + stuRec.Fields[key];
                                        break;
                                    }
                                }
                            }
                            //類別2
                            if (studentTag2Group.ContainsKey(studentID))
                            {
                                foreach (var tag in studentTags[studentID])
                                {
                                    if (tag.RefTagID == studentTag2Group[studentID])
                                    {
                                        key = "學期學業成績" + tag.Name + "排名";
                                        if (stuRec.Fields.ContainsKey(key))
                                            row["學期學業成績類別2排名"] = "" + stuRec.Fields[key];
                                        key = "學期學業成績" + tag.Name + "排名母數";
                                        if (stuRec.Fields.ContainsKey(key))
                                            row["學期學業成績類別2排名母數"] = "" + stuRec.Fields[key];
                                        break;
                                    }
                                }
                            }
                            #endregion

                            if (conf.Semester == "2")
                            {
                                #region 學年學業成績及排名
                                if (conf.WithSchoolYearScore)
                                {
                                    foreach (var schoolYearEntryScore in stuRec.SchoolYearEntryScoreList)
                                    {
                                        if (("" + schoolYearEntryScore.SchoolYear) == conf.SchoolYear)
                                        {
                                            row["學年" + schoolYearEntryScore.Entry + "成績"] = schoolYearEntryScore.Score;
                                        }
                                    }
                                    if (stuRec.Fields.ContainsKey("SchoolYearEntryClassRating"))
                                    {
                                        System.Xml.XmlElement _sems_ratings = stuRec.Fields["SchoolYearEntryClassRating"] as System.Xml.XmlElement;
                                        string path = string.Format("SchoolYearEntryScore[SchoolYear='{0}']/ClassRating/Rating/Item[@分項='學業']/@排名", conf.SchoolYear);
                                        System.Xml.XmlNode result = _sems_ratings.SelectSingleNode(path);
                                        if (result != null)
                                        {
                                            row["學年學業成績班排名"] = result.InnerText;
                                        }
                                    }
                                }
                                #endregion
                                if (conf.WithPrevSemesterScore)
                                {
                                    foreach (var semesterEntryScore in stuRec.SemesterEntryScoreList)
                                    {
                                        if (semesterEntryScore.Semester == 1 && semesterEntryScore.GradeYear == currentGradeYear)
                                        {
                                            row["上學期" + semesterEntryScore.Entry + "成績"] = semesterEntryScore.Score;
                                        }
                                    }
                                }
                            }






                            #endregion
                            #region 整理科目順序
                            List<string> subjects1 = new List<string>();//本學期
                            List<string> subjects2 = new List<string>();//上學期
                            List<string> subjects3 = new List<string>();//學年
                            foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                            {
                                if (("" + semesterSubjectScore.SchoolYear) == conf.SchoolYear && ("" + semesterSubjectScore.Semester) == conf.Semester)
                                {
                                    if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
                                    {
                                        subjects1.Add(semesterSubjectScore.Subject);
                                        currentGradeYear = semesterSubjectScore.GradeYear;
                                    }
                                }
                            }
                            if (studentExamSores.ContainsKey(stuRec.StudentID))
                            {
                                foreach (var subjectName in studentExamSores[studentID].Keys)
                                {
                                    foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                    {
                                        if (conf.PrintSubjectList.Contains(subjectName))
                                        {
                                            #region 跟學期成績做差異新增
                                            bool match = false;
                                            foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                            {
                                                if (("" + semesterSubjectScore.SchoolYear) == conf.SchoolYear
                                                    && ("" + semesterSubjectScore.Semester) == conf.Semester
                                                    && semesterSubjectScore.Subject == subjectName
                                                    && semesterSubjectScore.Level == accessHelper.CourseHelper.GetCourse(courseID)[0].SubjectLevel)
                                                {
                                                    match = true;
                                                    break;
                                                }
                                            }
                                            if (!match)
                                            {
                                                subjects1.Add(subjectName);
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                            if (conf.Semester == "2")
                            {
                                if (conf.WithPrevSemesterScore)
                                {
                                    foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                    {
                                        if (semesterSubjectScore.Semester == 1 && semesterSubjectScore.GradeYear == currentGradeYear)
                                        {
                                            if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
                                                subjects2.Add(semesterSubjectScore.Subject);
                                        }
                                    }
                                }
                                if (conf.WithSchoolYearScore)
                                {
                                    foreach (var schoolYearSubjectScore in stuRec.SchoolYearSubjectScoreList)
                                    {
                                        if (("" + schoolYearSubjectScore.SchoolYear) == conf.SchoolYear)
                                        {
                                            subjects3.Add(schoolYearSubjectScore.Subject);
                                        }
                                    }
                                }
                            }
                            var subjectNameList = new List<string>();
                            subjectNameList.AddRange(subjects1);
                            foreach (var subject in subjects1)
                            {
                                if (subjects2.Contains(subject)) subjects2.Remove(subject);
                                if (subjects3.Contains(subject)) subjects3.Remove(subject);
                            }
                            subjectNameList.AddRange(subjects2);
                            foreach (var subject in subjects2)
                            {
                                if (subjects3.Contains(subject)) subjects3.Remove(subject);
                            }
                            subjectNameList.AddRange(subjects3);
                            subjectNameList.Sort(new StringComparer("國文"
                                            , "英文"
                                            , "數學"
                                            , "理化"
                                            , "生物"
                                            , "社會"
                                            , "物理"
                                            , "化學"
                                            , "歷史"
                                            , "地理"
                                            , "公民"));
                            #endregion


                            // 處理本學期取得學分與累計取得學分
                            foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                            {
                                if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
                                {

                                    // 本學期取得
                                    if (semesterSubjectScore.SchoolYear.ToString() == conf.SchoolYear && semesterSubjectScore.Semester.ToString() == conf.Semester && semesterSubjectScore.Pass)
                                        _studPassSumCreditDict1[stuRec.StudentID] += semesterSubjectScore.CreditDec();

                                    // 累計取得
                                    if (semesterSubjectScore.Pass)
                                    {
                                        _studPassSumCreditDictAll[stuRec.StudentID] += semesterSubjectScore.CreditDec();

                                        if (semesterSubjectScore.Require)
                                            _studPassSumCreditDictC1[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                        else
                                            _studPassSumCreditDictC2[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                    }
                                }
                            }

                            row["本學期取得學分數"] = _studPassSumCreditDict1[stuRec.StudentID];
                            row["累計取得學分數"] = _studPassSumCreditDictAll[stuRec.StudentID];
                            row["累計取得必修學分"] = _studPassSumCreditDictC1[stuRec.StudentID];
                            row["累計取得選修學分"] = _studPassSumCreditDictC2[stuRec.StudentID];

                            // 取得學生及格與補考標準
                            // 及格
                            decimal scA = 0;
                            // 補考
                            decimal scB = 0;
                            if (StudentApplyLimitDict.ContainsKey(stuRec.StudentID))
                            {
                                string sA = stuRec.RefClass.GradeYear + "_及";
                                string sB = stuRec.RefClass.GradeYear + "_補";

                                if (StudentApplyLimitDict[stuRec.StudentID].ContainsKey(sA))
                                    scA = StudentApplyLimitDict[stuRec.StudentID][sA];

                                if (StudentApplyLimitDict[stuRec.StudentID].ContainsKey(sB))
                                    scB = StudentApplyLimitDict[stuRec.StudentID][sB];
                            }

                            int subjectIndex = 1;
                            // 學期科目與定期評量
                            foreach (string subjectName in subjectNameList)
                            {
                                if (subjectIndex <= conf.SubjectLimit)
                                {
                                    decimal? subjectNumber = null;
                                    bool findInSemesterSubjectScore = false;
                                    bool findInSemester1SubjectScore = false;
                                    bool findInExamScores = false;
                                    #region 本學期學期成績
                                    foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                    {
                                        if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是"
                                            && semesterSubjectScore.Subject == subjectName
                                            && ("" + semesterSubjectScore.SchoolYear) == conf.SchoolYear
                                            && ("" + semesterSubjectScore.Semester) == conf.Semester)
                                        {
                                            findInSemesterSubjectScore = true;


                                            decimal level;
                                            subjectNumber = decimal.TryParse(semesterSubjectScore.Level, out level) ? (decimal?)level : null;
                                            row["科目名稱" + subjectIndex] = semesterSubjectScore.Subject + GetNumber(subjectNumber);
                                            row["學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
                                            row["科目必選修" + subjectIndex] = semesterSubjectScore.Require ? "必修" : "選修";
                                            row["科目校部定" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("修課校部訂");
                                            row["科目註記" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("註記");
                                            row["科目取得學分" + subjectIndex] = semesterSubjectScore.Pass ? "是" : "否";
                                            row["科目未取得學分註記" + subjectIndex] = semesterSubjectScore.Pass ? "" : "\f";

                                            //"原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績"
                                            if (semesterSubjectScore.Detail.GetAttribute("不需評分") != "是")
                                            {
                                                row["學期科目原始成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("原始成績");
                                                row["學期科目補考成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("補考成績");
                                                row["學期科目重修成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("重修成績");
                                                row["學期科目手動調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("擇優採計成績");
                                                row["學期科目學年調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("學年調整成績");
                                                row["學期科目成績" + subjectIndex] = semesterSubjectScore.Score;

                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("原始成績"))
                                                    row["學期科目原始成績註記" + subjectIndex] = "\f";
                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("補考成績"))
                                                    row["學期科目補考成績註記" + subjectIndex] = "\f";
                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("重修成績"))
                                                    row["學期科目重修成績註記" + subjectIndex] = "\f";
                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("擇優採計成績"))
                                                    row["學期科目手動成績註記" + subjectIndex] = "\f";
                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("學年調整成績"))
                                                    row["學期科目學年成績註記" + subjectIndex] = "\f";

                                                // 不及格
                                                if (semesterSubjectScore.Score < scA)
                                                {
                                                    // 可補考
                                                    if (semesterSubjectScore.Score >= scB)
                                                    {

                                                        row["學期科目需要補考註記" + subjectIndex] = "\f";
                                                    }
                                                    else
                                                    {
                                                        // 不可補考，須重修
                                                        row["學期科目需要重修註記" + subjectIndex] = "\f";
                                                    }
                                                }
                                            }
                                            #region 學期科目班、科、校、類別1、類別2排名
                                            key = "學期科目排名成績" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目排名成績" + subjectIndex] = "" + stuRec.Fields[key];
                                            //班
                                            key = "學期科目班排名" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目班排名" + subjectIndex] = "" + stuRec.Fields[key];
                                            key = "學期科目班排名母數" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目班排名母數" + subjectIndex] = "" + stuRec.Fields[key];
                                            //科
                                            key = "學期科目科排名" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目科排名" + subjectIndex] = "" + stuRec.Fields[key];
                                            key = "學期科目班科名母數" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目科排名母數" + subjectIndex] = "" + stuRec.Fields[key];
                                            //校
                                            key = "學期科目校排名" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目全校排名" + subjectIndex] = "" + stuRec.Fields[key];
                                            key = "學期科目科校名母數" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目全校排名母數" + subjectIndex] = "" + stuRec.Fields[key];


                                            //類別1
                                            if (studentTag1Group.ContainsKey(studentID))
                                            {
                                                foreach (var tag in studentTags[studentID])
                                                {
                                                    if (tag.RefTagID == studentTag1Group[studentID])
                                                    {
                                                        key = "學期科目成績" + tag.Name + "排名" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                                        if (stuRec.Fields.ContainsKey(key))
                                                            row["學期科目類別1排名" + subjectIndex] = "" + stuRec.Fields[key];
                                                        key = "學期科目成績" + tag.Name + "排名母數" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                                        if (stuRec.Fields.ContainsKey(key))
                                                            row["學期科目類別1排名母數" + subjectIndex] = "" + stuRec.Fields[key];
                                                        break;
                                                    }
                                                }
                                            }


                                            //類別2
                                            if (studentTag2Group.ContainsKey(studentID))
                                            {
                                                foreach (var tag in studentTags[studentID])
                                                {
                                                    if (tag.RefTagID == studentTag2Group[studentID])
                                                    {
                                                        key = "學期科目成績" + tag.Name + "排名" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                                        if (stuRec.Fields.ContainsKey(key))
                                                            row["學期科目類別2排名" + subjectIndex] = "" + stuRec.Fields[key];
                                                        key = "學期科目成績" + tag.Name + "排名母數" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                                        if (stuRec.Fields.ContainsKey(key))
                                                            row["學期科目類別2排名母數" + subjectIndex] = "" + stuRec.Fields[key];
                                                        break;
                                                    }
                                                }
                                            }
                                            #endregion
                                            stuRec.SemesterSubjectScoreList.Remove(semesterSubjectScore);
                                            break;
                                        }
                                    }
                                    #endregion
                                    #region 定期評量成績
                                    // 檢查畫面上定期評量列印科目
                                    if (conf.PrintSubjectList.Contains(subjectName))
                                    {
                                        if (studentExamSores.ContainsKey(studentID))
                                        {
                                            if (studentExamSores[studentID].ContainsKey(subjectName))
                                            {
                                                foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                                {
                                                    var sceTakeRecord = studentExamSores[studentID][subjectName][courseID];
                                                    if (sceTakeRecord != null)
                                                    {//有輸入
                                                        if (findInSemesterSubjectScore)
                                                        {
                                                            if (sceTakeRecord.SubjectLevel != "" + subjectNumber)
                                                            {
                                                                continue;
                                                            }
                                                        }
                                                        findInExamScores = true;
                                                        if (!findInSemesterSubjectScore)
                                                        {
                                                            decimal level;
                                                            subjectNumber = decimal.TryParse(sceTakeRecord.SubjectLevel, out level) ? (decimal?)level : null;
                                                            row["科目名稱" + subjectIndex] = sceTakeRecord.Subject + GetNumber(subjectNumber);
                                                            row["學分數" + subjectIndex] = sceTakeRecord.CreditDec();
                                                        }
                                                        row["科目成績" + subjectIndex] = sceTakeRecord.SpecialCase == "" ? ("" + sceTakeRecord.ExamScore) : sceTakeRecord.SpecialCase;


                                                        #region 班排名及落點分析
                                                        if (stuRec.RefClass != null)
                                                        {
                                                            key = "班排名" + stuRec.RefClass.ClassID + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                            {

                                                                int cs2ColIndex = 0;
                                                                //母群Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);

                                                                //類別Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("科目");

                                                                //科目名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Subject);
                                                                //科目級別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SubjectLevel);
                                                                //母群類型	
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("班排名");
                                                                //母群名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RefClass.ClassName);
                                                                //總人數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                                                //頂標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //前標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                                                //均標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                                                //後標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                                                //底標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                                // 系統編號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentID);
                                                                //班級名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RefClass.ClassName);
                                                                //座號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SeatNo);
                                                                //學號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentNumber);
                                                                //姓名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentName);
                                                                //成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.ExamScore);
                                                                //排名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1);
                                                                //PR值
                                                                var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(sceTakeRecord.ExamScore) - 1m) / (decimal)ranks[key].Count);
                                                                if (pr == 0) pr = 1;
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                                                // 百分比
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(sceTakeRecord.ExamScore) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));
                                                                // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.CreditDec());
                                                                // 權數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //分項類別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Entry);
                                                                // 成績年級
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 必選修
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Required ? "必修" : "選修");
                                                                // 校部定
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RequiredBy);
                                                                // 科目成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.FinalScore);
                                                                //原始成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //補考成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //重修成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //手動調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //學年調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //取得學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 不計學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCredit ? "是" : "否");
                                                                //不算評分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCalc ? "是" : "否");
                                                                //註記
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                                cs2RowIndex++;

                                                            }

                                                            #region 標準差、組距(先註解起來以後可能會用得到)
                                                            //if (rankStudents.ContainsKey(key))
                                                            //{
                                                            //    row["班高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                            //    row["班均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                            //    row["班低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                            //    row["班標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                            //    row["班組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                            //    row["班組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                            //    row["班組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                            //    row["班組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                            //    row["班組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                            //    row["班組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                            //    row["班組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                            //    row["班組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                            //    row["班組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                            //    row["班組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                            //    row["班組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                            //    row["班組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                            //    row["班組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                            //    row["班組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                            //    row["班組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                            //    row["班組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                            //    row["班組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                            //    row["班組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                            //    row["班組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                            //    row["班組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                            //    row["班組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                            //    row["班組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                            //    row["班組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                            //    row["班組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                            //    row["班組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                            //    row["班組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                            //    row["班組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                            //    row["班組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                            //} 
                                                            #endregion
                                                        }
                                                        #endregion

                                                        #region 科排名及落點分析
                                                        if (stuRec.Department != "")
                                                        {
                                                            key = "科排名" + stuRec.Department + "^^^" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                            {
                                                                int cs2ColIndex = 0;
                                                                //母群Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);

                                                                //類別Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("科目");

                                                                //科目名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Subject);
                                                                //科目級別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SubjectLevel);
                                                                //母群類型	
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("科排名");
                                                                //母群名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Department);
                                                                //總人數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                                                //頂標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //前標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                                                //均標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                                                //後標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                                                //底標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                                // 系統編號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentID);
                                                                //班級名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RefClass.ClassName);
                                                                //座號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SeatNo);
                                                                //學號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentNumber);
                                                                //姓名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentName);
                                                                //成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.ExamScore);
                                                                //排名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1);
                                                                //PR值
                                                                var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(sceTakeRecord.ExamScore) - 1m) / (decimal)ranks[key].Count);
                                                                if (pr == 0) pr = 1;
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                                                // 百分比
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(sceTakeRecord.ExamScore) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));
                                                                // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.CreditDec());
                                                                // 權數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //分項類別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Entry);
                                                                // 成績年級
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 必選修
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Required ? "必修" : "選修");
                                                                // 校部定
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RequiredBy);
                                                                // 科目成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.FinalScore);
                                                                //原始成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //補考成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //重修成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //手動調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //學年調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //取得學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 不計學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCredit ? "是" : "否");
                                                                //不算評分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCalc ? "是" : "否");
                                                                //註記
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                                cs2RowIndex++;
                                                            }
                                                            #region 標準差、組距(先註解起來，以後可能會用到)
                                                            //if (rankStudents.ContainsKey(key))
                                                            //{
                                                            //    row["科高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                            //    row["科均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                            //    row["科低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                            //    row["科標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                            //    row["科組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                            //    row["科組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                            //    row["科組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                            //    row["科組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                            //    row["科組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                            //    row["科組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                            //    row["科組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                            //    row["科組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                            //    row["科組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                            //    row["科組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                            //    row["科組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                            //    row["科組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                            //    row["科組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                            //    row["科組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                            //    row["科組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                            //    row["科組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                            //    row["科組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                            //    row["科組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                            //    row["科組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                            //    row["科組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                            //    row["科組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                            //    row["科組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                            //    row["科組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                            //    row["科組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                            //    row["科組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                            //    row["科組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                            //    row["科組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                            //    row["科組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                            //} 
                                                            #endregion
                                                        }
                                                        #endregion

                                                        #region 全校排名及落點分析
                                                        key = "全校排名" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                        if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                        {
                                                            int cs2ColIndex = 0;
                                                            //母群Key
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);

                                                            //類別Key
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("科目");

                                                            //科目名稱
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Subject);
                                                            //科目級別
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SubjectLevel);
                                                            //母群類型	
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("校排名");
                                                            //母群名稱
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RefClass.GradeYear + "年級");
                                                            //總人數
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                                            //頂標
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            //前標
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                                            //均標
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                                            //後標
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                                            //底標
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                            // 系統編號
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentID);
                                                            //班級名稱
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RefClass.ClassName);
                                                            //座號
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SeatNo);
                                                            //學號
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentNumber);
                                                            //姓名
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentName);
                                                            //成績
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.ExamScore);
                                                            //排名
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1);
                                                            //PR值
                                                            var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(sceTakeRecord.ExamScore) - 1m) / (decimal)ranks[key].Count);
                                                            if (pr == 0) pr = 1;
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                                            // 百分比
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(sceTakeRecord.ExamScore) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));
                                                            // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.CreditDec());
                                                            // 權數
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            //分項類別
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Entry);
                                                            // 成績年級
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            // 必選修
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Required ? "必修" : "選修");
                                                            // 校部定
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RequiredBy);
                                                            // 科目成績
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.FinalScore);
                                                            //原始成績
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            //補考成績
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            //重修成績
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            //手動調整成績
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            //學年調整成績
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            //取得學分
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                            // 不計學分
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCredit ? "是" : "否");
                                                            //不算評分
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCalc ? "是" : "否");
                                                            //註記
                                                            cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                            cs2RowIndex++;
                                                        }

                                                        #region 標準差、組距(先註解起來，以後可以用)

                                                        //if (rankStudents.ContainsKey(key))
                                                        //{
                                                        //    row["校高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                        //    row["校均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                        //    row["校低標" + subjectIndex] = analytics[key + "^^^低標"];


                                                        //    row["校標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                        //    row["校組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                        //    row["校組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                        //    row["校組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                        //    row["校組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                        //    row["校組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                        //    row["校組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                        //    row["校組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                        //    row["校組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                        //    row["校組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                        //    row["校組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                        //    row["校組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                        //    row["校組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                        //    row["校組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                        //    row["校組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                        //    row["校組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                        //    row["校組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                        //    row["校組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                        //    row["校組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                        //    row["校組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                        //    row["校組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                        //    row["校組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                        //    row["校組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                        //    row["校組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                        //    row["校組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                        //    row["校組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                        //    row["校組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                        //    row["校組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                        //    row["校組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];

                                                        //}
                                                        #endregion

                                                        #endregion

                                                        #region 類別1排名及落點分析

                                                        if (studentTag1Group.ContainsKey(studentID) && conf.TagRank1SubjectList.Contains(subjectName))
                                                        {
                                                            key = "類別1排名" + studentTag1Group[studentID] + "^^^" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                            {

                                                                int cs2ColIndex = 0;
                                                                //母群Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);

                                                                //類別Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("科目");

                                                                //科目名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Subject);
                                                                //科目級別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SubjectLevel);
                                                                //母群類型	
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別1");

                                                                var gName = "";
                                                                if (studentTag1Group.ContainsKey(studentID))
                                                                {
                                                                    foreach (var tag in studentTags[studentID])
                                                                    {
                                                                        if (tag.RefTagID == studentTag1Group[studentID])
                                                                        {
                                                                            gName = tag.Name;
                                                                        }
                                                                    }
                                                                }
                                                                //母群名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                                                //總人數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                                                //頂標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //前標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                                                //均標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                                                //後標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                                                //底標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");


                                                                #region 標準差、組距 (先註解起來，以後統計可以用                                                                )
                                                                //row["類1標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                                //row["類1組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                                //row["類1組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                                //row["類1組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                                //row["類1組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                                //row["類1組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                                //row["類1組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                                //row["類1組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                                //row["類1組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                                //row["類1組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                                //row["類1組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                                //row["類1組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                                //row["類1組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                                //row["類1組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                                //row["類1組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                                //row["類1組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                                //row["類1組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                                //row["類1組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                                //row["類1組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                                //row["類1組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                                //row["類1組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                                //row["類1組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                                //row["類1組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                                //row["類1組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                                //row["類1組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                                //row["類1組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                                //row["類1組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                                //row["類1組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                                //row["類1組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"]; 
                                                                #endregion



                                                                // 系統編號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentID);
                                                                //班級名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RefClass.ClassName);
                                                                //座號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SeatNo);
                                                                //學號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentNumber);
                                                                //姓名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentName);
                                                                //成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.ExamScore);
                                                                //排名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1);
                                                                //PR值
                                                                var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(sceTakeRecord.ExamScore) - 1m) / (decimal)ranks[key].Count);
                                                                if (pr == 0) pr = 1;
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                                                // 百分比
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(sceTakeRecord.ExamScore) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));
                                                                // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.CreditDec());
                                                                // 權數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //分項類別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Entry);
                                                                // 成績年級
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 必選修
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Required ? "必修" : "選修");
                                                                // 校部定
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RequiredBy);
                                                                // 科目成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.FinalScore);
                                                                //原始成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //補考成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //重修成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //手動調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //學年調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //取得學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 不計學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCredit ? "是" : "否");
                                                                //不算評分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCalc ? "是" : "否");
                                                                //註記
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                                cs2RowIndex++;
                                                            }
                                                        }
                                                        #endregion

                                                        #region 類別2排名及落點分析
                                                        if (studentTag2Group.ContainsKey(studentID) && conf.TagRank2SubjectList.Contains(subjectName))
                                                        {
                                                            key = "類別2排名" + studentTag2Group[studentID] + "^^^" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                            {
                                                                int cs2ColIndex = 0;

                                                                //母群Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);

                                                                //類別Key
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("科目");

                                                                //科目名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Subject);
                                                                //科目級別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SubjectLevel);
                                                                //母群類型	
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別2");

                                                                var gName = "";
                                                                if (studentTag2Group.ContainsKey(studentID))
                                                                {
                                                                    foreach (var tag in studentTags[studentID])
                                                                    {
                                                                        if (tag.RefTagID == studentTag2Group[studentID])
                                                                        {
                                                                            gName = tag.Name;
                                                                        }
                                                                    }
                                                                }
                                                                //母群名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                                                //總人數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                                                //頂標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //前標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                                                //均標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                                                //後標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                                                //底標
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");


                                                                #region 標準差、組距(先註解起來，以後可能可以用)
                                                                //if (rankStudents.ContainsKey(key))
                                                                //{
                                                                //    row["類2高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                                //    row["類2均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                                //    row["類2低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                                //    row["類2標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                                //    row["類2組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                                //    row["類2組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                                //    row["類2組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                                //    row["類2組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                                //    row["類2組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                                //    row["類2組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                                //    row["類2組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                                //    row["類2組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                                //    row["類2組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                                //    row["類2組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                                //    row["類2組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                                //    row["類2組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                                //    row["類2組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                                //    row["類2組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                                //    row["類2組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                                //    row["類2組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                                //    row["類2組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                                //    row["類2組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                                //    row["類2組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                                //    row["類2組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                                //    row["類2組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                                //    row["類2組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                                //    row["類2組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                                //    row["類2組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                                //    row["類2組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                                //    row["類2組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                                //    row["類2組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                                //    row["類2組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                                //} 
                                                                #endregion

                                                                // 系統編號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentID);
                                                                //班級名稱
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RefClass.ClassName);
                                                                //座號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.SeatNo);
                                                                //學號
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentNumber);
                                                                //姓名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.StudentName);
                                                                //成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.ExamScore);
                                                                //排名
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1);
                                                                //PR值
                                                                var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(sceTakeRecord.ExamScore) - 1m) / (decimal)ranks[key].Count);
                                                                if (pr == 0) pr = 1;
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                                                // 百分比
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(sceTakeRecord.ExamScore) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));
                                                                // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.CreditDec());
                                                                // 權數
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //分項類別
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Entry);
                                                                // 成績年級
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 必選修
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.Required ? "必修" : "選修");
                                                                // 校部定
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.RequiredBy);
                                                                // 科目成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.FinalScore);
                                                                //原始成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //補考成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //重修成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //手動調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //學年調整成績
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                //取得學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                                                // 不計學分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCredit ? "是" : "否");
                                                                //不算評分
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue(sceTakeRecord.NotIncludedInCalc ? "是" : "否");
                                                                //註記
                                                                cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                                                cs2RowIndex++;

                                                            }
                                                        }
                                                        #endregion


                                                    }
                                                    else
                                                    {//修課有該考試但沒有成績資料
                                                        var courseRecs = accessHelper.CourseHelper.GetCourse(courseID);
                                                        if (courseRecs.Count > 0)
                                                        {
                                                            var courseRec = courseRecs[0];
                                                            if (findInSemesterSubjectScore)
                                                            {
                                                                if (courseRec.SubjectLevel != "" + subjectNumber)
                                                                {
                                                                    continue;
                                                                }
                                                            }
                                                            findInExamScores = true;
                                                            if (!findInSemesterSubjectScore)
                                                            {
                                                                decimal level;
                                                                subjectNumber = decimal.TryParse(courseRec.SubjectLevel, out level) ? (decimal?)level : null;
                                                                row["科目名稱" + subjectIndex] = courseRec.Subject + GetNumber(subjectNumber);
                                                                row["學分數" + subjectIndex] = courseRec.CreditDec();
                                                            }
                                                            row["科目成績" + subjectIndex] = "未輸入";
                                                        }
                                                    }
                                                    if (studentRefExamSores.ContainsKey(studentID) && studentRefExamSores[studentID].ContainsKey(courseID))
                                                    {
                                                        row["前次成績" + subjectIndex] =
                                                            studentRefExamSores[studentID][courseID].SpecialCase == ""
                                                            ? ("" + studentRefExamSores[studentID][courseID].ExamScore)
                                                            : studentRefExamSores[studentID][courseID].SpecialCase;
                                                    }
                                                    studentExamSores[studentID][subjectName].Remove(courseID);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region 上學期學期成績
                                    if (conf.Semester == "2" && conf.WithPrevSemesterScore)
                                    {
                                        foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                        {
                                            if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是"
                                                && semesterSubjectScore.Subject == subjectName
                                                && semesterSubjectScore.Semester == 1
                                                && semesterSubjectScore.GradeYear == currentGradeYear)
                                            {
                                                findInSemester1SubjectScore = true;
                                                if (!findInSemesterSubjectScore
                                                    && !findInExamScores)
                                                {
                                                    decimal level;
                                                    subjectNumber = decimal.TryParse(semesterSubjectScore.Level, out level) ? (decimal?)level : null;
                                                    row["科目名稱" + subjectIndex] = semesterSubjectScore.Subject + GetNumber(subjectNumber);
                                                    row["學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
                                                    row["科目必選修" + subjectIndex] = semesterSubjectScore.Require ? "必修" : "選修";
                                                    row["科目校部定" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("修課校部訂");
                                                    row["科目註記" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("註記");
                                                }
                                                row["上學期科目取得學分" + subjectIndex] = semesterSubjectScore.Pass ? "是" : "否";
                                                row["上學期科目未取得學分註記" + subjectIndex] = semesterSubjectScore.Pass ? "" : "\f";

                                                //"原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績"
                                                if (semesterSubjectScore.Detail.GetAttribute("不需評分") != "是")
                                                {
                                                    row["上學期科目原始成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("原始成績");
                                                    row["上學期科目補考成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("補考成績");
                                                    row["上學期科目重修成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("重修成績");
                                                    row["上學期科目手動調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("擇優採計成績");
                                                    row["上學期科目學年調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("學年調整成績");
                                                    row["上學期科目成績" + subjectIndex] = semesterSubjectScore.Score;

                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("原始成績"))
                                                        row["上學期科目原始成績註記" + subjectIndex] = "\f";
                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("補考成績"))
                                                        row["上學期科目補考成績註記" + subjectIndex] = "\f";
                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("重修成績"))
                                                        row["上學期科目重修成績註記" + subjectIndex] = "\f";
                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("擇優採計成績"))
                                                        row["上學期科目手動成績註記" + subjectIndex] = "\f";
                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("學年調整成績"))
                                                        row["上學期科目學年成績註記" + subjectIndex] = "\f";

                                                    // 不及格
                                                    if (semesterSubjectScore.Score < scA)
                                                    {
                                                        // 可補考
                                                        if (semesterSubjectScore.Score >= scB)
                                                        {
                                                            row["上學期科目需要補考註記" + subjectIndex] = "\f";
                                                        }
                                                        else
                                                        {
                                                            // 不可補考需要重修
                                                            row["上學期科目需要重修註記" + subjectIndex] = "\f";
                                                        }
                                                    }
                                                }
                                                stuRec.SemesterSubjectScoreList.Remove(semesterSubjectScore);
                                                break;
                                            }
                                        }
                                    }
                                    #endregion
                                    #region 學年成績
                                    if (conf.Semester == "2" && conf.WithSchoolYearScore)
                                    {
                                        foreach (var schoolYearSubjectScore in stuRec.SchoolYearSubjectScoreList)
                                        {
                                            if (("" + schoolYearSubjectScore.SchoolYear) == conf.SchoolYear
                                                && schoolYearSubjectScore.Subject == subjectName)
                                            {
                                                if (!findInSemesterSubjectScore
                                                    && !findInSemester1SubjectScore
                                                    && !findInExamScores)
                                                {
                                                    row["科目名稱" + subjectIndex] = schoolYearSubjectScore.Subject;
                                                }
                                                row["學年科目成績" + subjectIndex] = schoolYearSubjectScore.Score;
                                                stuRec.SchoolYearSubjectScoreList.Remove(schoolYearSubjectScore);
                                                break;
                                            }
                                        }
                                    }
                                    #endregion
                                    subjectIndex++;
                                }
                                else
                                {
                                    //重要!!發現資料在樣板中印不下時一定要記錄起來，否則使用者自己不會去發現的
                                    if (!overflowRecords.Contains(stuRec))
                                        overflowRecords.Add(stuRec);
                                }
                            }
                            #endregion

                            #region 總分
                            if (studentPrintSubjectSum.ContainsKey(studentID))
                            {
                                #region 總分班排名
                                //總分班排名                                
                                key = "總分班排名" + stuRec.RefClass.ClassID;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是總分排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);

                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");

                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總分");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("班排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectSum[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectSum[studentID]) + 1);

                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectSum[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectSum[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));

                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;


                                }


                                #region 標準差、組距 (先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["總分班高標"] = analytics[key + "^^^高標"];
                                //    row["總分班均標"] = analytics[key + "^^^均標"];
                                //    row["總分班低標"] = analytics[key + "^^^低標"];
                                //    row["總分班標準差"] = analytics[key + "^^^標準差"];
                                //    row["總分班組距count90"] = analytics[key + "^^^count90"];
                                //    row["總分班組距count80"] = analytics[key + "^^^count80"];
                                //    row["總分班組距count70"] = analytics[key + "^^^count70"];
                                //    row["總分班組距count60"] = analytics[key + "^^^count60"];
                                //    row["總分班組距count50"] = analytics[key + "^^^count50"];
                                //    row["總分班組距count40"] = analytics[key + "^^^count40"];
                                //    row["總分班組距count30"] = analytics[key + "^^^count30"];
                                //    row["總分班組距count20"] = analytics[key + "^^^count20"];
                                //    row["總分班組距count10"] = analytics[key + "^^^count10"];
                                //    row["總分班組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["總分班組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["總分班組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["總分班組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["總分班組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["總分班組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["總分班組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["總分班組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["總分班組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["總分班組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["總分班組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["總分班組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["總分班組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["總分班組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["總分班組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["總分班組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["總分班組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["總分班組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["總分班組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion

                                #region 總分科排名
                                //總分科排名
                                key = "總分科排名" + stuRec.Department + "^^^" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {



                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是總分排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總分");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("科排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.Department);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectSum[studentID]);

                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectSum[studentID]) + 1);

                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectSum[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectSum[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));

                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;


                                }
                                #region  標準差、組距(先註解起來以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["總分科高標"] = analytics[key + "^^^高標"];
                                //    row["總分科均標"] = analytics[key + "^^^均標"];
                                //    row["總分科低標"] = analytics[key + "^^^低標"];
                                //    row["總分科標準差"] = analytics[key + "^^^標準差"];
                                //    row["總分科組距count90"] = analytics[key + "^^^count90"];
                                //    row["總分科組距count80"] = analytics[key + "^^^count80"];
                                //    row["總分科組距count70"] = analytics[key + "^^^count70"];
                                //    row["總分科組距count60"] = analytics[key + "^^^count60"];
                                //    row["總分科組距count50"] = analytics[key + "^^^count50"];
                                //    row["總分科組距count40"] = analytics[key + "^^^count40"];
                                //    row["總分科組距count30"] = analytics[key + "^^^count30"];
                                //    row["總分科組距count20"] = analytics[key + "^^^count20"];
                                //    row["總分科組距count10"] = analytics[key + "^^^count10"];
                                //    row["總分科組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["總分科組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["總分科組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["總分科組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["總分科組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["總分科組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["總分科組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["總分科組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["總分科組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["總分科組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["總分科組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["總分科組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["總分科組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["總分科組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["總分科組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["總分科組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["總分科組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["總分科組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["總分科組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion

                                #region 總分全校排名
                                //總分全校排名
                                key = "總分全校排名" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["總分全校排名"] = ranks[key].IndexOf(studentPrintSubjectSum[studentID]) + 1;
                                    //row["總分全校排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是總分排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總分");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("校排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.GradeYear + "年級");
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectSum[studentID]);

                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectSum[studentID]) + 1);

                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectSum[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectSum[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));

                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;

                                }

                                #region 標準差、組距(先註解起來以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["總分校高標"] = analytics[key + "^^^高標"];
                                //    row["總分校均標"] = analytics[key + "^^^均標"];
                                //    row["總分校低標"] = analytics[key + "^^^低標"];
                                //    row["總分校標準差"] = analytics[key + "^^^標準差"];
                                //    row["總分校組距count90"] = analytics[key + "^^^count90"];
                                //    row["總分校組距count80"] = analytics[key + "^^^count80"];
                                //    row["總分校組距count70"] = analytics[key + "^^^count70"];
                                //    row["總分校組距count60"] = analytics[key + "^^^count60"];
                                //    row["總分校組距count50"] = analytics[key + "^^^count50"];
                                //    row["總分校組距count40"] = analytics[key + "^^^count40"];
                                //    row["總分校組距count30"] = analytics[key + "^^^count30"];
                                //    row["總分校組距count20"] = analytics[key + "^^^count20"];
                                //    row["總分校組距count10"] = analytics[key + "^^^count10"];
                                //    row["總分校組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["總分校組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["總分校組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["總分校組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["總分校組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["總分校組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["總分校組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["總分校組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["總分校組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["總分校組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["總分校組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["總分校組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["總分校組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["總分校組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["總分校組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["總分校組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["總分校組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["總分校組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["總分校組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion
                                #endregion
                            }
                            #endregion

                            #region 平均
                            if (studentPrintSubjectAvg.ContainsKey(studentID))
                            {
                                #region 平均班排名
                                key = "平均班排名" + stuRec.RefClass.ClassID;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["平均班排名"] = ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1;
                                    //row["平均班排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("平均");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("班排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectAvg[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1);


                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectAvg[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;


                                }

                                #region 標準差、組距(先註解起來，以後可能會用得到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["平均班高標"] = analytics[key + "^^^高標"];
                                //    row["平均班均標"] = analytics[key + "^^^均標"];
                                //    row["平均班低標"] = analytics[key + "^^^低標"];
                                //    row["平均班標準差"] = analytics[key + "^^^標準差"];
                                //    row["平均班組距count90"] = analytics[key + "^^^count90"];
                                //    row["平均班組距count80"] = analytics[key + "^^^count80"];
                                //    row["平均班組距count70"] = analytics[key + "^^^count70"];
                                //    row["平均班組距count60"] = analytics[key + "^^^count60"];
                                //    row["平均班組距count50"] = analytics[key + "^^^count50"];
                                //    row["平均班組距count40"] = analytics[key + "^^^count40"];
                                //    row["平均班組距count30"] = analytics[key + "^^^count30"];
                                //    row["平均班組距count20"] = analytics[key + "^^^count20"];
                                //    row["平均班組距count10"] = analytics[key + "^^^count10"];
                                //    row["平均班組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["平均班組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["平均班組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["平均班組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["平均班組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["平均班組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["平均班組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["平均班組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["平均班組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["平均班組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["平均班組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["平均班組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["平均班組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["平均班組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["平均班組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["平均班組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["平均班組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["平均班組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["平均班組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion


                                #endregion

                                #region  平均科排名

                                key = "平均科排名" + stuRec.Department + "^^^" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["平均科排名"] = ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1;
                                    //row["平均科排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("平均");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("科排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.Department);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectAvg[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1);


                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectAvg[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;



                                }

                                #region 標準差、組距(先註解起來，以後可能會用得到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["平均科高標"] = analytics[key + "^^^高標"];
                                //    row["平均科均標"] = analytics[key + "^^^均標"];
                                //    row["平均科低標"] = analytics[key + "^^^低標"];
                                //    row["平均科標準差"] = analytics[key + "^^^標準差"];
                                //    row["平均科組距count90"] = analytics[key + "^^^count90"];
                                //    row["平均科組距count80"] = analytics[key + "^^^count80"];
                                //    row["平均科組距count70"] = analytics[key + "^^^count70"];
                                //    row["平均科組距count60"] = analytics[key + "^^^count60"];
                                //    row["平均科組距count50"] = analytics[key + "^^^count50"];
                                //    row["平均科組距count40"] = analytics[key + "^^^count40"];
                                //    row["平均科組距count30"] = analytics[key + "^^^count30"];
                                //    row["平均科組距count20"] = analytics[key + "^^^count20"];
                                //    row["平均科組距count10"] = analytics[key + "^^^count10"];
                                //    row["平均科組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["平均科組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["平均科組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["平均科組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["平均科組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["平均科組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["平均科組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["平均科組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["平均科組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["平均科組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["平均科組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["平均科組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["平均科組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["平均科組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["平均科組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["平均科組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["平均科組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["平均科組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["平均科組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion


                                #endregion

                                #region 平均校排名


                                key = "平均全校排名" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["平均全校排名"] = ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1;
                                    //row["平均全校排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("平均");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("校排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.GradeYear + "年級");
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectAvg[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1);


                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectAvg[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;


                                }

                                #region  標準差、組距(先註解起來，以後可能會用的到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["平均校高標"] = analytics[key + "^^^高標"];
                                //    row["平均校均標"] = analytics[key + "^^^均標"];
                                //    row["平均校低標"] = analytics[key + "^^^低標"];
                                //    row["平均校標準差"] = analytics[key + "^^^標準差"];
                                //    row["平均校組距count90"] = analytics[key + "^^^count90"];
                                //    row["平均校組距count80"] = analytics[key + "^^^count80"];
                                //    row["平均校組距count70"] = analytics[key + "^^^count70"];
                                //    row["平均校組距count60"] = analytics[key + "^^^count60"];
                                //    row["平均校組距count50"] = analytics[key + "^^^count50"];
                                //    row["平均校組距count40"] = analytics[key + "^^^count40"];
                                //    row["平均校組距count30"] = analytics[key + "^^^count30"];
                                //    row["平均校組距count20"] = analytics[key + "^^^count20"];
                                //    row["平均校組距count10"] = analytics[key + "^^^count10"];
                                //    row["平均校組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["平均校組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["平均校組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["平均校組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["平均校組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["平均校組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["平均校組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["平均校組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["平均校組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["平均校組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["平均校組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["平均校組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["平均校組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["平均校組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["平均校組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["平均校組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["平均校組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["平均校組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["平均校組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion


                                #endregion
                            }
                            #endregion

                            #region 加權總分
                            if (studentPrintSubjectSumW.ContainsKey(studentID))
                            {
                                //row["加權總分"] = studentPrintSubjectSumW[studentID];

                                #region 加權總分班排名

                                key = "加權總分班排名" + stuRec.RefClass.ClassID;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["加權總分班排名"] = ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1;
                                    //row["加權總分班排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是加權總分排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權總分");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("班排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectSumW[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1);


                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectSumW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;

                                }


                                #region 標準差、組距(先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權總分班高標"] = analytics[key + "^^^高標"];
                                //    row["加權總分班均標"] = analytics[key + "^^^均標"];
                                //    row["加權總分班低標"] = analytics[key + "^^^低標"];
                                //    row["加權總分班標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權總分班組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權總分班組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權總分班組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權總分班組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權總分班組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權總分班組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權總分班組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權總分班組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權總分班組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權總分班組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權總分班組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權總分班組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權總分班組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權總分班組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權總分班組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權總分班組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權總分班組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權總分班組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權總分班組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權總分班組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權總分班組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權總分班組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權總分班組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權總分班組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權總分班組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權總分班組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權總分班組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權總分班組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion


                                #region  加權總分科排名

                                key = "加權總分科排名" + stuRec.Department + "^^^" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["加權總分科排名"] = ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1;
                                    //row["加權總分科排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是加權總分排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權總分");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("科排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.Department);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectSumW[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1);

                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectSumW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;

                                }

                                #region 標準差、組距(先註解起來，以後可能會用得到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權總分科高標"] = analytics[key + "^^^高標"];
                                //    row["加權總分科均標"] = analytics[key + "^^^均標"];
                                //    row["加權總分科低標"] = analytics[key + "^^^低標"];
                                //    row["加權總分科標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權總分科組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權總分科組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權總分科組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權總分科組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權總分科組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權總分科組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權總分科組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權總分科組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權總分科組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權總分科組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權總分科組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權總分科組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權總分科組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權總分科組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權總分科組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權總分科組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權總分科組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權總分科組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權總分科組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權總分科組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權總分科組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權總分科組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權總分科組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權總分科組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權總分科組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權總分科組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權總分科組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權總分科組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion


                                #endregion

                                #region 加權總分全校排名

                                key = "加權總分全校排名" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    row["加權總分全校排名"] = ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1;
                                    row["加權總分全校排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是加權總分排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權總分");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("校排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.GradeYear + "年級");
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectSumW[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1);


                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectSumW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;

                                }


                                #region 標準差、組距(先註解起來，以後可能會用得到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權總分校高標"] = analytics[key + "^^^高標"];
                                //    row["加權總分校均標"] = analytics[key + "^^^均標"];
                                //    row["加權總分校低標"] = analytics[key + "^^^低標"];
                                //    row["加權總分校標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權總分校組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權總分校組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權總分校組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權總分校組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權總分校組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權總分校組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權總分校組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權總分校組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權總分校組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權總分校組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權總分校組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權總分校組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權總分校組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權總分校組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權總分校組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權總分校組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權總分校組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權總分校組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權總分校組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權總分校組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權總分校組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權總分校組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權總分校組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權總分校組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權總分校組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權總分校組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權總分校組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權總分校組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion


                                #endregion
                            }


                            #endregion

                            #region 加權平均
                            if (studentPrintSubjectAvgW.ContainsKey(studentID))
                            {
                                //row["加權平均"] = studentPrintSubjectAvgW[studentID];


                                #region 加權平均班排名

                                key = "加權平均班排名" + stuRec.RefClass.ClassID;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["加權平均班排名"] = ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1;
                                    //row["加權平均班排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權平均");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("班排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectAvgW[studentID]);

                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1);

                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectAvgW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;




                                }

                                #region 標準差、組距(先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均班高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均班均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均班低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均班標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均班組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均班組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均班組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均班組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均班組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均班組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均班組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均班組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均班組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均班組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均班組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均班組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均班組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均班組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均班組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均班組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均班組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均班組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均班組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均班組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均班組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均班組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均班組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均班組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均班組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均班組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均班組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均班組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion

                                #region 加權平均科排名

                                key = "加權平均科排名" + stuRec.Department + "^^^" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["加權平均科排名"] = ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1;
                                    //row["加權平均科排名母數"] = ranks[key].Count;

                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權平均");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("科排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.Department);
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectAvgW[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1);


                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectAvgW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;



                                }

                                #region 標準差、組距(先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均科高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均科均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均科低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均科標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均科組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均科組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均科組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均科組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均科組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均科組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均科組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均科組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均科組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均科組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均科組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均科組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均科組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均科組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均科組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均科組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均科組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均科組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均科組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均科組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均科組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均科組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均科組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均科組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均科組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均科組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均科組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均科組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion

                                #region 加權平均校排名

                                key = "加權平均全校排名" + gradeYear;
                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                {
                                    //row["加權平均全校排名"] = ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1;
                                    //row["加權平均全校排名母數"] = ranks[key].Count;
                                    int cs2ColIndex = 0;

                                    // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                    //母群Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                    //類別Key
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                    //科目名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權平均");
                                    //科目級別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //母群類型	
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("校排名");
                                    //母群名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.GradeYear + "年級");
                                    //總人數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                    //頂標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //前標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                    //均標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                    //後標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                    //底標
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    // 系統編號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                    //班級名稱
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                    //座號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                    //學號
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                    //姓名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                    //成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentPrintSubjectAvgW[studentID]);
                                    //排名
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1);

                                    //PR值
                                    var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentPrintSubjectAvgW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                    if (pr == 0) pr = 1;
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                    // 百分比
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));

                                    // 學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 權數
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //分項類別
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 成績年級
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 必選修
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 校部定
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 科目成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //原始成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //補考成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //重修成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //手動調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //學年調整成績
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //取得學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    // 不計學分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //不算評分
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                    //註記
                                    cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                    cs2RowIndex++;



                                }

                                #region 標準差、組距(先註解起來，以後可能會用的到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均校高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均校均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均校低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均校標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均校組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均校組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均校組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均校組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均校組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均校組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均校組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均校組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均校組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均校組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均校組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均校組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均校組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均校組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均校組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均校組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均校組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均校組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均校組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均校組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均校組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均校組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均校組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均校組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均校組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均校組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均校組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均校組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion

                            }
                            #endregion

                            #region  回歸科目
                            //2016/7/5 穎驊新增，回歸科目，班科校填值
                            if (studentIDRegressSubjectAvgW.ContainsKey(studentID))
                            {
                                #region  回歸科目班排名

                                foreach (String RegressSubject in studentIDRegressSubjectAvgW[studentID].Keys)
                                {

                                    key = "回歸科目" + "_" + RegressSubject + "_" + "班排名" + stuRec.RefClass.ClassID;


                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                    {


                                        int cs2ColIndex = 0;

                                        // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("回歸科目");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(RegressSubject);
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("班排名");
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentIDRegressSubjectAvgW[studentID][RegressSubject]);

                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) + 1);

                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;


                                    }

                                }

                                #region 標準差、組距(先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均班高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均班均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均班低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均班標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均班組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均班組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均班組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均班組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均班組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均班組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均班組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均班組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均班組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均班組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均班組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均班組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均班組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均班組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均班組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均班組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均班組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均班組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均班組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均班組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均班組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均班組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均班組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均班組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均班組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均班組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均班組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均班組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion

                                #region  回歸科目加權平均科排名


                                foreach (String RegressSubject in studentIDRegressSubjectAvgW[studentID].Keys)
                                {

                                    key = "回歸科目" + "_" + RegressSubject + "_" + "科排名" + "_" + stuRec.Department + "^^^" + gradeYear;


                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                    {
                                        //row["加權平均科排名"] = ranks[key].IndexOf(studentRegressSubjectAvgW[studentID]) + 1;
                                        //row["加權平均科排名母數"] = ranks[key].Count;

                                        int cs2ColIndex = 0;

                                        // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("回歸科目");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(RegressSubject);
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("科排名");
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.Department);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentIDRegressSubjectAvgW[studentID][RegressSubject]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;



                                    }

                                }

                                #region 標準差、組距(先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均科高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均科均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均科低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均科標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均科組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均科組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均科組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均科組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均科組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均科組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均科組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均科組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均科組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均科組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均科組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均科組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均科組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均科組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均科組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均科組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均科組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均科組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均科組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均科組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均科組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均科組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均科組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均科組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均科組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均科組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均科組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均科組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion

                                #region  回歸科目加權平均校排名


                                foreach (String RegressSubject in studentIDRegressSubjectAvgW[studentID].Keys)
                                {

                                    key = "回歸科目" + "_" + RegressSubject + "_" + "全校排名" + gradeYear;
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                    {
                                        //row["加權平均全校排名"] = ranks[key].IndexOf(studentRegressSubjectAvgW[studentID]) + 1;
                                        //row["加權平均全校排名母數"] = ranks[key].Count;
                                        int cs2ColIndex = 0;

                                        // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("回歸科目");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(RegressSubject);
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("校排名");
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.GradeYear + "年級");
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentIDRegressSubjectAvgW[studentID][RegressSubject]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) + 1);

                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentIDRegressSubjectAvgW[studentID][RegressSubject]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));

                                        // 學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;



                                    }

                                }

                                #region 標準差、組距(先註解起來，以後可能會用的到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均校高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均校均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均校低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均校標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均校組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均校組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均校組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均校組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均校組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均校組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均校組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均校組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均校組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均校組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均校組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均校組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均校組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均校組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均校組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均校組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均校組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均校組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均校組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均校組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均校組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均校組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均校組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均校組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均校組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均校組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均校組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均校組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion
                            }
                            #endregion

                            //2016/7/7 穎驊新增，回歸科目，類別1填值
                            if (studentIDRegressSubjectAvgW_Tag1.ContainsKey(studentID))
                            {
                                #region  回歸科目類別1排名

                                foreach (String RegressSubject in studentIDRegressSubjectAvgW_Tag1[studentID].Keys)
                                {

                                    key = "類別1回歸科目排名" + "_" + RegressSubject + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];


                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                    {

                                        int cs2ColIndex = 0;

                                        // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("回歸科目");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(RegressSubject);
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");


                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別1");

                                        var gName = "";
                                        if (studentTag1Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag1Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentIDRegressSubjectAvgW_Tag1[studentID][RegressSubject]);

                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentIDRegressSubjectAvgW_Tag1[studentID][RegressSubject]) + 1);

                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentIDRegressSubjectAvgW_Tag1[studentID][RegressSubject]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentIDRegressSubjectAvgW_Tag1[studentID][RegressSubject]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;


                                    }

                                }

                                #region 標準差、組距(先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均班高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均班均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均班低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均班標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均班組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均班組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均班組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均班組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均班組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均班組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均班組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均班組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均班組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均班組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均班組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均班組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均班組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均班組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均班組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均班組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均班組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均班組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均班組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均班組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均班組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均班組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均班組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均班組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均班組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均班組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均班組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均班組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion
                            }

                            //2016/7/7 穎驊新增，回歸科目，類別2填值
                            if (studentIDRegressSubjectAvgW_Tag2.ContainsKey(studentID))
                            {
                                #region  回歸科目類別2排名

                                foreach (String RegressSubject in studentIDRegressSubjectAvgW_Tag2[studentID].Keys)
                                {

                                    key = "類別2回歸科目排名" + "_" + RegressSubject + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];


                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                    {
                                        //row["加權平均班排名"] = ranks[key].IndexOf(studentRegressSubjectAvgW[studentID]) + 1;
                                        //row["加權平均班排名母數"] = ranks[key].Count;



                                        int cs2ColIndex = 0;

                                        // 特別註記，由於這邊是加權平均排名，會有一堆欄位是空白的，純屬正常現象，不必擔心
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("回歸科目");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(RegressSubject);
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別2");


                                        var gName = "";
                                        if (studentTag2Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag2Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }

                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentIDRegressSubjectAvgW_Tag2[studentID][RegressSubject]);

                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentIDRegressSubjectAvgW_Tag2[studentID][RegressSubject]) + 1);

                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentIDRegressSubjectAvgW_Tag2[studentID][RegressSubject]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentIDRegressSubjectAvgW_Tag2[studentID][RegressSubject]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;


                                    }

                                }

                                #region 標準差、組距(先註解起來，以後可能會用到)
                                //if (rankStudents.ContainsKey(key))
                                //{
                                //    row["加權平均班高標"] = analytics[key + "^^^高標"];
                                //    row["加權平均班均標"] = analytics[key + "^^^均標"];
                                //    row["加權平均班低標"] = analytics[key + "^^^低標"];
                                //    row["加權平均班標準差"] = analytics[key + "^^^標準差"];
                                //    row["加權平均班組距count90"] = analytics[key + "^^^count90"];
                                //    row["加權平均班組距count80"] = analytics[key + "^^^count80"];
                                //    row["加權平均班組距count70"] = analytics[key + "^^^count70"];
                                //    row["加權平均班組距count60"] = analytics[key + "^^^count60"];
                                //    row["加權平均班組距count50"] = analytics[key + "^^^count50"];
                                //    row["加權平均班組距count40"] = analytics[key + "^^^count40"];
                                //    row["加權平均班組距count30"] = analytics[key + "^^^count30"];
                                //    row["加權平均班組距count20"] = analytics[key + "^^^count20"];
                                //    row["加權平均班組距count10"] = analytics[key + "^^^count10"];
                                //    row["加權平均班組距count100Up"] = analytics[key + "^^^count100Up"];
                                //    row["加權平均班組距count90Up"] = analytics[key + "^^^count90Up"];
                                //    row["加權平均班組距count80Up"] = analytics[key + "^^^count80Up"];
                                //    row["加權平均班組距count70Up"] = analytics[key + "^^^count70Up"];
                                //    row["加權平均班組距count60Up"] = analytics[key + "^^^count60Up"];
                                //    row["加權平均班組距count50Up"] = analytics[key + "^^^count50Up"];
                                //    row["加權平均班組距count40Up"] = analytics[key + "^^^count40Up"];
                                //    row["加權平均班組距count30Up"] = analytics[key + "^^^count30Up"];
                                //    row["加權平均班組距count20Up"] = analytics[key + "^^^count20Up"];
                                //    row["加權平均班組距count10Up"] = analytics[key + "^^^count10Up"];
                                //    row["加權平均班組距count90Down"] = analytics[key + "^^^count90Down"];
                                //    row["加權平均班組距count80Down"] = analytics[key + "^^^count80Down"];
                                //    row["加權平均班組距count70Down"] = analytics[key + "^^^count70Down"];
                                //    row["加權平均班組距count60Down"] = analytics[key + "^^^count60Down"];
                                //    row["加權平均班組距count50Down"] = analytics[key + "^^^count50Down"];
                                //    row["加權平均班組距count40Down"] = analytics[key + "^^^count40Down"];
                                //    row["加權平均班組距count30Down"] = analytics[key + "^^^count30Down"];
                                //    row["加權平均班組距count20Down"] = analytics[key + "^^^count20Down"];
                                //    row["加權平均班組距count10Down"] = analytics[key + "^^^count10Down"];
                                //} 
                                #endregion

                                #endregion
                            }

                            #region 類別1綜合成績
                            if (studentTag1Group.ContainsKey(studentID))
                            {
                                //foreach (var tag in studentTags[studentID])
                                //{
                                //    if (tag.RefTagID == studentTag1Group[studentID])
                                //    {
                                //        row["類別排名1"] = tag.Name;


                                //    }
                                //}

                                #region 類別1總分

                                if (studentTag1SubjectSum.ContainsKey(studentID))
                                {
                                    //row["類別1總分"] = studentTag1SubjectSum[studentID];

                                    key = "類別1總分排名" + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {
                                        //row["類別1總分排名"] = ranks[key].IndexOf(studentTag1SubjectSum[studentID]) + 1;
                                        //row["類別1總分排名母數"] = ranks[key].Count;

                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總分");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別1");

                                        var gName = "";
                                        if (studentTag1Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag1Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag1SubjectSum[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag1SubjectSum[studentID]) + 1);

                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag1SubjectSum[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag1SubjectSum[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;


                                    }


                                    #region 標準差、組距(先註解起來，以後可能會用到)
                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類1總分高標"] = analytics[key + "^^^高標"];
                                    //    row["類1總分均標"] = analytics[key + "^^^均標"];
                                    //    row["類1總分低標"] = analytics[key + "^^^低標"];
                                    //    row["類1總分標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類1總分組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類1總分組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類1總分組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類1總分組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類1總分組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類1總分組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類1總分組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類1總分組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類1總分組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類1總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類1總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類1總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類1總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類1總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類1總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類1總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類1總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類1總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類1總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類1總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類1總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類1總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類1總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類1總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類1總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類1總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類1總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類1總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //} 
                                    #endregion

                                }

                                #endregion

                                #region 類別1平均

                                if (studentTag1SubjectAvg.ContainsKey(studentID))
                                {
                                    //row["類別1平均"] = studentTag1SubjectAvg[studentID];
                                    key = "類別1平均排名" + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {
                                        //row["類別1平均排名"] = ranks[key].IndexOf(studentTag1SubjectAvg[studentID]) + 1; ;
                                        //row["類別1平均排名母數"] = ranks[key].Count;


                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("平均");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別1");

                                        var gName = "";
                                        if (studentTag1Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag1Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag1SubjectAvg[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag1SubjectAvg[studentID]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag1SubjectAvg[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag1SubjectAvg[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;

                                    }

                                    #region 標準差、組距(先註解起來，以後可能會用得到)
                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類1平均高標"] = analytics[key + "^^^高標"];
                                    //    row["類1平均均標"] = analytics[key + "^^^均標"];
                                    //    row["類1平均低標"] = analytics[key + "^^^低標"];
                                    //    row["類1平均標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類1平均組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類1平均組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類1平均組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類1平均組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類1平均組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類1平均組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類1平均組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類1平均組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類1平均組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類1平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類1平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類1平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類1平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類1平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類1平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類1平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類1平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類1平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類1平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類1平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類1平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類1平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類1平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類1平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類1平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類1平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類1平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類1平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //} 
                                    #endregion
                                }

                                #endregion

                                #region 類別1加權總分

                                if (studentTag1SubjectSumW.ContainsKey(studentID))
                                {
                                    //row["類別1加權總分"] = studentTag1SubjectSumW[studentID];

                                    key = "類別1加權總分排名" + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {
                                        //row["類別1加權總分排名"] = ranks[key].IndexOf(studentTag1SubjectSumW[studentID]) + 1; ;
                                        //row["類別1加權總分排名母數"] = ranks[key].Count;

                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權總分");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別1");

                                        var gName = "";
                                        if (studentTag1Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag1Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag1SubjectSumW[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag1SubjectSumW[studentID]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag1SubjectSumW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag1SubjectSumW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;




                                    }

                                    #region 標準差、組距(先註解起來，以後可能會用得到)
                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類1加權總分高標"] = analytics[key + "^^^高標"];
                                    //    row["類1加權總分均標"] = analytics[key + "^^^均標"];
                                    //    row["類1加權總分低標"] = analytics[key + "^^^低標"];
                                    //    row["類1加權總分標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類1加權總分組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類1加權總分組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類1加權總分組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類1加權總分組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類1加權總分組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類1加權總分組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類1加權總分組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類1加權總分組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類1加權總分組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類1加權總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類1加權總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類1加權總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類1加權總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類1加權總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類1加權總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類1加權總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類1加權總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類1加權總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類1加權總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類1加權總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類1加權總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類1加權總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類1加權總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類1加權總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類1加權總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類1加權總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類1加權總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類1加權總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //} 
                                    #endregion



                                }
                                #endregion

                                #region 類1加權平均


                                if (studentTag1SubjectAvgW.ContainsKey(studentID))
                                {
                                    //row["類別1加權平均"] = studentTag1SubjectAvgW[studentID];

                                    key = "類別1加權平均排名" + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {
                                        //row["類別1加權平均排名"] = ranks[key].IndexOf(studentTag1SubjectAvgW[studentID]) + 1; ;
                                        //row["類別1加權平均排名母數"] = ranks[key].Count;

                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權平均");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別1");

                                        var gName = "";
                                        if (studentTag1Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag1Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag1SubjectAvgW[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag1SubjectAvgW[studentID]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag1SubjectAvgW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag1SubjectAvgW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));

                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;



                                    }
                                    #region 標準差、組距(先註解起來，以後可能會用得到)

                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類1加權平均高標"] = analytics[key + "^^^高標"];
                                    //    row["類1加權平均均標"] = analytics[key + "^^^均標"];
                                    //    row["類1加權平均低標"] = analytics[key + "^^^低標"];
                                    //    row["類1加權平均標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類1加權平均組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類1加權平均組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類1加權平均組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類1加權平均組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類1加權平均組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類1加權平均組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類1加權平均組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類1加權平均組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類1加權平均組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類1加權平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類1加權平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類1加權平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類1加權平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類1加權平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類1加權平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類1加權平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類1加權平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類1加權平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類1加權平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類1加權平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類1加權平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類1加權平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類1加權平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類1加權平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類1加權平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類1加權平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類1加權平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類1加權平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //} 
                                    #endregion




                                }
                                #endregion
                            }
                            #endregion

                            #region 類別2綜合成績
                            if (studentTag2Group.ContainsKey(studentID))
                            {
                                //foreach (var tag in studentTags[studentID])
                                //{
                                //    if (tag.RefTagID == studentTag2Group[studentID])
                                //    {
                                //        row["類別排名2"] = tag.Name;
                                //    }
                                //}

                                #region 類別2總分

                                if (studentTag2SubjectSum.ContainsKey(studentID))
                                {
                                    //row["類別2總分"] = studentTag2SubjectSum[studentID];

                                    key = "類別2總分排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {

                                        //row["類別2總分排名"] = ranks[key].IndexOf(studentTag2SubjectSum[studentID]) + 1;
                                        //row["類別2總分排名母數"] = ranks[key].Count;



                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總分");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別2");

                                        var gName = "";
                                        if (studentTag2Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag2Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag2SubjectSum[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag2SubjectSum[studentID]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag2SubjectSum[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag2SubjectSum[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));



                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;

                                    }


                                    #region 標準差、組距(先註解起來，以後可能會用的到)
                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類2總分高標"] = analytics[key + "^^^高標"];
                                    //    row["類2總分均標"] = analytics[key + "^^^均標"];
                                    //    row["類2總分低標"] = analytics[key + "^^^低標"];
                                    //    row["類2總分標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類2總分組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類2總分組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類2總分組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類2總分組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類2總分組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類2總分組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類2總分組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類2總分組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類2總分組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類2總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類2總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類2總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類2總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類2總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類2總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類2總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類2總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類2總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類2總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類2總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類2總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類2總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類2總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類2總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類2總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類2總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類2總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類2總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //} 
                                    #endregion


                                }
                                #endregion


                                #region 類別2平均

                                if (studentTag2SubjectAvg.ContainsKey(studentID))
                                {
                                    //row["類別2平均"] = studentTag2SubjectAvg[studentID];
                                    key = "類別2平均排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {
                                        //row["類別2平均排名"] = ranks[key].IndexOf(studentTag2SubjectAvg[studentID]) + 1; ;
                                        //row["類別2平均排名母數"] = ranks[key].Count;



                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("平均");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別2");

                                        var gName = "";
                                        if (studentTag2Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag2Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag2SubjectAvg[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag2SubjectAvg[studentID]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag2SubjectAvg[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag2SubjectAvg[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;



                                    }

                                    #region 標準差、組距(先註解起來，以後可能會用得到)
                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類2平均高標"] = analytics[key + "^^^高標"];
                                    //    row["類2平均均標"] = analytics[key + "^^^均標"];
                                    //    row["類2平均低標"] = analytics[key + "^^^低標"];
                                    //    row["類2平均標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類2平均組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類2平均組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類2平均組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類2平均組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類2平均組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類2平均組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類2平均組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類2平均組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類2平均組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類2平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類2平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類2平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類2平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類2平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類2平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類2平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類2平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類2平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類2平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類2平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類2平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類2平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類2平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類2平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類2平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類2平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類2平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類2平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //}

                                    #endregion

                                }

                                #endregion

                                #region 類別2加權總分

                                if (studentTag2SubjectSumW.ContainsKey(studentID))
                                {
                                    //row["類別2加權總分"] = studentTag2SubjectSumW[studentID];
                                    key = "類別2加權總分排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {
                                        //row["類別2加權總分排名"] = ranks[key].IndexOf(studentTag2SubjectSumW[studentID]) + 1; ;
                                        //row["類別2加權總分排名母數"] = ranks[key].Count;


                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權總分");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別2");

                                        var gName = "";
                                        if (studentTag2Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag2Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag2SubjectSumW[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag2SubjectSumW[studentID]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag2SubjectSumW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag2SubjectSumW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));



                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;



                                    }
                                    #region 標準差、組距(先註解起來，以後可能會用得到)

                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類2加權總分高標"] = analytics[key + "^^^高標"];
                                    //    row["類2加權總分均標"] = analytics[key + "^^^均標"];
                                    //    row["類2加權總分低標"] = analytics[key + "^^^低標"];
                                    //    row["類2加權總分標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類2加權總分組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類2加權總分組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類2加權總分組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類2加權總分組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類2加權總分組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類2加權總分組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類2加權總分組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類2加權總分組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類2加權總分組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類2加權總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類2加權總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類2加權總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類2加權總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類2加權總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類2加權總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類2加權總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類2加權總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類2加權總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類2加權總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類2加權總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類2加權總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類2加權總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類2加權總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類2加權總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類2加權總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類2加權總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類2加權總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類2加權總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //} 
                                    #endregion
                                }

                                #endregion

                                #region 類別2加權平均

                                if (studentTag2SubjectAvgW.ContainsKey(studentID))
                                {
                                    row["類別2加權平均"] = studentTag2SubjectAvgW[studentID];
                                    key = "類別2加權平均排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                    if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                    {
                                        row["類別2加權平均排名"] = ranks[key].IndexOf(studentTag2SubjectAvgW[studentID]) + 1; ;
                                        row["類別2加權平均排名母數"] = ranks[key].Count;


                                        // 由於這邊是綜合排名，會有一堆資料欄位空白很正常
                                        int cs2ColIndex = 0;
                                        //母群Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(key);
                                        //類別Key
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("總計");
                                        //科目名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("加權平均");
                                        //科目級別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //母群類型	
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("類別2");

                                        var gName = "";
                                        if (studentTag2Group.ContainsKey(studentID))
                                        {
                                            foreach (var tag in studentTags[studentID])
                                            {
                                                if (tag.RefTagID == studentTag2Group[studentID])
                                                {
                                                    gName = tag.Name;
                                                }
                                            }
                                        }
                                        //母群名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(gName);
                                        //總人數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].Count);
                                        //頂標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //前標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^高標"]);
                                        //均標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^均標"]);
                                        //後標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(analytics[key + "^^^低標"]);
                                        //底標
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        // 系統編號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentID);
                                        //班級名稱
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.RefClass.ClassName);
                                        //座號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.SeatNo);
                                        //學號
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentNumber);
                                        //姓名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(stuRec.StudentName);
                                        //成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(studentTag2SubjectAvgW[studentID]);
                                        //排名
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(ranks[key].IndexOf(studentTag2SubjectAvgW[studentID]) + 1);


                                        //PR值
                                        var pr = Math.Floor(100m * ((decimal)ranks[key].Count - (decimal)ranks[key].LastIndexOf(studentTag2SubjectAvgW[studentID]) - 1m) / (decimal)ranks[key].Count);
                                        if (pr == 0) pr = 1;
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(pr);
                                        // 百分比
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue(System.Math.Round(100m * (decimal)ranks[key].IndexOf(studentTag2SubjectAvgW[studentID]) / (decimal)ranks[key].Count + 0.5m, 0, MidpointRounding.AwayFromZero));


                                        // 學分(使用新的CreditDec()方法，可以得到decimal的學分數，避免以後出現0.學分5這種小數位學分數的悲劇)
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 權數
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //分項類別
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 成績年級
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 必選修
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 校部定
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 科目成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //原始成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //補考成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //重修成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //手動調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //學年調整成績
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //取得學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        // 不計學分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //不算評分
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");
                                        //註記
                                        cs2[cs2RowIndex, cs2ColIndex++].PutValue("");

                                        cs2RowIndex++;



                                    }

                                    #region 標準差、組距(先註解起來以後可能會用的到)
                                    //if (rankStudents.ContainsKey(key))
                                    //{
                                    //    row["類2加權平均高標"] = analytics[key + "^^^高標"];
                                    //    row["類2加權平均均標"] = analytics[key + "^^^均標"];
                                    //    row["類2加權平均低標"] = analytics[key + "^^^低標"];
                                    //    row["類2加權平均標準差"] = analytics[key + "^^^標準差"];
                                    //    row["類2加權平均組距count90"] = analytics[key + "^^^count90"];
                                    //    row["類2加權平均組距count80"] = analytics[key + "^^^count80"];
                                    //    row["類2加權平均組距count70"] = analytics[key + "^^^count70"];
                                    //    row["類2加權平均組距count60"] = analytics[key + "^^^count60"];
                                    //    row["類2加權平均組距count50"] = analytics[key + "^^^count50"];
                                    //    row["類2加權平均組距count40"] = analytics[key + "^^^count40"];
                                    //    row["類2加權平均組距count30"] = analytics[key + "^^^count30"];
                                    //    row["類2加權平均組距count20"] = analytics[key + "^^^count20"];
                                    //    row["類2加權平均組距count10"] = analytics[key + "^^^count10"];
                                    //    row["類2加權平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                    //    row["類2加權平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                    //    row["類2加權平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                    //    row["類2加權平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                    //    row["類2加權平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                    //    row["類2加權平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                    //    row["類2加權平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                    //    row["類2加權平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                    //    row["類2加權平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                    //    row["類2加權平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                    //    row["類2加權平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                    //    row["類2加權平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                    //    row["類2加權平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                    //    row["類2加權平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                    //    row["類2加權平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                    //    row["類2加權平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                    //    row["類2加權平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                    //    row["類2加權平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                    //    row["類2加權平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                    //} 
                                    #endregion
                                }

                                #endregion

                            }
                            #endregion
                            #endregion

                            progressCount++;

                            bkw.ReportProgress(70 + progressCount * 20 / selectedStudents.Count);

                        }

                        #endregion

                        bkw.ReportProgress(90);


                        bkw.ReportProgress(100);


                        //document = conf.Template;

                        //document.MailMerge.Execute(table);

                    }
                    catch (Exception exception)
                    {
                        exc = exception;
                    }

                    e.Result = wb;
                };


                bkw.RunWorkerAsync();
            }
        }


        static private void ImportWorkbook(Workbook wb)
        {
            Cells cs0 = wb.Worksheets[0].Cells;
            Cells cs1 = wb.Worksheets[1].Cells;
            Cells cs2 = wb.Worksheets[2].Cells;


            #region 處理Excel cs0 第一頁 Rank Table

            Rank rank = new Rank();
            int cindex = 0;
            rank.SchoolYear = int.Parse(cs0[2, cindex++].StringValue);
            rank.Semester = int.Parse(cs0[2, cindex++].StringValue);
            rank.GradeYear = int.Parse(cs0[2, cindex++].StringValue);
            rank.RankType = cs0[2, cindex++].StringValue;

            if (cs0[2, cindex].Value != null && cs0[2, cindex].StringValue.Trim() != "")
                rank.RankSequence = int.Parse(cs0[2, cindex].StringValue);
            else
                rank.RankSequence = null;


            cindex++;
            rank.DisplayName = cs0[2, cindex++].StringValue;
            rank.Memo = cs0[2, cindex++].StringValue;
            rank.CreateTime = DateTime.Now;

            rank.Active = true;

            #region disable舊資料
            FISCA.UDT.AccessHelper accessHelper = new FISCA.UDT.AccessHelper();

            var list = accessHelper.Select<Rank>(string.Format("school_year={0} and semester={1} and grade_year={2} and rank_type='{3}' and rank_sequence={4}",
                rank.SchoolYear,
                rank.Semester,
                rank.GradeYear,
                rank.RankType,
                rank.RankSequence == null ? "null" : ("" + rank.RankSequence)
                ));
            foreach (var item in list)
            {
                item.Active = false;
            }
            list.SaveAll();
            #endregion
            rank.Save();
            
            #endregion



            #region 處理Excel cs1 第二頁 RankStudent Table

            List<RankStudent> rankStudentList = new List<RankStudent>();



            for (int x = 2; x < cs1.Rows.Count - 1; x++)
            {
                int col_index = 0;

                var stuRec = new RankStudent();

                stuRec.RefStudentID = int.Parse(cs1[x, col_index++].StringValue);

                stuRec.RefRankID = int.Parse(rank.UID);

                stuRec.StudentNumber = int.Parse(cs1[x, col_index++].StringValue);

                stuRec.Name = cs1[x, col_index++].StringValue;

                stuRec.ClassName = cs1[x, col_index++].StringValue;

                stuRec.SeatNo = cs1[x, col_index++].StringValue;

                stuRec.HRTeacherName = cs1[x, col_index++].StringValue;

                stuRec.DeptName = cs1[x, col_index++].StringValue;

                stuRec.CatName1 = cs1[x, col_index++].StringValue;

                stuRec.CatName2 = cs1[x, col_index++].StringValue;

                //判斷是、否 產生bool值
                if (cs1[x, col_index].StringValue == "是")
                {
                    stuRec.RankInclude = true;
                    col_index++;
                }
                else {
                    stuRec.RankInclude = false;
                    col_index++;                
                }

               
                if (cs1[x, col_index].Value != null && cs1[x, col_index].StringValue.Trim() != "")
                {
                    stuRec.ClassID = int.Parse(cs1[x, col_index++].StringValue);
                }
                else
                {
                    stuRec.ClassID = null;

                    col_index++;
                }

                if (cs1[x, col_index].Value != null && cs1[x, col_index].StringValue.Trim() != "")
                {
                    stuRec.Cat1ID = int.Parse(cs1[x, col_index++].StringValue);
                }
                else
                {
                    stuRec.Cat1ID = null;

                    col_index++;
                }


                if (cs1[x, col_index].Value != null && cs1[x, col_index].StringValue.Trim() != "")
                {
                    stuRec.Cat2ID = int.Parse(cs1[x, col_index++].StringValue);
                }
                else
                {
                    stuRec.Cat2ID = null;

                    col_index++;
                }

                rankStudentList.Add(stuRec);
            }


            rankStudentList.SaveAll();

            
            #endregion

            #region 處理Excel cs2 第三頁 RankGroup Table

            Dictionary<String, RankGroup> RankGroup_Dict = new Dictionary<string, RankGroup>();

            for (int x = 2; x < cs2.Rows.Count - 1; x++)
            {
                int col_index = 0;


                //2016/7/13 穎驊筆記，以下為將每一種的母群狀況加進去，
                //重覆的不加，舉例而言，以任一科目的的班排名，可能有50列在Excel，但其有一樣的母群狀態(平均、高均低標等等)，此時只需要加其中一項代表就可以了

                if (!RankGroup_Dict.ContainsKey(cs2[x, col_index].StringValue)) {

                    RankGroup_Dict.Add(cs2[x, 0].StringValue, new RankGroup());

                    RankGroup_Dict[cs2[x, 0].StringValue].RefRankID = int.Parse(rank.UID);

                    RankGroup_Dict[cs2[x, 0].StringValue].GroupHashKey = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].SubjectType = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].SubjectName = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].SubjectLevel = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].GroupType = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].GroupName = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].MemberCount = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].PERCENTILE88 = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].PERCENTILE75 = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].PERCENTILE50 = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].PERCENTILE25 = cs2[x, col_index++].StringValue;

                    RankGroup_Dict[cs2[x, 0].StringValue].PERCENTILE12 = cs2[x, col_index++].StringValue;
              

                }

            }



            RankGroup_Dict.Values.SaveAll();



            #region 抓取剛剛上傳的 RankGroups，取得DB系統自動產生的uid， 再存回RankGroup_Dict ，供等等RankDeyail 使用

            Dictionary<String, RankGroup> RankGroup_Dict_ReGet = new Dictionary<string, RankGroup>();

            foreach (string key in RankGroup_Dict.Keys)
            {

                var list_RankGroup_Dict = accessHelper.Select<RankGroup>(string.Format("group_hash_key='{0}'and ref_rank_id={1}", key,rank.UID));

                // 因為此List 每次用Key、rank.UID 只會抓到一筆資料，因此大膽使用list_RankGroup_Dict[0] 指定第一項加入RankGroup_Dict_ReGet                
                RankGroup_Dict_ReGet.Add(key,list_RankGroup_Dict[0]);

            }

           
            #endregion


            #endregion


            #region 處理Excel cs2 第三頁 RankDetail Table

            List<RankDetail> RankDetail_List = new List<RankDetail>();

            for (int x = 2; x < cs2.Rows.Count - 1; x++)
            {
                int col_index = 12;

                RankDetail rank_detail_record = new RankDetail();

                // 將第三頁的Excel sheet cs2 前半 RankGroup table 和後半 RankDetail table 以RefRankGroupID 連結

                rank_detail_record.RefRankGroupID = int.Parse(RankGroup_Dict_ReGet[cs2[x, 0].StringValue].UID);

                rank_detail_record.RefStudentID = int.Parse(cs2[x, col_index++].StringValue);

                rank_detail_record.班級 = cs2[x, col_index++].StringValue;

                rank_detail_record.座號 = cs2[x, col_index++].StringValue;

                rank_detail_record.學號 = cs2[x, col_index++].StringValue;

                rank_detail_record.姓名 = cs2[x, col_index++].StringValue;

                rank_detail_record.Score = cs2[x, col_index++].StringValue;

                rank_detail_record.Rank = cs2[x, col_index++].StringValue;

                rank_detail_record.PR = cs2[x, col_index++].StringValue;

                rank_detail_record.Percentage = cs2[x, col_index++].StringValue;

                rank_detail_record.Credit = cs2[x, col_index++].StringValue;

                rank_detail_record.Peroid = cs2[x, col_index++].StringValue;

                rank_detail_record.Entry = cs2[x, col_index++].StringValue;

                rank_detail_record.GradeYear = cs2[x, col_index++].StringValue;

                rank_detail_record.必選修 = cs2[x, col_index++].StringValue;

                rank_detail_record.校部訂 = cs2[x, col_index++].StringValue;

                rank_detail_record.科目成績 = cs2[x, col_index++].StringValue;

                rank_detail_record.原始成績 = cs2[x, col_index++].StringValue;

                rank_detail_record.補考成績 = cs2[x, col_index++].StringValue;

                rank_detail_record.重修成績 = cs2[x, col_index++].StringValue;

                rank_detail_record.手動調整成績 = cs2[x, col_index++].StringValue;

                rank_detail_record.學年調整成績 = cs2[x, col_index++].StringValue;

                rank_detail_record.Pass = cs2[x, col_index++].StringValue;

                rank_detail_record.不計學分 = cs2[x, col_index++].StringValue;

                rank_detail_record.不需評分 = cs2[x, col_index++].StringValue;

                rank_detail_record.Remark = cs2[x, col_index++].StringValue;



                RankDetail_List.Add(rank_detail_record);


                // 批次儲存，因為這邊的資料量隨便動輒上三~四萬筆，一口氣SaveAll 會爆掉，故一次分100份上傳，以目前實測40000筆資料大約需要180秒

                if (RankDetail_List.Count == 100)
                {

                    RankDetail_List.SaveAll();

                    RankDetail_List.Clear();
                }
            }

            //最後一份，總資料被100除不盡的資料在此上傳

            RankDetail_List.SaveAll();
            
            #endregion


        }


    }
}


