using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data;
using System.IO;

namespace 固定排名中繼Excel整理
{
    // 2016/6/23 穎驊筆記，繼續處理UI的部分，發現原本的btnSaveConfig 沒有作用...(好像只是騙人的功能QQ)


    public partial class ConfigForm : FISCA.Presentation.Controls.BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
        private Dictionary<string, List<string>> _ExamSubjects = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> _ExamSubjectFull = new Dictionary<string, List<string>>();
        private List<TagConfigRecord> _TagConfigRecords = new List<TagConfigRecord>();
        private List<Configure> _Configures = new List<固定排名中繼Excel整理.Configure>();
        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        public ConfigForm()
        {
            InitializeComponent();
            List<ExamRecord> exams = new List<ExamRecord>();
            BackgroundWorker bkw = new BackgroundWorker();
            bkw.DoWork += delegate
            {
                bkw.ReportProgress(1);
                //預設學年度學期
                _DefalutSchoolYear = "" + K12.Data.School.DefaultSchoolYear;
                _DefaultSemester = "" + K12.Data.School.DefaultSemester;
                bkw.ReportProgress(10);
                //試別清單
                exams = K12.Data.Exam.SelectAll();
                bkw.ReportProgress(20);
                //學生類別清單
                _TagConfigRecords = K12.Data.TagConfig.SelectByCategory(TagCategory.Student);
                #region 整理所有試別對應科目
                var AEIncludeRecords = K12.Data.AEInclude.SelectAll();
                bkw.ReportProgress(30);
                var AssessmentSetupRecords = K12.Data.AssessmentSetup.SelectAll();
                bkw.ReportProgress(40);
                List<string> courseIDs = new List<string>();
                foreach (var scattentRecord in K12.Data.SCAttend.SelectByStudentIDs(K12.Presentation.NLDPanels.Student.SelectedSource))
                {
                    if (!courseIDs.Contains(scattentRecord.RefCourseID))
                        courseIDs.Add(scattentRecord.RefCourseID);
                }
                bkw.ReportProgress(60);
                foreach (var courseRecord in K12.Data.Course.SelectAll())
                {
                    foreach (var aeIncludeRecord in AEIncludeRecords)
                    {
                        if (aeIncludeRecord.RefAssessmentSetupID == courseRecord.RefAssessmentSetupID)
                        {
                            string key = courseRecord.SchoolYear + "^^" + courseRecord.Semester + "^^" + aeIncludeRecord.RefExamID;
                            if (!_ExamSubjectFull.ContainsKey(key))
                            {
                                _ExamSubjectFull.Add(key, new List<string>());
                            }
                            if (!_ExamSubjectFull[key].Contains(courseRecord.Subject))
                                _ExamSubjectFull[key].Add(courseRecord.Subject);
                            if (courseIDs.Contains(courseRecord.ID))
                            {
                                if (!_ExamSubjects.ContainsKey(key))
                                {
                                    _ExamSubjects.Add(key, new List<string>());
                                }
                                if (!_ExamSubjects[key].Contains(courseRecord.Subject))
                                    _ExamSubjects[key].Add(courseRecord.Subject);
                            }
                        }
                    }
                }
                bkw.ReportProgress(70);
                foreach (var list in _ExamSubjectFull.Values)
                {
                    #region 排序
                    list.Sort(new StringComparer("國文"
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
                }
                #endregion
                bkw.ReportProgress(80);
                _Configures = _AccessHelper.Select<Configure>();
                bkw.ReportProgress(100);

            };
            bkw.WorkerReportsProgress = true;
            bkw.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                circularProgress1.Value = e.ProgressPercentage;
            };
            bkw.RunWorkerCompleted += delegate
            {
                cboConfigure.Items.Clear();
                foreach (var item in _Configures)
                {
                    cboConfigure.Items.Add(item);
                }
                cboConfigure.Items.Add(new Configure() { Name = "新增" });
                int i;
                if (int.TryParse(_DefalutSchoolYear, out i))
                {
                    for (int j = 0; j < 5; j++)
                    {
                        cboSchoolYear.Items.Add("" + (i - j));
                    }
                }

                //2016/6/23 穎驊新增"年級"cbo ，使用者可以自行決定要印1~ 3 年級的固定排名

                cboGradeYear.Items.Add("1");
                cboGradeYear.Items.Add("2");
                cboGradeYear.Items.Add("3");

                cboSemester.Items.Add("1");
                cboSemester.Items.Add("2");
                cboExam.Items.Clear();
                
                cboExam.Items.AddRange(exams.ToArray());
       
                List<string> prefix = new List<string>();
                List<string> tag = new List<string>();
                foreach (var item in _TagConfigRecords)
                {
                    if (item.Prefix != "")
                    {
                        if (!prefix.Contains(item.Prefix))
                            prefix.Add(item.Prefix);
                    }
                    else
                    {
                        tag.Add(item.Name);
                    }
                }
                cboRankRilter.Items.Clear();
                cboTagRank1.Items.Clear();
                cboTagRank2.Items.Clear();
                cboRankRilter.Items.Add("");
                cboTagRank1.Items.Add("");
                cboTagRank2.Items.Add("");
                foreach (var s in prefix)
                {
                    cboRankRilter.Items.Add("[" + s + "]");
                    cboTagRank1.Items.Add("[" + s + "]");
                    cboTagRank2.Items.Add("[" + s + "]");
                }
                foreach (var s in tag)
                {
                    cboRankRilter.Items.Add(s);
                    cboTagRank1.Items.Add(s);
                    cboTagRank2.Items.Add(s);
                }
                circularProgress1.Hide();
                if (_Configures.Count > 0)
                {
                    cboConfigure.SelectedIndex = 0;
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            };
            bkw.RunWorkerAsync();
        }

        public Configure Configure { get; private set; }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // 檢查設定學年度學期與目前系統預設是否相同
            string s1 = cboSchoolYear.Text + cboSemester.Text;
            string s2 = School.DefaultSchoolYear + School.DefaultSemester;

            if (s1 != s2)
                if (FISCA.Presentation.Controls.MsgBox.Show("畫面上學年度學期與系統學年度學期不相同，請問是否繼續?", "學年度學期不同", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                    return;

            if (cboGradeYear.SelectedItem ==null) {
                FISCA.Presentation.Controls.MsgBox.Show("尚未選擇年級，請選擇", "錯誤", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                this.DialogResult = System.Windows.Forms.DialogResult.No;
                return;
            
            }


            SaveTemplate(null, null);
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
     
        private void ExamChanged(object sender, EventArgs e)
        {
            string key = cboSchoolYear.Text + "^^" + cboSemester.Text + "^^" +
                (cboExam.SelectedItem == null ? "" : ((ExamRecord)cboExam.SelectedItem).ID);
            listViewEx1.SuspendLayout();
            listViewEx2.SuspendLayout();
            listViewEx3.SuspendLayout();
            listViewEx1.Items.Clear();
            listViewEx2.Items.Clear();
            listViewEx3.Items.Clear();
            if (_ExamSubjectFull.ContainsKey(key))
            {
                foreach (var subject in _ExamSubjectFull[key])
                {
                    var i1 = listViewEx1.Items.Add(subject);
                    var i2 = listViewEx2.Items.Add(subject);
                    var i3 = listViewEx3.Items.Add(subject);
                    if (Configure != null && Configure.PrintSubjectList.Contains(subject))
                        i1.Checked = true;
                    if (Configure != null && Configure.TagRank1SubjectList.Contains(subject))
                        i2.Checked = true;
                    if (Configure != null && Configure.TagRank2SubjectList.Contains(subject))
                        i3.Checked = true;
                    if (_ExamSubjects.ContainsKey(key) && !_ExamSubjects[key].Contains(subject))
                    {
                        i1.ForeColor = Color.DarkGray;
                        i2.ForeColor = Color.DarkGray;
                        i3.ForeColor = Color.DarkGray;
                    }
                }
            }
            listViewEx1.ResumeLayout(true);
            listViewEx2.ResumeLayout(true);
            listViewEx3.ResumeLayout(true);
        }

      


        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = btnPrint.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Configure = new Configure();

                    Configure.Name = dialog.ConfigName;

                    //Configure.Template = dialog.Template;

                    Configure.SubjectLimit = dialog.SubjectLimit;
                    Configure.SchoolYear = _DefalutSchoolYear;
                    Configure.Semester = _DefaultSemester;
                    if (cboExam.Items.Count > 0)
                        Configure.ExamRecord = (ExamRecord)cboExam.Items[0];
                    _Configures.Add(Configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, Configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;
                    Configure.WithSchoolYearScore = dialog.WithSchoolYearScore;
                    Configure.WithPrevSemesterScore = dialog.WithPrevSemesterScore;
                    Configure.Encode();
                    Configure.Save();
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnSaveConfig.Enabled = btnPrint.Enabled = true;
                    Configure = _Configures[cboConfigure.SelectedIndex];


                    //Decode很重要，沒有這個，舊設定不會出來。比如說:PrintSubjectList，可以將上次有勾選的列印科目顯示，
                    Configure.Decode();


                    if (!cboSchoolYear.Items.Contains(Configure.SchoolYear))
                        cboSchoolYear.Items.Add(Configure.SchoolYear);
                    cboSchoolYear.Text = Configure.SchoolYear;
                    cboSemester.Text = Configure.Semester;

                    // 2016/6/24 穎驊新增，讓系統讀進上次同一模板的輸入年級資料
                    cboGradeYear.Text = Configure.GradeYear;

                    if (Configure.ExamRecord != null)
                    {
                        foreach (var item in cboExam.Items)
                        {
                            if (((ExamRecord)item).ID == Configure.ExamRecord.ID)
                            {
                                cboExam.SelectedIndex = cboExam.Items.IndexOf(item);
                                break;
                            }
                        }
                    }
               

                    cboRankRilter.Text = Configure.RankFilterTagName;
                    foreach (ListViewItem item in listViewEx1.Items)
                    {
                        item.Checked = Configure.PrintSubjectList.Contains(item.Text);
                    }
                    cboTagRank1.Text = Configure.TagRank1TagName;
                    foreach (ListViewItem item in listViewEx2.Items)
                    {
                        item.Checked = Configure.TagRank1SubjectList.Contains(item.Text);
                    }
                    cboTagRank2.Text = Configure.TagRank2TagName;
                    foreach (ListViewItem item in listViewEx3.Items)
                    {
                        item.Checked = Configure.TagRank2SubjectList.Contains(item.Text);
                    }
                }
                else
                {
                    Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;
                    cboExam.SelectedIndex = -1;
                    
                    cboRankRilter.SelectedIndex = -1;
                    cboTagRank1.SelectedIndex = -1;
                    cboTagRank2.SelectedIndex = -1;
                    foreach (ListViewItem item in listViewEx1.Items)
                    {
                        item.Checked = false;
                    }
                    foreach (ListViewItem item in listViewEx2.Items)
                    {
                        item.Checked = false;
                    }
                    foreach (ListViewItem item in listViewEx3.Items)
                    {
                        item.Checked = false;
                    }
                }
            }
        }

    

        private void SaveTemplate(object sender, EventArgs e)
        {
            if (Configure == null) return;
            Configure.SchoolYear = cboSchoolYear.Text;
            Configure.Semester = cboSemester.Text;

            // 2016/6/23 穎驊新增，增加Cofigure 輸入"年級"的選項
            Configure.GradeYear = cboGradeYear.Text;
            
            // 2016/6/27 穎驊新增，增加 動態自動更改儲存檔名，存在config 中
            Configure.FileName = "固定排名中繼Excel整理" + "(" + cboConfigure.Text +" "+ cboGradeYear.Text + "年級" +" "+cboSchoolYear.Text+"-"+cboSemester.Text+ ")" + ".Xlsx";

            Configure.ExamRecord = ((ExamRecord)cboExam.SelectedItem);

            

            if (Configure.RefenceExamRecord != null && Configure.RefenceExamRecord.Name == "")
                Configure.RefenceExamRecord = null;
            foreach (ListViewItem item in listViewEx1.Items)
            {
                if (item.Checked)
                {
                    if (!Configure.PrintSubjectList.Contains(item.Text))
                        Configure.PrintSubjectList.Add(item.Text);
                }
                else
                {
                    if (Configure.PrintSubjectList.Contains(item.Text))
                        Configure.PrintSubjectList.Remove(item.Text);
                }
            }
            Configure.TagRank1TagName = cboTagRank1.Text;
            Configure.TagRank1TagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (cboTagRank1.Text == "[" + item.Prefix + "]")
                        Configure.TagRank1TagList.Add(item.ID);
                }
                else
                {
                    if (cboTagRank1.Text == item.Name)
                        Configure.TagRank1TagList.Add(item.ID);
                }
            }
            foreach (ListViewItem item in listViewEx2.Items)
            {
                if (item.Checked)
                {
                    if (!Configure.TagRank1SubjectList.Contains(item.Text))
                        Configure.TagRank1SubjectList.Add(item.Text);
                }
                else
                {
                    if (Configure.TagRank1SubjectList.Contains(item.Text))
                        Configure.TagRank1SubjectList.Remove(item.Text);
                }
            }

            Configure.TagRank2TagName = cboTagRank2.Text;
            Configure.TagRank2TagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (cboTagRank2.Text == "[" + item.Prefix + "]")
                        Configure.TagRank2TagList.Add(item.ID);
                }
                else
                {
                    if (cboTagRank2.Text == item.Name)
                        Configure.TagRank2TagList.Add(item.ID);
                }
            }
            foreach (ListViewItem item in listViewEx3.Items)
            {
                if (item.Checked)
                {
                    if (!Configure.TagRank2SubjectList.Contains(item.Text))
                        Configure.TagRank2SubjectList.Add(item.Text);
                }
                else
                {
                    if (Configure.TagRank2SubjectList.Contains(item.Text))
                        Configure.TagRank2SubjectList.Remove(item.Text);
                }
            }

            Configure.RankFilterTagName = cboRankRilter.Text;
            Configure.RankFilterTagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (cboRankRilter.Text == "[" + item.Prefix + "]")
                        Configure.RankFilterTagList.Add(item.ID);
                }
                else
                {
                    if (cboRankRilter.Text == item.Name)
                        Configure.RankFilterTagList.Add(item.ID);
                }
            }

            Configure.Encode();
            Configure.Save();
        }


        // 當使用者要新增樣板時，可以直接用現有樣本進行複製，'方便操作
        private void linkLabel3_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;
                conf.ExamRecord = Configure.ExamRecord;
                conf.PrintSubjectList.AddRange(Configure.PrintSubjectList);
                conf.RankFilterTagList.AddRange(Configure.RankFilterTagList);
                conf.RankFilterTagName = Configure.RankFilterTagName;

               

                conf.SchoolYear = Configure.SchoolYear;
                conf.Semester = Configure.Semester;
                conf.SubjectLimit = Configure.SubjectLimit;
                conf.TagRank1SubjectList.AddRange(Configure.TagRank1SubjectList);
                conf.TagRank1TagList.AddRange(Configure.TagRank1TagList);
                conf.TagRank1TagName = Configure.TagRank1TagName;
                conf.TagRank2SubjectList.AddRange(Configure.TagRank2SubjectList);
                conf.TagRank2TagList.AddRange(Configure.TagRank2TagList);
                conf.TagRank2TagName = Configure.TagRank2TagName;               

                conf.WithPrevSemesterScore = Configure.WithPrevSemesterScore;
                conf.WithSchoolYearScore = Configure.WithSchoolYearScore;
                conf.Encode();
                conf.Save();
                _Configures.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }


        // 供使用者可以刪除過多的樣板設定
        private void linkLabel4_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _Configures.Remove(Configure);
                if (Configure.UID != "")
                {
                    Configure.Deleted = true;
                    Configure.Save();
                }
                var conf = Configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }
        }


        //2016/7/4 穎驊製作，與恩正溝通後，開始要新增回歸科目設定，下面開始實作讓使用者設定類別、科目名稱

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            
            // 將Configure傳進去，就可以在裡面直接set改變囉~ 且改變後離開，就可以在本頁直接儲存 超方便的啦!
            SubjectTypeSettingForm stsf = new SubjectTypeSettingForm(Configure);

            stsf.ShowDialog();
        }

       
    }
}
