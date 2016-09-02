using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 個人學期成績單.UDT
{

    //2016/7/12，穎驊與恩正新增Table，做固定排名使用，代號:sunflower，向日葵

    [TableName("ischool.sunflower.rankstudent")]
    class RankStudent:ActiveRecord
    {
        
        //下面RefRankID欄位會綁定上傳的Rank，由系統自動產生的uid


        /// <summary>
        /// RefRankID
        /// </summary>
        [Field(Field = "ref_rank_id", Indexed = true)]
        public int RefRankID { get; set; }

        /// <summary>
        /// 學生系統編號 
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = true)]
        public int RefStudentID { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        [Field(Field = "student_number", Indexed = true)]
        public int StudentNumber { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        [Field(Field = "name", Indexed = true)]
        public string Name { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        [Field(Field = "class_name", Indexed = true)]
        public string ClassName { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        [Field(Field = "seat_no", Indexed = true)]
        public string SeatNo { get; set; }

        /// <summary>
        /// 班導師
        /// </summary>
        [Field(Field = "hr_teacher_name", Indexed = true)]
        public string HRTeacherName { get; set; }

        /// <summary>
        /// 科別
        /// </summary>
        [Field(Field = "dept_name", Indexed = true)]
        public string DeptName { get; set; }

        /// <summary>
        /// 類別1
        /// </summary>
        [Field(Field = "cat_name1", Indexed = true)]
        public string CatName1 { get; set; }

        /// <summary>
        /// 類別2
        /// </summary>
        [Field(Field = "cat_name2", Indexed = true)]
        public string CatName2 { get; set; }


        //2016/7/13 穎驊筆記，bool 的欄位不可以indexed，故設定Indexed = false

        /// <summary>
        /// 參與排名
        /// </summary>
        [Field(Field = "rank_include", Indexed = false)]
        public bool RankInclude { get; set; }

        /// <summary>
        /// ClassID
        /// </summary>
        [Field(Field = "class_id", Indexed = true)]
        public int? ClassID { get; set; }

        /// <summary>
        /// Cat1ID
        /// </summary>
        [Field(Field = "cat1_id", Indexed = true)]
        public int? Cat1ID { get; set; }

        /// <summary>
        /// Cat2ID
        /// </summary>
        [Field(Field = "cat2_id", Indexed = true)]
        public int? Cat2ID { get; set; }



   

    }
}
							
