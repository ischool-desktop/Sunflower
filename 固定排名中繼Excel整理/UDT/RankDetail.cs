using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 固定排名中繼Excel整理.UDT
{

    //2016/7/12，穎驊與恩正新增Table，做固定排名使用，代號:sunflower，向日葵

    [TableName("ischool.sunflower.rankdetail")]
    class RankDetail:ActiveRecord
    {

        //下面RefRankID欄位會綁定上傳的Rank，由系統自動產生的uid

        /// <summary>
        /// Rankp系統編號
        /// </summary>
        [Field(Field = "ref_rank_group_id", Indexed = true)]
        public int RefRankGroupID { get; set; }


        /// <summary>
        /// 學生系統編號
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = true)]
        public int RefStudentID { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        [Field(Field = "class", Indexed = true)]
        public string 班級 { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        [Field(Field = "seat_no", Indexed = true)]
        public string 座號 { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        [Field(Field = "student_no", Indexed = true)]
        public string 學號 { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        [Field(Field = "name", Indexed = true)]
        public string 姓名 { get; set; }


        /// <summary>
        /// 成績
        /// </summary>
        [Field(Field = "score", Indexed = true)]
        public string Score { get; set; }

        /// <summary>
        /// 排名
        /// </summary>
        [Field(Field = "rank", Indexed = true)]
        public string Rank { get; set; }

        /// <summary>
        /// PR
        /// </summary>
        [Field(Field = "pr", Indexed = true)]
        public string PR { get; set; }

        /// <summary>
        /// 百分比
        /// </summary>
        [Field(Field = "percentage", Indexed = true)]
        public string Percentage { get; set; }


        /// <summary>
        /// 學分數
        /// </summary>
        [Field(Field = "credit", Indexed = true)]
        public string Credit { get; set; }

        /// <summary>
        /// 權數
        /// </summary>
        [Field(Field = "peroid", Indexed = true)]
        public string Peroid { get; set; }

        /// <summary>
        /// 分項類別
        /// </summary>
        [Field(Field = "entry", Indexed = true)]
        public string Entry { get; set; }

        /// <summary>
        /// 年級成績
        /// </summary>
        [Field(Field = "grade_year", Indexed = true)]
        public string GradeYear { get; set; }


        /// <summary>
        /// 必選修
        /// </summary>
        [Field(Field = "必選修", Indexed = true)]
        public string 必選修 { get; set; }

        /// <summary>
        /// 校部訂
        /// </summary>
        [Field(Field = "校部訂", Indexed = true)]
        public string 校部訂 { get; set; }

        /// <summary>
        /// 母群Key
        /// </summary>
        [Field(Field = "科目成績", Indexed = true)]
        public string 科目成績 { get; set; }

        /// <summary>
        /// 母群Key
        /// </summary>
        [Field(Field = "原始成績", Indexed = true)]
        public string 原始成績 { get; set; }


        /// <summary>
        /// 母群Key
        /// </summary>
        [Field(Field = "補考成績", Indexed = true)]
        public string 補考成績 { get; set; }

        /// <summary>
        /// 重修成績
        /// </summary>
        [Field(Field = "重修成績", Indexed = true)]
        public string 重修成績 { get; set; }

        /// <summary>
        /// 手動調整成績
        /// </summary>
        [Field(Field = "手動調整成績", Indexed = true)]
        public string 手動調整成績 { get; set; }

        /// <summary>
        /// 學年調整成績
        /// </summary>
        [Field(Field = "學年調整成績", Indexed = true)]
        public string 學年調整成績 { get; set; }

        /// <summary>
        /// 取得學分
        /// </summary>
        [Field(Field = "pass", Indexed = true)]
        public string Pass { get; set; }

        /// <summary>
        /// 不計學分
        /// </summary>
        [Field(Field = "不計學分", Indexed = true)]
        public string 不計學分 { get; set; }

        /// <summary>
        /// 不需評分
        /// </summary>
        [Field(Field = "不需評分", Indexed = true)]
        public string 不需評分 { get; set; }

        /// <summary>
        /// 註記
        /// </summary>
        [Field(Field = "remark", Indexed = true)]
        public string Remark { get; set; }



    }


    



}
