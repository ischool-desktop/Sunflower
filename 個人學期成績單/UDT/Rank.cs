using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 個人學期成績單.UDT
{


    //2016/7/12，穎驊與恩正新增Table，做固定排名使用，代號:sunflower，向日葵

    [TableName("ischool.sunflower.rank")]
    class Rank:ActiveRecord
    {

        
        /// <summary>
        /// 學年度
        /// </summary>
        [Field(Field="school_year",Indexed=true)]
        public int SchoolYear { get; set; }

        
        /// <summary>
        /// 學期
        /// </summary>
        [Field(Field = "semester", Indexed = true)]
        public int Semester { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        [Field(Field = "grade_year", Indexed = true)]
        public int GradeYear { get; set; }

        /// <summary>
        /// 排名類型
        /// </summary>
        [Field(Field = "rank_type", Indexed = true)]
        public string RankType { get; set; }

        /// <summary>
        /// 排名次序
        /// </summary>
        [Field(Field = "rank_sequence", Indexed = true)]
        public int? RankSequence { get; set; }

        /// <summary>
        /// 顯示名稱
        /// </summary>
        [Field(Field = "display_name", Indexed = false)]
        public string DisplayName { get; set; }

        /// <summary>
        /// 註記
        /// </summary>
        [Field(Field = "memo", Indexed = false)]
        public string Memo { get; set; }

        /// <summary>
        /// 建立時間
        /// </summary>
        [Field(Field = "create_time", Indexed = false)]
        public DateTime CreateTime { get; set; }



        //是否啟動(此欄位不會定義在Excel中，會隨著每一次的上傳將以前的Table 設為False，最新的為True)
        [Field(Field = "active", Indexed = false)]
        public bool Active { get; set; }
    }
}
