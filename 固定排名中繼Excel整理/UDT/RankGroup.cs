using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 固定排名中繼Excel整理.UDT
{

    //2016/7/12，穎驊與恩正新增Table，做固定排名使用，代號:sunflower，向日葵

    [TableName("ischool.sunflower.rankgroup")]
    class RankGroup:ActiveRecord
    
    {



        //下面RefRankID欄位會綁定上傳的Rank，由系統自動產生的uid

        /// <summary>
        /// Rankp系統編號
        /// </summary>
        [Field(Field = "ref_rank_id", Indexed = true)]
        public int RefRankID { get; set; }


        /// <summary>
        /// 母群Key
        /// </summary>
        [Field(Field = "group_hash_key", Indexed = true)]
        public string GroupHashKey { get; set; }


        /// <summary>
        /// 類別
        /// </summary>
        [Field(Field = "subject_type", Indexed = true)]
        public string SubjectType { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        [Field(Field = "subject_name", Indexed = true)]
        public string SubjectName { get; set; }

        /// <summary>
        /// 科目級別
        /// </summary>
        [Field(Field = "subject_level", Indexed = true)]
        public string SubjectLevel { get; set; }

        /// <summary>
        /// 母群類型
        /// </summary>
        [Field(Field = "group_type", Indexed = true)]
        public string GroupType { get; set; }

        /// <summary>
        /// 母群名稱
        /// </summary>
        [Field(Field = "group_name", Indexed = true)]
        public string GroupName { get; set; }

        /// <summary>
        /// 總人數
        /// </summary>
        [Field(Field = "member_count", Indexed = true)]
        public string MemberCount { get; set; }

        /// <summary>
        /// 頂標
        /// </summary>
        [Field(Field = "percentile_88", Indexed = true)]
        public string PERCENTILE88 { get; set; }

        /// <summary>
        /// 前標
        /// </summary>
        [Field(Field = "percentile_75", Indexed = true)]
        public string PERCENTILE75 { get; set; }


        /// <summary>
        /// 均標
        /// </summary>
        [Field(Field = "percentile_50", Indexed = true)]
        public string PERCENTILE50 { get; set; }

        /// <summary>
        /// 後標
        /// </summary>
        [Field(Field = "percentile_25", Indexed = true)]
        public string PERCENTILE25 { get; set; }

        /// <summary>
        /// 底標
        /// </summary>
        [Field(Field = "percentile_12", Indexed = true)]
        public string PERCENTILE12 { get; set; }




    }
}

