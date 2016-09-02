using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 固定排名中繼Excel整理
{
    class Permissions
    {
        

        public static string 固定排名中繼Excel整理test { get { return "OK"; } }
        public static bool 固定排名中繼Excel整理test權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[固定排名中繼Excel整理test].Executable;
            }
        }


    }
}
