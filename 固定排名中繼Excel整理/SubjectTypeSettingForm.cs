using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace 固定排名中繼Excel整理
{
    public partial class SubjectTypeSettingForm : BaseForm
    {
        public Configure Configure { get; private set; }

        List<DataGridViewRow> rowList = new List<DataGridViewRow>();

        public SubjectTypeSettingForm(Configure _Configure)
        {
            InitializeComponent();

            // 取得由ConfigForm傳進來的 Cofigure 物件
            Configure = _Configure;

            // 解碼
            Configure.Decode();           

            // 以下進行初始化，將上一次的紀錄Load進來
            List<string> _SubjectTypeList = Configure.SubjectTypeList;

            List<string> _SubjectNameList = Configure.SubjectNameList;

            for (int x = 0; x < _SubjectTypeList.Count; x++)
            {

                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);

                if (_SubjectTypeList[x] != null)
                {
                    row.Cells[0].Value = _SubjectTypeList[x];
                }

                if (_SubjectNameList[x] != null)
                {
                    row.Cells[1].Value = _SubjectNameList[x];

                }


                rowList.Add(row);

            }
            dataGridViewX1.Rows.AddRange(rowList.ToArray());
        }

        //private List<DataGridViewRow> GetRow(string k, string p, int year, XmlNode boy_仰臥起坐)
        //{
        //    List<DataGridViewRow> rowList = new List<DataGridViewRow>();

        //    for (int x = year; x <= 23; x++)
        //    {
        //        foreach (XmlNode year_node in boy_仰臥起坐.SelectNodes("_" + x))
        //        {
        //            if (year_node != null)
        //            {
        //                string value1 = ((XmlElement)year_node).GetAttribute(name1);
        //                string value2 = ((XmlElement)year_node).GetAttribute(name2);
        //                string value3 = ((XmlElement)year_node).GetAttribute(name3);
        //                string value4 = ((XmlElement)year_node).GetAttribute(name4);
        //                string value5 = ((XmlElement)year_node).GetAttribute(name5);

        //                DataGridViewRow row = new DataGridViewRow();
        //                row.CreateCells(dataGridViewX1);
        //                row.Cells[0].Value = k;
        //                row.Cells[1].Value = x;
        //                row.Cells[2].Value = p;
        //                row.Cells[3].Value = value1.Split(',')[0] + "至" + value1.Split(',')[1];
        //                row.Cells[4].Value = name1;
        //                row.Cells[5].Value = value2.Split(',')[0] + "至" + value2.Split(',')[1];
        //                row.Cells[6].Value = name2;
        //                row.Cells[7].Value = value3.Split(',')[0] + "至" + value3.Split(',')[1];
        //                row.Cells[8].Value = name3;
        //                row.Cells[9].Value = value4.Split(',')[0] + "至" + value4.Split(',')[1];
        //                row.Cells[10].Value = name4;
        //                row.Cells[11].Value = value5.Split(',')[0] + "至" + value5.Split(',')[1];
        //                row.Cells[12].Value = name5;


        //                rowList.Add(row);
        //            }
        //        }
        //    }
        //    return rowList;
        //}

      
        // 取消
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        // 確認，將會進行set 新Configure的工作後，關閉，實際的Save 會在ConfigRorm.cs
        private void buttonX1_Click(object sender, EventArgs e)
        {
             List<string> _SubjectTypeList_Save = new List<string>();

            List<string> _SubjectNameList_Save = new List<string>();

            foreach (DataGridViewRow dgvr in dataGridViewX1.Rows)
            {


                
                _SubjectTypeList_Save.Add( (string) dgvr.Cells[0].Value);
                _SubjectNameList_Save.Add( (string)dgvr.Cells[1].Value);
                                    
            }



            Configure.SubjectTypeList = _SubjectTypeList_Save;
            Configure.SubjectNameList = _SubjectNameList_Save;

         

            this.Close();

        }
    }
}
