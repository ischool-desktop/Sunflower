using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace 固定排名中繼Excel整理
{
    public partial class NewConfigure : FISCA.Presentation.Controls.BaseForm
    {
        public Aspose.Words.Document Template { get; private set; }
        public int SubjectLimit { get; private set; }
        public string ConfigName { get; private set; }
        public bool WithSchoolYearScore { get; private set; }
        public bool WithPrevSemesterScore { get; private set; }

        public NewConfigure()
        {
            InitializeComponent();
           
        }
      

        private void checkReady(object sender, EventArgs e)
        {
            bool ready = true;
            if (txtName.Text == "")
                ready = false;
            else
                ConfigName = txtName.Text;


         
            btnSubmit.Enabled = ready;
        }

        

        private void btnSubmit_Click(object sender, EventArgs e)
        {            
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }
    }
}
