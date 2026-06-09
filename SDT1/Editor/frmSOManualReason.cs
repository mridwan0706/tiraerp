using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    public partial class frmSOManualReason : Form
    {
        private clsGlobal _clsGlobal = new clsGlobal();

        private string strSQL;
        //frmSOManualEntryList frm = new frmSOManualEntryList();
        private int _xHaveClickR;
        private string _xReasonR;
        public int xHaveClickR
        {
            get { return _xHaveClickR; }
            set { _xHaveClickR = value; }
        }

        public string xReasonR
        {
            get { return _xReasonR; }
            set { _xReasonR = value; }
        }

        public frmSOManualReason()
        {
            InitializeComponent();
        }

        #region WinForm

        private void btnCancel_Click_1(object sender, EventArgs e)
        {
            xHaveClickR = 0;
            xReasonR = "";
            this.Close();
        }

        private void frmSOManualReason_Load_1(object sender, EventArgs e)
        {
            xHaveClickR = 0;
            xReasonR = "";
            BindSOReason();
        }

        private void btnOk_Click_1(object sender, EventArgs e)
        {
            xHaveClickR = 1;
            xReasonR = cbSOReason.SelectedValue.ToString();
            this.Close();
        }

        #endregion

        #region Function

        private void BindSOReason()
        {
            strSQL = "";
            strSQL = "select rm_reason_code, rm_reason_desc from TBL_GS_REASON_MASTER WITH (NOLOCK) ";
            strSQL = strSQL + " where rm_sub_modul_id = 'SALES_CANCEL'";
            strSQL = strSQL + " and rm_auto_flag = 'N'";
            strSQL = strSQL + " order by rm_reason_code";

            cbSOReason.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSOReason.ValueMember = "rm_reason_code";
            cbSOReason.DisplayMember = "rm_reason_desc";
        }

        #endregion

        
    }
}
