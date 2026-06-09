using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    public partial class frmSOManualEntryEditor2 : Form
    {

        clsGlobal _clsGLobal = new clsGlobal();
        public DataTable dt = new DataTable();

        private string _query;
        private bool _btnoke;
        private string _btnNo;

        private int[] _hideIndex;
        private bool g_bYesNo;

        public int[] HideIndex
        {
            get { return _hideIndex; }
            set { _hideIndex = value; }
        }

        public string Query
        {
            get { return _query; }
            set { _query = value; }
        }

        public string btnNo
        {
            get { return _btnNo; }
            set { _btnNo = value; }
        }

        public bool btnoke
        {
            get { return _btnoke; }
            set { _btnoke = value; }
        }

        public frmSOManualEntryEditor2()
        {
            InitializeComponent();
        }

        private void frmSOManualEntryEditor2_Load(object sender, EventArgs e)
        {
            btnOk.Visible = btnoke;
            btnCancel.Text = btnNo;
            FillGridNew();
        }

        private void FillGridNew()
        {
            using (SqlConnection cnData = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {
                cnData.Open();
                SqlDataAdapter da = new SqlDataAdapter(Query, cnData);
                da.Fill(dt);
                cnData.Close();
                dgvPopUp.DataSource = dt;

                dgvPopUp.Columns[0].DefaultCellStyle.BackColor = Color.Aquamarine;
                //khusus kolom terakhir supaya jadi full
                dgvPopUp.Columns[dgvPopUp.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                this.Width = dgvPopUp.Columns.GetColumnsWidth(DataGridViewElementStates.Visible) + 110;
                if (HideIndex.Count() != 0)
                {
                    for (int i = 0; i < HideIndex.Length; i++)
                    {
                        int a = HideIndex[i];
                        dgvPopUp.Columns[a].Visible = false;
                    }

                }


            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            g_bYesNo = false;
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            g_bYesNo = true;
            this.Close();
        }


    }
}
