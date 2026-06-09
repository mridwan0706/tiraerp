using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TIRASnDNet.INV.INVProductMaster
{
    public partial class frmINVProductMasterList : Form
    {
        LineHighlight __cursor = null;
        const int WidthList = 750;
        const int HeightList = 460;

        private clsGlobal _clsGlobal = new clsGlobal();
       
        string ProdID,Grade,Size;

        frmPopUp frm;
        private string strSQL, paramMenuId;
        bool dariPopup = false;
        public frmINVProductMasterList()
        {
            InitializeComponent();
        }
        private void AccessButton()
        {
            tsb_new.Enabled = clsLogin.BTNNEW;
            tsb_edit.Enabled = clsLogin.BTNEDIT;
            tsb_delete.Enabled = clsLogin.BTNDELETE;
            tsb_print.Enabled = clsLogin.BTNPRINT;
        }
        private void FillGridSP() 
        {
            int CEK = 0;
            if (CHKTgl.Checked)
            {
                CEK = 1;
             }
            
            try
            {
                strSQL = "EXEC SP_INV_PRODUK_MASTER_LIST 0, '" + txtProdMaster.Text + "' , " +
                    "'"+txtBrandDesc.Text + "', " +
                    "'" + txtGroupDesc.Text + "'," + 
                    "'"+ txtSubGDesc.Text + "'," + 
                    "'"+ txtModelDesc.Text + "'," +
                    "'" + txtRawMCodeId.Text + "'," +
                    "'" + txtbarcode.Text + "'," +
                     "'" + txtdesc.Text + "'," +
                    "'" + txtColorId.Text + "'," +
                    "'" + CBProType.SelectedValue.ToString().Trim() + "'," +
                    "'"+ CBWarehose.SelectedValue.ToString().Trim() + "',"+
                     "'" + CbNonStock.SelectedValue.ToString().Trim() + "',"+
                    "'" + CbTech.SelectedValue.ToString().Trim() + "',"+
                    "'" + CbAktif.SelectedValue.ToString().Trim() + "'," +
                    "'" + txtTrend.Text + "',"+
                     "'" + txtmarket.Text + "'," +
                     "'" + txtseason.Text + "'," +
                    "'" + txtKategori.Text + "'," +
                     //"'" + CbOrder.SelectedValue.ToString().Trim() + "',"+
                    "'" + CEK.ToString() + "',"+
                    "'" + Convert.ToDateTime(dateTimePicker1.Value.ToString().Trim()).ToString("yyyyMMdd") + "'," +
                    "'" + Convert.ToDateTime(dateTimePicker2.Value.ToString().Trim()).ToString("yyyyMMdd") + "'," +
                    
                     " '"+ CbSex.SelectedValue.ToString().Trim() + "' ";

                
                dataGridView1.AutoGenerateColumns = false;
                dataGridView1.DataSource = _clsGlobal.ExecDT(strSQL);
                strSQL = "";

               

                if (dataGridView1.Rows.Count > 0)
                {
                    DataGridViewCellEventArgs DataGridViewCellEventArgs = new DataGridViewCellEventArgs(dataGridView1.Columns["prm_prd_master_code"].Index, dataGridView1.CurrentRow.Index);
                    this.dataGridView1_CellClick(dataGridView1.CurrentRow.Index, DataGridViewCellEventArgs);


                  

                }
             }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        


        void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                ProdID = dataGridView1.Rows[e.RowIndex].Cells["prm_prd_master_code"].Value.ToString().Trim();
                Grade = dataGridView1.Rows[e.RowIndex].Cells["prm_grade"].Value.ToString().Trim();
                Size = dataGridView1.Rows[e.RowIndex].Cells["prm_prd_size"].Value.ToString().Trim();
            }
        }

        private void btnTrend_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Trend ";
            frm.Query = "SELECT pp_function_code AS Trend_Code ,pp_function_desc as Descriptions   FROM IM_PRD_PARAMETER WHERE pp_function_name ='TREND' ORDER BY pp_sequence_no ";

            frm.ShowDialog();

            if (frm.ArrField != null)
            {
                txtTrend.Text = frm.ArrField[0].Trim();
                txtTrenddesc.Text = frm.ArrField[1].Trim();
            }
        }

        private void btnMarket_Click(object sender, EventArgs e)
        {
              
        frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Target ";
            frm.Query = "SELECT pp_function_code AS Target_Code ,pp_function_desc as Descriptions   FROM IM_PRD_PARAMETER WHERE pp_function_name ='TARGET' ORDER BY pp_sequence_no ";

            frm.ShowDialog();

            if (frm.ArrField != null)
            {
                txtmarket.Text = frm.ArrField[0].Trim();
                txtmarketDesc.Text = frm.ArrField[1].Trim();
            }
        }

        private void btnSeason_Click(object sender, EventArgs e)
        {
           frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Season ";
            frm.Query = " SELECT pp_function_code AS Season_Code ,pp_function_desc as Descriptions   FROM IM_PRD_PARAMETER WHERE pp_function_name ='SEASON' ORDER BY pp_sequence_no ";

            frm.ShowDialog();

            if (frm.ArrField != null)
            {
                txtseason.Text = frm.ArrField[0].Trim();
                txtseasonDesc.Text = frm.ArrField[1].Trim();
            } 
        }

        private void btnKategori_Click(object sender, EventArgs e)
        {
          
        frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Season ";
            frm.Query = "SELECT pch_year AS Category_Code1 ,pch_kelompok AS Category_Code1 , pch_description as Descriptions  FROM IM_PRD_CATEGORY_HEADER where pch_year LIKE '%' AND   pch_kelompok LIKE '%'";

            frm.ShowDialog();

            if (frm.ArrField != null)
            {
                txtKategori.Text = frm.ArrField[0].Trim();
                txtKategoriDesc.Text = frm.ArrField[1].Trim();
                txtkategoriTo.Text = frm.ArrField[2].Trim();
            } 
        }

        private void frmINVProductMasterList_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            __cursor = new LineHighlight(this.grupBox1);
            foreach (Control __child in this.grupBox1.Controls)
            {
                if (__child.TabStop) __cursor.Add(__child);
            }
            txtBrandId.Focus();
            txtBrandId.Select();
            this.Text = clsLogin.MENUID + " - " + this.Text;
            paramMenuId = clsLogin.MENUID;
            AccessButton();


            DataTable dataTabel3 = new DataTable();
            dataTabel3.Columns.Add("id");
            dataTabel3.Columns.Add("Nama");
            dataTabel3.Rows.Add("All", "<ALL>");
            dataTabel3.Rows.Add("Y", "Y-Yes");
            dataTabel3.Rows.Add("N", "N-No");

            CBWarehose.DataSource = dataTabel3;
            CBWarehose.DisplayMember = "Nama";
            CBWarehose.ValueMember = "id";


            DataTable dataTabel = new DataTable();
            dataTabel.Columns.Add("id");
            dataTabel.Columns.Add("Nama");
            dataTabel.Rows.Add("All", "<ALL>");
            dataTabel.Rows.Add("Y", "Y-Yes");
            dataTabel.Rows.Add("N", "N-No");

            CbNonStock.DataSource = dataTabel;
            CbNonStock.DisplayMember = "Nama";
            CbNonStock.ValueMember = "id";


            DataTable dataTabel2 = new DataTable();
            dataTabel2.Columns.Add("id");
            dataTabel2.Columns.Add("Nama");
            dataTabel2.Rows.Add("All", "<ALL>");
            dataTabel2.Rows.Add("Y", "Y-Yes");
            dataTabel2.Rows.Add("N", "N-No");

            DataTable dataTabel4 = new DataTable();
            dataTabel4.Columns.Add("id");
            dataTabel4.Columns.Add("Nama");
            dataTabel4.Rows.Add("All", "<ALL>");
            dataTabel4.Rows.Add("Y", "Y-Yes");
            dataTabel4.Rows.Add("N", "N-No");

            CbTech.DataSource = dataTabel4;
            CbTech.DisplayMember = "Nama";
            CbTech.ValueMember = "id";

            DataTable dataTabel5 = new DataTable();
            dataTabel5.Columns.Add("id");
            dataTabel5.Columns.Add("Nama");
            dataTabel5.Rows.Add("All", "<ALL>");
            dataTabel5.Rows.Add("Y", "Y-Yes");
            dataTabel5.Rows.Add("N", "N-No");

            CbAktif.DataSource = dataTabel5;
            CbAktif.DisplayMember = "Nama";
            CbAktif.ValueMember = "id";

           

            DataTable dataTabel6 = new DataTable();
            dataTabel6.Columns.Add("id");
            dataTabel6.Columns.Add("Nama");
            dataTabel6.Rows.Add(0,"0-Standard");
            dataTabel6.Rows.Add(1,"1-Group|Market Target");
            dataTabel6.Rows.Add(2,"2-Market Target|Group");
            dataTabel6.Rows.Add(3,"3-Group");
            dataTabel6.Rows.Add(4,"4-Market Target");
            dataTabel6.Rows.Add(5,"5-Raw Material|Market Target");
            dataTabel6.Rows.Add(6,"6-Season");
            dataTabel6.Rows.Add(7,"7-Collection");

            CbOrder.DataSource = dataTabel6;
            CbOrder.DisplayMember = "Nama";
            CbOrder.ValueMember = "id";

            DataTable dataTabel7 = new DataTable();
            dataTabel7.Columns.Add("id");
            dataTabel7.Columns.Add("Nama");
            dataTabel7.Rows.Add("All", "<ALL>");
            dataTabel7.Rows.Add("M", "M-Male");
            dataTabel7.Rows.Add("F", "F-Female");
            dataTabel7.Rows.Add("U", "U-Unisex");

            CbSex.DataSource = dataTabel7;
            CbSex.DisplayMember = "Nama";
            CbSex.ValueMember = "id";

            try
            {
                //produk type
                strSQL = "SELECT * FROM (" +
                        "(select 0 AS gh_sequence_no, '' as gh_function_code, '<ALL>' as gh_function_desc) " + 
                        "UNION " +
                        "(select gh_sequence_no, gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc " + 
                        " from GS_GEN_HARDCODED " +
                        " where gh_sys = 'H' and gh_function_name = 'PRDTYPE'))VW order by gh_sequence_no";
                CBProType.DataSource = _clsGlobal.ExecDT(strSQL);
                CBProType.ValueMember = "gh_function_code";
                CBProType.DisplayMember = "gh_function_desc"; 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
              
        private void CHKTgl_CheckedChanged(object sender, EventArgs e)
        {
            if (CHKTgl.Checked)
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
            else
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
        }

        private void CBKLong_CheckedChanged(object sender, EventArgs e)
        {
            if (CBKLong.Checked)
            {
                CBKLong.Text = "Short Type";
            }
            else
            {
                CBKLong.Text = "Long Type";
            }
        }

        private void btnExec_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            FillGridSP();
            Cursor.Current = Cursors.Default;
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void tsb_edit_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                int rowindex = dataGridView1.CurrentCell.RowIndex;


                clsGlobal.MODE_TRX = 2;
                frmINVProductMasterEditor frm = new frmINVProductMasterEditor();
                frm.PrdBrandCode = dataGridView1.Rows[rowindex].Cells["prm_prd_master_code"].Value.ToString().Trim();
                frm.PrdGrade = dataGridView1.Rows[rowindex].Cells["prm_grade"].Value.ToString().Trim();
                frm.ShowDialog();
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                int rowindex = dataGridView1.CurrentCell.RowIndex;
               

                clsGlobal.MODE_TRX = 2;
                frmINVProductMasterEditor frm = new frmINVProductMasterEditor();
                frm.PrdBrandCode = dataGridView1.Rows[rowindex].Cells["prm_prd_master_code"].Value.ToString().Trim();
                frm.PrdGrade = dataGridView1.Rows[rowindex].Cells["prm_grade"].Value.ToString().Trim();
                frm.ShowDialog();
               // FillGrid();

            }
        }

        private void tsb_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tsb_delete_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    DialogResult dr;
                    dr = MessageBox.Show(_clsGlobal.ApplMessage(10001), clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.OK)
                    {
                        strSQL = "";
                        strSQL = "EXEC SP_INV_PRODUK_MASTER '3', " +
                            
                            "'" + ProdID + "', " +
                            "'" + Grade + "', " +
                            "'" + Size + "', " +
                            "'', '', '', '', '','','','','','', " +
                            "'', '', '', '', '','','','','',''," +
                            "'', '', '', '', '','','','','','', " +
                            "'" + clsLogin.USERID + "','', '', '', '', '','', " +
                            "'', '', '', '', '','','','','','', " +
                            "'', '', '', '', '', '', '', '', '' ";

                        _clsGlobal.BeginTrans();
                        _clsGlobal.ExecuteTrans(strSQL);
                        _clsGlobal.CommitTrans();

                        FillGridSP();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsb_new_Click(object sender, EventArgs e)
        {
            clsGlobal.MODE_TRX = 1;
            frmINVProductMasterEditor frm = new frmINVProductMasterEditor();
            frm.ShowDialog(); 
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txtTrend_TextChanged(object sender, EventArgs e)
        {
            if (txtTrend.Text.ToString() == "")
            {
                txtTrenddesc.Clear();
            }
        }

        private void txtTrenddesc_TextChanged(object sender, EventArgs e)
        {
            if (txtTrenddesc.Text.ToString() == "")
            {
                txtTrend.Clear();
            }
        }

        private void txtmarket_TextChanged(object sender, EventArgs e)
        {
            if (txtmarket.Text.ToString() == "")
            {
                txtmarketDesc.Clear();
            }
        }

        private void txtmarketDesc_TextChanged(object sender, EventArgs e)
        {
            if (txtmarketDesc.Text.ToString() == "")
            {
                txtmarket.Clear();
            }
        }

        private void txtKategori_TextChanged(object sender, EventArgs e)
        {
            if (txtkategoriTo.Text.ToString() == "")
            {
                txtKategoriDesc.Clear();
                txtkategoriTo.Clear();
            }
        }

        private void txtKategoriDesc_TextChanged(object sender, EventArgs e)
        {
            if (txtKategoriDesc.Text.ToString() == "")
            {
                txtkategoriTo.Clear();
                txtKategori.Clear();
            }
        }

        private void txtkategoriTo_TextChanged(object sender, EventArgs e)
        {
            if (txtkategoriTo.Text.ToString() == "")
            {
                txtKategoriDesc.Clear();
                txtKategori.Clear();
            }
        }


        private void tsb_print_Click(object sender, EventArgs e)
        {

            int CEK = 0;
            if (CHKTgl.Checked)
            {
                CEK = 1;
            }
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                DataTable dt = _clsGlobal.GetEntityBranch();
                string paramkode;

                frmReport frm = new frmReport();
                frm.ReportName("INVProductMasterList.rpt");

                strSQL = "SELECT * FROM VW_INV_PRODUK_MASTER where 1=1 ";
                if (txtBrandDesc.Text != "")
                {
                    strSQL += " AND prm_prd_line_code='" + txtBrandDesc.Text + "'";
                }
                if (txtGroupDesc.Text != "")
                {
                    strSQL += " AND prm_prd_group_code ='" + txtGroupDesc.Text + "'";
                }
                if (txtSubGDesc.Text != "")
                {
                    strSQL += " AND prm_prd_sgroup_code ='" + txtSubGDesc.Text + "'";
                }
                if (txtModelId.Text != "")
                {
                    strSQL += " AND pm_prd_model_desc='" + txtModelId.Text + "'";
                }
                if (txtRawMCodeId.Text != "")
                {
                    strSQL += " AND rmu_raw_mat_used_desc='" + txtRawMCodeId.Text + "'";
                }
                if (txtTrend.Text != "")
                {
                    strSQL += " AND prm_trend like '%" + txtTrend.Text + "%'";
                }
                if (txtTrenddesc.Text != "")
                {
                    strSQL += " AND desc_trend like '%" + txtTrenddesc.Text + "%'";
                }
                if (txtmarket.Text != "")
                {
                    strSQL += " AND prm_target like '%" + txtmarket.Text + "%'";
                }
                if (txtmarketDesc.Text != "")
                {
                    strSQL += " AND desc_target like '%" + txtmarketDesc.Text + "%'";
                }
                if (txtmarketDesc.Text != "")
                {
                    strSQL += " AND desc_target like '%" + txtmarketDesc.Text + "%'";
                }
                if (txtProdMaster.Text != "")
                {
                    strSQL += " AND prm_prd_master_code like '%" + txtProdMaster.Text + "%'";
                }
                if (txtseason.Text != "")
                {
                    strSQL += " AND prm_season like '%" + txtseason.Text + "%'";
                }
                if (txtseasonDesc.Text != "")
                {
                    strSQL += " AND desc_season like '%" + txtseasonDesc.Text + "%'";
                }
                if (txtKategori.Text != "")
                {
                    strSQL += " AND prm_mfg_periode_mm_yy like '%" + txtKategori.Text + "%'";
                }
                if (txtKategoriDesc.Text != "")
                {
                    strSQL += " AND so_kelompok  like '%" + txtKategoriDesc.Text + "%'";
                }
                if (txtkategoriTo.Text != "")
                {
                    strSQL += " AND pch_description like '%" + txtkategoriTo.Text + "%'";
                }
                if (txtdesc.Text != "")
                {
                    strSQL += " AND prm_prd_desc like '%" + txtdesc.Text + "%'";
                }

                if (txtColorId.Text != "")
                {
                    strSQL += " AND col_washing_collor_desc ='" + txtColorId.Text + "'";
                }
                if (CBProType.Text.Length > 0)
                {
                    strSQL += "and prm_prd_type like '%" + CBProType.SelectedValue.ToString().Trim() + "%' ";
                }

                if (CBWarehose.SelectedIndex != 0)
                {
                    strSQL = strSQL + "and prm_wh_control_flag = '" + CBWarehose.SelectedValue.ToString().Trim() + "' ";

                }
                if (CbNonStock.SelectedIndex != 0)
                {

                    strSQL += "and prm_non_stock_flag = '" + CbNonStock.SelectedValue.ToString().Trim() + "' ";
                }
                if (CbTech.SelectedIndex != 0)
                {
                    strSQL = strSQL + "and prm_tech_constrain_flag = '" + CbTech.SelectedValue.ToString().Trim() + "' ";
                }
                if (CbAktif.SelectedIndex != 0)
                {
                    strSQL = strSQL + "and prm_active_flag = '" + CbAktif.SelectedValue.ToString().Trim() + "' ";
                }
                //if (CbOrderSelectedIndex != 0)
                //{
                //    strSQL = strSQL + "and prm_active_flag = '" + CbOrder.SelectedValue.ToString().Trim() + "' ";
                //}
                if (CbSex.SelectedIndex != 0)
                {
                    strSQL = strSQL + "and prm_sex = '" + CbSex.SelectedValue.ToString().Trim() + "' ";
                }

                if (CHKTgl.Checked)
                {
                    strSQL = strSQL + "AND CONVERT(VARCHAR(8),prm_creation_date,112) between '" + _clsGlobal.DateToDB(dateTimePicker1.Text) + "' AND '" + _clsGlobal.DateToDB(dateTimePicker2.Text) + "' ";
                }
                if (CBKLong.Checked)
                {
                    strSQL = strSQL + "AND prm_active_flag ='Y'";
                }

                //strSQL = "EXEC SP_INV_PRODUK_MASTER_LIST 0, '" + txtProdMaster.Text + "' , " +
                //     "'"+popUpBrand1.txtBrandDesc.Text + "', " +
                //     "'" + popUpBrand1.txtGroupDesc.Text + "'," + 
                //     "'"+ popUpBrand1.txtSubGDesc.Text + "'," + 
                //     "'"+ popUpModel1.txtModelDesc.Text + "'," +
                //     "'" + popUpRawMCode1.txtRawMCodeId.Text + "'," +
                //     "'" + txtbarcode.Text + "'," +
                //      "'" + txtdesc.Text + "'," +
                //     "'" + popUpColorCode1.txtColorDesc.Text + "'," +
                //     "'" + CBProType.SelectedValue.ToString().Trim() + "'," +
                //     "'"+ CBWarehose.SelectedValue.ToString().Trim() + "',"+
                //      "'" + CbNonStock.SelectedValue.ToString().Trim() + "',"+
                //     "'" + CbTech.SelectedValue.ToString().Trim() + "',"+
                //     "'" + CbAktif.SelectedValue.ToString().Trim() + "'," +
                //     "'" + txtTrend.Text + "',"+
                //      "'" + txtmarket.Text + "'," +
                //      "'" + txtseason.Text + "'," +
                //     "'" + txtKategori.Text + "'," +
                //      //"'" + CbOrder.SelectedValue.ToString().Trim() + "',"+
                //     "'" + CEK.ToString() + "',"+
                //     "'" + Convert.ToDateTime(dateTimePicker1.Value.ToString().Trim()).ToString("yyyyMMdd") + "'," +
                //     "'" + Convert.ToDateTime(dateTimePicker2.Value.ToString().Trim()).ToString("yyyyMMdd") + "'," +

                //      " '"+ CbSex.SelectedValue.ToString().Trim() + "' ";


                frm.QueryString(strSQL);
                frm.Parameter("ProgId", paramMenuId);
                frm.Parameter("Group", "" + (txtGroupId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtGroupId.Text.Trim()));
                frm.Parameter("SGroup", "" + (txtSubGId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtSubGId.Text.Trim()));
                frm.Parameter("Category", "" + (txtKategori.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtKategori.Text.Trim()));
                frm.Parameter("Style", "" + (txtTrend.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtTrend.Text.Trim()));
                frm.Parameter("LevelOrder", "" + CbOrder.SelectedValue.ToString().Trim());//CbOrder.SelectedValue.ToString().Trim()
                frm.Parameter("Sex", "" + CbSex.SelectedValue.ToString().Trim());
                frm.Parameter("Market", "" + (txtmarket.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtmarket.Text.Trim()));
                frm.Parameter("Color", "" + (txtColorId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtColorId.Text.Trim()));
                frm.Parameter("RawMat", "" + (txtRawMCodeId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtRawMCodeId.Text.Trim()));
                frm.Parameter("LevelDetail", "" + CbAktif.SelectedValue.ToString().Trim());
                frm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnPopUpBrandIdE_Click(object sender, EventArgs e)
        {
            try
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Brand";
                frm.Query = "SELECT pl_prd_line_code as 'Prod. Line', pl_prd_line_desc as 'Description' " +
                            "  From IM_PRD_LINE " +
                            " ORDER BY pl_prd_line_code";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    txtBrandDesc.Text = frm.ArrField[0].Trim();
                    txtBrandId.Text = frm.ArrField[1].Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void txtBrandId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                strSQL = "SELECT pl_prd_line_code as 'Prod. Line', pl_prd_line_desc as 'Description' " +
                        "  From IM_PRD_LINE " +
                        " Where pl_prd_line_desc='" + txtBrandId.Text.Trim() + "'" +
                        " ORDER BY pl_prd_line_code";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    txtBrandDesc.Text = dt.Rows[0]["Prod. Line"].ToString().Trim();
                }
                else
                {
                    txtBrandDesc.Text = string.Empty;
                }
            }
            catch (Exception ex)
            { }
        }

        private void btnPopUpGroup_Click(object sender, EventArgs e)
        {
            try
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Group";
                frm.Query = "SELECT distinct pg_prd_group_code as 'Prod. Group', pg_prd_group_desc as 'Description' " +
                            "  From IM_PRD_GROUP " +
                            "  WHERE pg_prd_line_code like '" + txtBrandDesc.Text.Trim() + "%' " +
                            " ORDER BY pg_prd_group_code";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    txtGroupDesc.Text = frm.ArrField[0].Trim();
                    txtGroupId.Text = frm.ArrField[1].Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnPopUpSubGr_Click(object sender, EventArgs e)
        {
            try
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Sub Group";
                frm.Query = "SELECT distinct psg_prd_sgroup_code as 'Prod. Subgroup', psg_prd_sgroup_desc as 'Description ' " +
                            "  From IM_PRD_SGROUP " +
                            " WHERE psg_prd_line like '" + txtBrandDesc.Text.Trim() + "%' " +
                            "   AND psg_prd_group_code like '" + txtGroupDesc.Text.Trim() + "%' " +
                            " ORDER BY psg_prd_sgroup_code";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    txtSubGDesc.Text = frm.ArrField[0].Trim();
                    txtSubGId.Text = frm.ArrField[1].Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnPopUpModel_Click(object sender, EventArgs e)
        {
            try
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Model";
                frm.Query = "SELECT pm_prd_model_code as 'Model Code',pm_prd_model_desc as 'Description' " +
                            "  From IM_PRD_MODEL " +
                            " ORDER BY pm_prd_model_code";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    txtModelDesc.Text = frm.ArrField[0].Trim();
                    txtModelId.Text = frm.ArrField[1].Trim();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnPopUpRawM_Click(object sender, EventArgs e)
        {
            try
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Model";
                frm.Query = "SELECT rmu_raw_mat_used_code as 'Raw Material', rmu_raw_mat_used_desc as 'Description' " +
                            "  From IM_RM_USED " +
                            " ORDER BY rmu_raw_mat_used_code";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    txtRawMCodeDesc.Text = frm.ArrField[0].Trim();
                    txtRawMCodeId.Text = frm.ArrField[1].Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void txtGroupId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                    strSQL = "SELECT distinct pg_prd_group_code as 'Prod. Group', pg_prd_group_desc as 'Description' " +
                            "  From IM_PRD_GROUP " +
                            "  WHERE pg_prd_line_code like '" + txtBrandDesc.Text.Trim() + "%' and  pg_prd_group_desc='" + txtGroupId.Text.Trim() + "'" +
                            " ORDER BY pg_prd_group_code";

                    DataTable dt = _clsGlobal.ExecDT(strSQL);

                    if (dt.Rows.Count > 0)
                    {
                        txtGroupDesc.Text = dt.Rows[0]["Prod. Group"].ToString().Trim();
                    }
                    else
                    {
                        txtGroupDesc.Text = string.Empty;
                    }                
            }
            catch (Exception ex)
            { }
        }

        private void txtSubGId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                strSQL = "";
                strSQL = "SELECT distinct psg_prd_sgroup_code as 'Prod. Subgroup', psg_prd_sgroup_desc as 'Description ' " +
                        "  From IM_PRD_SGROUP " +
                        " WHERE psg_prd_line like '" + txtBrandDesc.Text.Trim() + "%' " +
                        "   AND psg_prd_group_code like '" + txtGroupDesc.Text.Trim() + "%' " +
                        "   AND psg_prd_sgroup_desc = '" + txtSubGId.Text.Trim() + "' " +
                        " ORDER BY psg_prd_sgroup_code";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    txtSubGDesc.Text = dt.Rows[0]["Prod. Subgroup"].ToString().Trim();
                }
                else
                {
                    txtSubGDesc.Text = string.Empty;
                }

            }
            catch (Exception ex)
            { }
        }

        private void txtModelId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if(txtModelId.Text =="")
                {
                    txtModelDesc.Text = "";
                }
                else
                {
                    strSQL = "";
                    strSQL = "SELECT pm_prd_model_code as 'Model Code',pm_prd_model_desc as 'Description' " +
                            "  From IM_PRD_MODEL " +
                            " where pm_prd_model_code ='" + txtModelDesc.Text.Trim() + "' " +
                            " ORDER BY pm_prd_model_code";

                    DataTable dt = _clsGlobal.ExecDT(strSQL);

                    if (dt.Rows.Count > 0)
                    {
                        txtModelDesc.Text = dt.Rows[0]["Model Code"].ToString().Trim();
                    }
                    else
                    {
                        txtModelDesc.Text = string.Empty;
                    }  
                }
                             
            }
            catch (Exception ex)
            { }
        }

        private void txtRawMCodeId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                    strSQL = "";
                    strSQL = "SELECT rmu_raw_mat_used_code as 'Raw Material', rmu_raw_mat_used_desc as 'Description' " +
                            "  From IM_RM_USED " +
                            " where rmu_raw_mat_used_desc= '" + txtRawMCodeId.Text.Trim() + "'" +
                            " ORDER BY rmu_raw_mat_used_code";

                    DataTable dt = _clsGlobal.ExecDT(strSQL);

                    if (dt.Rows.Count > 0)
                    {
                        txtRawMCodeDesc.Text = dt.Rows[0]["Raw Material"].ToString().Trim();
                    }
                    else
                    {
                        txtRawMCodeDesc.Text = string.Empty;
                    }                
            }
            catch (Exception ex)
            { }
        }

        private void btnPopUpColor_Click(object sender, EventArgs e)
        {
            try
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Color";
                strSQL = "";
                strSQL = "SELECT col_washing_collor_code as 'Color', col_washing_collor_desc as 'Description' " +
                            "  From IM_W_COLLOR " +
                            " ORDER BY col_washing_collor_code";
                frm.Query = strSQL;
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    txtColorId.Text = frm.ArrField[1].Trim();
                    txtColorDesc.Text = frm.ArrField[0].Trim();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void txtColorId_TextChanged(object sender, EventArgs e)
        {
            try
            {
                strSQL = "";
                strSQL = "SELECT col_washing_collor_code as 'Color', col_washing_collor_desc as 'Description' " +
                             "  From IM_W_COLLOR where col_washing_collor_code='" + txtColorId.Text.Trim() + "'" +
                             " ORDER BY col_washing_collor_code";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    txtColorDesc.Text = dt.Rows[0]["Description"].ToString().Trim();
                }
                else
                {
                    txtColorDesc.Text = string.Empty;
                }
            }
            catch (Exception ex)
            { }
        }
        }
  
    
}
