using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TIRASnDNet.UsrControls;

namespace TIRASnDNet.AR.ARCustMaster
{
    public partial class eARCustMasterList : Form
    {
        const int WidthList = 940;
        const int HeightList = 610;
        private clsGlobal _clsGlobal = new clsGlobal();
        private string strSQL;
        string paramMenuId;

        const string QUERY = @"SELECT gh_function_code AS [code],gh_function_code+'-'+gh_function_desc AS [value] FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name LIKE 'OUTLETSTATUS'";
        readonly CheckListProvider<StatusItem> __statusOutlet = null;
        public eARCustMasterList()
        {
            InitializeComponent();
            new StackedHeader.StackedHeaderDecorator(this.dgvCustList);
            __statusOutlet = new CheckListProvider<StatusItem>(statusCLB, QUERY);
            __statusOutlet.DisplayMember = "value";
            __statusOutlet.ValueMember = "code";
        }

        #region "Toolstrip Buttons"

        private void tsb_new_Click(object sender, EventArgs e)
        {
            clsGlobal.MODE_TRX = 1;
            eARCustMasterEntry frm = new eARCustMasterEntry();
            frm.State = eARCustMasterEntry.StateEntry.New;
            frm.ShowDialog();
        }

        private void tsb_edit_Click(object sender, EventArgs e)
        {
            if (dgvCustList.Rows.Count > 0)
            {
                int rowindex = dgvCustList.CurrentCell.RowIndex;

                //david 17 Mei 2018
                clsGlobal.MODE_TRX = 2;
                eARCustMasterEntry frm = new eARCustMasterEntry();
                frm.PrdBrandCode = dgvCustList.Rows[rowindex].Cells["cm_branch"].Value.ToString().Trim();
                frm.PrdEntityCode = "01";              

                //frm.PrdGrade = dgvCustList.Rows[rowindex].Cells["cm_cust_code"].Value.ToString().Trim().Substring(7);
                string[] __code2 = dgvCustList.Rows[rowindex].Cells["cm_cust_code"].Value.ToString().Trim().Split('~');
                frm.PrdGrade = __code2.Length > 1 ? __code2[1] : string.Empty;
                frm.PrdCustCode1 = dgvCustList.Rows[rowindex].Cells["cm_cust_code1"].Value.ToString().Trim();
                //frm.PrdCustCode2 = dgvCustList.Rows[rowindex].Cells["cm_cust_code2"].Value.ToString().Trim();
                frm.PrmSalesman = popUpDistrikCM1.txtSalesman.Text;
                frm.PrmSalesmanTo = popUpDistrikCM1.txtSalesmanTo.Text;
                frm.State = eARCustMasterEntry.StateEntry.Edit;
                frm.ShowDialog();
            }
        }

        private void tsb_delete_Click(object sender, EventArgs e)
        {

        }

        #endregion

        #region "Execute"

        private void FillGrid()
        {
            string statFlag = "NULL";
            try
            {
                strSQL = "exec SP_AR_CUST_MASTER_LIST ";

                strSQL += ((popUpDistrikCM1.txtEntityId.Text != "") ? FmtStr(popUpDistrikCM1.txtEntityId.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtEntityIdTo.Text != "") ? FmtStr(popUpDistrikCM1.txtEntityIdTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtBranchId.Text != "") ? FmtStr(popUpDistrikCM1.txtBranchId.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtBranchIdTo.Text != "") ? FmtStr(popUpDistrikCM1.txtBranchIdTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtDistrik.Text != "") ? FmtStr(popUpDistrikCM1.txtDistrik.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtDistrikTo.Text != "") ? FmtStr(popUpDistrikCM1.txtDistrikTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtBeat.Text != "") ? FmtStr(popUpDistrikCM1.txtBeat.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtBeatTo.Text != "") ? FmtStr(popUpDistrikCM1.txtBeatTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtSubBeat.Text != "") ? FmtStr(popUpDistrikCM1.txtSubBeat.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtSubBeatTo.Text != "") ? FmtStr(popUpDistrikCM1.txtSubBeatTo.Text) : "NULL") + ",";

                strSQL += ((popUpDistrikCM1.txtTypeOutlet.Text != "") ? FmtStr(popUpDistrikCM1.txtTypeOutlet.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtTypeOutletTo.Text != "") ? FmtStr(popUpDistrikCM1.txtTypeOutletTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtLokasi.Text != "") ? FmtStr(popUpDistrikCM1.txtLokasi.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtLokasiTo.Text != "") ? FmtStr(popUpDistrikCM1.txtLokasiTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtKodePasar.Text != "") ? FmtStr(popUpDistrikCM1.txtKodePasar.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtKodePasarTo.Text != "") ? FmtStr(popUpDistrikCM1.txtKodePasarTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtKlasifikasi.Text != "") ? FmtStr(popUpDistrikCM1.txtKlasifikasi.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtKlasifikasiTo.Text != "") ? FmtStr(popUpDistrikCM1.txtKlasifikasiTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtGroup.Text != "") ? FmtStr(popUpDistrikCM1.txtGroup.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtGroupTo.Text != "") ? FmtStr(popUpDistrikCM1.txtGroupTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtKategori.Text != "") ? FmtStr(popUpDistrikCM1.txtKategori.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtKategoriTo.Text != "") ? FmtStr(popUpDistrikCM1.txtKategoriTo.Text) : "NULL") + ", ";

                strSQL += ((popUpDistrikCM1.txtSalesman.Text != "") ? FmtStr(popUpDistrikCM1.txtSalesman.Text) : "NULL") + ", ";
                strSQL += ((popUpDistrikCM1.txtSalesmanTo.Text != "") ? FmtStr(popUpDistrikCM1.txtSalesmanTo.Text) : "NULL") + ", ";

                strSQL += ((ctrlUpRegionsCM1.txtPropinsi.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtPropinsi.Text) : "NULL") + ", ";
                strSQL += ((ctrlUpRegionsCM1.txtPropinsiTo.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtPropinsiTo.Text) : "NULL") + ", ";

                strSQL += ((ctrlUpRegionsCM1.txtKabupaten.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtKabupaten.Text) : "NULL") + ", ";
                strSQL += ((ctrlUpRegionsCM1.txtKabupatenTo.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtKabupatenTo.Text) : "NULL") + ", ";

                strSQL += ((ctrlUpRegionsCM1.txtKecamatan.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtKecamatan.Text) : "NULL") + ", ";
                strSQL += ((ctrlUpRegionsCM1.txtKecamatanTo.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtKecamatanTo.Text) : "NULL") + ", ";

                strSQL += ((ctrlUpRegionsCM1.txtKelurahan.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtKelurahan.Text) : "NULL") + ", ";
                strSQL += ((ctrlUpRegionsCM1.txtKelurahanTo.Text != "") ? FmtStr(ctrlUpRegionsCM1.txtKelurahanTo.Text) : "NULL") + ", ";

                strSQL += ((txtCustCode.Text != "") ? FmtStr(txtCustCode.Text) : "NULL") + ", ";
                strSQL += ((txtCustCodeTo.Text != "") ? FmtStr(txtCustCodeTo.Text) : "NULL") + ", ";

                strSQL += ((CtrlProduklineCust1.txtProductLine.Text != "") ? FmtStr(CtrlProduklineCust1.txtProductLine.Text) : "NULL") + ", ";
                strSQL += ((CtrlProduklineCust1.txtProductLineTo.Text != "") ? FmtStr(CtrlProduklineCust1.txtProductLineTo.Text) : "NULL") + ", ";

                strSQL += ((CtrlProduklineCust1.txtTipeOutletCLU.Text != "") ? FmtStr(CtrlProduklineCust1.txtTipeOutletCLU.Text) : "NULL") + ", ";
                strSQL += ((CtrlProduklineCust1.txtTipeOutletCLUTo.Text != "") ? FmtStr(CtrlProduklineCust1.txtTipeOutletCLUTo.Text) : "NULL") + ", ";

                strSQL += ((checkBoxAllNpwp.Checked) ? "NULL" : "N") + ", ";
                strSQL += (ddlPajak.SelectedIndex) + ", ";

                //if (radioBtnSemuaOutlet.Checked)
                //{
                //    statFlag = "NULL";
                //}
                //if (radioBtnConvertedOutlet.Checked)
                //{
                //    statFlag = "C";
                //}
                //if (radioBtnOnoOutlet.Checked)
                //{
                //    statFlag = "O";
                //}
                //if (radioBtnRegisterOutlet.Checked)
                //{
                //    statFlag = "R";
                //}
                //if (radioButtonNonActiveOutlet.Checked)
                //{
                //    statFlag = "N";
                //}
                //if (radioBtnClosedOutlet.Checked)
                //{
                //    statFlag = "D";
                //}
                //if (radioBtnPotOutlet.Checked)
                //{
                //    statFlag = "P";
                //}
                statFlag = __statusOutlet.GetCheckedAsString("'',''", n => n.Code);
                strSQL += string.Format("'{0}', ", statFlag);

                strSQL += ((txtCustName.Text != "") ? FmtStr(txtCustName.Text) : "NULL") + ", ";
                strSQL += ((txtAddess.Text != "") ? FmtStr(txtAddess.Text) : "NULL") + ",";
                strSQL += string.Format("'{0}',", clsLogin.USERID);
                strSQL += ((txtDestSite.Text != "") ? FmtStr(txtDestSite.Text) : "NULL") + "";

                dgvCustList.AutoGenerateColumns = false;

                dgvCustList.DataSource = _clsGlobal.ExecDT(strSQL);

                if (dgvCustList.Rows.Count > 0)
                {
                    DataGridViewCellEventArgs DataGridViewCellEventArgs = new DataGridViewCellEventArgs(dgvCustList.Columns["cm_cust_code1"].Index, dgvCustList.CurrentRow.Index);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //private void dgvCustList_CellClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    if (e.RowIndex > -1)
        //    {
        //        //pCustID = dgvCustList.Rows[rowindex].Cells["cm_cust_code1"].Value.ToString().Trim();
        //        //pCustID2 = dgvCustList.Rows[rowindex].Cells["cm_cust_group"].Value.ToString().Trim();
        //    }
        //}

        //private void dgvCustList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    if (clsLogin.BTNEDIT)
        //    {
        //        if (e.RowIndex > -1)
        //        {
        //            dgvCustList_CellClick(sender, e);
        //            tsb_edit_Click(sender, e);
        //        }
        //    }
        //}

        #endregion

        private void checkBoxAllhirarki_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAllhirarki.Checked)
            {
                checkBoxYa.Enabled = false;
                CtrlProduklineCust1.txtProductLine.Enabled = false;
                CtrlProduklineCust1.txtProductLineTo.Enabled = false;
                CtrlProduklineCust1.txtTipeOutletCLU.Enabled = false;
                CtrlProduklineCust1.txtTipeOutletCLUTo.Enabled = false;
                CtrlProduklineCust1.btnPopUpProductLine.Enabled = false;
                CtrlProduklineCust1.btnPopUpProductLineTo.Enabled = false;
                CtrlProduklineCust1.btnPopUpTipeOutlet.Enabled = false;
                CtrlProduklineCust1.btnPopUpTipeOutletTo.Enabled = false;

            }
            else
            {
                checkBoxYa.Enabled = true;
                CtrlProduklineCust1.txtProductLine.Enabled = true;
                CtrlProduklineCust1.txtProductLineTo.Enabled = true;
                CtrlProduklineCust1.txtTipeOutletCLU.Enabled = true;
                CtrlProduklineCust1.txtTipeOutletCLUTo.Enabled = true;
                CtrlProduklineCust1.btnPopUpProductLine.Enabled = true;
                CtrlProduklineCust1.btnPopUpProductLineTo.Enabled = true;
                CtrlProduklineCust1.btnPopUpTipeOutlet.Enabled = true;
                CtrlProduklineCust1.btnPopUpTipeOutletTo.Enabled = true;
            }
        }

        private void eARCustMasterList_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = clsLogin.MENUID + " - " + clsLogin.MENUNAME;
                this.WindowState = FormWindowState.Maximized;
                paramMenuId = clsLogin.MENUID;
                AccessButton();

                tsb_delete.Enabled = false;

                checkBoxAllhirarki.Checked = true;
                radioBtnRegisterOutlet.Checked = true;
                checkBoxAllNpwp.Checked = true;
                ddlPajak.SelectedText = "0 ~ Standart";

                //strSQL = " select top 1 br_branch_short,br_branch_desc from GS_BRANCH  ";
                strSQL = "SELECT A.gu_entity,A.gu_branch, B.br_branch_short,B.br_branch_desc FROM VW_USERS_SECURITY A inner join GS_BRANCH B WITH (NOLOCK) on A.gu_entity = B.br_gl_entity_initial  AND A.gu_branch = B.br_branch_id AND A.gu_user_id = '" + clsLogin.USERID + "'";
                DataTable dt = new DataTable();
                dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    lblHeader.Text = "         " + this.Text + " - " + dt.Rows[0]["br_branch_desc"].ToString().Trim() + " - " + dt.Rows[0]["br_branch_short"].ToString().Trim() + "";
                    popUpDistrikCM1.txtEntityId.Text = dt.Rows[0]["gu_entity"].ToString();
                    popUpDistrikCM1.txtBranchId.Text = dt.Rows[0]["gu_branch"].ToString();
                }

                __statusOutlet.Execute();
                //__statusOutlet.InsertItems(0, new StatusItem("R", "R-REGISTER"));
                //__statusOutlet.Refresh();
                __statusOutlet.CheckAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AccessButton()
        {
            tsb_new.Enabled = clsLogin.BTNNEW;
            tsb_edit.Enabled = clsLogin.BTNEDIT;
            tsb_delete.Enabled = clsLogin.BTNDELETE;
            tsb_print.Enabled = clsLogin.BTNPRINT;
        }

        private void checkBoxAllNpwp_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAllNpwp.Checked)
            {
                ddlPajak.Enabled = false;
                checkBoxTdkNpwp.Enabled = false;
            }
            else
            {
                checkBoxTdkNpwp.Enabled = true;
                ddlPajak.Enabled = true;
            }
        }

        private void checkBoxTdkNpwp_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxTdkNpwp.Checked)
            {
                checkBoxTdkNpwp.Text = "Ya";
            }
            else
            {
                checkBoxTdkNpwp.Text = "Tidak";
            }
        }

        private void checkBoxYa_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxYa.Checked)
            {
                checkBoxYa.Text = "Ya";

            }
            else
            {
                checkBoxYa.Text = "Tidak";
            }

        }

        private void tsb_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExec_Click(object sender, EventArgs e)
        {
            MessageBoxIcon __msg = MessageBoxIcon.Exclamation;
            try
            {
                if (popUpDistrikCM1.txtEntityId.Text.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception("Entity is empty.");

                if (popUpDistrikCM1.txtBranchId.Text.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception("Branch is empty.");

                if (sender != null)
                {
                    if (txtCustCode.Text.ToString() != "" && txtCustCodeTo.Text == "")
                        txtCustCodeTo.Text = txtCustCode.Text;

                    if (popUpDistrikCM1.txtSalesman.Text.ToString() != "" && popUpDistrikCM1.txtSalesmanTo.Text == "")
                        popUpDistrikCM1.txtSalesmanTo.Text = popUpDistrikCM1.txtSalesman.Text;
                }

                __msg = MessageBoxIcon.Error;
                FillGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, __msg);
            }
        }

        //bool __isFirst = true;
        private void dgvCustList_Paint(object sender, PaintEventArgs e)
        {
            //if (!__isFirst) return;
            //Atur Format
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            format.LineAlignment = StringAlignment.Center;

            try
            {
                //----------------------------------- BRANCH ------------------------------------- 
                Rectangle r1a = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_branch"].Index, -1, true);
                int r2a = this.dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["gbc_branch"].Index, -1, true).Width;

                r1a.X += 1;
                r1a.Y += 1;
                r1a.Width = (r1a.Width + r2a) - 2;
                r1a.Height = r1a.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r1a);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Branch", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r1a, sf);
                }

                //----------------------------------- OUTLET ------------------------------------- 
                Rectangle r1 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_cust_code1"].Index, -1, true);
                int r2 = this.dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_cust_name"].Index, -1, true).Width;
                int r3 = this.dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_cust_status"].Index, -1, true).Width;

                r1.X += 1;
                r1.Y += 1;
                r1.Width = (r1.Width + r2 + r3) - 2;
                r1.Height = r1.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r1);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Outlet", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r1, sf);
                }

                //------------------------------------ ALAMAT -------------------------------------
                Rectangle r4 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_bill_adrress1"].Index, -1, true);
                int r5 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_bill_adrress2"].Index, -1, true).Width;

                r4.X += 1;
                r4.Y += 1;
                r4.Width = r4.Width + r5 - 2;
                r4.Height = r4.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r4);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Alamat", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r4, sf);
                }

                //------------------------------------ DISTRIK -------------------------------------
                Rectangle r6 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_area"].Index, -1, true);
                int r7 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["rg_region_desc"].Index, -1, true).Width;

                r6.X += 1;
                r6.Y += 1;
                r6.Width = r6.Width + r7 - 2;
                r6.Height = r6.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r6);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Distrik", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r6, sf);
                }


                //------------------------------------ BEAT -------------------------------------
                Rectangle r8 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_wilayah"].Index, -1, true);
                int r9 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["wy_wilayah_desc"].Index, -1, true).Width;

                r8.X += 1;
                r8.Y += 1;
                r8.Width = r8.Width + r9 - 2;
                r8.Height = r8.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r8);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Beat", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r8, sf);
                }

                //------------------------------------ SUB BEAT -------------------------------------
                Rectangle r10 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_rayon"].Index, -1, true);
                int r11 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["ry_rayon_desc"].Index, -1, true).Width;

                r10.X += 1;
                r10.Y += 1;
                r10.Width = r10.Width + r11 - 2;
                r10.Height = r10.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r10);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Sub Beat", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r10, sf);
                }

                //------------------------------------ TYPE OUTLET -------------------------------------
                Rectangle r12 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_cust_type"].Index, -1, true);
                int r13 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["ct_cust_type_desc"].Index, -1, true).Width;

                r12.X += 1;
                r12.Y += 1;
                r12.Width = r12.Width + r13 - 2;
                r12.Height = r12.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r12);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Type Outlet", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r12, sf);
                }

                //------------------------------------ LOKASI -------------------------------------
                Rectangle r14 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_location"].Index, -1, true);
                int r15 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["loc_long_desc"].Index, -1, true).Width;

                r14.X += 1;
                r14.Y += 1;
                r14.Width = r12.Width + r15 - 2;
                r14.Height = r14.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r14);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Lokasi", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r14, sf);
                }

                //------------------------------------ PASAR -------------------------------------
                Rectangle r16 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_pasar_code"].Index, -1, true);
                int r17 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["psr_long_desc"].Index, -1, true).Width;

                r16.X += 1;
                r16.Y += 1;
                r16.Width = r16.Width + r17 - 2;
                r16.Height = r16.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r16);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Pasar", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r16, sf);
                }


                //------------------------------------ KLASIFIKASI -------------------------------------
                Rectangle r18 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_cust_class"].Index, -1, true);
                int r19 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cc_cust_class_desc"].Index, -1, true).Width;

                r18.X += 1;
                r18.Y += 1;
                r18.Width = r18.Width + r19 - 2;
                r18.Height = r18.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r18);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Klasifikasi", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r18, sf);
                }


                //------------------------------------ KATEGORI -------------------------------------
                Rectangle r20 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_cust_catg_code"].Index, -1, true);
                int r21 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["ctg_cust_catg_desc"].Index, -1, true).Width;

                r20.X += 1;
                r20.Y += 1;
                r20.Width = r20.Width + r21 - 2;
                r20.Height = r20.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r20);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Kategori", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r20, sf);
                }


                //------------------------------------ Group Outlet -------------------------------------
                Rectangle r22 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_cust_group"].Index, -1, true);
                int r23 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cg_cust_group_desc"].Index, -1, true).Width;

                r22.X += 1;
                r22.Y += 1;
                r22.Width = r22.Width + r23 - 2;
                r22.Height = r22.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r22);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Group Outlet", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r22, sf);
                }

                //------------------------------------ Pemerintahan -------------------------------------
                Rectangle r24 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_delv_proviency"].Index, -1, true);
                int r25 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_delv_city"].Index, -1, true).Width;
                int r26 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_devl_kecamatan"].Index, -1, true).Width;
                int r27 = dgvCustList.GetCellDisplayRectangle(dgvCustList.Columns["cm_devl_kelurahan"].Index, -1, true).Width;

                r24.X += 1;
                r24.Y += 1;
                r24.Width += r24.Width + r25 + r26 + r27 - 2;
                r24.Height = r24.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(dgvCustList.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r24);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.dgvCustList.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Pemerintahan", dgvCustList.ColumnHeadersDefaultCellStyle.Font, br, r24, sf);
                }
            }
            finally
            {
                //__isFirst = false;
            }
        }

        private string FmtStr(string value_str)
        {
            string formatStr;
            formatStr = "'" + value_str.Trim().Replace("'", "''") + "'";

            return formatStr;
        }

        private void tsb_print_Click(object sender, EventArgs e)
        {
            eARCustPrint frm = new eARCustPrint();

            frm.PrmEntity = popUpDistrikCM1.txtEntityId.Text;
            frm.PrmEntityTo = popUpDistrikCM1.txtEntityIdTo.Text;

            frm.PrmBranch = popUpDistrikCM1.txtBranchId.Text;
            frm.PrmBranchTo = popUpDistrikCM1.txtBranchIdTo.Text;

            frm.PrmDistrik = popUpDistrikCM1.txtDistrik.Text;
            frm.PrmDistrikTo = popUpDistrikCM1.txtDistrikTo.Text;

            frm.PrmBeat = popUpDistrikCM1.txtBeat.Text;
            frm.PrmBeatTo = popUpDistrikCM1.txtBeatTo.Text;

            frm.PrmSubBeat = popUpDistrikCM1.txtSubBeat.Text;
            frm.PrmSubBeatTo = popUpDistrikCM1.txtSubBeatTo.Text;

            frm.PrmTypeOutlet = popUpDistrikCM1.txtTypeOutlet.Text;
            frm.PrmTypeOutletTo = popUpDistrikCM1.txtTypeOutletTo.Text;

            frm.PrmLokasi = popUpDistrikCM1.txtLokasi.Text;
            frm.PrmLokasiTo = popUpDistrikCM1.txtLokasiTo.Text;

            frm.PrmKodePasar = popUpDistrikCM1.txtKodePasar.Text;
            frm.PrmKodePasarTo = popUpDistrikCM1.txtKodePasarTo.Text;

            frm.PrmKlasifikasi = popUpDistrikCM1.txtKlasifikasi.Text;
            frm.PrmKlasifikasiTo = popUpDistrikCM1.txtKlasifikasiTo.Text;

            frm.PrmKategori = popUpDistrikCM1.txtKategori.Text;
            frm.PrmKategoriTo = popUpDistrikCM1.txtKategoriTo.Text;

            frm.PrmGroupOutlet = popUpDistrikCM1.txtGroup.Text;
            frm.PrmGroupOutletTo = popUpDistrikCM1.txtGroupTo.Text;

            frm.PrmSalesman = popUpDistrikCM1.txtSalesman.Text;
            frm.PrmSalesmanTo = popUpDistrikCM1.txtSalesmanTo.Text;

            frm.PrmProvince = ctrlUpRegionsCM1.txtPropinsi.Text;
            frm.PrmProvinceTo = ctrlUpRegionsCM1.txtPropinsiTo.Text;

            frm.PrmCity = ctrlUpRegionsCM1.txtKabupaten.Text;
            frm.PrmCityTo = ctrlUpRegionsCM1.txtKabupatenTo.Text;

            frm.Prmkecamatan = ctrlUpRegionsCM1.txtKecamatan.Text;
            frm.PrmkecamatanTo = ctrlUpRegionsCM1.txtKecamatanTo.Text;

            frm.PrmKelurahan = ctrlUpRegionsCM1.txtKelurahan.Text;
            frm.PrmKelurahanTo = ctrlUpRegionsCM1.txtKelurahanTo.Text;

            frm.PrmCustomerCode = txtCustCode.Text;
            frm.PrmCustomerCodeTo = txtCustCodeTo.Text;

            frm.PrmCustomerName = txtCustName.Text;
            frm.PrmBillAddress = txtAddess.Text;


            string stOutlet = "";
            //if (radioBtnSemuaOutlet.Checked)
            //{
            //    stOutlet = "NULL";
            //}
            //if (radioBtnConvertedOutlet.Checked)
            //{
            //    stOutlet = "C";
            //}
            //if (radioBtnOnoOutlet.Checked)
            //{
            //    stOutlet = "O";
            //}
            //if (radioBtnRegisterOutlet.Checked)
            //{
            //    stOutlet = "R";
            //}
            //if (radioButtonNonActiveOutlet.Checked)
            //{
            //    stOutlet = "N";
            //}
            //if (radioBtnClosedOutlet.Checked)
            //{
            //    stOutlet = "D";
            //}
            //if (radioBtnPotOutlet.Checked)
            //{
            //    stOutlet = "P";
            //}

            stOutlet = __statusOutlet.GetCheckedAsString("'',''", m => m.Code);
            frm.PrmStatusOutlet = stOutlet;
            frm.PrmStatusOutletDesc = __statusOutlet.GetCheckedAsString(", ", x => x.Value);

            frm.PrmPajak = ddlPajak.SelectedIndex.ToString();

            frm.PrmPrdLine = CtrlProduklineCust1.txtProductLine.Text;
            frm.PrmPrdLineTo = CtrlProduklineCust1.txtProductLineTo.Text;

            frm.PrmTypeOutletLine = CtrlProduklineCust1.txtTipeOutletCLU.Text;
            frm.PrmTypeOutletLineTo = CtrlProduklineCust1.txtTipeOutletCLUTo.Text;

            frm.PrmDestinationZone = txtDestSite.Text;

            frm.ShowDialog();
        }

        private void eARCustMasterList_Resize(object sender, EventArgs e)
        {
            if (this.Width < WidthList)
            {
                this.Width = WidthList;
            }

            if (this.Height < HeightList)
            {
                this.Height = HeightList;
            }
        }

        int max = 0;
        private void dgvCustList_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dgvCustList.RowHeadersDefaultCellStyle.ForeColor))
            {
                //e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
                SizeF sif = e.Graphics.MeasureString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font);
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
                max = Math.Max(max, (int)sif.Width + 20);
                ((DataGridView)sender).RowHeadersWidth = max;
            }
        }

        private void eARCustMasterList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //FillGrid();
                btnExec_Click(null, null);
            }
        }

        private void dgvCustList_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            _clsGlobal.GridColorCell(sender, e);
        }

        private void dgvCustList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (clsLogin.BTNEDIT == true || clsLogin.BTNPRINT == true)
            {
                tsb_edit_Click(sender, e);
            }
           
        }

        private void dgvCustList_Scroll(object sender, ScrollEventArgs e)
        {
            ((DataGridView)sender).Refresh();
        }

        private void txtCustName_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = (DataTable)dgvCustList.DataSource;
                if (dt == null) return;
                dt.FilterDTable(string.Format("cm_cust_name like '{0}%'", txtCustName.Text));
            }
            catch (Exception) { }
        }

        private void dgvCustList_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void txtDestSite_Leave(object sender, EventArgs e)
        {

        }

        private void btnDestSite_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Delivery Site ";
            frm.Query = " select Distinct itzr_destination_site [Dest Zone],itzr_desc_destination_site [Desc] from IM_TRANS_ZONE_ROUTE WITH (NOLOCK) where itzr_entity = '" + popUpDistrikCM1.txtEntityId.Text + "' and itzr_branch = '" + popUpDistrikCM1.txtBranchId.Text + "' ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtDestSite.Text = frm.ArrField[0].Trim();
            }
        }

        //private void dgvCustList_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        //{
        //    if (clsLogin.BTNEDIT)
        //    {
        //        if (e.RowIndex > -1)
        //        {
        //            dgvVendorMasterList_SelectionChanged(sender, e);
        //            tsb_edit_Click(sender, e);
        //        }
        //    }
        //}                          
    }

    public class StatusItem
    {
        public string Code { get; set; }
        public string Value { get; set; }
        public StatusItem() { }
        public StatusItem(string code, string value)
        {
            this.Code = code;
            this.Value = value;
        }
    }
}
