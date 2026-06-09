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
    public partial class eARCustMasterList : TIRASnDNet.BaseListForm
    {
        const int WidthList = 1680;
        const int HeightList = 790;
        private clsGlobal _clsGlobal = new clsGlobal();
        private string strSQL;
        string paramMenuId;

        const string QUERY = @"SELECT gh_function_code AS [code],gh_function_code+'-'+gh_function_desc AS [value] FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name LIKE 'OUTLETSTATUS'";

        public eARCustMasterList()
        {
            InitializeComponent();
            InitializeDevExpressGrid();
            ApplyGridHeaderCenterAlignment();
            ApplyFullLayout();
        }


        private void ApplyGridHeaderCenterAlignment()
        {
            dgvCustListView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            dgvCustListView.Appearance.HeaderPanel.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;

            foreach (DevExpress.XtraGrid.Views.BandedGrid.GridBand band in dgvCustListView.Bands)
            {
                band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                band.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                band.AppearanceHeader.Options.UseTextOptions = true;
            }

            foreach (DevExpress.XtraGrid.Columns.GridColumn column in dgvCustListView.Columns)
            {
                column.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                column.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                column.AppearanceHeader.Options.UseTextOptions = true;
            }
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
            if (dgvCustListView.RowCount <= 0 || dgvCustListView.FocusedRowHandle < 0) return;

            clsGlobal.MODE_TRX = 2;
            eARCustMasterEntry frm = new eARCustMasterEntry();
            frm.PrdBrandCode = GetFocusedGridValue("cm_branch");
            frm.PrdEntityCode = "01";

            string[] __code2 = GetFocusedGridValue("cm_cust_code").Split('~');
            frm.PrdGrade = __code2.Length > 1 ? __code2[1] : string.Empty;
            frm.PrdCustCode1 = GetFocusedGridValue("cm_cust_code1");
            frm.PrmSalesman = popUpDistrikCM1.txtSalesman.Text;
            frm.PrmSalesmanTo = popUpDistrikCM1.txtSalesmanTo.Text;
            frm.State = eARCustMasterEntry.StateEntry.Edit;
            frm.ShowDialog();
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

                // Pada form lama, parameter Type Outlet filter kiri tidak dikirim ke SP ini.
                // Type Outlet yang dipakai untuk clusterisasi dikirim melalui CtrlProduklineCust1 di parameter berikutnya.
                strSQL += "NULL, ";
                strSQL += "NULL, ";

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
                statFlag = GetCheckedStatusCodes("'',''");
                strSQL += string.Format("'{0}', ", statFlag);

                strSQL += ((txtCustName.Text != "") ? FmtStr(txtCustName.Text) : "NULL") + ", ";
                strSQL += ((txtAddess.Text != "") ? FmtStr(txtAddess.Text) : "NULL") + ",";
                strSQL += string.Format("'{0}',", clsLogin.USERID);
                strSQL += ((txtDestSite.Text != "") ? FmtStr(txtDestSite.Text) : "NULL") + "";

                DataTable dtResult = _clsGlobal.ExecDT(strSQL);
                dgvCustList.BeginUpdate();
                try
                {
                    dgvCustList.DataSource = null;
                    dgvCustList.DataSource = dtResult;
                    dgvCustListView.RefreshData();
                    dgvCustListView.BestFitColumns();
                }
                finally
                {
                    dgvCustList.EndUpdate();
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

        //private void dgvCustList_CellDoubleClick(object sender, EventArgs e)
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

                barBtnDelete.Enabled = false;

                checkBoxAllhirarki.Checked = true;
                radioBtnRegisterOutlet.Checked = true;
                checkBoxAllNpwp.Checked = true;
                ddlPajak.SelectedIndex = -1;
                ddlPajak.EditValue = null;

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

                LoadStatusOutlet();
                CheckAllStatusOutlet();
                ApplyFullLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void AccessButton()
        {
            barBtnNew.Enabled = clsLogin.BTNNEW;
            barBtnEdit.Enabled = clsLogin.BTNEDIT;
            barBtnDelete.Enabled = clsLogin.BTNDELETE;
            barBtnPrint.Enabled = clsLogin.BTNPRINT;
            barBtnClose.Enabled = true;
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
                checkBoxTdkNpwp.Properties.Caption = "Ya";
            }
            else
            {
                checkBoxTdkNpwp.Properties.Caption = "Tidak";
            }
        }

        private void checkBoxYa_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxYa.Checked)
            {
                checkBoxYa.Properties.Caption = "Ya";

            }
            else
            {
                checkBoxYa.Properties.Caption = "Tidak";
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
            // DevExpress BandedGridView replaces the custom DataGridView stacked-header painting.
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

            stOutlet = GetCheckedStatusCodes("'',''");
            frm.PrmStatusOutlet = stOutlet;
            frm.PrmStatusOutletDesc = GetCheckedStatusDescriptions(", ");

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
            ApplyFullLayout();
        }

        private void ApplyFullLayout()
        {
            if (contentPanel == null || groupBox1 == null || panel1 == null || dgvCustList == null) return;

            int availableWidth = Math.Max(1, contentPanel.ClientSize.Width);
            int availableHeight = Math.Max(1, contentPanel.ClientSize.Height);
            const int bottomMargin = 8;

            lblHeader.Width = availableWidth;
            groupBox1.Width = availableWidth;

            // Tombol Execute selalu diposisikan ulang ke pojok kanan area filter yang terlihat.
            // Tidak memakai Anchor Right agar designer tetap aman dan runtime tetap konsisten.
            btnExec.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            btnExec.Left = Math.Max(0, groupBox1.ClientSize.Width - btnExec.Width - 80);
            btnExec.Top = 251;
            btnExec.Visible = true;
            btnExec.BringToFront();

            // Sisakan sedikit ruang bawah agar scrollbar horizontal DevExpress terlihat, tetapi grid tetap lebih turun.
            panel1.Left = 0;
            panel1.Width = availableWidth;
            panel1.Height = Math.Max(180, availableHeight - panel1.Top - bottomMargin);

            dgvCustList.Dock = DockStyle.Fill;
            dgvCustListView.OptionsView.ColumnAutoWidth = false;
            dgvCustListView.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            dgvCustListView.VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            dgvCustListView.LayoutChanged();
        }

        private void eARCustMasterList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //FillGrid();
                btnExec_Click(null, null);
            }
        }

        private void dgvCustList_CellPainting(object sender, EventArgs e)
        {
            // No-op: DevExpress grid appearance is configured in InitializeDevExpressGrid.
        }

        private void dgvCustList_CellDoubleClick(object sender, EventArgs e)
        {
            if (clsLogin.BTNEDIT == true || clsLogin.BTNPRINT == true)
            {
                tsb_edit_Click(sender, e);
            }

        }

        private void dgvCustList_Scroll(object sender, EventArgs e)
        {
            dgvCustList.Refresh();
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

        private void dgvCustList_DataError(object sender, EventArgs e)
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

        protected override void OnNew()
        {
            tsb_new_Click(this, EventArgs.Empty);
        }

        protected override void OnEdit()
        {
            tsb_edit_Click(this, EventArgs.Empty);
        }

        protected override void OnDelete()
        {
            tsb_delete_Click(this, EventArgs.Empty);
        }

        protected override void OnPrint()
        {
            tsb_print_Click(this, EventArgs.Empty);
        }

        protected override void OnCloseForm()
        {
            tsb_close_Click(this, EventArgs.Empty);
        }

        private void InitializeDevExpressGrid()
        {
            dgvCustListView.OptionsBehavior.Editable = false;
            dgvCustListView.OptionsSelection.EnableAppearanceFocusedCell = false;
            dgvCustListView.OptionsView.ColumnAutoWidth = false;
            dgvCustListView.OptionsView.ShowGroupPanel = false;
            dgvCustListView.OptionsView.ShowColumnHeaders = true;
            dgvCustListView.OptionsView.ShowBands = true;
        }

        private string GetFocusedGridValue(string fieldName)
        {
            object value = dgvCustListView.GetFocusedRowCellValue(fieldName);
            return value == null || value == DBNull.Value ? string.Empty : value.ToString().Trim();
        }

        private void dgvCustListView_DoubleClick(object sender, EventArgs e)
        {
            if (clsLogin.BTNEDIT == true || clsLogin.BTNPRINT == true)
            {
                tsb_edit_Click(sender, e);
            }
        }

        private void dgvCustListView_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        private void LoadStatusOutlet()
        {
            statusCLB.Items.Clear();
            DataTable dt = _clsGlobal.ExecDT(QUERY);
            foreach (DataRow row in dt.Rows)
            {
                string code = row["code"] == DBNull.Value ? string.Empty : row["code"].ToString();
                string value = row["value"] == DBNull.Value ? string.Empty : row["value"].ToString();
                statusCLB.Items.Add(new DevExpress.XtraEditors.Controls.CheckedListBoxItem(new StatusItem(code, value), false));
            }
        }

        private void CheckAllStatusOutlet()
        {
            for (int i = 0; i < statusCLB.ItemCount; i++)
            {
                statusCLB.SetItemChecked(i, true);
            }
        }

        private string GetCheckedStatusCodes(string separator)
        {
            return string.Join(separator, GetCheckedStatusItems().Select(x => x.Code));
        }

        private string GetCheckedStatusDescriptions(string separator)
        {
            return string.Join(separator, GetCheckedStatusItems().Select(x => x.Value));
        }

        private IEnumerable<StatusItem> GetCheckedStatusItems()
        {
            for (int i = 0; i < statusCLB.ItemCount; i++)
            {
                if (statusCLB.GetItemChecked(i))
                {
                    object rawItem = statusCLB.GetItem(i);
                    DevExpress.XtraEditors.Controls.CheckedListBoxItem checkedItem = rawItem as DevExpress.XtraEditors.Controls.CheckedListBoxItem;
                    StatusItem item = null;

                    if (checkedItem != null)
                    {
                        item = checkedItem.Value as StatusItem;
                    }
                    else
                    {
                        item = rawItem as StatusItem;
                    }

                    if (item != null) yield return item;
                }
            }
        }

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

        public override string ToString()
        {
            return Value;
        }
    }
}
