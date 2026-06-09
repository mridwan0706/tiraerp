using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Base;
using System.Windows.Forms; // diperlukan oleh XtraForm untuk enum FormWindowState, Keys, MessageBoxButtons, MessageBoxIcon
using DevExpress.XtraEditors.Controls;

namespace TIRASnDNet.AR.ARCustMaster
{
    public partial class eARCustMasterList : TIRASnDNet.BaseListForm
    {
        const int WidthList = 940;
        const int HeightList = 610;
        private clsGlobal _clsGlobal = new clsGlobal();
        private string strSQL;
        string paramMenuId;

        const string QUERY = @"SELECT gh_function_code AS [code],gh_function_code+'-'+gh_function_desc AS [value] FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name LIKE 'OUTLETSTATUS'";
        readonly DxCheckListProvider<StatusItem> __statusOutlet = null;
        public eARCustMasterList()
        {
            InitializeComponent();
            // DevExpress GridView menggunakan banded-style caption via GridColumn.Caption; StackedHeader WinForms tidak dipakai lagi.
            __statusOutlet = new DxCheckListProvider<StatusItem>(statusCLB, QUERY);
            __statusOutlet.DisplayMember = "value";
            __statusOutlet.ValueMember = "code";
        }

        #region "Toolstrip Buttons"

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

        private void tsb_new_Click(object sender, EventArgs e)
        {
            clsGlobal.MODE_TRX = 1;
            eARCustMasterEntry frm = new eARCustMasterEntry();
            frm.State = eARCustMasterEntry.StateEntry.New;
            frm.ShowDialog();
        }

        private void tsb_edit_Click(object sender, EventArgs e)
        {
            if (gridViewCustList.RowCount > 0 && gridViewCustList.FocusedRowHandle >= 0)
            {
                int rowindex = gridViewCustList.FocusedRowHandle;

                //david 17 Mei 2018
                clsGlobal.MODE_TRX = 2;
                eARCustMasterEntry frm = new eARCustMasterEntry();
                frm.PrdBrandCode = Convert.ToString(gridViewCustList.GetRowCellValue(rowindex, "cm_branch")).Trim();
                frm.PrdEntityCode = "01";

                //frm.PrdGrade = dgvCustList.Rows[rowindex].Cells["cm_cust_code"].Value.ToString().Trim().Substring(7);
                string[] __code2 = Convert.ToString(gridViewCustList.GetRowCellValue(rowindex, "cm_cust_code")).Trim().Split('~');
                frm.PrdGrade = __code2.Length > 1 ? __code2[1] : string.Empty;
                frm.PrdCustCode1 = Convert.ToString(gridViewCustList.GetRowCellValue(rowindex, "cm_cust_code1")).Trim();
                //frm.PrdCustCode2 = dgvCustList.Rows[rowindex].Cells["cm_cust_code2"].Value.ToString().Trim();
                frm.PrmSalesman = txtSalesman.Text;
                frm.PrmSalesmanTo = txtSalesmanTo.Text;
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

                strSQL += ((txtEntityId.Text != "") ? FmtStr(txtEntityId.Text) : "NULL") + ", ";
                strSQL += ((txtEntityIdTo.Text != "") ? FmtStr(txtEntityIdTo.Text) : "NULL") + ", ";

                strSQL += ((txtBranchId.Text != "") ? FmtStr(txtBranchId.Text) : "NULL") + ", ";
                strSQL += ((txtBranchIdTo.Text != "") ? FmtStr(txtBranchIdTo.Text) : "NULL") + ", ";

                strSQL += ((txtDistrik.Text != "") ? FmtStr(txtDistrik.Text) : "NULL") + ", ";
                strSQL += ((txtDistrikTo.Text != "") ? FmtStr(txtDistrikTo.Text) : "NULL") + ", ";

                strSQL += ((txtBeat.Text != "") ? FmtStr(txtBeat.Text) : "NULL") + ", ";
                strSQL += ((txtBeatTo.Text != "") ? FmtStr(txtBeatTo.Text) : "NULL") + ", ";

                strSQL += ((txtSubBeat.Text != "") ? FmtStr(txtSubBeat.Text) : "NULL") + ", ";
                strSQL += ((txtSubBeatTo.Text != "") ? FmtStr(txtSubBeatTo.Text) : "NULL") + ",";

                strSQL += ((txtTypeOutlet.Text != "") ? FmtStr(txtTypeOutlet.Text) : "NULL") + ", ";
                strSQL += ((txtTypeOutletTo.Text != "") ? FmtStr(txtTypeOutletTo.Text) : "NULL") + ", ";

                strSQL += ((txtLokasi.Text != "") ? FmtStr(txtLokasi.Text) : "NULL") + ", ";
                strSQL += ((txtLokasiTo.Text != "") ? FmtStr(txtLokasiTo.Text) : "NULL") + ", ";

                strSQL += ((txtKodePasar.Text != "") ? FmtStr(txtKodePasar.Text) : "NULL") + ", ";
                strSQL += ((txtKodePasarTo.Text != "") ? FmtStr(txtKodePasarTo.Text) : "NULL") + ", ";

                strSQL += ((txtKlasifikasi.Text != "") ? FmtStr(txtKlasifikasi.Text) : "NULL") + ", ";
                strSQL += ((txtKlasifikasiTo.Text != "") ? FmtStr(txtKlasifikasiTo.Text) : "NULL") + ", ";

                strSQL += ((txtGroup.Text != "") ? FmtStr(txtGroup.Text) : "NULL") + ", ";
                strSQL += ((txtGroupTo.Text != "") ? FmtStr(txtGroupTo.Text) : "NULL") + ", ";

                strSQL += ((txtKategori.Text != "") ? FmtStr(txtKategori.Text) : "NULL") + ", ";
                strSQL += ((txtKategoriTo.Text != "") ? FmtStr(txtKategoriTo.Text) : "NULL") + ", ";

                strSQL += ((txtSalesman.Text != "") ? FmtStr(txtSalesman.Text) : "NULL") + ", ";
                strSQL += ((txtSalesmanTo.Text != "") ? FmtStr(txtSalesmanTo.Text) : "NULL") + ", ";

                strSQL += ((txtPropinsi.Text != "") ? FmtStr(txtPropinsi.Text) : "NULL") + ", ";
                strSQL += ((txtPropinsiTo.Text != "") ? FmtStr(txtPropinsiTo.Text) : "NULL") + ", ";

                strSQL += ((txtKabupaten.Text != "") ? FmtStr(txtKabupaten.Text) : "NULL") + ", ";
                strSQL += ((txtKabupatenTo.Text != "") ? FmtStr(txtKabupatenTo.Text) : "NULL") + ", ";

                strSQL += ((txtKecamatan.Text != "") ? FmtStr(txtKecamatan.Text) : "NULL") + ", ";
                strSQL += ((txtKecamatanTo.Text != "") ? FmtStr(txtKecamatanTo.Text) : "NULL") + ", ";

                strSQL += ((txtKelurahan.Text != "") ? FmtStr(txtKelurahan.Text) : "NULL") + ", ";
                strSQL += ((txtKelurahanTo.Text != "") ? FmtStr(txtKelurahanTo.Text) : "NULL") + ", ";

                strSQL += ((txtCustCode.Text != "") ? FmtStr(txtCustCode.Text) : "NULL") + ", ";
                strSQL += ((txtCustCodeTo.Text != "") ? FmtStr(txtCustCodeTo.Text) : "NULL") + ", ";

                strSQL += ((txtProductLine.Text != "") ? FmtStr(txtProductLine.Text) : "NULL") + ", ";
                strSQL += ((txtProductLineTo.Text != "") ? FmtStr(txtProductLineTo.Text) : "NULL") + ", ";

                strSQL += ((txtTipeOutletCLU.Text != "") ? FmtStr(txtTipeOutletCLU.Text) : "NULL") + ", ";
                strSQL += ((txtTipeOutletCLUTo.Text != "") ? FmtStr(txtTipeOutletCLUTo.Text) : "NULL") + ", ";

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

                dgvCustList.DataSource = _clsGlobal.ExecDT(strSQL);
                gridViewCustList.BestFitColumns();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        private void checkBoxAllhirarki_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAllhirarki.Checked)
            {
                checkBoxYa.Enabled = false;
                txtProductLine.Enabled = false;
                txtProductLineTo.Enabled = false;
                txtTipeOutletCLU.Enabled = false;
                txtTipeOutletCLUTo.Enabled = false;
            }
            else
            {
                checkBoxYa.Enabled = true;
                txtProductLine.Enabled = true;
                txtProductLineTo.Enabled = true;
                txtTipeOutletCLU.Enabled = true;
                txtTipeOutletCLUTo.Enabled = true;
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
                ddlPajak.SelectedText = "0 ~ Standart";

                //strSQL = " select top 1 br_branch_short,br_branch_desc from GS_BRANCH  ";
                strSQL = "SELECT A.gu_entity,A.gu_branch, B.br_branch_short,B.br_branch_desc FROM VW_USERS_SECURITY A inner join GS_BRANCH B WITH (NOLOCK) on A.gu_entity = B.br_gl_entity_initial  AND A.gu_branch = B.br_branch_id AND A.gu_user_id = '" + clsLogin.USERID + "'";
                DataTable dt = new DataTable();
                dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    lblHeader.Text = "         " + this.Text + " - " + dt.Rows[0]["br_branch_desc"].ToString().Trim() + " - " + dt.Rows[0]["br_branch_short"].ToString().Trim() + "";
                    txtEntityId.Text = dt.Rows[0]["gu_entity"].ToString();
                    txtBranchId.Text = dt.Rows[0]["gu_branch"].ToString();
                }

                __statusOutlet.Execute();
                //__statusOutlet.InsertItems(0, new StatusItem("R", "R-REGISTER"));
                //__statusOutlet.Refresh();
                __statusOutlet.CheckAll();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void AccessButton()
        {
            base.AccessButton();
            barBtnNew.Enabled = clsLogin.BTNNEW;
            barBtnEdit.Enabled = clsLogin.BTNEDIT;
            barBtnDelete.Enabled = clsLogin.BTNDELETE;
            barBtnPrint.Enabled = clsLogin.BTNPRINT;
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
                if (txtEntityId.Text.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception("Entity is empty.");

                if (txtBranchId.Text.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception("Branch is empty.");

                if (sender != null)
                {
                    if (txtCustCode.Text.ToString() != "" && txtCustCodeTo.Text == "")
                        txtCustCodeTo.Text = txtCustCode.Text;

                    if (txtSalesman.Text.ToString() != "" && txtSalesmanTo.Text == "")
                        txtSalesmanTo.Text = txtSalesman.Text;
                }

                __msg = MessageBoxIcon.Error;
                FillGrid();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, __msg);
            }
        }
        // Header bertingkat WinForms diganti dengan caption kolom DevExpress di GridView.

        private string FmtStr(string value_str)
        {
            string formatStr;
            formatStr = "'" + value_str.Trim().Replace("'", "''") + "'";

            return formatStr;
        }

        private void tsb_print_Click(object sender, EventArgs e)
        {
            eARCustPrint frm = new eARCustPrint();

            frm.PrmEntity = txtEntityId.Text;
            frm.PrmEntityTo = txtEntityIdTo.Text;

            frm.PrmBranch = txtBranchId.Text;
            frm.PrmBranchTo = txtBranchIdTo.Text;

            frm.PrmDistrik = txtDistrik.Text;
            frm.PrmDistrikTo = txtDistrikTo.Text;

            frm.PrmBeat = txtBeat.Text;
            frm.PrmBeatTo = txtBeatTo.Text;

            frm.PrmSubBeat = txtSubBeat.Text;
            frm.PrmSubBeatTo = txtSubBeatTo.Text;

            frm.PrmTypeOutlet = txtTypeOutlet.Text;
            frm.PrmTypeOutletTo = txtTypeOutletTo.Text;

            frm.PrmLokasi = txtLokasi.Text;
            frm.PrmLokasiTo = txtLokasiTo.Text;

            frm.PrmKodePasar = txtKodePasar.Text;
            frm.PrmKodePasarTo = txtKodePasarTo.Text;

            frm.PrmKlasifikasi = txtKlasifikasi.Text;
            frm.PrmKlasifikasiTo = txtKlasifikasiTo.Text;

            frm.PrmKategori = txtKategori.Text;
            frm.PrmKategoriTo = txtKategoriTo.Text;

            frm.PrmGroupOutlet = txtGroup.Text;
            frm.PrmGroupOutletTo = txtGroupTo.Text;

            frm.PrmSalesman = txtSalesman.Text;
            frm.PrmSalesmanTo = txtSalesmanTo.Text;

            frm.PrmProvince = txtPropinsi.Text;
            frm.PrmProvinceTo = txtPropinsiTo.Text;

            frm.PrmCity = txtKabupaten.Text;
            frm.PrmCityTo = txtKabupatenTo.Text;

            frm.Prmkecamatan = txtKecamatan.Text;
            frm.PrmkecamatanTo = txtKecamatanTo.Text;

            frm.PrmKelurahan = txtKelurahan.Text;
            frm.PrmKelurahanTo = txtKelurahanTo.Text;

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

            frm.PrmPrdLine = txtProductLine.Text;
            frm.PrmPrdLineTo = txtProductLineTo.Text;

            frm.PrmTypeOutletLine = txtTipeOutletCLU.Text;
            frm.PrmTypeOutletLineTo = txtTipeOutletCLUTo.Text;

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

        private void eARCustMasterList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnExec_Click(null, null);
            }
        }

        private void gridViewCustList_DoubleClick(object sender, EventArgs e)
        {
            if (clsLogin.BTNEDIT == true || clsLogin.BTNPRINT == true)
            {
                tsb_edit_Click(sender, e);
            }
        }

        private void gridViewCustList_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            // Jika diperlukan, styling cell lama dari _clsGlobal.GridColorCell bisa dipetakan di sini.
        }

        private void txtCustName_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = dgvCustList.DataSource as DataTable;
                if (dt == null) return;
                dt.FilterDTable(string.Format("cm_cust_name like '{0}%'", txtCustName.Text));
            }
            catch (Exception) { }
        }

        private void txtDestSite_Leave(object sender, EventArgs e)
        {

        }

        private void txtDestSite_QueryPopUp(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                txtDestSite.Properties.DataSource = _clsGlobal.ExecDT(" select Distinct itzr_destination_site [Dest Zone],itzr_desc_destination_site [Desc] from IM_TRANS_ZONE_ROUTE WITH (NOLOCK) where itzr_entity = '" + txtEntityId.Text + "' and itzr_branch = '" + txtBranchId.Text + "' ");
                txtDestSite.Properties.DisplayMember = "Dest Zone";
                txtDestSite.Properties.ValueMember = "Dest Zone";
                txtDestSite.Properties.NullText = string.Empty;
                txtDestSite.Properties.PopupView.PopulateColumns();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtDestSite_EditValueChanged(object sender, EventArgs e)
        {
            if (txtDestSite.EditValue != null)
            {
                txtDestSite.Text = txtDestSite.EditValue.ToString();
            }
        }
    }


    /// <summary>
    /// Adapter DevExpress untuk menggantikan CheckListProvider WinForms.
    /// Dipakai supaya statusCLB tetap menggunakan DevExpress.XtraEditors.CheckedListBoxControl,
    /// tanpa mengubah business logic pemanggil (Execute, CheckAll, GetCheckedAsString).
    /// </summary>
    public class DxCheckListProvider<T> where T : class, new()
    {
        private readonly DevExpress.XtraEditors.CheckedListBoxControl _control;
        private readonly string _query;
        private readonly clsGlobal _clsGlobal = new clsGlobal();
        private string _displayMember;
        private string _valueMember;

        public DxCheckListProvider(DevExpress.XtraEditors.CheckedListBoxControl control, string query)
        {
            _control = control;
            _query = query;
        }

        public string DisplayMember
        {
            get { return _displayMember; }
            set
            {
                _displayMember = value;
                _control.DisplayMember = value;
            }
        }

        public string ValueMember
        {
            get { return _valueMember; }
            set
            {
                _valueMember = value;
                _control.ValueMember = value;
            }
        }

        public void Execute()
        {
            DataTable dt = _clsGlobal.ExecDT(_query);
            List<T> items = new List<T>();

            foreach (DataRow row in dt.Rows)
            {
                T item = new T();
                SetPropertyValue(item, "Code", row.Table.Columns.Contains("code") ? row["code"] : null);
                SetPropertyValue(item, "Value", row.Table.Columns.Contains("value") ? row["value"] : null);
                items.Add(item);
            }

            _control.DataSource = items;
            if (!string.IsNullOrEmpty(_displayMember)) _control.DisplayMember = _displayMember;
            if (!string.IsNullOrEmpty(_valueMember)) _control.ValueMember = _valueMember;
        }

        public void CheckAll()
        {
            for (int i = 0; i < _control.ItemCount; i++)
            {
                _control.SetItemChecked(i, true);
            }
        }

        public string GetCheckedAsString(string separator, Func<T, string> selector)
        {
            List<string> values = new List<string>();
            for (int i = 0; i < _control.ItemCount; i++)
            {
                if (!_control.GetItemChecked(i)) continue;
                T item = _control.GetItem(i) as T;
                if (item == null) continue;

                string value = selector(item);
                if (!string.IsNullOrEmpty(value)) values.Add(value);
            }

            return string.Join(separator, values.ToArray());
        }

        private static void SetPropertyValue(T item, string propertyName, object value)
        {
            if (value == null || value == DBNull.Value) return;

            System.Reflection.PropertyInfo property = typeof(T).GetProperty(propertyName);
            if (property == null || !property.CanWrite) return;

            property.SetValue(item, Convert.ToString(value).Trim(), null);
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
    }

}
