using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    public partial class frmSOManualSalesman : XtraForm
    {
        private clsGlobal _clsGlobal = new clsGlobal();
        private string strSQL;

        private string _xSalesID;

        public string xSalesID
        {
            get { return _xSalesID; }
            set { _xSalesID = value; }
        }

        private string _xEdit;

        public string xEdit
        {
            get { return _xEdit; }
            set { _xEdit = value; }
        }

        private string _xFlagDCS;
        public string xFlagDCS
        {
            get { return _xFlagDCS; }
            set { _xFlagDCS = value; }
        }


        public frmSOManualSalesman()
        {
            InitializeComponent();
        }

        #region WinForm

        private void frmSOManualSalesman_Load(object sender, EventArgs e)
        {
            dtTglTrans.DateTime = DateTime.Today;
            BindDetailEntity();
            BindDetailBranch();
            BindSalesmanLookup();
            this.ActiveControl = txtSalesID;
            this.cbEntityID.EditValueChanged += new System.EventHandler(this.cbEntityID_SelectedIndexChanged);
            //this.WindowState = FormWindowState.Maximized;
            this.CenterToParent();
        }

        private void rdAktif_CheckedChanged(object sender, EventArgs e)
        {
            rdAllAktif.Checked = !rdAktif.Checked;
        }

        private void rdAllAktif_CheckedChanged(object sender, EventArgs e)
        {
            rdAktif.Checked = !rdAllAktif.Checked;
        }

        private void btnPopUpSalesCode_Click(object sender, EventArgs e)
        {
            string _strsql = "";
            _strsql = "select sgm_spgm_id as [Sales ID], sgm_spgm_name as [Sales Name], sgm_type_operasi as [Type Operasi] from SO_SPG_GIRL_MAN WITH (NOLOCK) where sgm_spgm_status = '2' AND sgm_entity_id = '" + cbEntityID.EditValue + "' AND sgm_branch_id = '" + cbBranchID.EditValue + "' ";
            if (rdAktif.Checked == true)
            {
                _strsql = _strsql + "  and sgm_active_flag = 'A' ";
            }
            _strsql = _strsql + " order by sgm_spgm_id";

            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Salesman";
            frm.Query = _strsql;
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtSalesID.EditValue = frm.ArrField[0].Trim();
                txtSalesDesc.Text = frm.ArrField[1].Trim();
                SetSalesmanDetailReadOnly(true);
            }

            DataTable dtFill = new DataTable();


            strSQL = " select isnull(gh_function_desc, '') gh_function_desc, isnull(wh_loc_id1, '') wh_loc_id1, isnull(wh_loc_id2, '') wh_loc_id2, wh_loc_last_work_date," +
                " isnull(sgm_type_operasi, '') sgm_type_operasi, isnull(sgm_spgm_group, '') sgm_spgm_group, isnull(wh_loc_name, '') wh_loc_name, msow_so_type, isnull(sgm_employee, '') sgm_employee " +
                " From dbo.SO_SPG_GIRL_MAN WITH (NOLOCK) inner join (SELECT * from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) where msow_default = '1' and msow_entity_id = '" + cbEntityID.EditValue + "' and msow_branch_id = '" + cbBranchID.EditValue + "') MAP on msow_sld_id = sgm_spgm_id and msow_entity_id = sgm_entity_id and msow_branch_id = sgm_branch_id inner join SO_ORDER_TYPE WITH (NOLOCK) " +
                " on ot_order_type = msow_so_type and ot_transaction_type = 'D' Left Join (select * from GS_GEN_HARDCODED WITH (NOLOCK) where gh_sys = 'H' and gh_function_name = 'TYPEOPERASISLS' ) GS" +
                " on sgm_type_operasi = gh_function_code Left Join IM_WH_LOC WITH (NOLOCK) on msow_wh_loc1 = wh_loc_id1  and msow_wh_loc2 = wh_loc_id2 and msow_entity_id = wh_loc_entity and msow_branch_id = wh_branch_id inner join IM_WH_LOC_SECURITY WITH (NOLOCK) on ws_entity_id=msow_entity_id and ws_branch_id=msow_branch_id and ws_loc_id1=msow_wh_loc1 and ws_loc_id2=msow_wh_loc2 AND ws_user_id = '" + clsLogin.USERID + "' where sgm_spgm_id ='" + GetSalesIDText() + "' and sgm_entity_id = '" + cbEntityID.EditValue + "' and sgm_branch_id = '" + cbBranchID.EditValue + "'  ";

            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                lblKdTipe.Text = dtFill.Rows[0]["sgm_type_operasi"].ToString().Trim();
                lblOrderType.Text = dtFill.Rows[0]["msow_so_type"].ToString().Trim();
                lblEmployee.Text = dtFill.Rows[0]["sgm_employee"].ToString().Trim();
                lblWHLoc1.Text = dtFill.Rows[0]["wh_loc_id1"].ToString().Trim();
                lblWHLoc2.Text = dtFill.Rows[0]["wh_loc_id2"].ToString().Trim();
                txtTipeSales.Text = dtFill.Rows[0]["gh_function_desc"].ToString().Trim();
                dtTglTrans.Text = dtFill.Rows[0]["wh_loc_last_work_date"].ToString() == string.Empty ? "" : Convert.ToDateTime(dtFill.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                txtTeam.Text = dtFill.Rows[0]["sgm_spgm_group"].ToString().Trim();
                txtGudang.Text = dtFill.Rows[0]["wh_loc_name"].ToString().Trim();
            }
            else
            {
                lblKdTipe.Text = "";
                lblOrderType.Text = "";
                lblEmployee.Text = "";
                lblWHLoc1.Text = "";
                lblWHLoc2.Text = "";
                txtTipeSales.Text = "";
                txtTeam.Text = "";
                txtGudang.Text = "";
            }
        }

        private void txtSalesID_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(GetSalesIDText()))
            {
                ClearSalesmanDetail();
                return;
            }

            DataTable dtFill = new DataTable();

            strSQL = " select sgm_spgm_name, isnull(gh_function_desc, '') gh_function_desc, isnull(wh_loc_id1, '') wh_loc_id1, isnull(wh_loc_id2, '') wh_loc_id2, wh_loc_last_work_date," +
                " isnull(sgm_type_operasi, '') sgm_type_operasi, isnull(sgm_spgm_group, '') sgm_spgm_group, isnull(wh_loc_name, '') wh_loc_name, msow_so_type, isnull(sgm_employee, '') sgm_employee " +
                " From dbo.SO_SPG_GIRL_MAN WITH (NOLOCK) inner join (SELECT * from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) where msow_default = '1' and msow_entity_id = '" + cbEntityID.EditValue + "' and msow_branch_id = '" + cbBranchID.EditValue + "') MAP on msow_sld_id = sgm_spgm_id inner join SO_ORDER_TYPE WITH (NOLOCK) " +
                " on ot_order_type = msow_so_type and ot_transaction_type = 'D' Left Join (select * from GS_GEN_HARDCODED WITH (NOLOCK) where gh_sys = 'H' and gh_function_name = 'TYPEOPERASISLS' ) GS" +
                " on sgm_type_operasi = gh_function_code Left Join IM_WH_LOC WITH (NOLOCK) on msow_wh_loc1 = wh_loc_id1  and msow_wh_loc2 = wh_loc_id2 and msow_entity_id = wh_loc_entity and msow_branch_id = wh_branch_id inner join IM_WH_LOC_SECURITY WITH (NOLOCK) on ws_entity_id=msow_entity_id and ws_branch_id=msow_branch_id and ws_loc_id1=msow_wh_loc1 and ws_loc_id2=msow_wh_loc2 AND ws_user_id = '" + clsLogin.USERID + "'  where sgm_spgm_id ='" + GetSalesIDText() + "' ";

            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                txtSalesDesc.Text = dtFill.Rows[0]["sgm_spgm_name"].ToString().Trim();
                lblKdTipe.Text = dtFill.Rows[0]["sgm_type_operasi"].ToString().Trim();
                lblOrderType.Text = dtFill.Rows[0]["msow_so_type"].ToString().Trim();
                lblEmployee.Text = dtFill.Rows[0]["sgm_employee"].ToString().Trim();
                lblWHLoc1.Text = dtFill.Rows[0]["wh_loc_id1"].ToString().Trim();
                lblWHLoc2.Text = dtFill.Rows[0]["wh_loc_id2"].ToString().Trim();
                txtTipeSales.Text = dtFill.Rows[0]["gh_function_desc"].ToString().Trim();
                if (dtFill.Rows[0]["wh_loc_last_work_date"].ToString() != "")
                {
                    dtTglTrans.Text = Convert.ToDateTime(dtFill.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                }

                txtTeam.Text = dtFill.Rows[0]["sgm_spgm_group"].ToString().Trim();
                txtGudang.Text = dtFill.Rows[0]["wh_loc_name"].ToString().Trim();
                txtSalesID.EditValue = GetSalesIDText().ToUpper();
                SetSalesmanDetailReadOnly(true);
                btnOk.Enabled = true;
            }
            else
            {
                txtSalesDesc.Text = "";
                lblKdTipe.Text = "";
                lblOrderType.Text = "";
                lblEmployee.Text = "";
                lblWHLoc1.Text = "";
                lblWHLoc2.Text = "";
                txtTipeSales.Text = "";
                txtTeam.Text = "";
                txtGudang.Text = "";
                //btnOk.Enabled = false;
            }
        }

        private void txtSalesID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4)
            {
                txtSalesID.ShowPopup();
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (GetSalesIDText() != "")
            {
                if (txtGudang.Text != "" && lblWHLoc1.Text != "" && lblWHLoc2.Text != "")
                {
                    if (lblOrderType.Text != "")
                    {
                        //frmSOManualSalesman.ActiveForm.Hide();

                        frmSOManualEntryEditor frm = new frmSOManualEntryEditor();
                        frm.xEntityID = cbEntityID.EditValue.ToString();
                        frm.xBranchID = cbBranchID.EditValue.ToString();
                        frm.xSalesID = GetSalesIDText();
                        frm.xEmployee = lblEmployee.Text;
                        frm.xOrderType = lblOrderType.Text;
                        frm.xSalesDesc = txtSalesDesc.Text;
                        frm.xTglOrder = dtTglTrans.DateTime;
                        frm.xTglTransaksi = dtTglTrans.DateTime;
                        frm.xTipeSales = txtTipeSales.Text;
                        frm.xWHLoc1 = lblWHLoc1.Text;
                        frm.xWHLoc2 = lblWHLoc2.Text;
                        frm.xFlagDC = xFlagDCS;
                        //this.Close();
                        //if (!frm.Visible)
                        //{
                        //    frm.Show();
                        //}
                        //else
                        //{
                        //    frm.BringToFront();
                        //}

                        //iForm = 0;
                        //for (int i = 2; i < Application.OpenForms.Count; i++)
                        //{

                        foreach (Form f in Application.OpenForms)
                        {
                            if (f is frmSOManualEntryEditor)
                            {
                                //this.Hide();
                                //this.Close();
                                f.Hide();
                                f.Close();
                                //frm.ShowDialog();
                                //return;
                            }
                            else if (f is frmSOManualSalesman)
                            {
                                //frm.ShowDialog();
                                //this.Hide();
                                //this.Close();
                                if (f != this)
                                {
                                    f.Hide();
                                    f.Close();
                                }
                                //else
                                //{
                                //    f.Hide();
                                //}
                                //return;

                                //frm.ShowDialog();
                                //return;
                            }


                        }
                        frm.ShowDialog();
                        this.Hide();
                        this.Close();
                        //frmSOManualEntryList frmMain = new frmSOManualEntryList();
                        //frmMain.Activate();

                        //}

                        //if (iForm > 1)
                        //{
                        //    MessageBox.Show("Close Form yang lain  ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //    //hasil = false;
                        //    return;
                        //}
                        //frm.WindowState = FormWindowState.Maximized;

                        //if(xEdit == null)
                        //{

                        //}


                    }
                    else
                    {
                        MessageBox.Show("Tipe Order" + " " + _clsGlobal.ApplMessage(40003), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show("Gudang" + " " + _clsGlobal.ApplMessage(40003), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("Sales Code" + " " + _clsGlobal.ApplMessage(40003), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                btnOk.Enabled = false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void cbEntityID_SelectedIndexChanged(object sender, EventArgs e)
        {
            BindDetailBranch();
            BindSalesmanLookup();
        }

        #endregion

        #region Function

        private string GetSalesIDText()
        {
            return Convert.ToString(txtSalesID.EditValue ?? txtSalesID.Text).Trim();
        }

        private void ConfigureSearchLookUpViews()
        {
            cbBranchIDView.Columns.Clear();
            cbBranchIDView.Columns.AddVisible("id", "Branch");
            cbBranchIDView.Columns.AddVisible("descc", "Description");
            cbBranchIDView.BestFitColumns();

            txtSalesIDView.Columns.Clear();
            txtSalesIDView.Columns.AddVisible("Sales ID", "Sales ID");
            txtSalesIDView.Columns.AddVisible("Sales Name", "Sales Name");
            txtSalesIDView.Columns.AddVisible("Type Operasi", "Type Operasi");
            txtSalesIDView.BestFitColumns();
        }

        private void BindSalesmanLookup()
        {
            if (cbEntityID.EditValue == null || cbBranchID.EditValue == null)
            {
                txtSalesID.Properties.DataSource = null;
                return;
            }

            string sql = "select sgm_spgm_id as [Sales ID], sgm_spgm_name as [Sales Name], sgm_type_operasi as [Type Operasi] " +
                         "from SO_SPG_GIRL_MAN WITH (NOLOCK) " +
                         "where sgm_spgm_status = '2' " +
                         "AND sgm_entity_id = '" + cbEntityID.EditValue + "' " +
                         "AND sgm_branch_id = '" + cbBranchID.EditValue + "' ";

            if (rdAktif.Checked == true)
            {
                sql += " and sgm_active_flag = 'A' ";
            }

            sql += " order by sgm_spgm_id";

            txtSalesID.Properties.DataSource = _clsGlobal.ExecDT(sql);
            txtSalesID.Properties.ValueMember = "Sales ID";
            txtSalesID.Properties.DisplayMember = "Sales ID";
            ConfigureSearchLookUpViews();
        }

        private void ClearSalesmanDetail()
        {
            txtSalesDesc.Text = "";
            lblKdTipe.Text = "";
            lblOrderType.Text = "";
            lblEmployee.Text = "";
            lblWHLoc1.Text = "";
            lblWHLoc2.Text = "";
            txtTipeSales.Text = "";
            txtTeam.Text = "";
            txtGudang.Text = "";
            dtTglTrans.DateTime = DateTime.Today;
            SetSalesmanDetailReadOnly(false);
        }

        private void SetSalesmanDetailReadOnly(bool readOnly)
        {
            txtTipeSales.Properties.ReadOnly = readOnly;
            dtTglTrans.Properties.ReadOnly = readOnly;
            txtTeam.Properties.ReadOnly = readOnly;
            txtGudang.Properties.ReadOnly = readOnly;
        }

        private void SetFirstValue(DevExpress.XtraEditors.LookUpEdit lookup)
        {
            SetIndexValue(lookup, 0);
        }

        private void SetFirstValue(DevExpress.XtraEditors.SearchLookUpEdit lookup)
        {
            SetIndexValue(lookup, 0);
        }

        private void SetIndexValue(DevExpress.XtraEditors.LookUpEdit lookup, int rowIndex)
        {
            DataTable dt = lookup.Properties.DataSource as DataTable;
            if (dt != null && dt.Rows.Count > 0)
            {
                int index = dt.Rows.Count > rowIndex ? rowIndex : 0;
                lookup.EditValue = dt.Rows[index][lookup.Properties.ValueMember];
            }
        }

        private void SetIndexValue(DevExpress.XtraEditors.SearchLookUpEdit lookup, int rowIndex)
        {
            DataTable dt = lookup.Properties.DataSource as DataTable;
            if (dt != null && dt.Rows.Count > 0)
            {
                int index = dt.Rows.Count > rowIndex ? rowIndex : 0;
                lookup.EditValue = dt.Rows[index][lookup.Properties.ValueMember];
            }
        }

        private void BindDetailEntity()
        {
            strSQL = "";
            //strSQL = "SELECT '0' as id, 'ALL' as descc ";
            //strSQL = strSQL + " union ALL ";
            //strSQL = strSQL + "select distinct e.ge_entity_id  as id,  e.ge_entity_id + ' - ' + e.ge_entity as descc from GS_ENTITY e INNER JOIN GS_USERS_SECURITY u on u.gu_entity = e.ge_entity_id WHERE gu_user_id = '" + clsLogin.USERID + "' order by id ";
            strSQL = strSQL + "select distinct gu_entity  as id,  gu_entity + ' - ' + ge_entity as descc from VW_USERS_SECURITY where gu_user_id = '" + clsLogin.USERID + "' order by gu_entity ";

            DataTable dtEntity = _clsGlobal.ExecDT(strSQL);
            cbEntityID.Properties.DataSource = dtEntity;
            cbEntityID.Properties.ValueMember = "id";
            cbEntityID.Properties.DisplayMember = "descc";
            SetIndexValue(cbEntityID, 1);
            // default selected index 1

        }

        private void BindDetailBranch()
        {
            strSQL = "";
            //strSQL = "SELECT '0' as id, 'ALL' as descc ";
            //strSQL = strSQL + " union ALL ";
            // strSQL = strSQL + "Select distinct br_branch_id as id, br_branch_id + ' - ' + br_branch_desc as descc from GS_BRANCH inner join GS_ENTITY on br_gl_entity_initial = ge_entity_id INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id where gu_user_id = '" + clsLogin.USERID + "' and ge_entity_id = '" + cbEntityID.EditValue + "' ";
            strSQL = strSQL + "Select distinct gu_branch as id, gu_branch + ' - ' + br_branch_desc as descc from VW_USERS_SECURITY where gu_user_id = '" + clsLogin.USERID + "' and gu_entity = '" + cbEntityID.EditValue + "' ";

            DataTable dtBranch = _clsGlobal.ExecDT(strSQL);
            cbBranchID.Properties.DataSource = dtBranch;
            cbBranchID.Properties.ValueMember = "id";
            cbBranchID.Properties.DisplayMember = "descc";
            ConfigureSearchLookUpViews();
            SetIndexValue(cbBranchID, 0);
            // default selected index 0
        }


        #endregion


        private void cbBranchID_EditValueChanged(object sender, EventArgs e)
        {
            BindSalesmanLookup();
            txtSalesID.EditValue = null;
            ClearSalesmanDetail();
        }

        private void txtSalesID_EditValueChanged(object sender, EventArgs e)
        {
            DataRowView row = txtSalesID.Properties.View.GetFocusedRow() as DataRowView;
            if (row != null)
            {
                txtSalesDesc.Text = Convert.ToString(row["Sales Name"]).Trim();
            }

            if (!string.IsNullOrWhiteSpace(GetSalesIDText()))
            {
                txtSalesID_Leave(sender, EventArgs.Empty);
            }
            else
            {
                ClearSalesmanDetail();
            }
        }

        private void cbEntityID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void cbBranchID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = false;
        }




    }
}
