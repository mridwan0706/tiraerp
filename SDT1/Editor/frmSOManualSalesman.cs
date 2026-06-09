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
    public partial class frmSOManualSalesman : Form
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
            BindDetailEntity();
            BindDetailBranch();
            this.ActiveControl = txtSalesID;
            this.cbEntityID.SelectedIndexChanged += new System.EventHandler(this.cbEntityID_SelectedIndexChanged);
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
            _strsql = "select sgm_spgm_id as [Sales ID], sgm_spgm_name as [Sales Name], sgm_type_operasi as [Type Operasi] from SO_SPG_GIRL_MAN WITH (NOLOCK) where sgm_spgm_status = '2' AND sgm_entity_id = '" + cbEntityID.SelectedValue + "' AND sgm_branch_id = '" + cbBranchID.SelectedValue + "' ";
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
                txtSalesID.Text = frm.ArrField[0].Trim();
                txtSalesDesc.Text = frm.ArrField[1].Trim();
                txtTipeSales.Enabled = false;
                dtTglTrans.Enabled = false;
                txtTeam.Enabled = false;
                txtGudang.Enabled = false;
            }

            DataTable dtFill = new DataTable();


            strSQL = " select isnull(gh_function_desc, '') gh_function_desc, isnull(wh_loc_id1, '') wh_loc_id1, isnull(wh_loc_id2, '') wh_loc_id2, wh_loc_last_work_date," +
                " isnull(sgm_type_operasi, '') sgm_type_operasi, isnull(sgm_spgm_group, '') sgm_spgm_group, isnull(wh_loc_name, '') wh_loc_name, msow_so_type, isnull(sgm_employee, '') sgm_employee " +
                " From dbo.SO_SPG_GIRL_MAN WITH (NOLOCK) inner join (SELECT * from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) where msow_default = '1' and msow_entity_id = '" + cbEntityID.SelectedValue + "' and msow_branch_id = '" + cbBranchID.SelectedValue + "') MAP on msow_sld_id = sgm_spgm_id and msow_entity_id = sgm_entity_id and msow_branch_id = sgm_branch_id inner join SO_ORDER_TYPE WITH (NOLOCK) " +
                " on ot_order_type = msow_so_type and ot_transaction_type = 'D' Left Join (select * from GS_GEN_HARDCODED WITH (NOLOCK) where gh_sys = 'H' and gh_function_name = 'TYPEOPERASISLS' ) GS" +
                " on sgm_type_operasi = gh_function_code Left Join IM_WH_LOC WITH (NOLOCK) on msow_wh_loc1 = wh_loc_id1  and msow_wh_loc2 = wh_loc_id2 and msow_entity_id = wh_loc_entity and msow_branch_id = wh_branch_id inner join IM_WH_LOC_SECURITY WITH (NOLOCK) on ws_entity_id=msow_entity_id and ws_branch_id=msow_branch_id and ws_loc_id1=msow_wh_loc1 and ws_loc_id2=msow_wh_loc2 AND ws_user_id = '" + clsLogin.USERID + "' where sgm_spgm_id ='" + txtSalesID.Text + "' and sgm_entity_id = '" + cbEntityID.SelectedValue + "' and sgm_branch_id = '" + cbBranchID.SelectedValue + "'  ";

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
            DataTable dtFill = new DataTable();

            strSQL = " select sgm_spgm_name, isnull(gh_function_desc, '') gh_function_desc, isnull(wh_loc_id1, '') wh_loc_id1, isnull(wh_loc_id2, '') wh_loc_id2, wh_loc_last_work_date," +
                " isnull(sgm_type_operasi, '') sgm_type_operasi, isnull(sgm_spgm_group, '') sgm_spgm_group, isnull(wh_loc_name, '') wh_loc_name, msow_so_type, isnull(sgm_employee, '') sgm_employee " +
                " From dbo.SO_SPG_GIRL_MAN WITH (NOLOCK) inner join (SELECT * from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) where msow_default = '1' and msow_entity_id = '" + cbEntityID.SelectedValue + "' and msow_branch_id = '" + cbBranchID.SelectedValue + "') MAP on msow_sld_id = sgm_spgm_id inner join SO_ORDER_TYPE WITH (NOLOCK) " +
                " on ot_order_type = msow_so_type and ot_transaction_type = 'D' Left Join (select * from GS_GEN_HARDCODED WITH (NOLOCK) where gh_sys = 'H' and gh_function_name = 'TYPEOPERASISLS' ) GS" +
                " on sgm_type_operasi = gh_function_code Left Join IM_WH_LOC WITH (NOLOCK) on msow_wh_loc1 = wh_loc_id1  and msow_wh_loc2 = wh_loc_id2 and msow_entity_id = wh_loc_entity and msow_branch_id = wh_branch_id inner join IM_WH_LOC_SECURITY WITH (NOLOCK) on ws_entity_id=msow_entity_id and ws_branch_id=msow_branch_id and ws_loc_id1=msow_wh_loc1 and ws_loc_id2=msow_wh_loc2 AND ws_user_id = '" + clsLogin.USERID + "'  where sgm_spgm_id ='" + txtSalesID.Text + "' ";

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
                txtSalesID.Text = txtSalesID.Text.ToUpper();
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
                btnPopUpSalesCode_Click(sender, e);
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (txtSalesID.Text != "")
            {
                if (txtGudang.Text != "" && lblWHLoc1.Text != "" && lblWHLoc2.Text != "")
                {
                    if (lblOrderType.Text != "")
                    {
                        //frmSOManualSalesman.ActiveForm.Hide();

                        frmSOManualEntryEditor frm = new frmSOManualEntryEditor();
                        frm.xEntityID = cbEntityID.SelectedValue.ToString();
                        frm.xBranchID = cbBranchID.SelectedValue.ToString();
                        frm.xSalesID = txtSalesID.Text;
                        frm.xEmployee = lblEmployee.Text;
                        frm.xOrderType = lblOrderType.Text;
                        frm.xSalesDesc = txtSalesDesc.Text;
                        frm.xTglOrder = dtTglTrans.Value;
                        frm.xTglTransaksi = dtTglTrans.Value;
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
        }

        #endregion

        #region Function

        private void BindDetailEntity()
        {
            strSQL = "";
            //strSQL = "SELECT '0' as id, 'ALL' as descc ";
            //strSQL = strSQL + " union ALL ";
            //strSQL = strSQL + "select distinct e.ge_entity_id  as id,  e.ge_entity_id + ' - ' + e.ge_entity as descc from GS_ENTITY e INNER JOIN GS_USERS_SECURITY u on u.gu_entity = e.ge_entity_id WHERE gu_user_id = '" + clsLogin.USERID + "' order by id ";
            strSQL = strSQL + "select distinct gu_entity  as id,  gu_entity + ' - ' + ge_entity as descc from VW_USERS_SECURITY where gu_user_id = '" + clsLogin.USERID + "' order by gu_entity ";

            cbEntityID.DataSource = _clsGlobal.ExecDT(strSQL);
            cbEntityID.ValueMember = "id";
            cbEntityID.DisplayMember = "descc";
            //cbSOType.SelectedIndex = 0;

        }

        private void BindDetailBranch()
        {
            strSQL = "";
            //strSQL = "SELECT '0' as id, 'ALL' as descc ";
            //strSQL = strSQL + " union ALL ";
           // strSQL = strSQL + "Select distinct br_branch_id as id, br_branch_id + ' - ' + br_branch_desc as descc from GS_BRANCH inner join GS_ENTITY on br_gl_entity_initial = ge_entity_id INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id where gu_user_id = '" + clsLogin.USERID + "' and ge_entity_id = '" + cbEntityID.SelectedValue + "' ";
            strSQL = strSQL + "Select distinct gu_branch as id, gu_branch + ' - ' + br_branch_desc as descc from VW_USERS_SECURITY where gu_user_id = '" + clsLogin.USERID + "' and gu_entity = '" + cbEntityID.SelectedValue + "' ";

            cbBranchID.DataSource = _clsGlobal.ExecDT(strSQL);
            cbBranchID.ValueMember = "id";
            cbBranchID.DisplayMember = "descc";
            //cbSOType.SelectedIndex = 0;
        }


        #endregion


        private void cbEntityID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void cbBranchID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }




    }
}
