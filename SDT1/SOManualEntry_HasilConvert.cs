using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.Utils;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;

namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    public partial class frmSOManualEntryList : BaseListForm
    {
        const bool __isAsk27757 = true;

        private clsGlobal _clsGlobal = new clsGlobal();
        private DataTable dtGridDisc = new DataTable();
        private DataTable dtGridTPRU = new DataTable();

        private int lRow;
        private string strSQL;
        private string paramMenuId;
        private string NoOrder;
        private int _xHaveClick;
        private string _xReason;
        private DateTime g_dTglGudang;
        public string g_FlagDC;
        private int lCounter;
        private bool isKAM = false;

        private DataTable dtSD = new DataTable();
        private DataTable dtS = new DataTable();
        private DataTable dtU = new DataTable();
        private DataTable dtDD = new DataTable();

        private bool isCrossSite = false;
        private bool isMultisource = false;

        /** DDE PCP **/
        private bool isExcludeNonDDEPCP = false;
        private string SourceNonDDE = "";
        /** END OF DDE PCP **/

        public int xHaveClick
        {
            get { return _xHaveClick; }
            set { _xHaveClick = value; }
        }

        public string xReason
        {
            get { return _xReason; }
            set { _xReason = value; }
        }

        public frmSOManualEntryList()
        {
            InitializeComponent();
        }

        #region BaseListForm Button Mapping

        // Jika BaseListForm Anda memakai nama method berbeda, cukup sesuaikan nama override ini.
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
            this.Close();
        }

        #endregion

        #region Load

        private void frmSOManualEntryList_Load(object sender, EventArgs e)
        {
            this.Text = clsLogin.MENUID + " - " + clsLogin.MENUNAME + " " + this.Text;
            paramMenuId = clsLogin.MENUID;
            this.WindowState = FormWindowState.Maximized;

            AccessButton();

            isKAM = _clsGlobal.IsNefoForKam;
            isMultisource = _clsGlobal.IsMultiSource;
            isCrossSite = _clsGlobal.IsCrossSite;

            BindDetailSOType();
            BindDetailSOStatus();
            BindDetailDateSelection();
            BindDefaultBranch(txtEntityCode, txtBranchCode);
            BindSource();
            BindSloc();
            FLagDC();

            SetGridHeader();
            SetGridDetail();
            SetGridTPRB();
            SetGridTPRU();
            SetGridDisc();

            this.ActiveControl = txtEntityCode;

            if (isCrossSite)
            {
                txtShippingPlantId.Text = "ALL";
                txtShippingPlant.Text = "ALL";
                txtShippingPlant.Enabled = true;
                btnShipping.Enabled = true;
            }
            else
            {
                txtShippingPlant.Enabled = false;
                btnShipping.Enabled = false;
            }

            isExcludeNonDDEPCP = IsExcludeNoNDDEPCP();
        }

        #endregion

        #region Toolbar Logic Lama

        private void btnExec_Click(object sender, EventArgs e)
        {
            ShowLoadingMarquee("Loading SO Manual Entry...");
            try
            {
                FillGridList();
            }
            finally
            {
                HideLoading();
            }
        }

        private void tsb_new_Click(object sender, EventArgs e)
        {
            clsGlobal.MODE_TRX = clsGlobal.TRX_NEW;
            frmSOManualSalesman frm = new frmSOManualSalesman();
            frm.xFlagDCS = g_FlagDC;
            frm.ShowDialog();
        }

        private void tsb_edit_Click(object sender, EventArgs e)
        {
            DataRow row = GetFocusedHeaderRow();
            if (row == null)
            {
                XtraMessageBox.Show("Mohon Pilih SO yang akan diedit ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string sStatusCode = "";

            DataTable dtFill1 = new DataTable();
            strSQL = " select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_sys = 'H' ";
            strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
            strSQL = strSQL + " and gh_sequence_no = 3";
            dtFill1 = _clsGlobal.ExecDT(strSQL);

            if (dtFill1.Rows.Count > 0)
            {
                sStatusCode = dtFill1.Rows[0]["gh_function_code"].ToString();
            }
            else
            {
                XtraMessageBox.Show("Status Approve tidak ada pada GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DataTable dtFill2 = new DataTable();
            strSQL = " select wsh_seq_no,wsh_qasir_flag,wsh_extract_sap,wsh_runcode_ext  from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
            strSQL = strSQL + " WHERE wsh_seq_no = '" + GetRowValue(row, "wsh_seq_no") + "' ";
            strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            if (dtFill2.Rows.Count <= 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = " select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_sys = 'H' ";
                strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
                strSQL = strSQL + " and gh_sequence_no = 1";
                dtFill3 = _clsGlobal.ExecDT(strSQL);

                if (dtFill3.Rows.Count > 0)
                {
                    sStatusCode = dtFill3.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    XtraMessageBox.Show("Status Open tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                DataTable dtFill4 = new DataTable();
                strSQL = " select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + GetRowValue(row, "wsh_seq_no") + "' ";
                strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
                dtFill4 = _clsGlobal.ExecDT(strSQL);

                if (dtFill4.Rows.Count <= 0)
                {
                    XtraMessageBox.Show("Status SO sudah bukan pada Approve atau Open ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            checkStatus();

            string g_iStatus = GetRowValue(row, "wsh_status_so").Trim();
            if (g_iStatus == "O" || g_iStatus == "A")
            {
                clsGlobal.MODE_TRX = clsGlobal.TRX_EDIT;

                frmSOManualEntryEditor frm = new frmSOManualEntryEditor();
                frm.xEntityID = GetRowValue(row, "wsh_entity_id").Trim();
                frm.xBranchID = GetRowValue(row, "wsh_branch_id").Trim();
                frm.xSONo = GetRowValue(row, "wsh_seq_no").Trim();
                frm.xSalesID = GetRowValue(row, "wsh_spgm_id").Trim();
                frm.xFlagDC = g_FlagDC;
                frm.xTipeSales = GetRowValue(row, "wsh_status_so").Trim();
                frm.ShowDialog();
            }
        }

        private void tsb_delete_Click(object sender, EventArgs e)
        {
            checkStatus();

            try
            {
                DataRow row = GetFocusedHeaderRow();
                if (row == null) return;

                string g_iStatus = GetRowValue(row, "wsh_status_so").Trim();
                string g_entity = GetRowValue(row, "wsh_entity_id").Trim();
                string g_branch = GetRowValue(row, "wsh_branch_id").Trim();

                if (g_iStatus == "O" || g_iStatus == "A")
                {
                    DataTable dtFill5 = new DataTable();
                    strSQL = "select wsh_seq_no,wsh_runcode_ext,wsh_extract_sap from SO_WEB_SALES_HEADER  WITH (NOLOCK) ";
                    strSQL = strSQL + " WHERE wsh_seq_no = '" + NoOrder + "' ";
                    strSQL = strSQL + " AND wsh_status_so = 'O' ";
                    dtFill5 = _clsGlobal.ExecDT(strSQL);

                    if (dtFill5.Rows.Count > 0)
                    {
                        if (g_iStatus == "O" && isMultisource)
                        {
                            string extractsap = dtFill5.Rows[0]["wsh_extract_sap"].ToString();
                            string runcodeext = dtFill5.Rows[0]["wsh_runcode_ext"].ToString();

                            if (extractsap.Equals("Y") && !string.IsNullOrEmpty(runcodeext))
                            {
                                XtraMessageBox.Show("SO " + GetRowValue(row, "wsh_seq_no") + " tidak bisa di Cancel lakukan Proses Order terlebih Dahulu",
                                    clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                        }
                    }

                    if (CekFlagTMS(NoOrder))
                    {
                        XtraMessageBox.Show("No Order ter-interface ke TMS, proses cancel dibatalkan",
                            clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        if (g_iStatus == "O")
                        {
                            frmSOManualReason frm = new frmSOManualReason();
                            frm.ShowDialog();

                            if (frm.xHaveClickR == 1)
                            {
                                DialogResult dr = XtraMessageBox.Show(_clsGlobal.ApplMessage(10001),
                                    clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                                if (dr == DialogResult.OK)
                                {
                                    pCancelSO(frm.xReasonR, g_entity, g_branch);
                                    FillGridList();
                                }
                            }
                        }
                        else
                        {
                            DialogResult dr = XtraMessageBox.Show(_clsGlobal.ApplMessage(10001),
                                clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if (dr == DialogResult.OK)
                            {
                                pCancelSO("", g_entity, g_branch);
                                FillGridList();
                            }
                        }
                    }
                }
                else
                {
                    XtraMessageBox.Show("SO status tidak bisa di Cancel ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsb_print_Click(object sender, EventArgs e)
        {
            // Logic lama kosong.
        }

        #endregion

        #region Popup dan Validasi Control

        private void btnPopUpEntityCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Entity";
            frm.Query = "select distinct ge_entity_id [Entity], ge_entity [Desc]  from GS_ENTITY WITH(NOLOCK) INNER JOIN GS_USERS_SECURITY u WITH (NOLOCK) on u.gu_entity = ge_entity_id where gu_user_id = '" + clsLogin.USERID + "' order by ge_entity_id ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtEntityCode.Text = frm.ArrField[0].Trim();
                txtEntityDesc.Text = frm.ArrField[1].Trim();
            }
        }

        private void txtEntityCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4) btnPopUpEntityCode_Click(sender, e);
        }

        private void btnPopUpBranchCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Branch";
            frm.Query = "select gu_branch [Branch], br_branch_desc [Desc] FROM VW_USERS_SECURITY WITH (NOLOCK)  where gu_entity = '" + txtEntityCode.Text.Trim() + "' and gu_user_id = '" + clsLogin.USERID + "'";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtBranchCode.Text = frm.ArrField[0].Trim();
                txtBranchDesc.Text = frm.ArrField[1].Trim();
                WorkDateAccess();
            }
        }

        private void txtBranchCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4) btnPopUpBranchCode_Click(sender, e);
        }

        private void btnPopUpSalesCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Salesman";
            frm.Query = "select sgm_spgm_id as [Sales ID], sgm_spgm_name as [Sales Name], sgm_type_operasi as [Type Operasi] from SO_SPG_GIRL_MAN WITH (NOLOCK)  where sgm_spgm_status = '2' AND sgm_entity_id = '" + txtEntityCode.Text + "' AND sgm_branch_id = '" + txtBranchCode.Text + "'";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtSalesCode.Text = frm.ArrField[0].Trim();
                txtSalesDesc.Text = frm.ArrField[1].Trim();
            }
        }

        private void txtSalesCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4) btnPopUpSalesCode_Click(sender, e);
        }

        private void btnPopUpOutlet_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Outlet";
            frm.Query = "select cm_cust_code1 as [Cust ID1], cm_cust_code2 as [Cust ID2], cm_cust_name as [Cust Name] from SO_CUST_MASTER WITH (NOLOCK)  WHERE cm_active_flag <> 'D' AND cm_entity = '" + txtEntityCode.Text + "' AND cm_branch = '" + txtBranchCode.Text + "'";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtOutletCode.Text = frm.ArrField[0].Trim();
                txtOutlet2.Text = frm.ArrField[1].Trim();
                txtOutletDesc.Text = frm.ArrField[2].Trim();
            }
        }

        private void txtOutletCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4) btnPopUpOutlet_Click(sender, e);
        }

        private void btnPopUpNoOrder_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search No Order";
            frm.Query = "select wsh_seq_no [No Order], cast(convert(varchar(12), wsh_so_date, 106) as datetime) [Order Date], isnull(wsh_so_ref_no, '') [Ref No],wsh_po_no [Po No]  from SO_WEB_SALES_HEADER WITH (NOLOCK) inner join SO_ORDER_TYPE WITH (NOLOCK) on wsh_so_type = ot_order_type where ot_transaction_type = 'D' AND wsh_entity_id = '" + txtEntityCode.Text + "' AND wsh_branch_id = '" + txtBranchCode.Text + "'";
            frm.ShowDialog();

            if (frm.ArrField != null)
            {
                txtNoPO.Enabled = false;
                dtTglOrder.Enabled = false;
                dtTglOrderTo.Enabled = false;
                btnPopUpNoPO.Enabled = false;
                cbdateselection.ItemIndex = 0;
                cbdateselection.Enabled = false;

                txtNoOrder.Text = frm.ArrField[0].Trim();
                txtNoPO.Text = "";

                DateTime tglO = Convert.ToDateTime(frm.ArrField[1].Trim());
                dtTglOrder.DateTime = tglO;
                dtTglOrderTo.DateTime = tglO;
            }
            else
            {
                txtNoPO.Enabled = true;
                dtTglOrder.Enabled = true;
                dtTglOrderTo.Enabled = true;
                cbdateselection.Enabled = true;
                btnPopUpNoPO.Enabled = true;
                txtNoPO.Text = "";
            }
        }

        private void txtNoOrder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4) btnPopUpNoOrder_Click(sender, e);
        }

        private void btnPopUpNoPO_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search No PO";
            frm.Query = "select wsh_po_no [No PO],wsh_seq_no,wsh_so_date from SO_WEB_SALES_HEADER WITH (NOLOCK)  inner join SO_ORDER_TYPE WITH (NOLOCK)  on wsh_so_type = ot_order_type where wsh_po_no is not null and ot_transaction_type = 'D' AND wsh_entity_id = '" + txtEntityCode.Text + "' AND wsh_branch_id = '" + txtBranchCode.Text + "'";
            frm.ShowDialog();

            if (frm.ArrField != null)
            {
                txtNoOrder.Enabled = false;
                dtTglOrder.Enabled = false;
                dtTglOrderTo.Enabled = false;
                btnPopUpNoOrder.Enabled = false;
                cbdateselection.Enabled = false;
                cbdateselection.ItemIndex = 0;

                txtNoPO.Text = frm.ArrField[0].Trim();
                txtNoOrder.Text = "";

                DateTime tglO = Convert.ToDateTime(frm.ArrField[2].Trim());
                dtTglOrder.DateTime = tglO;
                dtTglOrderTo.DateTime = tglO;
            }
            else
            {
                txtNoOrder.Enabled = true;
                dtTglOrder.Enabled = true;
                dtTglOrderTo.Enabled = true;
                btnPopUpNoOrder.Enabled = false;
                cbdateselection.Enabled = false;
                txtNoOrder.Text = "";
            }
        }

        private void txtNoPO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4) btnPopUpNoPO_Click(sender, e);
        }

        private void txtEntityCode_Leave(object sender, EventArgs e)
        {
            DataTable dtFill = new DataTable();
            strSQL = " select distinct ge_entity_id as Entity, ge_entity as EntityDesc  from GS_ENTITY WITH(NOLOCK) INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id where ge_entity_id ='" + txtEntityCode.Text + "' AND gu_user_id = '" + clsLogin.USERID + "' order by ge_entity_id ";
            dtFill = _clsGlobal.ExecDT(strSQL);

            txtEntityDesc.Text = dtFill.Rows.Count > 0 ? dtFill.Rows[0]["EntityDesc"].ToString().Trim() : "";
        }

        private void txtBranchCode_Leave(object sender, EventArgs e)
        {
            DataTable dtFill = new DataTable();
            strSQL = " Select distinct br_branch_id as branch, br_branch_desc as BranchDesc from GS_BRANCH WITH (NOLOCK)  inner join GS_ENTITY WITH (NOLOCK)  on br_gl_entity_initial = ge_entity_id INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id where br_branch_id= '" + txtBranchCode.Text + "' AND gu_user_id = '" + clsLogin.USERID + "' and ge_entity_id = '" + txtEntityCode.Text.Trim() + "' ";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                txtBranchDesc.Text = dtFill.Rows[0]["BranchDesc"].ToString().Trim();
                txtShippingPlantId.Text = "";
                txtShippingPlant.Text = "";
                WorkDateAccess();
            }
            else
            {
                txtBranchDesc.Text = "";
                txtShippingPlantId.Text = "";
                txtShippingPlant.Text = "";
            }
        }

        private void txtSalesCode_Leave(object sender, EventArgs e)
        {
            DataTable dtFill = new DataTable();
            strSQL = " select sgm_spgm_id as salesID, sgm_spgm_name as salesName, sgm_type_operasi as [Type Operasi] from SO_SPG_GIRL_MAN WITH (NOLOCK) where sgm_spgm_status = '2' AND sgm_spgm_id = '" + txtSalesCode.Text + "' AND sgm_entity_id = '" + txtEntityCode.Text + "' AND sgm_branch_id = '" + txtBranchCode.Text + "'";
            dtFill = _clsGlobal.ExecDT(strSQL);

            txtSalesDesc.Text = dtFill.Rows.Count > 0 ? dtFill.Rows[0]["salesName"].ToString().Trim() : "";
        }

        private void txtOutletCode_Leave(object sender, EventArgs e)
        {
            DataTable dtFill = new DataTable();
            strSQL = " select cm_cust_code1 as custID1, cm_cust_code2 as CustID2, cm_cust_name as CustName from SO_CUST_MASTER WITH (NOLOCK) WHERE cm_active_flag <> 'D' AND cm_cust_code1 ='" + txtOutletCode.Text + "' AND cm_entity = '" + txtEntityCode.Text + "' AND cm_branch = '" + txtBranchCode.Text + "' ";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                txtOutlet2.Text = dtFill.Rows[0]["CustID2"].ToString().Trim();
                txtOutletDesc.Text = dtFill.Rows[0]["CustName"].ToString().Trim();
            }
            else
            {
                txtOutlet2.Text = "";
                txtOutletDesc.Text = "";
            }
        }

        private void txtNoOrder_Leave(object sender, EventArgs e)
        {
            if (txtNoOrder.Text != "")
            {
                DataTable dtFill = new DataTable();
                strSQL = " select wsh_seq_no as NoOrder, cast(convert(varchar(12), wsh_so_date, 106) as datetime) as OrderDate, isnull(wsh_so_ref_no, '') [Ref No]  from SO_WEB_SALES_HEADER WITH (NOLOCK) inner join SO_ORDER_TYPE WITH (NOLOCK) on wsh_so_type = ot_order_type where ot_transaction_type = 'D' AND wsh_seq_no ='" + txtNoOrder.Text + "' AND wsh_entity_id = '" + txtEntityCode.Text + "' AND wsh_branch_id = '" + txtBranchCode.Text + "' ";
                dtFill = _clsGlobal.ExecDT(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    cbdateselection.ItemIndex = 0;
                    cbdateselection.Enabled = false;
                    txtNoPO.Enabled = false;
                    btnPopUpNoPO.Enabled = false;
                    dtTglOrder.Enabled = false;
                    dtTglOrderTo.Enabled = false;

                    txtNoPO.Text = "";

                    DateTime tglO = Convert.ToDateTime(dtFill.Rows[0]["OrderDate"].ToString().Trim());
                    dtTglOrder.DateTime = tglO;
                    dtTglOrderTo.DateTime = tglO;
                }
                else
                {
                    cbdateselection.Enabled = true;
                    txtNoPO.Enabled = true;
                    btnPopUpNoPO.Enabled = true;
                    dtTglOrder.Enabled = true;
                    dtTglOrderTo.Enabled = true;
                    txtNoPO.Text = "";
                    txtNoOrder.Text = "";
                }
            }
            else
            {
                cbdateselection.Enabled = true;
                txtNoPO.Enabled = true;
                btnPopUpNoPO.Enabled = true;
                dtTglOrder.Enabled = true;
                dtTglOrderTo.Enabled = true;
                txtNoPO.Text = "";
            }
        }

        private void txtNoPO_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtNoPO.Text))
            {
                DataTable dtFill = new DataTable();
                strSQL = " select wsh_po_no as NoPO, wsh_seq_no as OrderNo,wsh_so_date OrderDate from SO_WEB_SALES_HEADER WITH (NOLOCK) inner join SO_ORDER_TYPE WITH (NOLOCK) on wsh_so_type = ot_order_type where wsh_po_no is not null and ot_transaction_type = 'D' AND wsh_po_no ='" + txtNoPO.Text + "' AND wsh_entity_id = '" + txtEntityCode.Text + "' AND wsh_branch_id = '" + txtBranchCode.Text + "' ";
                dtFill = _clsGlobal.ExecDT(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    btnPopUpNoOrder.Enabled = false;
                    txtNoOrder.Enabled = false;
                    dtTglOrder.Enabled = false;
                    dtTglOrderTo.Enabled = false;
                    cbdateselection.Enabled = false;
                    cbdateselection.ItemIndex = 0;

                    txtNoPO.Text = dtFill.Rows[0]["NoPO"].ToString().Trim();
                    txtNoOrder.Text = "";

                    DateTime tglO = Convert.ToDateTime(dtFill.Rows[0]["OrderDate"].ToString().Trim());
                    dtTglOrder.DateTime = tglO;
                    dtTglOrderTo.DateTime = tglO;
                }
                else
                {
                    btnPopUpNoOrder.Enabled = true;
                    txtNoOrder.Enabled = true;
                    dtTglOrder.Enabled = true;
                    dtTglOrderTo.Enabled = true;
                    cbdateselection.Enabled = true;

                    txtNoOrder.Text = "";
                    txtNoPO.Text = "";
                }
            }
            else
            {
                btnPopUpNoOrder.Enabled = true;
                txtNoOrder.Enabled = true;
                dtTglOrder.Enabled = true;
                dtTglOrderTo.Enabled = true;
                cbdateselection.Enabled = true;

                txtNoOrder.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            strSQL = "SELECT 'ALL' AS [Shipping Plant Id], 'ALL' AS [Shipping Plant] " +
             "UNION " +
             "select wsh_ship_plant AS [Shipping Plant Id], br_branch_desc AS [Shipping Plant] " +
             "from SO_WEB_SALES_HEADER WITH(NOLOCK) " +
             "INNER JOIN GS_BRANCH WITH(NOLOCK) ON br_branch_ud2 = wsh_ship_plant " +
             "WHERE wsh_branch_id = '" + txtBranchCode.Text + "'  ";

            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Shipping Plant";
            frm.Query = strSQL;
            frm.ShowDialog();

            if (frm.ArrField != null)
            {
                txtShippingPlantId.Text = frm.ArrField[0].Trim();
                txtShippingPlant.Text = frm.ArrField[1].Trim();
            }
        }

        #endregion

        #region Grid Event

        private void gvSalesHeader_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            LoadSelectedHeader();
        }

        private void gvSalesHeader_DoubleClick(object sender, EventArgs e)
        {
            if (clsLogin.BTNEDIT)
                tsb_edit_Click(sender, e);
        }

        private void gvSalesHeader_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle < 0) return;
            e.Appearance.BackColor = (e.RowHandle % 2 == 0) ? Color.Aqua : Color.White;
        }

        private void LoadSelectedHeader()
        {
            DataRow row = GetFocusedHeaderRow();
            if (row == null) return;

            NoOrder = GetRowValue(row, "wsh_seq_no").Trim();
            txtBebanCashDisc.Text = GetRowValue(row, "wsh_beban_cashdisc_desc").Trim();

            Cursor.Current = Cursors.WaitCursor;

            FillGridDetail(NoOrder);
            FillGridTPRB(NoOrder);
            FillGridTPRU(NoOrder);
            FillGridDisc(NoOrder);
            CalcAmount();

            Cursor.Current = Cursors.Default;
        }

        #endregion

        #region Grid Setup

        private void SetDefaultGrid(GridView view, bool editable)
        {
            view.OptionsBehavior.Editable = editable;
            view.OptionsView.ShowGroupPanel = false;
            view.OptionsView.ColumnAutoWidth = false;
            view.OptionsSelection.EnableAppearanceFocusedCell = false;
            view.OptionsView.ShowIndicator = true;
            view.OptionsView.EnableAppearanceEvenRow = true;
            view.OptionsView.EnableAppearanceOddRow = true;
        }

        private void AddColumn(GridView view, string fieldName, string caption, int width, bool visible)
        {
            var col = view.Columns.AddVisible(fieldName, caption);
            col.FieldName = fieldName;
            col.Caption = caption;
            col.Width = width;
            col.Visible = visible;
        }

        private void AddNumericColumn(GridView view, string fieldName, string caption, int width, bool visible, string format)
        {
            AddColumn(view, fieldName, caption, width, visible);
            var col = view.Columns[fieldName];
            col.DisplayFormat.FormatType = FormatType.Numeric;
            col.DisplayFormat.FormatString = format;
            col.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
        }

        private void SetGridHeader()
        {
            SetDefaultGrid(gvSalesHeader, false);
            gvSalesHeader.Columns.Clear();

            AddColumn(gvSalesHeader, "wsh_entity_id", "Entity", 70, true);
            AddColumn(gvSalesHeader, "wsh_branch_id", "Branch", 70, true);
            AddColumn(gvSalesHeader, "wsh_seq_no", "No Order", 100, true);
            AddColumn(gvSalesHeader, "wsh_po_no", "PO No", 100, true);
            AddColumn(gvSalesHeader, "wsh_so_date", "Tanggal", 100, true);
            AddColumn(gvSalesHeader, "wsh_cust_code1", "No Outlet", 90, true);
            AddColumn(gvSalesHeader, "wsh_cust_code2", "No Outlet 2", 90, false);
            AddColumn(gvSalesHeader, "cm_cust_name", "Nama Outlet", 180, true);
            AddColumn(gvSalesHeader, "wsh_spgm_id", "Kd Sales", 80, true);
            AddColumn(gvSalesHeader, "sgm_spgm_name", "Sales", 150, true);
            AddColumn(gvSalesHeader, "wsh_user_id", "User", 90, true);
            AddColumn(gvSalesHeader, "so_status", "Status", 80, true);
            AddColumn(gvSalesHeader, "wsh_status_so", "status code", 90, false);
            AddColumn(gvSalesHeader, "wsh_loc_id1", "WHLoc1", 80, false);
            AddColumn(gvSalesHeader, "wsh_loc_id2", "WHLoc2", 80, false);
            AddColumn(gvSalesHeader, "wsh_flag_so_limit", "Status KL", 90, true);
            AddColumn(gvSalesHeader, "wsh_sch_date", "Delivery Date", 110, true);
            AddColumn(gvSalesHeader, "wsh_expired_date", "Expired Date", 110, true);
            AddColumn(gvSalesHeader, "wsh_beban_cashdisc", "Beban Cash Disc", 120, false);
            AddColumn(gvSalesHeader, "wsh_beban_cashdisc_desc", "Cash Disc Desc", 120, false);

            gvSalesHeader.FocusedRowChanged -= gvSalesHeader_FocusedRowChanged;
            gvSalesHeader.FocusedRowChanged += gvSalesHeader_FocusedRowChanged;
            gvSalesHeader.DoubleClick -= gvSalesHeader_DoubleClick;
            gvSalesHeader.DoubleClick += gvSalesHeader_DoubleClick;
            gvSalesHeader.RowStyle -= gvSalesHeader_RowStyle;
            gvSalesHeader.RowStyle += gvSalesHeader_RowStyle;
        }

        private void SetGridDetail()
        {
            SetDefaultGrid(gvSalesDetail, false);
            gvSalesDetail.Columns.Clear();

            AddColumn(gvSalesDetail, "wsd_prd_master_code", "PCODE", 100, true);
            AddColumn(gvSalesDetail, "wsd_real_order_amt", "wsd_real_order_amt", 100, false);
            AddColumn(gvSalesDetail, "wsd_total_so_amount", "wsd_total_so_amount", 100, false);
            AddColumn(gvSalesDetail, "wsd_grade", "grade", 70, false);
            AddColumn(gvSalesDetail, "wsd_prd_size", "prd size", 70, false);
            AddColumn(gvSalesDetail, "wsd_uom_convert_big", "Conv1", 70, false);
            AddColumn(gvSalesDetail, "wsd_uom_convert_mid", "Conv2", 70, false);
            AddColumn(gvSalesDetail, "prm_prd_desc", "Nama Barang", 220, true);
            AddNumericColumn(gvSalesDetail, "wsd_het_unit_price", "Harga", 120, true, "#,##0.00");
            AddColumn(gvSalesDetail, "wsd_real_order_xqty", "RealOrder Qty", 120, true);
            AddColumn(gvSalesDetail, "wsd_so_xqty", "Order Qty", 100, true);
            AddNumericColumn(gvSalesDetail, "ijumlahharga", "Jumlah Harga", 130, true, "#,##0.00");
            AddNumericColumn(gvSalesDetail, "wsd_tot_disc_pc", "Discount", 120, true, "#,##0.00");
        }

        private void SetGridTPRB()
        {
            SetDefaultGrid(gvPromosiQty, false);
            gvPromosiQty.Columns.Clear();

            AddColumn(gvPromosiQty, "wsd_prd_master_code", "PCODE", 100, true);
            AddColumn(gvPromosiQty, "prm_prd_desc", "Nama Barang", 220, true);
            AddColumn(gvPromosiQty, "wsd_grade", "Grade", 80, false);
            AddColumn(gvPromosiQty, "wsd_prd_size", "Size", 80, false);
            AddColumn(gvPromosiQty, "wsd_uom_convert_big", "Conv1", 80, false);
            AddColumn(gvPromosiQty, "wsd_uom_convert_mid", "Conv2", 80, false);
            AddNumericColumn(gvPromosiQty, "wsd_het_unit_price", "Harga", 120, true, "#,##0.00");
            AddColumn(gvPromosiQty, "wsd_so_xqty", "Qty Promosi", 120, true);
        }

        private void SetGridTPRU()
        {
            if (!dtGridTPRU.Columns.Contains("pt_promo_id")) dtGridTPRU.Columns.Add("pt_promo_id", typeof(string));
            if (!dtGridTPRU.Columns.Contains("pt_product_id")) dtGridTPRU.Columns.Add("pt_product_id", typeof(string));
            if (!dtGridTPRU.Columns.Contains("prm_prd_desc")) dtGridTPRU.Columns.Add("prm_prd_desc", typeof(string));
            if (!dtGridTPRU.Columns.Contains("pt_amount")) dtGridTPRU.Columns.Add("pt_amount", typeof(string));
            if (!dtGridTPRU.Columns.Contains("pt_pct_promo")) dtGridTPRU.Columns.Add("pt_pct_promo", typeof(string));

            gcPromosiRp.DataSource = dtGridTPRU;

            SetDefaultGrid(gvPromosiRp, false);
            gvPromosiRp.Columns.Clear();

            AddColumn(gvPromosiRp, "pt_promo_id", "Promo Code", 100, true);
            AddColumn(gvPromosiRp, "pt_product_id", "PCODE", 100, true);
            AddColumn(gvPromosiRp, "prm_prd_desc", "Nama Barang", 200, true);
            AddNumericColumn(gvPromosiRp, "pt_amount", "Promosi Rp", 120, true, "#,##0.00");
            AddColumn(gvPromosiRp, "pt_pct_promo", "Nilai Promo (%)", 120, true);
        }

        private void SetGridDisc()
        {
            string[] cols =
            {
                "tdt_product_master_id",
                "g_iPctD1","g_iRpD1","g_iPctK1","g_iRpK1",
                "g_iPctD2","g_iRpD2","g_iPctK2","g_iRpK2",
                "g_iPctD3","g_iRpD3","g_iPctK3","g_iRpK3",
                "g_iPctD4","g_iRpD4","g_iPctK4","g_iRpK4",
                "g_iPctD5","g_iRpD5","g_iPctK5","g_iRpK5",
                "g_iPctD6","g_iRpD6","g_iPctK6","g_iRpK6",
                "g_iPctD7","g_iRpD7","g_iPctK7","g_iRpK7",
                "g_iPctD8","g_iRpD8","g_iPctK8","g_iRpK8",
                "g_iPctD9","g_iRpD9","g_iPctK9","g_iRpK9",
                "g_iPctD10","g_iRpD10","g_iPctK10","g_iRpK10"
            };

            foreach (string c in cols)
                if (!dtGridDisc.Columns.Contains(c))
                    dtGridDisc.Columns.Add(c, typeof(string));

            gcDisc.DataSource = dtGridDisc;

            SetDefaultGrid(gvDisc, false);
            gvDisc.Columns.Clear();

            AddColumn(gvDisc, "tdt_product_master_id", "PCODE", 100, true);

            for (int i = 1; i <= 10; i++)
            {
                AddNumericColumn(gvDisc, "g_iPctD" + i, "%D" + i, 70, true, "#,##0.00");
                AddNumericColumn(gvDisc, "g_iRpD" + i, "RpD" + i, 90, true, "#,##0.00");
                AddNumericColumn(gvDisc, "g_iPctK" + i, "%K" + i, 70, true, "#,##0.00");
                AddNumericColumn(gvDisc, "g_iRpK" + i, "RpK" + i, 90, true, "#,##0.00");
            }
        }

        #endregion

        #region Function

        private void CalcAmount()
        {
            decimal promU = 0;
            decimal disc110 = 0;
            decimal cashdisc = 0;
            decimal rpcashdisc = 0;
            decimal ppn = 0;
            decimal ttlinv = 0;
            decimal disc = 0;
            decimal curHarga = 0;

            try
            {
                DataTable detail = gcSalesDetail.DataSource as DataTable;
                if (detail != null)
                {
                    foreach (DataRow gRow in detail.Rows)
                    {
                        if (Convert.ToString(gRow["ijumlahharga"]) != "")
                        {
                            disc = Convert.ToString(gRow["wsd_tot_disc_pc"]) == "" ? 0 : Convert.ToDecimal(gRow["wsd_tot_disc_pc"]);
                            curHarga = curHarga + (Convert.ToDecimal(gRow["ijumlahharga"]) - disc);
                        }
                    }
                }

                txtSubTotal.Text = curHarga.ToString("#,##0.00");

                DataTable dtFill = new DataTable();

                strSQL = "select isnull(wsh_tot_tpr_amt,0) wsh_tot_tpr_amt, wsh_tot_disc_1, wsh_tot_disc_2, " +
                         " wsh_tot_disc_3, wsh_tot_disc_4, wsh_tot_disc_5, wsh_tot_disc_6, wsh_tot_disc_7, wsh_tot_disc_8,  " +
                         "wsh_tot_disc_9, wsh_tot_disc_10, isnull(wsh_pct_cash_disc,0) wsh_pct_cash_disc, isnull(wsh_cash_disc,0) wsh_cash_disc, isnull(wsh_tax_amount,0) wsh_tax_amount, isnull(wsh_net_amount,0) wsh_net_amount " +
                         "From dbo.SO_WEB_SALES_HEADER WITH (NOLOCK) " +
                         "where wsh_seq_no = '" + NoOrder + "' ";

                dtFill = _clsGlobal.ExecDT(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    decimal ttldisc1 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_1"]);
                    decimal ttldisc2 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_2"]);
                    decimal ttldisc3 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_3"]);
                    decimal ttldisc4 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_4"]);
                    decimal ttldisc5 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_5"]);
                    decimal ttldisc6 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_6"]);
                    decimal ttldisc7 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_7"]);
                    decimal ttldisc8 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_8"]);
                    decimal ttldisc9 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_9"]);
                    decimal ttldisc10 = ToDecimal(dtFill.Rows[0]["wsh_tot_disc_10"]);

                    promU = ToDecimal(dtFill.Rows[0]["wsh_tot_tpr_amt"]);
                    disc110 = ttldisc1 + ttldisc2 + ttldisc3 + ttldisc4 + ttldisc5 + ttldisc6 + ttldisc7 + ttldisc8 + ttldisc9 + ttldisc10;
                    cashdisc = ToDecimal(dtFill.Rows[0]["wsh_pct_cash_disc"]);
                    rpcashdisc = ToDecimal(dtFill.Rows[0]["wsh_cash_disc"]);
                    ppn = ToDecimal(dtFill.Rows[0]["wsh_tax_amount"]);
                    ttlinv = ToDecimal(dtFill.Rows[0]["wsh_net_amount"]);

                    txtPromosiU.Text = promU.ToString("#,##0.00");
                    txtCashDiscP.Text = cashdisc.ToString("#,##0.#####");
                    txtCashDiscR.Text = rpcashdisc.ToString("#,##0.00");
                    txtPPN.Text = ppn.ToString("#,##0.00");
                    txtTtlInvoice.Text = ttlinv.ToString("#,##0.00");
                    txtDisc1.Text = disc110.ToString("#,##0.00");
                }
                else
                {
                    txtPromosiU.Text = "";
                    txtCashDiscP.Text = "";
                    txtCashDiscR.Text = "";
                    txtPPN.Text = "";
                    txtTtlInvoice.Text = "";
                    txtDisc1.Text = "";
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AccessButton()
        {
            SetActionButton(false, false, false, false);
        }

        private void WorkDateAccess()
        {
            try
            {
                DataTable dtFill = new DataTable();

                strSQL = "EXEC SP_CURRENT_BOOK_MONTH '" + txtEntityCode.Text + "','" + txtBranchCode.Text + "'";
                dtFill = _clsGlobal.ExecDT(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    if (dtFill.Rows[0]["wh_loc_last_work_date"].ToString() != "")
                    {
                        DateTime workDate = Convert.ToDateTime(dtFill.Rows[0]["wh_loc_last_work_date"].ToString());
                        dtTglOrder.DateTime = workDate;
                        dtTglOrderTo.DateTime = workDate;
                        g_dTglGudang = workDate;

                        SetActionButton(
                            clsLogin.BTNNEW,
                            clsLogin.BTNEDIT,
                            clsLogin.BTNDELETE,
                            clsLogin.BTNPRINT
                        );
                    }
                    else
                    {
                        SetActionButton(false, false, false, false);
                    }
                }
                else
                {
                    g_dTglGudang = DateTime.Now;
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FLagDC()
        {
            try
            {
                DataTable dtFill = new DataTable();

                strSQL = "SELECT gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_function_name = 'FLAGDC'";
                dtFill = _clsGlobal.ExecDT(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    if (dtFill.Rows[0]["gh_function_code"].ToString() != "")
                    {
                        g_FlagDC = dtFill.Rows[0]["gh_function_code"].ToString();
                    }
                    else
                    {
                        XtraMessageBox.Show("Silahkan tentukan value dari parameter FLAGDC terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    XtraMessageBox.Show("Belum ada parameter DC dan proses dibatalkan", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool pCheckValueU(string sValue, string sValue2, int iIndex)
        {
            bool keyFound = true;

            for (lCounter = 0; lCounter < dtGridTPRU.Rows.Count; lCounter++)
            {
                if (Convert.ToString(dtGridTPRU.Rows[lCounter]["pt_product_id"]).Trim() == sValue.Trim()
                    && Convert.ToString(dtGridTPRU.Rows[lCounter]["pt_promo_id"]).Trim() == sValue.Trim())
                {
                    lRow = lCounter;
                    keyFound = false;
                    return keyFound;
                }
            }

            return keyFound;
        }

        private bool pCheckValueD(string sValue, int iIndex, long lRowParam)
        {
            bool keyFound = true;

            for (lCounter = 0; lCounter < dtGridDisc.Rows.Count; lCounter++)
            {
                if (Convert.ToString(dtGridDisc.Rows[lCounter]["tdt_product_master_id"]).Trim() == sValue.Trim())
                {
                    lRow = lCounter;
                    keyFound = false;
                    return keyFound;
                }
            }

            return keyFound;
        }

        private bool CekFlagTMS(string vSoNo)
        {
            bool keyFound = false;
            DataTable dtFill = new DataTable();

            strSQL = "select isnull(wsh_interface_tms, 'N') as flag_tms from SO_WEB_SALES_HEADER WITH (NOLOCK) ";
            strSQL = strSQL + " where wsh_seq_no = '" + vSoNo + "'";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
                keyFound = dtFill.Rows[0]["flag_tms"].ToString() == "Y";

            return keyFound;
        }

        private void checkStatus()
        {
            string sStatusCode = "";
            DataTable dtFill = new DataTable();

            strSQL = " select gh_function_code";
            strSQL = strSQL + " from GS_GEN_HARDCODED  WITH (NOLOCK)";
            strSQL = strSQL + " where gh_sys = 'H' ";
            strSQL = strSQL + " and gh_function_name = 'SOSTATUS' ";
            strSQL = strSQL + " and gh_sequence_no = 3";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                sStatusCode = dtFill.Rows[0]["gh_function_code"].ToString();
            }
            else
            {
                XtraMessageBox.Show("Status Approve tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            DataTable dtFill2 = new DataTable();
            strSQL = "select wsh_seq_no,wsh_runcode_ext,wsh_extract_sap from SO_WEB_SALES_HEADER  WITH (NOLOCK) ";
            strSQL = strSQL + " WHERE wsh_seq_no = '" + NoOrder + "' ";
            strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            if (dtFill2.Rows.Count <= 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = " select gh_function_code";
                strSQL = strSQL + " from GS_GEN_HARDCODED WITH (NOLOCK)  ";
                strSQL = strSQL + " where gh_sys = 'H'";
                strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
                strSQL = strSQL + " and gh_sequence_no = 1";
                dtFill3 = _clsGlobal.ExecDT(strSQL);

                if (dtFill3.Rows.Count > 0)
                {
                    sStatusCode = dtFill3.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    XtraMessageBox.Show("Status Open tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                DataTable dtFill4 = new DataTable();
                strSQL = "select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + NoOrder + "' ";
                strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
                dtFill4 = _clsGlobal.ExecDT(strSQL);

                if (dtFill4.Rows.Count <= 0)
                {
                    XtraMessageBox.Show("Status SO sudah bukan pada Approve atau Open ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void pCancelSO(string _xReasonR, string _xEntity, string _xBranch)
        {
            try
            {
                if (_clsGlobal.IsNefoForKam)
                {
                    strSQL = " UPDATE SDCUST SET sdc_firstbuy_flag = 'N',	sdc_update_date=getdate() FROM SD_SALES_DEAL_TRX WITH (NOLOCK) ";
                    strSQL = strSQL + " INNER JOIN SO_WEB_SALES_HEADER WITH (NOLOCK) ON wsh_entity_id = dt_entity_id AND wsh_branch_id = dt_branch_id AND wsh_seq_no = dt_order_no ";
                    strSQL = strSQL + " INNER JOIN SD_SALES_DEAL_CUST SDCUST WITH (UPDLOCK) ON sdc_entity_id = wsh_entity_id	AND sdc_branch_id = wsh_branch_id AND sdc_outlet = wsh_cust_code1 ";
                    strSQL = strSQL + " AND sdc_outlet2 = wsh_cust_code2 AND sdc_promo_code = dt_promo_id AND sdc_firstbuy_flag = 'Y' ";
                    strSQL = strSQL + " WHERE  wsh_entity_id='" + _xEntity + "' and wsh_branch_id='" + _xBranch + "' AND wsh_seq_no='" + NoOrder + "'";
                    _clsGlobal.ExecDT(strSQL);
                }

                strSQL = " EXEC SP_CANCEL_SO '" + NoOrder + "','" + clsLogin.USERID + "','" + _xEntity + "','" + _xBranch + "'";
                _clsGlobal.ExecDT(strSQL);

                strSQL = " EXEC SP_CHECK_BALANCE_SO_MULTIBRANCH '1','" + NoOrder + "', '" + _xEntity + "','" + _xBranch + "'";
                _clsGlobal.ExecDT(strSQL);

                strSQL = "update SO_WEB_SALES_DETAIL set wsd_reason = '" + _xReasonR + "' where wsd_so_seq_no = '" + NoOrder + "'";
                _clsGlobal.ExecDT(strSQL);

                if (_clsGlobal.IsNefoForKam)
                {
                    strSQL = "EXEC IP_PROSES_RECONDITION_SUBS_CANCEL_KAM '" + NoOrder + "','" + _xEntity + "','" + _xBranch + "','" + clsLogin.USERID + "'";
                    _clsGlobal.ExecDT(strSQL);
                }
                else
                {
                    if (__isAsk27757)
                    {
                        strSQL = "EXEC IP_PROSES_RECONDITION_SUBS_CANCEL_KAM '" + NoOrder + "','" + _xEntity + "','" + _xBranch + "','" + clsLogin.USERID + "'";
                        _clsGlobal.ExecDT(strSQL);
                    }
                }

                XtraMessageBox.Show("No Order : " + NoOrder + " berhasil di cancel", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindSource()
        {
            strSQL = "";
            strSQL = "SELECT 'ALL' as fcode, 'ALL' as fdesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'SOURCE_CLNT'";

            cbSource.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSource.Properties.ValueMember = "fcode";
            cbSource.Properties.DisplayMember = "fdesc";
            cbSource.ItemIndex = 0;
        }

        private void BindSloc()
        {
            strSQL = "";
            strSQL = "SELECT 'ALL' as fcode, 'ALL' as fdesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'MULTISOURCE'";

            cbFlagSloc.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cbFlagSloc.Properties.ValueMember = "fcode";
            cbFlagSloc.Properties.DisplayMember = "fdesc";
            cbFlagSloc.ItemIndex = 0;
        }

        private void BindDetailSOType()
        {
            strSQL = "";
            strSQL = "SELECT '0' as otype, 'ALL' as odesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select ot_order_type as otype, ot_desc as odesc from SO_ORDER_TYPE WITH (NOLOCK)  where ot_transaction_type = 'D' order by otype";

            cbSOType.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSOType.Properties.ValueMember = "otype";
            cbSOType.Properties.DisplayMember = "odesc";
            cbSOType.ItemIndex = 0;
        }

        private void BindDetailSOStatus()
        {
            strSQL = "";
            strSQL = "SELECT '0' as fcode, 'ALL' as fdesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'SOSTATUS'";

            cbSOStatus.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSOStatus.Properties.ValueMember = "fcode";
            cbSOStatus.Properties.DisplayMember = "fdesc";
            cbSOStatus.ItemIndex = 0;
        }

        private void BindDetailDateSelection()
        {
            strSQL = "";
            strSQL = "SELECT 'Order Date' as display, 'wsh_so_date' as value ";
            strSQL += "Union All SELECT 'PO Date' as display, 'wsh_po_date' as value ";
            strSQL += "Union All SELECT 'PO Delv Date' as display, 'wsh_sch_date' as value ";
            strSQL += "Union All SELECT 'PO Exp Date' as display, 'wsh_expired_date' as value ";

            cbdateselection.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cbdateselection.Properties.DisplayMember = "display";
            cbdateselection.Properties.ValueMember = "value";
            cbdateselection.ItemIndex = 0;
        }

        private void FillGridList()
        {
            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {
                DataTable dt = new DataTable();
                con.Open();

                string where2 = string.Empty;
                string dateField = Convert.ToString(cbdateselection.EditValue);

                string where = "where convert(varchar(10), " + dateField + ", 120) >='" + dtTglOrder.DateTime.ToString("yyyy-MM-dd") + "' and convert(varchar(10), " + dateField + ", 120) <='" + dtTglOrderTo.DateTime.ToString("yyyy-MM-dd") + "' and ot_transaction_type = 'D' and ws_user_id = '" + clsLogin.USERID + "' ";

                if (txtEntityCode.Text != "")
                    where = where + " and wsh_entity_id = '" + txtEntityCode.Text + "' ";
                if (txtBranchCode.Text != "")
                    where = where + " and wsh_branch_id = '" + txtBranchCode.Text + "' ";
                if (txtSalesCode.Text != "")
                    where = where + " and wsh_spgm_id = '" + txtSalesCode.Text + "' ";
                if (txtOutletCode.Text != "")
                {
                    where = where + " and wsh_cust_code1 = '" + txtOutletCode.Text + "' ";
                    where = where + " and wsh_cust_code2 = '" + txtOutlet2.Text + "' ";
                }
                if (txtNoOrder.Text != "")
                    where = where + " and wsh_seq_no = '" + txtNoOrder.Text + "' ";
                if (txtNoPO.Text != "")
                    where = where + " and wsh_po_no = '" + txtNoPO.Text + "' ";
                if (cbSOStatus.Text != "ALL")
                    where = where + " and wsh_status_so = '" + cbSOStatus.EditValue + "' ";
                if (cbSOType.Text != "ALL")
                    where = where + " and wsh_so_type = '" + cbSOType.EditValue + "' ";
                if (cbSource.Text != "ALL")
                    where = where + " and wsh_qasir_flag = '" + cbSource.EditValue + "' ";
                if (cbFlagSloc.Text != "ALL")
                    where = where + " and isnull(wsh_flag_sloc,'') = '" + cbFlagSloc.EditValue + "' ";
                if (isCrossSite && txtShippingPlantId.Text != "ALL" && !String.IsNullOrEmpty(txtShippingPlantId.Text))
                {
                    where = where + " AND ISNULL(wsh_ship_plant,'') = '" + txtShippingPlantId.Text + "'";
                }
                if (isExcludeNonDDEPCP)
                {
                    where = where + " AND isnull(wsh_flag_sloc,'') = '" + SourceNonDDE + "' ";
                }

                if (isCrossSite && txtShippingPlantId.Text == "ALL" && txtBranchCode.Text != "")
                {
                    where2 = " where convert(varchar(10), " + dateField + ", 120) >='" + dtTglOrder.DateTime.ToString("yyyy-MM-dd") + "' and convert(varchar(10), " + dateField + ", 120) <='" + dtTglOrderTo.DateTime.ToString("yyyy-MM-dd") + "' and ot_transaction_type = 'D' and ws_user_id = '" + clsLogin.USERID + "' ";
                    where2 = where2 + " AND ISNULL(wsh_ship_plant,'') = (SELECT br_branch_ud2 FROM GS_BRANCH WITH(NOLOCK) WHERE br_branch_id = '" + txtBranchCode.Text + "' )";
                }
                else
                {
                    where2 = "where (1=0)";
                }

                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();

                try
                {
                    cmd = new SqlCommand("IP_SO_MANUAL_LIST_MBRANCH", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.Parameters.Add(new SqlParameter("@where2", where2));
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dt);

                    gcSalesHeader.DataSource = dt;
                    con.Close();

                    if (dt.Rows.Count > 0)
                    {
                        gvSalesHeader.FocusedRowHandle = 0;
                        LoadSelectedHeader();
                    }
                    else
                    {
                        txtBebanCashDisc.Text = "";
                        FillGridDetail("");
                        FillGridTPRB("");
                        FillGridTPRU("");
                        FillGridDisc("");
                    }
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();
                }
            }
        }

        private void FillGridDetail(string NoOrder)
        {
            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {
                con.Open();
                string where = "where wsd_line_type = 'N' and wsd_so_seq_no = '" + NoOrder + "' ";

                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();

                try
                {
                    dtSD.Rows.Clear();

                    cmd = new SqlCommand("SP_SO_MANUAL_DETAIL_LIST", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dtSD);

                    gcSalesDetail.DataSource = dtSD;
                    con.Close();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();
                }
            }
        }

        private void FillGridTPRB(string NoOrder)
        {
            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {
                con.Open();
                string where = "where wsd_line_type = 'B' and wsd_so_seq_no = '" + NoOrder + "' and wsd_qty_sales > 0 ";

                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();

                try
                {
                    dtS.Rows.Clear();

                    cmd = new SqlCommand("SP_SO_MANUAL_TPRB_LIST", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dtS);

                    gcPromosiQty.DataSource = dtS;
                    con.Close();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();
                }
            }
        }

        private void FillGridTPRU(string NoOrder)
        {
            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {
                string isKAMstatus = "N";
                string where = "";

                con.Open();

                if (isKAM)
                {
                    where = "where dt_flag_proc = '1' and dt_order_no = '" + NoOrder + "'";
                    isKAMstatus = "Y";
                }
                else
                {
                    where = "where pt_flag_proc = '1' and pt_order_no = '" + NoOrder + "' ";
                }

                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();

                try
                {
                    dtU.Rows.Clear();

                    cmd = new SqlCommand("SP_SO_MANUAL_TPRU_LIST", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.Parameters.Add(new SqlParameter("@isKAM", isKAMstatus));
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dtU);

                    if (dtU.Rows.Count > 0)
                    {
                        dtGridTPRU.Rows.Clear();
                        lCounter = 0;

                        foreach (DataRow rw in dtU.Rows)
                        {
                            if (pCheckValueU(rw["pt_product_id"].ToString(), rw["pt_promo_id"].ToString(), 3))
                            {
                                DataRow nr = dtGridTPRU.NewRow();
                                nr["pt_promo_id"] = rw["pt_promo_id"].ToString();
                                nr["pt_product_id"] = rw["pt_product_id"];
                                nr["prm_prd_desc"] = rw["prm_prd_desc"];
                                decimal pro = Convert.ToDecimal(rw["pt_amount"]);
                                nr["pt_amount"] = pro.ToString("#,##0.00");
                                nr["pt_pct_promo"] = rw["pt_pct_promo"];
                                dtGridTPRU.Rows.Add(nr);

                                lCounter++;
                            }
                            else
                            {
                                decimal pro = ToDecimal(dtGridTPRU.Rows[lRow]["pt_amount"]);
                                dtGridTPRU.Rows[lRow]["pt_amount"] = pro.ToString("#,##0.00");
                            }
                        }
                    }
                    else
                    {
                        dtGridTPRU.Rows.Clear();
                    }

                    gcPromosiRp.DataSource = dtGridTPRU;
                    con.Close();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();
                }
            }
        }

        private void FillGridDisc(string NoOrder)
        {
            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {
                con.Open();

                string where = "where tdt_disc_level <> '0' and tdt_order_no = '" + NoOrder + "' ";

                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();

                try
                {
                    long Lbrs = 0;
                    dtDD.Rows.Clear();

                    cmd = new SqlCommand("SP_SO_MANUAL_DISC_LIST", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dtDD);

                    if (dtDD.Rows.Count > 0)
                    {
                        dtGridDisc.Rows.Clear();
                        lCounter = 0;

                        foreach (DataRow rw in dtDD.Rows)
                        {
                            string pcode = rw["tdt_product_master_id"].ToString();

                            if (pCheckValueD(pcode, 3, Lbrs))
                            {
                                DataRow nr = dtGridDisc.NewRow();
                                nr["tdt_product_master_id"] = rw["tdt_product_master_id"];
                                SetDiscValue(nr, rw);
                                dtGridDisc.Rows.Add(nr);
                                lCounter++;
                            }
                            else
                            {
                                SetDiscValue(dtGridDisc.Rows[lRow], rw);
                            }
                        }
                    }
                    else
                    {
                        dtGridDisc.Rows.Clear();
                    }

                    gcDisc.DataSource = dtGridDisc;
                    con.Close();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();
                }
            }
        }

        private void SetDiscValue(DataRow target, DataRow source)
        {
            int level = Convert.ToInt32(source["tdt_disc_level"]);
            decimal pct = Convert.ToDecimal(source["tdt_disc_pct"]);
            decimal val = Convert.ToDecimal(source["tdt_disc_value"]);

            string pctCol = "";
            string rpCol = "";

            switch (level)
            {
                case 1: pctCol = "g_iPctD1"; rpCol = "g_iRpD1"; break;
                case 2: pctCol = "g_iPctK1"; rpCol = "g_iRpK1"; break;
                case 3: pctCol = "g_iPctD2"; rpCol = "g_iRpD2"; break;
                case 4: pctCol = "g_iPctK2"; rpCol = "g_iRpK2"; break;
                case 5: pctCol = "g_iPctD3"; rpCol = "g_iRpD3"; break;
                case 6: pctCol = "g_iPctK3"; rpCol = "g_iRpK3"; break;
                case 7: pctCol = "g_iPctD4"; rpCol = "g_iRpD4"; break;
                case 8: pctCol = "g_iPctK4"; rpCol = "g_iRpK4"; break;
                case 9: pctCol = "g_iPctD5"; rpCol = "g_iRpD5"; break;
                case 10: pctCol = "g_iPctK5"; rpCol = "g_iRpK5"; break;
                case 11: pctCol = "g_iPctD6"; rpCol = "g_iRpD6"; break;
                case 12: pctCol = "g_iPctK6"; rpCol = "g_iRpK6"; break;
                case 13: pctCol = "g_iPctD7"; rpCol = "g_iRpD7"; break;
                case 14: pctCol = "g_iPctK7"; rpCol = "g_iRpK7"; break;
                case 15: pctCol = "g_iPctD8"; rpCol = "g_iRpD8"; break;
                case 16: pctCol = "g_iPctK8"; rpCol = "g_iRpK8"; break;
                case 17: pctCol = "g_iPctD9"; rpCol = "g_iRpD9"; break;
                case 18: pctCol = "g_iPctK9"; rpCol = "g_iRpK9"; break;
                case 19: pctCol = "g_iPctD10"; rpCol = "g_iRpD10"; break;
                case 20: pctCol = "g_iPctK10"; rpCol = "g_iRpK10"; break;
            }

            if (pctCol != "")
                target[pctCol] = pct.ToString("#,##0.00");
            if (rpCol != "")
                target[rpCol] = val.ToString("#,##0.00");
        }

        #endregion

        #region Helper

        private DataRow GetFocusedHeaderRow()
        {
            if (gvSalesHeader.FocusedRowHandle < 0) return null;
            return gvSalesHeader.GetDataRow(gvSalesHeader.FocusedRowHandle);
        }

        private string GetRowValue(DataRow row, string columnName)
        {
            if (row == null) return "";
            if (!row.Table.Columns.Contains(columnName)) return "";
            return Convert.ToString(row[columnName]);
        }

        private decimal ToDecimal(object value)
        {
            if (value == null || value == DBNull.Value) return 0;
            string s = Convert.ToString(value);
            if (string.IsNullOrWhiteSpace(s)) return 0;

            decimal result;
            return decimal.TryParse(s, out result) ? result : 0;
        }




        #endregion

        private bool IsExcludeNoNDDEPCP()
        {
            DataTable dtGS = _clsGlobal.ExecDT(
                "select gh_function_code " +
                "from GS_GEN_HARDCODED  WITH (NOLOCK) " +
                "where gh_function_name = 'Exclude_SO_Non_DDEPCP' " +
                "and gh_sys = 'H' " +
                "and gh_function_desc = '" + clsLogin.USERID + "' ");

            if (dtGS.Rows.Count > 0)
            {
                SourceNonDDE = dtGS.Rows[0]["gh_function_code"].ToString();
                return true;
            }

            return false;
        }
        protected void BindDefaultBranch(TextEdit txtEntity, TextEdit txtBranch)
        {
            DataTable dt = new DataTable();

            strSQL = "select distinct ge_entity_id [Entity], ge_entity [Desc]  " +
                     "from GS_ENTITY WITH(NOLOCK) " +
                     "INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id " +
                     "where gu_user_id = '" + clsLogin.USERID + "' " +
                     "order by ge_entity_id ";

            dt = _clsGlobal.ExecDT(strSQL);

            if (dt.Rows.Count > 0)
            {
                txtEntity.Text = dt.Rows[0]["Entity"].ToString();
            }

            DataTable dtBranch = new DataTable();

            strSQL = "select gu_branch [Branch], br_branch_desc [Desc] " +
                     "FROM VW_USERS_SECURITY WITH (NOLOCK)  " +
                     "where gu_entity = '" + txtEntityCode.Text.Trim() + "' " +
                     "and gu_user_id = '" + clsLogin.USERID + "'";

            dtBranch = _clsGlobal.ExecDT(strSQL);

            if (dtBranch.Rows.Count > 0)
            {
                txtBranch.Text = dtBranch.Rows[0]["Branch"].ToString();
            }
        }

    }
}