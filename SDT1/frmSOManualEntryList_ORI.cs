using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    public partial class frmSOManualEntryList : Form
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
        DataTable dtSD = new DataTable();
        DataTable dtS = new DataTable();
        DataTable dtU = new DataTable();
        DataTable dtDD = new DataTable();
        //private Decimal curHarga;

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

        #region WinForm Select

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
            //_clsGlobal.BindDefaultBranch(txtEntityCode, txtBranchCode);
            //WorkDateAccess();
            FLagDC();

            SetGridTPRU(dgvPromosiRp);
            SetGridDisc(dgvDisc);

            //txtEntityCode.Focus();
            //txtEntityCode.Select(0, 0);

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

        private void btnExec_Click(object sender, EventArgs e)
        {
            //if (txtEntityCode.Text == "")
            //{
            //    MessageBox.Show("Mohon tentukan Entity terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}
            //if (txtBranchCode.Text == "")
            //{
            //    MessageBox.Show("Mohon tentukan Branch terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}

            FillGridList();

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
            if (dgvSalesHeader.CurrentCell == null)
            {
                MessageBox.Show("Mohon Pilih SO yang akan diedit ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            string sStatusCode = "";
            //check so status
            DataTable dtFill1 = new DataTable();
            strSQL = " select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_sys = 'H' ";
            strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
            strSQL = strSQL + " and gh_sequence_no = 3";
            dtFill1 = _clsGlobal.ExecDT(strSQL);

            if (dtFill1.Rows.Count > 0)
            {
                sStatusCode = dtFill1.Rows[0]["gh_function_code"].ToString();
                //keyFound = true;
                //return keyFound;
            }
            else
            {
                MessageBox.Show("Status Approve tidak ada pada GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //keyFound = false;
                return;
            }


            //DataTable dtCheckIsAPI = new DataTable();
            //strSQL = " select wsh_qasir_flag from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
            //strSQL = strSQL + " WHERE wsh_seq_no = '" + dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells["wsh_seq_no"].Value.ToString() + "' ";
            //dtCheckIsAPI = _clsGlobal.ExecDT(strSQL);
            //if (dtCheckIsAPI.Rows.Count > 0)
            //{
            //    bool BlockingAPI = false;
            //    if (dtCheckIsAPI.Rows[0]["wsh_qasir_flag"].ToString().Trim() == "N" || dtCheckIsAPI.Rows[0]["wsh_qasir_flag"].ToString().Trim() == "D")
            //    {
            //        BlockingAPI = false;
            //    }
            //    else
            //    {
            //        BlockingAPI = true;
            //    }
            //    if (BlockingAPI)
            //    {
            //        MessageBox.Show("SO No." + dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells["wsh_seq_no"].Value.ToString() + " flagnya SINBAD/BL tidak bisa di Edit.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        //keyFound = false;
            //        return;
            //    }
            //}

            DataTable dtFill2 = new DataTable();
            strSQL = " select wsh_seq_no,wsh_qasir_flag,wsh_extract_sap,wsh_runcode_ext  from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
            strSQL = strSQL + " WHERE wsh_seq_no = '" + dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells["wsh_seq_no"].Value.ToString() + "' ";
            strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            if (dtFill2.Rows.Count > 0)
            {

            }
            else
            {
                DataTable dtFill3 = new DataTable();
                strSQL = " select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_sys = 'H' ";
                strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
                strSQL = strSQL + " and gh_sequence_no = 1";
                dtFill3 = _clsGlobal.ExecDT(strSQL);

                if (dtFill3.Rows.Count > 0)
                {
                    sStatusCode = dtFill3.Rows[0]["gh_function_code"].ToString();
                    //keyFound = true;
                    //return keyFound;
                }
                else
                {
                    MessageBox.Show("Status Open tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //keyFound = false;
                    return;
                }

                DataTable dtFill4 = new DataTable();
                strSQL = " select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells["wsh_seq_no"].Value.ToString() + "' ";
                strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
                dtFill4 = _clsGlobal.ExecDT(strSQL);

                if (dtFill4.Rows.Count > 0)
                {

                }
                else
                {
                    MessageBox.Show("Status SO sudah bukan pada Approve atau Open ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //keyFound = false;
                    return;
                }
            }

            if (dgvSalesHeader.Rows.Count > 0)
            {
                checkStatus();
                int rowindex = dgvSalesHeader.CurrentCell.RowIndex;
                string g_iStatus = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();

                if (g_iStatus == "O" || g_iStatus == "A")
                {
                    clsGlobal.MODE_TRX = clsGlobal.TRX_EDIT;
                    frmSOManualEntryEditor frm = new frmSOManualEntryEditor();
                    frm.xEntityID = dgvSalesHeader.Rows[rowindex].Cells["wsh_entity_id"].Value.ToString().Trim();
                    frm.xBranchID = dgvSalesHeader.Rows[rowindex].Cells["wsh_branch_id"].Value.ToString().Trim();
                    frm.xSONo = dgvSalesHeader.Rows[rowindex].Cells["wsh_seq_no"].Value.ToString().Trim();
                    //frm.xBranchID = dgvSalesHeader.Rows[rowindex].Cells["wsh_branch_id"].Value.ToString().Trim();
                    frm.xSalesID = dgvSalesHeader.Rows[rowindex].Cells["wsh_spgm_id"].Value.ToString().Trim();
                    frm.xFlagDC = g_FlagDC;
                    //frm.xEmployee = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    //frm.xOrderType = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    //frm.xSalesDesc = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    //frm.xTglOrder = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    //frm.xTglTransaksi = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    frm.xTipeSales = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    //frm.xWHLoc1 = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    //frm.xWHLoc2 = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    frm.ShowDialog();
                }
            }
        }

        private void tsb_delete_Click(object sender, EventArgs e)
        {
            checkStatus();
            try
            {

                if (dgvSalesHeader.Rows.Count > 0)
                {
                    int rowindex = dgvSalesHeader.CurrentCell.RowIndex;
                    string g_iStatus = dgvSalesHeader.Rows[rowindex].Cells["wsh_status_so"].Value.ToString().Trim();
                    string g_entity = dgvSalesHeader.Rows[rowindex].Cells["wsh_entity_id"].Value.ToString().Trim();
                    string g_branch = dgvSalesHeader.Rows[rowindex].Cells["wsh_branch_id"].Value.ToString().Trim();
                    if (g_iStatus == "O" || g_iStatus == "A")
                    {
                        DataTable dtFill5 = new DataTable();
                        strSQL = "select wsh_seq_no,wsh_runcode_ext,wsh_extract_sap from SO_WEB_SALES_HEADER  WITH (NOLOCK) ";
                        strSQL = strSQL + " WHERE wsh_seq_no = '" + NoOrder + "' ";
                        strSQL = strSQL + " AND wsh_status_so = 'O' ";
                        dtFill5 = _clsGlobal.ExecDT(strSQL);

                        if (dtFill5.Rows.Count > 0)
                        {
                            /** UPdate Ticket 26011 **/
                            int rowindex1 = dgvSalesHeader.CurrentCell.RowIndex;
                            string g_iStatus1 = dgvSalesHeader.Rows[rowindex1].Cells["wsh_status_so"].Value.ToString().Trim();
                            if (g_iStatus1 == "O")
                            {
                                if (isMultisource)
                                {
                                    string extractsap = dtFill5.Rows[0]["wsh_extract_sap"].ToString();
                                    string runcodeext = dtFill5.Rows[0]["wsh_runcode_ext"].ToString();
                                    if (extractsap.Equals("Y") && !string.IsNullOrEmpty(runcodeext))
                                    {
                                        MessageBox.Show("SO " + dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells["wsh_seq_no"].Value.ToString() + " tidak bisa di Cancel lakukan Proses Order terlebih Dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        return;
                                    }
                                }
                            }
                            /** End Update Ticket 26011 **/
                        }

                        if (CekFlagTMS(NoOrder) == true)
                        {
                            MessageBox.Show("No Order ter-interface ke TMS, proses cancel dibatalkan", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            if (g_iStatus == "O")
                            {
                                frmSOManualReason frm = new frmSOManualReason();
                                frm.ShowDialog();
                                if (frm.xHaveClickR == 1)
                                {
                                    DialogResult dr;
                                    dr = MessageBox.Show(_clsGlobal.ApplMessage(10001), clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                                    if (dr == DialogResult.OK)
                                    {
                                        pCancelSO(frm.xReasonR, g_entity, g_branch);
                                        FillGridList();
                                    }
                                }
                            }
                            else
                            {
                                DialogResult dr;
                                dr = MessageBox.Show(_clsGlobal.ApplMessage(10001), clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
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
                        MessageBox.Show("SO status tidak bisa di Cancel ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                //_clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsb_print_Click(object sender, EventArgs e)
        {

        }

        private void tsb_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

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
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpEntityCode_Click(sender, e);
            }
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
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpBranchCode_Click(sender, e);
            }
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
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpSalesCode_Click(sender, e);
            }
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
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpOutlet_Click(sender, e);
            }
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
                cbdateselection.SelectedIndex = 0;
                cbdateselection.Enabled = false;

                txtNoOrder.Text = frm.ArrField[0].Trim();
                txtNoPO.Text = "";
                DateTime tglO = Convert.ToDateTime(frm.ArrField[1].Trim());
                dtTglOrder.Text = tglO.ToString();
                dtTglOrderTo.Text = tglO.ToString();
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
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpNoOrder_Click(sender, e);
            }
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
                cbdateselection.SelectedIndex = 0;

                txtNoPO.Text = frm.ArrField[0].Trim();
                txtNoOrder.Text = "";

                DateTime tglO = Convert.ToDateTime(frm.ArrField[2].Trim());
                dtTglOrder.Text = tglO.ToString();
                dtTglOrderTo.Text = tglO.ToString();
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
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpNoPO_Click(sender, e);
            }
        }

        private void dgvSalesHeader_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                NoOrder = dgvSalesHeader.Rows[e.RowIndex].Cells["wsh_seq_no"].Value.ToString().Trim();
                txtBebanCashDisc.Text = dgvSalesHeader.Rows[e.RowIndex].Cells["wsh_beban_cashdisc_desc"].Value.ToString().Trim();
                Cursor.Current = Cursors.WaitCursor;
                FillGridDetail(NoOrder);
                FillGridTPRB(NoOrder);
                FillGridTPRU(NoOrder);
                FillGridDisc(NoOrder);
                CalcAmount();

                Cursor.Current = Cursors.Default;
            }
        }

        private void dgvSalesHeader_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dgvSalesHeader.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void dgvSalesHeader_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex % 2 == 0)
            {
                e.CellStyle.BackColor = Color.Aqua;
            }
            else
            {
                e.CellStyle.BackColor = Color.White;
            }
        }

        private void dgvSalesHeader_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvSalesHeader.CurrentCell != null)
            {
                NoOrder = dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells["wsh_seq_no"].Value.ToString().Trim();

                Cursor.Current = Cursors.WaitCursor;
                FillGridDetail(NoOrder);
                FillGridTPRB(NoOrder);
                FillGridTPRU(NoOrder);
                FillGridDisc(NoOrder);
                CalcAmount();


                Cursor.Current = Cursors.Default;
            }
        }

        private void dgvSalesHeader_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (clsLogin.BTNEDIT)
            {
                if (e.RowIndex > -1)
                {
                    dgvSalesHeader_CellClick(sender, e);
                    tsb_edit_Click(sender, e);
                }
            }
        }

        private void txtEntityCode_Leave(object sender, EventArgs e)
        {

            DataTable dtFill = new DataTable();
            strSQL = " select distinct ge_entity_id as Entity, ge_entity as EntityDesc  from GS_ENTITY WITH(NOLOCK) INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id where ge_entity_id ='" + txtEntityCode.Text + "' AND gu_user_id = '" + clsLogin.USERID + "' order by ge_entity_id ";
            dtFill = _clsGlobal.ExecDT(strSQL);
            if (dtFill.Rows.Count > 0)
            {
                txtEntityDesc.Text = dtFill.Rows[0]["EntityDesc"].ToString().Trim();
            }
            else
            {
                txtEntityDesc.Text = "";
            }
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
            if (dtFill.Rows.Count > 0)
            {
                txtSalesDesc.Text = dtFill.Rows[0]["salesName"].ToString().Trim();
            }
            else
            {
                txtSalesDesc.Text = "";
            }
        }

        private void txtOutletCode_Leave(object sender, EventArgs e)
        {
            DataTable dtFill = new DataTable();
            strSQL = " select cm_cust_code1 as custID1, cm_cust_code2 as CustID2, cm_cust_name as CustName from SO_CUST_MASTER WITH (NOLOCK) WHERE cm_active_flag <> 'D' AND cm_cust_code1 ='" + txtOutletCode.Text + "' AND cm_entity = '" + txtEntityCode.Text + "' AND cm_branch = '" + txtBranchCode.Text + "' ";
            dtFill = _clsGlobal.ExecDT(strSQL);
            if (dtFill.Rows.Count > 0)
            {
                //txtOutletCode.Text = dtFill.Rows[0]["BranchDesc"].ToString().Trim();
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
                    cbdateselection.SelectedIndex = 0;
                    cbdateselection.Enabled = false;
                    txtNoPO.Enabled = false;
                    btnPopUpNoPO.Enabled = false;
                    dtTglOrder.Enabled = false;
                    dtTglOrderTo.Enabled = false;


                    txtNoPO.Text = "";
                    //txtNoOrder.Text = frm.ArrField[0].Trim();
                    DateTime tglO = Convert.ToDateTime(dtFill.Rows[0]["OrderDate"].ToString().Trim());
                    dtTglOrder.Text = tglO.ToString();
                    dtTglOrderTo.Text = tglO.ToString();
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
                    //MessageBox.Show(" NO Order tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                    cbdateselection.SelectedIndex = 0;

                    txtNoPO.Text = dtFill.Rows[0]["NoPO"].ToString().Trim();
                    txtNoOrder.Text = "";
                    DateTime tglO = Convert.ToDateTime(dtFill.Rows[0]["OrderDate"].ToString().Trim());
                    dtTglOrder.Text = tglO.ToString();
                    dtTglOrderTo.Text = tglO.ToString();
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
                    //MessageBox.Show(" NO PO tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        #endregion

        #region Function

        private void SetGridTPRU(DataGridView dgv)
        {
            dtGridTPRU.Columns.Add("pt_promo_id", typeof(string));
            dtGridTPRU.Columns.Add("pt_product_id", typeof(string));
            dtGridTPRU.Columns.Add("prm_prd_desc", typeof(string));
            dtGridTPRU.Columns.Add("pt_amount", typeof(string));
            dtGridTPRU.Columns.Add("pt_pct_promo", typeof(string));

            dgv.DataSource = dtGridTPRU;

            dgv.Columns["pt_promo_id"].HeaderText = "Promo Code";
            dgv.Columns["pt_product_id"].HeaderText = "PCODE";
            dgv.Columns["prm_prd_desc"].HeaderText = "Nama Barang";
            dgv.Columns["pt_amount"].HeaderText = "Promosi Rp";
            dgv.Columns["pt_pct_promo"].HeaderText = "Nilai Promo (%)";

            //hide column
            dgv.Columns["pt_promo_id"].Visible = true;
            dgv.Columns["pt_product_id"].Visible = true;
            dgv.Columns["prm_prd_desc"].Visible = true;
            dgv.Columns["pt_amount"].Visible = true;
            dgv.Columns["pt_pct_promo"].Visible = true;

            dgv.Columns["pt_promo_id"].ReadOnly = true;
            dgv.Columns["pt_product_id"].ReadOnly = true;
            dgv.Columns["prm_prd_desc"].ReadOnly = true;
            dgv.Columns["pt_amount"].ReadOnly = true;
            dgv.Columns["pt_pct_promo"].ReadOnly = true;


            //width header
            dgv.Columns["pt_amount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgv.Columns["opl_validity_date_to"].Width = 125;

            //alignment cellf
            dgv.Columns["pt_amount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            ////type data cell            
            dgv.Columns["pt_amount"].DefaultCellStyle.Format = "#,##0.00";
        }

        private void SetGridDisc(DataGridView dgv)
        {
            dtGridDisc.Columns.Add("tdt_product_master_id", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD1", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD1", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK1", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK1", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD2", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD2", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK2", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK2", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD3", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD3", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK3", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK3", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD4", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD4", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK4", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK4", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD5", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD5", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK5", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK5", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD6", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD6", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK6", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK6", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD7", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD7", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK7", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK7", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD8", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD8", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK8", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK8", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD9", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD9", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK9", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK9", typeof(string));
            dtGridDisc.Columns.Add("g_iPctD10", typeof(string));
            dtGridDisc.Columns.Add("g_iRpD10", typeof(string));
            dtGridDisc.Columns.Add("g_iPctK10", typeof(string));
            dtGridDisc.Columns.Add("g_iRpK10", typeof(string));


            dgv.DataSource = dtGridDisc;

            dgv.Columns["tdt_product_master_id"].HeaderText = "PCODE";
            dgv.Columns["g_iPctD1"].HeaderText = "%D1";
            dgv.Columns["g_iRpD1"].HeaderText = "RpD1";
            dgv.Columns["g_iPctK1"].HeaderText = "%K1";
            dgv.Columns["g_iRpK1"].HeaderText = "RpK1";
            dgv.Columns["g_iPctD2"].HeaderText = "%D2";
            dgv.Columns["g_iRpD2"].HeaderText = "RpD2";
            dgv.Columns["g_iPctK2"].HeaderText = "%K2";
            dgv.Columns["g_iRpK2"].HeaderText = "RpK2";
            dgv.Columns["g_iPctD3"].HeaderText = "%D3";
            dgv.Columns["g_iRpD3"].HeaderText = "RpD3";
            dgv.Columns["g_iPctK3"].HeaderText = "%K3";
            dgv.Columns["g_iRpK3"].HeaderText = "RpK3";
            dgv.Columns["g_iPctD4"].HeaderText = "%D4";
            dgv.Columns["g_iRpD4"].HeaderText = "RpD4";
            dgv.Columns["g_iPctK4"].HeaderText = "%K4";
            dgv.Columns["g_iRpK4"].HeaderText = "RpK4";
            dgv.Columns["g_iPctD5"].HeaderText = "%D5";
            dgv.Columns["g_iRpD5"].HeaderText = "RpD5";
            dgv.Columns["g_iPctK5"].HeaderText = "%K5";
            dgv.Columns["g_iRpK5"].HeaderText = "RpK5";
            dgv.Columns["g_iPctD6"].HeaderText = "%D6";
            dgv.Columns["g_iRpD6"].HeaderText = "RpD6";
            dgv.Columns["g_iPctK6"].HeaderText = "%K6";
            dgv.Columns["g_iRpK6"].HeaderText = "RpK6";
            dgv.Columns["g_iPctD7"].HeaderText = "%D7";
            dgv.Columns["g_iRpD7"].HeaderText = "RpD7";
            dgv.Columns["g_iPctK7"].HeaderText = "%K7";
            dgv.Columns["g_iRpK7"].HeaderText = "RpK7";
            dgv.Columns["g_iPctD8"].HeaderText = "%D8";
            dgv.Columns["g_iRpD8"].HeaderText = "RpD8";
            dgv.Columns["g_iPctK8"].HeaderText = "%K8";
            dgv.Columns["g_iRpK8"].HeaderText = "RpK8";
            dgv.Columns["g_iPctD9"].HeaderText = "%D9";
            dgv.Columns["g_iRpD9"].HeaderText = "RpD9";
            dgv.Columns["g_iPctK9"].HeaderText = "%K9";
            dgv.Columns["g_iRpK9"].HeaderText = "RpK9";
            dgv.Columns["g_iPctD10"].HeaderText = "%D10";
            dgv.Columns["g_iRpD10"].HeaderText = "RpD10";
            dgv.Columns["g_iPctK10"].HeaderText = "%K10";
            dgv.Columns["g_iRpK10"].HeaderText = "RpK10";

            //hide column
            //dgv.Columns["pt_product_id"].Visible = true;
            //dgv.Columns["prm_prd_desc"].Visible = true;
            //dgv.Columns["pt_amount"].Visible = true;
            dgv.Columns["tdt_product_master_id"].ReadOnly = true;
            dgv.Columns["g_iPctD1"].ReadOnly = true;
            dgv.Columns["g_iRpD1"].ReadOnly = true;
            dgv.Columns["g_iPctK1"].ReadOnly = true;
            dgv.Columns["g_iRpK1"].ReadOnly = true;
            dgv.Columns["g_iPctD2"].ReadOnly = true;
            dgv.Columns["g_iRpD2"].ReadOnly = true;
            dgv.Columns["g_iPctK2"].ReadOnly = true;
            dgv.Columns["g_iRpK2"].ReadOnly = true;
            dgv.Columns["g_iPctD3"].ReadOnly = true;
            dgv.Columns["g_iRpD3"].ReadOnly = true;
            dgv.Columns["g_iPctK3"].ReadOnly = true;
            dgv.Columns["g_iRpK3"].ReadOnly = true;
            dgv.Columns["g_iPctD4"].ReadOnly = true;
            dgv.Columns["g_iRpD4"].ReadOnly = true;
            dgv.Columns["g_iPctK4"].ReadOnly = true;
            dgv.Columns["g_iRpK4"].ReadOnly = true;
            dgv.Columns["g_iPctD5"].ReadOnly = true;
            dgv.Columns["g_iRpD5"].ReadOnly = true;
            dgv.Columns["g_iPctK5"].ReadOnly = true;
            dgv.Columns["g_iRpK5"].ReadOnly = true;
            dgv.Columns["g_iPctD6"].ReadOnly = true;
            dgv.Columns["g_iRpD6"].ReadOnly = true;
            dgv.Columns["g_iPctK6"].ReadOnly = true;
            dgv.Columns["g_iRpK6"].ReadOnly = true;
            dgv.Columns["g_iPctD7"].ReadOnly = true;
            dgv.Columns["g_iRpD7"].ReadOnly = true;
            dgv.Columns["g_iPctK7"].ReadOnly = true;
            dgv.Columns["g_iRpK7"].ReadOnly = true;
            dgv.Columns["g_iPctD8"].ReadOnly = true;
            dgv.Columns["g_iRpD8"].ReadOnly = true;
            dgv.Columns["g_iPctK8"].ReadOnly = true;
            dgv.Columns["g_iRpK8"].ReadOnly = true;
            dgv.Columns["g_iPctD9"].ReadOnly = true;
            dgv.Columns["g_iRpD9"].ReadOnly = true;
            dgv.Columns["g_iPctK9"].ReadOnly = true;
            dgv.Columns["g_iRpK9"].ReadOnly = true;
            dgv.Columns["g_iPctD10"].ReadOnly = true;
            dgv.Columns["g_iRpD10"].ReadOnly = true;
            dgv.Columns["g_iPctK10"].ReadOnly = true;
            dgv.Columns["g_iRpK10"].ReadOnly = true;

            //width header
            dgv.Columns["g_iRpD1"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD2"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD3"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD4"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD5"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD6"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD7"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD8"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD9"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpD10"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK1"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK2"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK3"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK4"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK5"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK6"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK7"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK8"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK9"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["g_iRpK10"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgv.Columns["opl_validity_date_to"].Width = 125;

            //alignment cell
            dgv.Columns["g_iRpD1"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD2"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD3"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD4"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD5"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD6"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD7"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD8"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD9"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpD10"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK1"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK2"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK3"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK4"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK5"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK6"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK7"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK8"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK9"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["g_iRpK10"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            ////type data cell            
            dgv.Columns["g_iRpD1"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD2"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD3"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD4"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD5"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD6"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD7"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD8"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD9"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpD10"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK1"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK2"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK3"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK4"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK5"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK6"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK7"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK8"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK9"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["g_iRpK10"].DefaultCellStyle.Format = "#,##0.00";
        }

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

                foreach (DataGridViewRow gRow in dgvSalesDetail.Rows)
                {
                    if (gRow.Cells["ijumlahharga"].Value.ToString() != "")
                    {
                        if (gRow.Cells["wsd_tot_disc_pc"].Value.ToString() == "")
                        {
                            disc = 0;
                        }
                        else
                        {
                            disc = Convert.ToDecimal(gRow.Cells["wsd_tot_disc_pc"].Value);
                        }
                        curHarga = curHarga + (Convert.ToDecimal(gRow.Cells["ijumlahharga"].Value.ToString()) - disc);
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
                    decimal ttldisc1 = 0;
                    decimal ttldisc2 = 0;
                    decimal ttldisc3 = 0;
                    decimal ttldisc4 = 0;
                    decimal ttldisc5 = 0;
                    decimal ttldisc6 = 0;
                    decimal ttldisc7 = 0;
                    decimal ttldisc8 = 0;
                    decimal ttldisc9 = 0;
                    decimal ttldisc10 = 0;

                    if (dtFill.Rows[0]["wsh_tot_disc_1"].ToString() != "")
                    {
                        ttldisc1 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_1"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_2"].ToString() != "")
                    {
                        ttldisc2 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_2"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_3"].ToString() != "")
                    {
                        ttldisc3 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_3"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_4"].ToString() != "")
                    {
                        ttldisc4 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_4"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_5"].ToString() != "")
                    {
                        ttldisc5 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_5"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_6"].ToString() != "")
                    {
                        ttldisc6 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_6"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_7"].ToString() != "")
                    {
                        ttldisc7 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_7"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_8"].ToString() != "")
                    {
                        ttldisc8 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_8"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_9"].ToString() != "")
                    {
                        ttldisc9 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_9"].ToString());
                    }
                    if (dtFill.Rows[0]["wsh_tot_disc_10"].ToString() != "")
                    {
                        ttldisc10 = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_disc_10"].ToString());
                    }


                    promU = Convert.ToDecimal(dtFill.Rows[0]["wsh_tot_tpr_amt"].ToString());
                    disc110 = ttldisc1 + ttldisc2 + ttldisc3;
                    disc110 = disc110 + ttldisc4 + ttldisc5 + ttldisc6;
                    disc110 = disc110 + ttldisc7 + ttldisc8 + ttldisc9 + ttldisc10;
                    cashdisc = Convert.ToDecimal(dtFill.Rows[0]["wsh_pct_cash_disc"].ToString());
                    rpcashdisc = Convert.ToDecimal(dtFill.Rows[0]["wsh_cash_disc"].ToString());
                    ppn = Convert.ToDecimal(dtFill.Rows[0]["wsh_tax_amount"].ToString() == "" ? "0" : dtFill.Rows[0]["wsh_tax_amount"].ToString());
                    ttlinv = Convert.ToDecimal(dtFill.Rows[0]["wsh_net_amount"].ToString() == "" ? "0" : dtFill.Rows[0]["wsh_net_amount"].ToString());

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
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void AccessButton()
        {
            tsb_new.Enabled = false;
            tsb_edit.Enabled = false;
            tsb_delete.Enabled = false;
            tsb_print.Enabled = false;
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
                        dtTglOrder.Text = Convert.ToDateTime(dtFill.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                        dtTglOrderTo.Text = Convert.ToDateTime(dtFill.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                        g_dTglGudang = Convert.ToDateTime(dtFill.Rows[0]["wh_loc_last_work_date"].ToString());

                        tsb_new.Enabled = clsLogin.BTNNEW;
                        tsb_edit.Enabled = clsLogin.BTNEDIT;
                        tsb_delete.Enabled = clsLogin.BTNDELETE;
                        tsb_print.Enabled = clsLogin.BTNPRINT;
                    }
                    else
                    {

                        tsb_new.Enabled = false;
                        tsb_edit.Enabled = false;
                        tsb_delete.Enabled = false;
                        tsb_print.Enabled = false;
                    }

                }
                else
                {
                    g_dTglGudang = DateTime.Now;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        //frmSOManualEntryEditor frm = new frmSOManualEntryEditor();
                        //frm.xFlagDC = dtFill.Rows[0]["gh_function_code"].ToString();
                    }
                    else
                    {
                        MessageBox.Show("Silahkan tentukan value dari parameter FLAGDC terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else
                {
                    MessageBox.Show("Belum ada parameter DC dan proses dibatalkan", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private bool pCheckValueU(string sValue, string sValue2, int iIndex)
        {
            bool keyFound = true;
            for (lCounter = 0; lCounter < dgvPromosiRp.Rows.Count; lCounter++)
            {
                if (dgvPromosiRp.Rows[lCounter].Cells["pt_product_id"].Value.ToString().Trim() == sValue.Trim() && dgvPromosiRp.Rows[lCounter].Cells["pt_promo_id"].Value.ToString().Trim() == sValue.Trim())
                {
                    lRow = lCounter;
                    keyFound = false;
                    return keyFound;
                }
            }

            return keyFound;
        }

        private bool pCheckValueD(string sValue, int iIndex, long lRow)
        {
            bool keyFound = true;
            for (lCounter = 0; lCounter < dgvDisc.Rows.Count; lCounter++)
            {
                if (dgvDisc.Rows[lCounter].Cells["tdt_product_master_id"].Value.ToString().Trim() == sValue.Trim())
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
            {
                if (dtFill.Rows[0]["flag_tms"].ToString() == "Y")
                {
                    keyFound = true;
                }
                else
                {
                    keyFound = false;
                }
            }
            else
            {
                keyFound = false;
            }


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
                MessageBox.Show("Status Approve tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            DataTable dtFill2 = new DataTable();
            strSQL = "select wsh_seq_no,wsh_runcode_ext,wsh_extract_sap from SO_WEB_SALES_HEADER  WITH (NOLOCK) ";
            strSQL = strSQL + " WHERE wsh_seq_no = '" + NoOrder + "' ";
            strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            if (dtFill2.Rows.Count > 0)
            {

            }
            else
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
                    MessageBox.Show("Status Open tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                DataTable dtFill4 = new DataTable();
                strSQL = "select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + NoOrder + "' ";
                strSQL = strSQL + " AND wsh_status_so = '" + sStatusCode + "'";
                dtFill4 = _clsGlobal.ExecDT(strSQL);

                if (dtFill3.Rows.Count > 0)
                {

                }
                else
                {
                    MessageBox.Show("Status SO sudah bukan pada Approve atau Open ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        private void pCancelSO(string _xReasonR, string _xEntity, string _xBranch)
        {
            try
            {
                string sStatusSO = "";
                string sErrDesc = "";
                string sStatusCode = "";
                long iCounter = 0;
                //clsLogin.USERID

                //_clsGlobal.BeginTrans();
                if (_clsGlobal.IsNefoForKam)
                {
                    //Kondisi Update Untuk Promo First Buy
                    strSQL = " UPDATE SDCUST SET sdc_firstbuy_flag = 'N',	sdc_update_date=getdate() FROM SD_SALES_DEAL_TRX WITH (NOLOCK) ";
                    strSQL = strSQL + " INNER JOIN SO_WEB_SALES_HEADER WITH (NOLOCK) ON wsh_entity_id = dt_entity_id AND wsh_branch_id = dt_branch_id AND wsh_seq_no = dt_order_no ";
                    strSQL = strSQL + " INNER JOIN SD_SALES_DEAL_CUST SDCUST WITH (UPDLOCK) ON sdc_entity_id = wsh_entity_id	AND sdc_branch_id = wsh_branch_id AND sdc_outlet = wsh_cust_code1 ";
                    strSQL = strSQL + " AND sdc_outlet2 = wsh_cust_code2 AND sdc_promo_code = dt_promo_id AND sdc_firstbuy_flag = 'Y' ";
                    strSQL = strSQL + " WHERE  wsh_entity_id='" + _xEntity + "' and wsh_branch_id='" + _xBranch + "' AND wsh_seq_no='" + NoOrder + "'";
                    _clsGlobal.ExecDT(strSQL);
                }

                strSQL = " EXEC SP_CANCEL_SO '" + NoOrder + "','" + clsLogin.USERID + "','" + _xEntity + "','" + _xBranch + "'";
                _clsGlobal.ExecDT(strSQL);
                //_clsGlobal.ExecDTTrans(strSQL);

                strSQL = " EXEC SP_CHECK_BALANCE_SO_MULTIBRANCH '1','" + NoOrder + "', '" + _xEntity + "','" + _xBranch + "'";
                _clsGlobal.ExecDT(strSQL);
                //_clsGlobal.ExecDTTrans(strSQL);

                strSQL = "update SO_WEB_SALES_DETAIL set wsd_reason = '" + _xReasonR + "' where wsd_so_seq_no = '" + NoOrder + "'";
                _clsGlobal.ExecDT(strSQL);
                //_clsGlobal.ExecDTTrans(strSQL);

                if (_clsGlobal.IsNefoForKam)
                {
                    //Untuk KAM RECONDITION SUBTITUSI 
                    strSQL = "EXEC IP_PROSES_RECONDITION_SUBS_CANCEL_KAM '" + NoOrder + "','" + _xEntity + "','" + _xBranch + "','" + clsLogin.USERID + "'";
                    _clsGlobal.ExecDT(strSQL);
                    //_clsGlobal.ExecDTTrans(strSQL);
                }
                else
                {
                    if (__isAsk27757)
                    {
                        strSQL = "EXEC IP_PROSES_RECONDITION_SUBS_CANCEL_KAM '" + NoOrder + "','" + _xEntity + "','" + _xBranch + "','" + clsLogin.USERID + "'";
                        _clsGlobal.ExecDT(strSQL);
                    }
                }

                //_clsGlobal.CommitTrans();

                MessageBox.Show("No Order : " + NoOrder + " berhasil di cancel", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                //_clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void BindSource()
        {
            strSQL = "";
            strSQL = "SELECT 'ALL' as fcode, 'ALL' as fdesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'SOURCE_CLNT'";

            cbSource.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSource.ValueMember = "fcode";
            cbSource.DisplayMember = "fdesc";
        }

        private void BindSloc()
        {
            strSQL = "";
            strSQL = "SELECT 'ALL' as fcode, 'ALL' as fdesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'MULTISOURCE'";

            cbFlagSloc.DataSource = _clsGlobal.ExecDT(strSQL);
            cbFlagSloc.ValueMember = "fcode";
            cbFlagSloc.DisplayMember = "fdesc";
        }

        private void BindDetailSOType()
        {
            strSQL = "";
            strSQL = "SELECT '0' as otype, 'ALL' as odesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select ot_order_type as otype, ot_desc as odesc from SO_ORDER_TYPE WITH (NOLOCK)  where ot_transaction_type = 'D' order by otype";


            cbSOType.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSOType.ValueMember = "otype";
            cbSOType.DisplayMember = "odesc";
            //cbSOType.SelectedIndex = 0;

        }

        private void BindDetailSOStatus()
        {
            strSQL = "";
            strSQL = "SELECT '0' as fcode, 'ALL' as fdesc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'SOSTATUS'";

            cbSOStatus.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSOStatus.ValueMember = "fcode";
            cbSOStatus.DisplayMember = "fdesc";
        }

        private void BindDetailDateSelection()
        {
            strSQL = "";
            strSQL = "SELECT 'Order Date' as display, 'wsh_so_date' as value ";
            strSQL += "Union All SELECT 'PO Date' as display, 'wsh_po_date' as value ";
            strSQL += "Union All SELECT 'PO Delv Date' as display, 'wsh_sch_date' as value ";
            strSQL += "Union All SELECT 'PO Exp Date' as display, 'wsh_expired_date' as value ";
            cbdateselection.DataSource = _clsGlobal.ExecDT(strSQL);
            cbdateselection.DisplayMember = "display";
            cbdateselection.ValueMember = "value";
            cbdateselection.SelectedIndex = 0;
        }

        private void FillGridList()
        {
            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {
                DataTable dt = new DataTable();
                con.Open();

                string field = "";
                string keyword = "";
                //string where = "where convert(varchar(10), wsh_so_date, 120) ='" + dtTglOrder.Value.ToString("yyyy-MM-dd") + "' and ot_transaction_type = 'D' and ws_user_id = '" + clsLogin.USERID + "' ";
                string where2 = String.Empty;
                string where = "where convert(varchar(10), " + cbdateselection.SelectedValue + ", 120) >='" + dtTglOrder.Value.ToString("yyyy-MM-dd") + "' and convert(varchar(10), " + cbdateselection.SelectedValue + ", 120) <='" + dtTglOrderTo.Value.ToString("yyyy-MM-dd") + "' and ot_transaction_type = 'D' and ws_user_id = '" + clsLogin.USERID + "' ";

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
                    where = where + " and wsh_status_so = '" + cbSOStatus.SelectedValue + "' ";
                if (cbSOType.Text != "ALL")
                    where = where + " and wsh_so_type = '" + cbSOType.SelectedValue + "' ";
                if (cbSource.Text != "ALL")
                    where = where + " and wsh_qasir_flag = '" + cbSource.SelectedValue + "' ";
                if (cbFlagSloc.Text != "ALL")
                    where = where + " and isnull(wsh_flag_sloc,'') = '" + cbFlagSloc.SelectedValue + "' ";
                if (isCrossSite && txtShippingPlantId.Text != "ALL" && !String.IsNullOrEmpty(txtShippingPlantId.Text))
                {
                    where = where + " AND ISNULL(wsh_ship_plant,'') = '" + txtShippingPlantId.Text + "'";
                }
                if (isExcludeNonDDEPCP)
                {
                    where = where + " AND isnull(wsh_flag_sloc,'') = '" + SourceNonDDE + "' ";
                }
                //else
                //{
                //    where = where + " AND isnull(wsh_flag_sloc,'') not in (select distinct gh_function_code from GS_GEN_HARDCODED where gh_function_name like 'Exclude_SO_Non_DDEPCP') ";
                //}

                if (isCrossSite && txtShippingPlantId.Text == "ALL" && txtBranchCode.Text != "")
                {
                    where2 = " where convert(varchar(10), " + cbdateselection.SelectedValue + ", 120) >='" + dtTglOrder.Value.ToString("yyyy-MM-dd") + "' and convert(varchar(10), " + cbdateselection.SelectedValue + ", 120) <='" + dtTglOrderTo.Value.ToString("yyyy-MM-dd") + "' and ot_transaction_type = 'D' and ws_user_id = '" + clsLogin.USERID + "' ";

                    where2 = where2 + " AND ISNULL(wsh_ship_plant,'') = (SELECT br_branch_ud2 FROM GS_BRANCH WITH(NOLOCK) WHERE br_branch_id = '" + txtBranchCode.Text + "' )";
                }
                else
                {
                    where2 = "where (1=0)";
                }
                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();
                ////DataTable dt = new DataTable();
                //cmd.CommandType = CommandType.StoredProcedure;
                //cmd.CommandText = "SP_SO_MANUAL_LIST";
                //SqlParameter Par1 = cmd.Parameters.Add("@where", SqlDbType.VarChar, 255);
                //Par1.Value = where;

                try
                {
                    //cmd.Connection = con;
                    //da.SelectCommand = cmd;
                    //da.Fill(dt);
                    //dgvSalesHeader.DataSource = dt;
                    //dgvSalesHeader.DataBind();

                    cmd = new SqlCommand("IP_SO_MANUAL_LIST_MBRANCH", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.Parameters.Add(new SqlParameter("@where2", where2));
                    cmd.CommandTimeout = 0;
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dt);
                    dgvSalesHeader.DataSource = dt;
                    con.Close();

                    if (dgvSalesHeader.Rows.Count > 0)
                    {
                        txtBebanCashDisc.Text = dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells["wsh_beban_cashdisc_desc"].Value.ToString().Trim();
                        FillGridDetail(dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells[0].Value.ToString());
                        FillGridTPRB(dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells[0].Value.ToString());
                        FillGridTPRU(dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells[0].Value.ToString());
                        FillGridDisc(dgvSalesHeader.Rows[dgvSalesHeader.CurrentCell.RowIndex].Cells[0].Value.ToString());
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
                    //alert.InnerHtml = "<p class='msg error'>" + ex + "</p>";
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();

                }
            }
        }

        private void FillGridDetail(string NoOrder)
        {
            #region OLD QUERY

            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {

                con.Open();

                string where = "where wsd_line_type = 'N' and wsd_so_seq_no = '" + NoOrder + "' ";

                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();
                //cmd.CommandType = CommandType.StoredProcedure;
                //cmd.CommandText = "SP_SO_MANUAL_DETAIL_LIST";
                //SqlParameter Par1 = cmd.Parameters.Add("@where", SqlDbType.VarChar, 255);
                //Par1.Value = where;

                try
                {
                    //cmd.Connection = con;
                    //dgvSalesDetail.DataSource = cmd.ExecuteReader();
                    //dgvSalesHeader.DataBind();
                    if (dgvSalesDetail.Rows.Count > 0)
                        dtSD.Rows.Clear();

                    cmd = new SqlCommand("SP_SO_MANUAL_DETAIL_LIST", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dtSD);
                    dgvSalesDetail.DataSource = dtSD;
                    con.Close();
                }
                catch (Exception ex)
                {
                    //alert.InnerHtml = "<p class='msg error'>" + ex + "</p>";
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();

                }
            }

            #endregion



        }

        private void FillGridTPRB(string NoOrder)
        {
            using (SqlConnection con = new SqlConnection(clsGlobal.CONNECTION_STRING))
            {

                con.Open();

                string where = "where wsd_line_type = 'B' and wsd_so_seq_no = '" + NoOrder + "' and wsd_qty_sales > 0 ";

                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter da = new SqlDataAdapter();
                //cmd.CommandType = CommandType.StoredProcedure;
                //cmd.CommandText = "SP_SO_MANUAL_TPRB_LIST";
                //SqlParameter Par1 = cmd.Parameters.Add("@where", SqlDbType.VarChar, 255);
                //Par1.Value = where;

                try
                {

                    if (dgvPromosiQty.Rows.Count > 0)
                        dtS.Rows.Clear();
                    //cmd.Connection = con;
                    //dgvPromosiQty.DataSource = cmd.ExecuteReader();
                    ////dgvSalesHeader.DataBind();
                    cmd = new SqlCommand("SP_SO_MANUAL_TPRB_LIST", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dtS);

                    if (dtS.Rows.Count > 0)
                    {
                        dgvPromosiQty.DataSource = dtS;
                    }
                    else
                    {
                        //dtGridTPRU.Rows.Clear();
                        //dgvPromosiQty.Rows.Clear();
                    }
                    con.Close();
                }
                catch (Exception ex)
                {
                    //alert.InnerHtml = "<p class='msg error'>" + ex + "</p>";
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //cmd.CommandType = CommandType.StoredProcedure;
                //cmd.CommandText = "SP_SO_MANUAL_TPRU_LIST";
                //SqlParameter Par1 = cmd.Parameters.Add("@where", SqlDbType.VarChar, 255);
                //Par1.Value = where;

                try
                {

                    if (dgvPromosiRp.Rows.Count > 0)
                        dtU.Rows.Clear();
                    //cmd.Connection = con;
                    //dgvPromosiRp.DataSource = cmd.ExecuteReader();
                    ////dgvSalesHeader.DataBind();
                    //DataTable dtr = dgvPromosiRp.DataSource as DataTable;
                    long Lbrs = 0;
                    cmd = new SqlCommand("SP_SO_MANUAL_TPRU_LIST", con);
                    cmd.Parameters.Add(new SqlParameter("@where", where));
                    cmd.Parameters.Add(new SqlParameter("@isKAM", isKAMstatus));
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.SelectCommand = cmd;
                    da.Fill(dtU);

                    if (dtU.Rows.Count > 0)
                    {
                        dtGridTPRU.Rows.Clear();
                        //foreach (DataRow rw in dt.Rows)
                        //{
                        //    DataRow newRow = dtGridTPRU.NewRow();

                        //        newRow["pt_product_id"] = rw["pt_product_id"];
                        //        newRow["prm_prd_desc"] = rw["prm_prd_desc"];
                        //        newRow["pt_amount"] = rw["pt_amount"];
                        //        dtGridTPRU.Rows.Add(newRow);

                        //}
                        lCounter = 0;
                        foreach (DataRow rw in dtU.Rows)
                        {
                            if (pCheckValueU(rw["pt_product_id"].ToString(), rw["pt_promo_id"].ToString(), 3) == true) //'Jika produk tidak ada di Grid
                            {
                                //DataRow newRow = dtGridTPRU.NewRow();
                                //newRow["pt_product_id"] = rw["pt_product_id"];
                                //newRow["prm_prd_desc"] = rw["prm_prd_desc"];
                                //newRow["pt_amount"] = rw["pt_amount"];
                                //dtGridTPRU.Rows.Add(newRow);
                                //dgvPromosiRp.Rows.Add(newRow);
                                dtGridTPRU.Rows.Add();
                                dgvPromosiRp.Rows[lCounter].Cells["pt_promo_id"].Value = rw["pt_promo_id"].ToString();
                                dgvPromosiRp.Rows[lCounter].Cells["pt_product_id"].Value = rw["pt_product_id"];
                                dgvPromosiRp.Rows[lCounter].Cells["prm_prd_desc"].Value = rw["prm_prd_desc"];
                                decimal pro = Convert.ToDecimal(rw["pt_amount"]);
                                dgvPromosiRp.Rows[lCounter].Cells["pt_amount"].Value = pro.ToString("#,##0.00");
                                dgvPromosiRp.Rows[lCounter].Cells["pt_pct_promo"].Value = rw["pt_pct_promo"];

                                lCounter++;
                            }
                            else //'Jika produk ada di grid
                            {
                                decimal pro = Convert.ToDecimal(dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value);
                                //decimal pro = Convert.ToDecimal(dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value) + Convert.ToDecimal(rw["pt_amount"]);
                                dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value = pro.ToString("#,##0.00");
                            }
                        }
                    }
                    else
                    {
                        dtGridTPRU.Rows.Clear();
                        //dgvPromosiRp.Rows.Clear();
                    }

                    //dgvPromosiRp.DataSource = dt;
                    con.Close();
                }
                catch (Exception ex)
                {
                    //alert.InnerHtml = "<p class='msg error'>" + ex + "</p>";
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //cmd.CommandType = CommandType.StoredProcedure;
                //cmd.CommandText = "SP_SO_MANUAL_DISC_LIST";
                //SqlParameter Par1 = cmd.Parameters.Add("@where", SqlDbType.VarChar, 255);
                //Par1.Value = where;

                try
                {

                    //cmd.Connection = con;
                    //dgvDisc.DataSource = cmd.ExecuteReader();
                    ////dgvSalesHeader.DataBind();
                    //DataTable dtd = dgvDisc.DataSource as DataTable;
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
                        //foreach (DataRow rw in dt.Rows)
                        //{
                        //    DataRow newRow = dtGridTPRU.NewRow();

                        //        newRow["pt_product_id"] = rw["pt_product_id"];
                        //        newRow["prm_prd_desc"] = rw["prm_prd_desc"];
                        //        newRow["pt_amount"] = rw["pt_amount"];
                        //        dtGridTPRU.Rows.Add(newRow);

                        //}
                        lCounter = 0;
                        foreach (DataRow rw in dtDD.Rows)
                        {
                            if (pCheckValueD(rw["tdt_product_master_id"].ToString(), 3, Lbrs) == true) //'Jika produk tidak ada di Grid
                            {
                                //DataRow newRow = dtGridDisc.NewRow();
                                dtGridDisc.Rows.Add();
                                dgvDisc.Rows[lCounter].Cells["tdt_product_master_id"].Value = rw["tdt_product_master_id"];
                                dgvDisc.Rows[lCounter].Cells["tdt_product_master_id"].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                                if (rw["tdt_disc_level"].ToString() == "1")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD1"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD1"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");

                                }
                                else if (rw["tdt_disc_level"].ToString() == "2")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK1"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK1"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "3")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD2"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD2"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "4")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK2"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK2"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "5")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD3"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD3"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "6")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK3"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK3"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "7")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD4"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD4"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "8")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK4"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK4"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "9")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD5"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD5"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "10")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK5"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK5"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "11")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD6"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD6"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "12")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK6"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK6"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "13")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD7"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD7"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "14")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK7"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK7"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "15")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD8"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD8"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "16")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK8"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK8"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "17")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD9"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD9"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "18")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK9"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK9"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "19")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD10"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD10"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "20")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK10"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK10"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }


                                //dgvDisc.Rows.Add(newRow);
                            }
                            else //'Jika produk ada di grid
                            {
                                if (rw["tdt_disc_level"].ToString() == "1")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD1"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD1"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "2")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK1"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK1"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "3")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD2"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD2"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "4")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK2"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK2"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "5")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD3"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD3"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "6")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK3"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK3"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "7")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD4"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD4"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "8")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK4"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK4"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "9")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD5"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD5"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "10")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK5"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK5"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "11")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD6"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD6"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "12")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK6"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK6"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "13")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD7"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD7"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "14")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK7"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK7"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "15")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD8"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD8"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "16")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK8"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK8"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "17")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD9"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD9"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "18")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK9"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK9"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "19")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctD10"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpD10"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                else if (rw["tdt_disc_level"].ToString() == "20")
                                {
                                    dgvDisc.Rows[lCounter].Cells["g_iPctK10"].Value = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                                    dgvDisc.Rows[lCounter].Cells["g_iRpK10"].Value = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                                }
                                //decimal pro = Convert.ToDecimal(dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value) + Convert.ToDecimal(rw["pt_amount"]);
                                //dgvPromosiRp.Rows[lCounter].Cells["pt_amount"].Value = pro.ToString("");
                            }
                            lCounter++;
                        }

                        bool bHiddenCol = true;
                        for (int f = 1; f < 41; f++)
                        {
                            bHiddenCol = true;
                            for (lRow = 0; lRow < dgvDisc.Rows.Count; lRow++)
                            {
                                if (dgvDisc.Rows[lRow].Cells[f].Value.ToString() == "")
                                {
                                    bHiddenCol = false;
                                    break;
                                }
                            }
                            dgvDisc.Columns[f].Visible = bHiddenCol;
                        }
                    }
                    else
                    {
                        dtGridDisc.Rows.Clear();
                        //dgvDisc.Rows.Clear();
                    }

                    //dgvDisc.DataSource = dt;
                    con.Close();
                }
                catch (Exception ex)
                {
                    //alert.InnerHtml = "<p class='msg error'>" + ex + "</p>";
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();

                }
            }
        }

        #endregion
        protected void BindDefaultBranch(TextBox txtEntity, TextBox txtBranch)
        {
            DataTable dt = new DataTable();
            strSQL = "select distinct ge_entity_id [Entity], ge_entity [Desc]  from GS_ENTITY WITH(NOLOCK) INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id where gu_user_id = '" + clsLogin.USERID + "' order by ge_entity_id ";
            dt = _clsGlobal.ExecDT(strSQL);
            if (dt.Rows.Count > 0)
            {
                txtEntity.Text = dt.Rows[0]["Entity"].ToString();
            }
            DataTable dtBranch = new DataTable();

            strSQL = "select gu_branch [Branch], br_branch_desc [Desc] FROM VW_USERS_SECURITY WITH (NOLOCK)  where gu_entity = '" + txtEntityCode.Text.Trim() + "' and gu_user_id = '" + clsLogin.USERID + "'";
            dtBranch = _clsGlobal.ExecDT(strSQL);
            if (dtBranch.Rows.Count > 0)
            {
                txtBranch.Text = dtBranch.Rows[0]["Branch"].ToString();
            }
        }

        private void dgvSalesDetail_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dgvSalesHeader.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            strSQL = "SELECT 'ALL' AS [Shipping Plant Id], 'ALL' AS [Shipping Plant] UNION select wsh_ship_plant AS [Shipping Plant Id], br_branch_desc AS [Shipping Plant] from SO_WEB_SALES_HEADER WITH(NOLOCK) INNER JOIN GS_BRANCH WITH(NOLOCK) ON br_branch_ud2 = wsh_ship_plant WHERE wsh_branch_id = '" + txtBranchCode.Text + "'  ";

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
        private bool IsExcludeNoNDDEPCP()
        {

            DataTable dtGS = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED  WITH (NOLOCK) where gh_function_name = 'Exclude_SO_Non_DDEPCP' and gh_sys = 'H' and gh_function_desc = '" + clsLogin.USERID + "' ");
            if (dtGS.Rows.Count > 0)
            {
                SourceNonDDE = dtGS.Rows[0]["gh_function_code"].ToString();
                return true;
            }


            return false;
        }

        private void dgvSalesHeader_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

       

      



    }
}
