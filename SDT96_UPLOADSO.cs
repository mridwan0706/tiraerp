using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Threading;
using System.Globalization;

namespace TIRASnDNet.PROCESS.SO.SOUpload
{
    public partial class frmSOUpload : Form
    {
        private clsGlobal _clsGlobal = new clsGlobal();
        private string strSQL;
        private string paramMenuId;
        DataTable dtView = new DataTable();
        string dataFile = "";
        private int lCounter;

        private DataTable dtGridData = new DataTable();
        private DataTable dtHeader = new DataTable();
        private DataTable dtHeaderGrid = new DataTable();
        private DataTable dtDetail = new DataTable();

        Color __COLOR = Color.FromArgb(245, 245, 255);

        private decimal g_currMaxSelisih;
        private string eFile;
        private string sCustNo = "";
        private string sSldID = "";
        private decimal eParSelisih = 0;
        private string gNoPO;

        private string sWHLoc1s;
        private string sWHLoc2s;
        private DateTime dtWHDates;
        private decimal _curCost = 0;
        private bool bFlagCr;
        private bool bExists;
        private string _fConvertQty;

        private string sSONo;

        private string g_sDefaultSONo = "XXXXXXXXXX";
        private decimal _curAvCredLimit;
        private int m_iColumnCount;

        //private string sCustNo;
        private string sCust1;
        private string sCust2;
        private string sCustName;
        private string sTopId;
        private string sTOPVal;
        private string sPayType;
        private string sLeadtime;
        private string sSalesmanCode;
        private string sSalesmanName;
        //private string sTopId;
        //private string sTOPVal;
        //private string sPayType;
        //private string sLeadtime;

        private string sNoPO;
        private string ConvFail;
        private string sLastNoPO;
        private string dCustPcode;
        private string dCustPrice;

        private string sLineCode;
        private string sLineDesc;
        private string sPcode;
        private string sGrade;
        private string sSize;
        private string sSled;
        private string sDesc;
        private string iConv1;
        private string iConv2;
        private string sTaxCode;
        private string siTaxPct;
        private string sStatus;
        private string sWHLoc;
        private string sProdLine;
        private string sTopDefault;
        private string sTopDefaultValue;
        private string typePembayaran;
        private string sPrdSLED;

        string g_sSOHDR = "SOHDR";

        private decimal dTRSPrice;
        private decimal curPrice;
        private string _fQtyFormat;
        private double lAFSQty = 0;
        private bool isRoundingMekanism = false;
        private bool isNefoKAM = false;
        private bool isDivisionTOP = false;
        private bool resultInitNefoKAM = true;
        private bool isUploadMigrasiToNEFOKAM = false;
        private bool isMultiSource = false;
        private bool isMultibranch = false;
        private DateTime g_dTglGudang;


        public frmSOUpload()
        {
            InitializeComponent();
        }

        #region WinForm

        private void frmSOUpload_Load(object sender, EventArgs e)
        {
            this.Text = clsLogin.MENUID + " - " + clsLogin.MENUNAME + " " + this.Text;
            paramMenuId = clsLogin.MENUID;
            isRoundingMekanism = IsRoundingMekanism();
            isNefoKAM = _clsGlobal.IsNefoForKam;
            isMultiSource = _clsGlobal.IsMultiSource;
            isDivisionTOP = IsDivisionTOP();
            isUploadMigrasiToNEFOKAM = IsUploadMigrasiSAPToNEFOKAM();
            isMultibranch = _clsGlobal.IsMultibranch;

            AccessButton();
            if (isUploadMigrasiToNEFOKAM)
            {
                dgvHeader.Columns["g_iTOPId"].Visible = false;
                dgvHeader.Columns["g_iTOP"].Visible = false;
            }
            //if (isNefoKAM)
            //{


            //}
            //else
            //{
            cbSource.Visible = true;
            label2.Visible = true;
            label3.Visible = false;
            cbFlagSloc.Visible = false;
            fLoadSourceData();
            if (isMultiSource)
            {
                label3.Visible = true;
                cbFlagSloc.Visible = true;
                fLoadFlagSloc();
            }



            //}
            BindDetailEntity();
            BindDetailBranch();
            BindDefaultBranch(cbEntityBranch.txtCM, cbEntityBranch.txtBranchIdCM);
            pEnabled(1);
            fGetWHDate();
            _bindCombo();
        }

        private bool IsDivisionTOP()
        {
            DataTable dtTOPDivision = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'TOPBYDIVISION' and gh_sys = 'H'");
            if (dtTOPDivision.Rows.Count > 0)
            {
                if (dtTOPDivision.Rows[0][0].ToString().Trim() == "Y")
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsRoundingMekanism()
        {

            DataTable dtGS = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name = 'FLAG_ROUNDING_MEKANISM' and gh_sys = 'H'");
            if (dtGS.Rows.Count > 0)
            {
                if (dtGS.Rows[0][0].ToString().Trim() == "Y")
                {
                    return true;
                }
            }


            return false;
        }

        private bool IsUploadMigrasiSAPToNEFOKAM()
        {

            DataTable dtGS = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name = 'FLAG_UPLOAD_SO_SAP_TO_NEFOKAM' and gh_sys = 'H'");
            if (dtGS.Rows.Count > 0)
            {
                if (dtGS.Rows[0][0].ToString().Trim() == "Y")
                {
                    return true;
                }
            }


            return false;
        }

        private void btnPopUpSalesman_Click(object sender, EventArgs e)
        {
            string _strSQL = "select sgm_spgm_id as [Sales ID], sgm_spgm_name as [Sales Name], sgm_type_operasi as [Type Operasi], sgm_employee as [Employee] ";
            _strSQL = _strSQL + "  From dbo.SO_SPG_GIRL_MAN WITH (NOLOCK)  where sgm_spgm_status = '2'  and sgm_active_flag = 'A' and sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' and sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' order by sgm_spgm_id";

            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Salesman";
            frm.Query = _strSQL;
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtSalesman.Text = frm.ArrField[0].Trim();
                txtSalesmanDesc.Text = frm.ArrField[1].Trim();
                lblEmployee.Text = frm.ArrField[3].Trim();
                sSldID = frm.ArrField[0].Trim();

                DataTable dtFill1 = new DataTable();
                strSQL = " select msow_so_type code, msow_so_type + ' ~ ' + ot_desc order_type From TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) ";
                strSQL = strSQL + " inner join SO_ORDER_TYPE WITH (NOLOCK) on msow_so_type = ot_order_type";
                strSQL = strSQL + " where msow_sld_id = '" + txtSalesman.Text + "' and msow_entity_id = '" + cbEntityBranch.txtCM.Text + "' and msow_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                strSQL = strSQL + " and ot_transaction_type = 'D'";
                strSQL = strSQL + " Order By msow_so_type";
                dtFill1 = _clsGlobal.ExecDT(strSQL);

                if (dtFill1.Rows.Count > 0)
                {
                    //cbTipeOrder.DataSource = null;
                    //cbTipeOrder.Items.Clear();
                    cbTipeOrder.DataSource = _clsGlobal.ExecDT(strSQL);
                    cbTipeOrder.ValueMember = "code";
                    cbTipeOrder.DisplayMember = "order_type";
                    cbTipeOrder.SelectedIndex = -1;
                }
            }
        }


        private void txtSalesman_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpSalesman_Click(sender, e);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            string trsCode = "Product TRS Code Not Found";
            //for (int i = 0; i < dgvCommerce.Rows.Count; i++)
            //{
            //    strSQL = "";
            //    strSQL = "SELECT prm_prd_line_code FROM IM_PRD_MASTER WHERE prm_prd_master_code = '" + dgvCommerce.Rows[i].Cells[4].Value + "'";

            //    dtp = _clsGlobal.ExecDT(strSQL);
            //    //Application.DoEvents();
            //    if (dtp.Rows.Count > 0)
            //    {
            //        dgvCommerce.Rows[i].Cells[5].Value = dtp.Rows[0]["prm_prd_line_code"].ToString().Trim();
            //        dgvCommerce.Invoke(new Action(() => dgvCommerce.Rows[i].Cells[9].Value = "OK"));
            //        //dgvCommerce.Rows[i].Cells[9].Value = "OK";
            //        dgvCommerce.Rows[i].Cells[4].Style.BackColor = Color.Empty;
            //        dgvCommerce.Rows[i].Cells[5].Style.BackColor = Color.Empty;
            //    }
            //    else
            //    {
            //        dgvCommerce.Rows[i].Cells[5].Value = string.Empty;
            //        dgvCommerce.Invoke(new Action(() => dgvCommerce.Rows[i].Cells[9].Value = trsCode));
            //        dgvCommerce.Rows[i].Cells[4].Style.BackColor = Color.LightGreen;
            //        dgvCommerce.Rows[i].Cells[5].Style.BackColor = Color.LightGreen;

            //    }

            //    strSQL = "";
            //    strSQL = "SELECT cm_cust_name,cm_cust_group FROM SO_CUST_MASTER WHERE cm_cust_code1 = '" + dgvCommerce.Rows[i].Cells[0].Value + "'";

            //    DataTable dtDesc = _clsGlobal.ExecDT(strSQL);
            //    if (dtDesc.Rows.Count > 0)
            //    {

            //        dgvCommerce.Rows[i].Cells[1].Value = dtDesc.Rows[0]["cm_cust_name"].ToString().Trim();
            //        dgvCommerce.Rows[i].Cells[8].Value = dtDesc.Rows[0]["cm_cust_group"].ToString().Trim();
            //    }
            //    else
            //    {
            //        dgvCommerce.Rows[i].Cells[1].Value = string.Empty;
            //    }

            //    int percentage = (i + 1) * 100 / dgvCommerce.Rows.Count;
            //    backgroundWorker1.ReportProgress(percentage);
            //}
            ////while (progress < dtView.Rows.Count)
            ////{
            ////    progress = progress + 1;
            ////    backgroundWorker1.ReportProgress(progress);
            ////}
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Visible = false;

            //foreach (DataGridViewRow row in dgvCommerce.Rows)
            //{
            //    btnOk.Enabled = true;
            //    if (dgvCommerce.Rows[row.Index].Cells[9].Value == "Product TRS Code Not Found")
            //    {
            //        MessageBox.Show("There is any incorrect data,please check on column remarks!!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        btnOk.Enabled = false;
            //        break;
            //    }
            //}

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            //if (lblAccount.Text == "")
            //{
            //    MessageBox.Show("Outlet Account belum diisi !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}
            backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;

            backgroundWorker1.WorkerSupportsCancellation = true;
            //dgvCommerce.Rows.Clear();
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog ofd = new OpenFileDialog();
            //if (txtAccountDesc.Text == "LAZADA" || txtAccountDesc.Text == "BILNA")
            //if (txtAccountDesc.Text == "JD.ID")
            //{
            //    ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //}
            //else if (txtAccountDesc.Text == "BILNA" || txtAccountDesc.Text == "LAZADA")
            //{
            if (chkEcommerce.Checked)
            {
                ofd.Filter = "*.xls|*.xls|*.xlsx|*.xlsx";
            }
            else
            {
                ofd.Filter = "*.txt|*.txt|*.edi|*.edi";
            }

            //}
            ofd.FilterIndex = 0;
            ofd.RestoreDirectory = true;
            //ofd.InitialDirectory = "C:";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                dataFile = ofd.FileName;
                lblPath.Text = dataFile;
                filePath = dataFile;
                fileExt = Path.GetExtension(filePath);//get the file extension
                btnOk.Enabled = true;

            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            eExecute();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            pEnabled(1);
            pClearGrid();
        }

        private void btnProses_Click(object sender, EventArgs e)
        {
            bool bProses = false;
            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cSled = "";
            string sTopId = "";
            double cMinQty = 0;
            int cConv1 = 0;
            int cConv2 = 0;
            string _branch = "";
            string _entity = "";
            string _salesman = "";
            string _wh1 = "PST";
            string _wh2 = "000";

            string sPONo = "";
            double lAFSQty = 0;
            string sCredLimit = "";
            //string sSoNo;
            //string sCustNo ="";
            string sSoNoProc = "";
            int iProcess = 0;
            int iClose = 0;
            try
            {
                for (lCounter = 0; lCounter < dgvHeader.Rows.Count; lCounter++)
                {
                    if (Convert.ToBoolean(dgvHeader.Rows[lCounter].Cells["g_iProses"] == null ? "False" : dgvHeader.Rows[lCounter].Cells["g_iProses"].Value.ToString()) == true)
                    {
                        if (dgvHeader.Rows[lCounter].Cells["g_iSONo"].Value.ToString() != "")
                        {
                            MessageBox.Show("Data already process, please unchecked data !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                        else if (dgvHeader.Rows[lCounter].Cells["g_iRemark"].Value.ToString().ToLower() != "ok")
                        {
                            if (string.IsNullOrEmpty(dgvHeader.Rows[lCounter].Cells["g_iValid"].Value.ToString()))
                            {
                                MessageBox.Show("Please unchecked failed row data !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                            else
                            {
                                if (dgvHeader.Rows[lCounter].Cells["g_iValid"].Value.ToString().Trim() == "0")
                                {
                                    MessageBox.Show("Please unchecked failed row data !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    return;
                                }
                                else if (dgvHeader.Rows[lCounter].Cells["g_iValid"].Value.ToString().Trim() == "1" && string.IsNullOrEmpty(dgvHeader.Rows[lCounter].Cells["g_iReasonRejection"].Value.ToString()))
                                {
                                    MessageBox.Show("Reason Rejection must choose at row " + (lCounter + 1).ToString() + " !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    return;
                                }
                                else if (dgvHeader.Rows[lCounter].Cells["g_iValid"].Value.ToString().Trim() == "1" && !string.IsNullOrEmpty(dgvHeader.Rows[lCounter].Cells["g_iReasonRejection"].Value.ToString()))
                                {
                                    iClose++;
                                }
                            }
                        }
                    }
                    if (Convert.ToBoolean(dgvHeader.Rows[lCounter].Cells["g_iProses"].Value) == true)
                    {
                        iProcess = iProcess + 1;
                    }
                }

                if (iProcess == 0)
                {
                    MessageBox.Show("Please select row data to process !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //'Checking Mandatory
                for (lCounter = 0; lCounter < dgvHeader.Rows.Count; lCounter++)
                {

                    if (fCekMandatoryProses(lCounter) == false)
                    {
                        //_clsGlobal.RollbackTrans();
                        return;
                    }
                }

                sSoNoProc = "";
                if (!isNefoKAM)
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Data yang akan di Upload Berasal dari " + cbSource.Text + " apakah anda yakin akan Melanjutkan ?", "Proses Upload Order", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }

                if (iClose > 0)
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Data yang akan di Upload ada yang akan di close apakah anda akan melanjutkan ?", "Proses Upload Order", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }

                _clsGlobal.BeginTrans();
                for (int iCounter = 0; iCounter < dgvHeader.Rows.Count; iCounter++)
                {
                    if (Convert.ToBoolean(dgvHeader.Rows[iCounter].Cells["g_iProses"].Value) == true)
                    {
                        if (dgvHeader.Rows[iCounter].Cells["g_iSONo"].Value.ToString() == "")
                        {
                            sPONo = dgvHeader.Rows[iCounter].Cells["g_iCustPONo"].Value.ToString();
                            sCustNo = dgvHeader.Rows[iCounter].Cells["g_iCustCode"].Value.ToString();
                            sTopDefault = dgvHeader.Rows[iCounter].Cells["g_iTOPId"].Value.ToString();
                            sTopDefaultValue = dgvHeader.Rows[iCounter].Cells["g_iTOP"].Value.ToString();
                            sPayType = dgvHeader.Rows[iCounter].Cells["g_iPaymentType"].Value.ToString();
                            if (chkEcommerce.Checked)
                            {
                                _entity = dgvHeader.Rows[iCounter].Cells["g_iEntity"].Value.ToString();
                                _branch = dgvHeader.Rows[iCounter].Cells["g_iBranch"].Value.ToString();
                            }

                            //'CHECKING AFS
                            #region Checking Stock
                            for (int i = 0; i < dgvDetail.Rows.Count; i++)
                            {
                                cPcode = dgvDetail2.Rows[i].Cells["g_iDetPCode1"].Value.ToString();
                                cGrade = dgvDetail2.Rows[i].Cells["g_iDetGrade1"].Value.ToString();
                                cSize = dgvDetail2.Rows[i].Cells["g_iDetSize1"].Value.ToString();
                                cSled = dgvDetail2.Rows[i].Cells["g_iDetSled1"].Value.ToString();
                                cMinQty = Convert.ToDouble(dgvDetail2.Rows[i].Cells["g_iDetQty1"].Value.ToString());
                                cConv1 = Convert.ToInt16(dgvDetail2.Rows[i].Cells["g_iDetConv11"].Value.ToString());
                                cConv2 = Convert.ToInt16(dgvDetail2.Rows[i].Cells["g_iDetConv21"].Value.ToString());

                                lAFSQty = 0;

                                if (chkEcommerce.Checked)
                                {
                                    if (fCheckBalanceEcommerce(_entity, _branch, cPcode, cGrade, cSize, _wh1, _wh2) == true)
                                    {
                                        if (fGetAFSeCommerce(_entity, _branch, cPcode, cGrade, cSize, _wh1, _wh2) == true)
                                        {
                                            if (cMinQty > lAFSQty)
                                            {
                                                if (lAFSQty > 0)
                                                {
                                                    fQtyFormat(lAFSQty.ToString(), Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));

                                                    dgvDetail2.Rows[i].Cells["g_iDetQtyOrd1"].Value = _fQtyFormat;
                                                }
                                                else
                                                {
                                                    if (isUploadMigrasiToNEFOKAM)
                                                    {
                                                        fQtyFormat(cMinQty.ToString(), Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));
                                                        dgvDetail2.Rows[i].Cells["g_iDetQtyOrd1"].Value = _fQtyFormat;
                                                    }
                                                    else
                                                    {
                                                        fQtyFormat("0", Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));
                                                        dgvDetail2.Rows[i].Cells["g_iDetQtyOrd1"].Value = _fQtyFormat;
                                                    }

                                                }
                                            }
                                        }
                                        else
                                        {
                                            _clsGlobal.RollbackTrans();
                                            MessageBox.Show("Stock PCode : " + cPcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        _clsGlobal.RollbackTrans();
                                        MessageBox.Show("Stock PCode : " + cPcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        return;
                                    }

                                }
                                else
                                {
                                    if (fCheckBalance(cPcode, cGrade, cSize, lblWHLoc1.Text, lblWHLoc2.Text) == true)
                                    {
                                        if (fGetAFS(cPcode, cGrade, cSize, lblWHLoc1.Text, lblWHLoc2.Text) == true)
                                        {
                                            if (cMinQty > lAFSQty)
                                            {
                                                if (lAFSQty > 0)
                                                {
                                                    fQtyFormat(lAFSQty.ToString(), Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));

                                                    dgvDetail2.Rows[i].Cells["g_iDetQtyOrd1"].Value = _fQtyFormat;
                                                }
                                                else
                                                {
                                                    if (isUploadMigrasiToNEFOKAM)
                                                    {
                                                        fQtyFormat(cMinQty.ToString(), Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));
                                                        dgvDetail2.Rows[i].Cells["g_iDetQtyOrd1"].Value = _fQtyFormat;
                                                    }
                                                    else
                                                    {
                                                        fQtyFormat("0", Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));
                                                        dgvDetail2.Rows[i].Cells["g_iDetQtyOrd1"].Value = _fQtyFormat;
                                                    }

                                                }
                                            }
                                        }
                                        else
                                        {
                                            _clsGlobal.RollbackTrans();
                                            MessageBox.Show("Stock PCode : " + cPcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        _clsGlobal.RollbackTrans();
                                        MessageBox.Show("Stock PCode : " + cPcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        return;
                                    }
                                }

                            }
                            #endregion
                            // jika bukan nefokam
                            #region bukan Nefo Kam
                            if (!isNefoKAM)
                            {
                                if (fViewDetail(iCounter) == false)
                                {
                                    MessageBox.Show("Process Detail Not Found...!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    _clsGlobal.RollbackTrans();
                                    return;
                                }

                                if (chkEcommerce.Checked)
                                {
                                    if (fSaveSOEcommerce(sSONo, sCredLimit, false, iCounter) == false)
                                    {
                                        _clsGlobal.RollbackTrans();
                                        return;
                                    }
                                }
                                else
                                {
                                    if (fSaveSO(sSONo, sCredLimit, false, iCounter) == false)
                                    {
                                        _clsGlobal.RollbackTrans();
                                        return;
                                    }
                                }


                                //'INSERT TABLE PO OL HEADER
                                strSQL = " EXEC SP_UPLOADSO_HEADER ";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustCode"].Value.ToString() + "',";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustCode2"].Value.ToString() + "',";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustName"].Value.ToString() + "',";
                                strSQL = strSQL + "'" + txtSalesman.Text + "',";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustPONo"].Value.ToString() + "',";
                                dgvHeader.Rows[iCounter].Cells["g_iCustPODate"].Value = dtSODate.Value.ToString("yyyy-MM-dd");
                                strSQL = strSQL + "'" + Convert.ToDateTime(dgvHeader.Rows[iCounter].Cells["g_iCustPODate"].Value).ToString("yyyy-MM-dd") + "',";
                                strSQL = strSQL + "'" + sSONo + "',";
                                strSQL = strSQL + "'" + eFile + "',";
                                strSQL = strSQL + "'" + clsLogin.USERID + "'";
                                _clsGlobal.ExecuteTrans(strSQL);

                                dgvHeader.Rows[iCounter].Cells["g_iSONo"].Value = sSONo;
                                sSoNoProc = sSoNoProc + "~" + sSONo;

                                //'INSERT TABLE PO OL DETAIL
                                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                                {
                                    if (sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() && dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() != "")
                                    {
                                        strSQL = " EXEC SP_UPLOADSO_DTL ";
                                        strSQL = strSQL + "'" + sCustNo + "',";
                                        strSQL = strSQL + "'" + "000',";
                                        strSQL = strSQL + "'" + sPONo + "',";
                                        strSQL = strSQL + "'" + sSONo + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetValue1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetValue1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + clsLogin.USERID + "'";
                                        _clsGlobal.ExecuteTrans(strSQL);
                                    }
                                }

                            }
                            //jika nefokam
                            #endregion
                            #region KAM
                            else
                            {

                                bool top_by_cust = false;
                                if (chkEcommerce.Checked)
                                {
                                    _entity = dgvHeader.Rows[iCounter].Cells["g_iEntity"].Value.ToString();
                                    _branch = dgvHeader.Rows[iCounter].Cells["g_iBranch"].Value.ToString();
                                    _salesman = dgvHeader.Rows[iCounter].Cells["g_iSalesmanCode"].Value.ToString();
                                }
                                if (chkEcommerce.Checked)
                                {
                                    strSQL = "select cm_top_by_cust from SO_CUST_MASTER WITH (NOLOCK)  where cm_cust_code1 = '" + sCustNo + "' and cm_entity = '" + _entity + "' and cm_branch = '" + _branch + "'";
                                }
                                else
                                {
                                    strSQL = "select cm_top_by_cust from SO_CUST_MASTER WITH (NOLOCK)  where cm_cust_code1 = '" + sCustNo + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                                }

                                DataTable DtCust = _clsGlobal.ExecDTTrans(strSQL);
                                if (DtCust.Rows[0]["cm_top_by_cust"].ToString().Trim() == "Y")
                                {
                                    top_by_cust = true;
                                }
                                resultInitNefoKAM = true;
                                DataTable dtPrdLine = new DataTable();
                                if (chkEcommerce.Checked)
                                {
                                    dtPrdLine = InitNefoKAMProsesKAM(_entity, _branch, _salesman, sPONo, top_by_cust);
                                }
                                else
                                {
                                    dtPrdLine = InitNefoKAMProsesKAM(sPONo, top_by_cust);
                                }

                                if (resultInitNefoKAM == false)
                                {
                                    _clsGlobal.RollbackTrans();
                                    return;
                                }

                                /*dibuat update kalo 1 top dijadiin 1*/

                                #region update 09-03-2021

                                //***** update untuk ngebalikin 1 po 1 so **/
                                //dtPrdLine = InitNefoKAMProsesMigrasiSAPKAM_HEADER(sPONo, top_by_cust);
                                //for (int i = 0; i < dtPrdLine.Rows.Count; i++)
                                //{
                                //******************************************/
                                if (chkEcommerce.Checked)
                                {
                                    if (fSaveSOeCom_KAM(sSONo, sCredLimit, false, iCounter, dtPrdLine.Rows[0]["TOPID"].ToString(), dtPrdLine.Rows[0]["TOPVALUE"].ToString()) == false)
                                    {
                                        _clsGlobal.RollbackTrans();
                                        return;
                                    }
                                }
                                else
                                {
                                    if (fSaveSOMigrasiSAP_KAM(sSONo, sCredLimit, false, iCounter, dtPrdLine.Rows[0]["TOPID"].ToString(), dtPrdLine.Rows[0]["TOPVALUE"].ToString()) == false)
                                    {
                                        _clsGlobal.RollbackTrans();
                                        return;
                                    }
                                }


                                //'INSERT TABLE PO OL HEADER
                                strSQL = " EXEC SP_UPLOADSO_HEADER ";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustCode"].Value.ToString() + "',";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustCode2"].Value.ToString() + "',";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustName"].Value.ToString() + "',";
                                strSQL = strSQL + "'" + txtSalesman.Text + "',";
                                strSQL = strSQL + "'" + dgvHeader.Rows[iCounter].Cells["g_iCustPONo"].Value.ToString() + "',";
                                dgvHeader.Rows[iCounter].Cells["g_iCustPODate"].Value = dtSODate.Value.ToString("yyyy-MM-dd");
                                strSQL = strSQL + "'" + Convert.ToDateTime(dgvHeader.Rows[iCounter].Cells["g_iCustPODate"].Value).ToString("yyyy-MM-dd") + "',";
                                strSQL = strSQL + "'" + sSONo + "',";
                                strSQL = strSQL + "'" + eFile + "',";
                                strSQL = strSQL + "'" + clsLogin.USERID + "'";
                                _clsGlobal.ExecuteTrans(strSQL);

                                dgvHeader.Rows[iCounter].Cells["g_iSONo"].Value = sSONo;
                                sSoNoProc = sSoNoProc + "~" + sSONo;

                                //'INSERT TABLE PO OL DETAIL
                                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                                {
                                    //if (sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() && dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() != "" && dtPrdLine.Rows[i]["TOPVALUE"].ToString() == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString())
                                    if (sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() && dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() != "")
                                    {
                                        strSQL = " EXEC SP_UPLOADSO_DTL ";
                                        strSQL = strSQL + "'" + sCustNo + "',";
                                        strSQL = strSQL + "'" + "000',";
                                        strSQL = strSQL + "'" + sPONo + "',";
                                        strSQL = strSQL + "'" + sSONo + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetValue1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetValue1"].Value.ToString() + "',";
                                        strSQL = strSQL + "'" + clsLogin.USERID + "'";
                                        _clsGlobal.ExecuteTrans(strSQL);
                                    }

                                }
                                //Update untuk ngebalikin 1 PO = 1 SO //
                                // }
                                //***********************************//
                                #endregion



                            }
                            #endregion
                        }

                    }
                }
                _clsGlobal.CommitTrans();
                //_clsGlobal.RollbackTrans();

                pClearGrid();
                MessageBox.Show("Berhasil upload SO ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private DataTable InitNefoKAMProsesKAM(string sPONO, bool top_by_cust)
        {

            string PCode1 = "";
            string PCode2 = "";


            DataTable DtTemp = new DataTable();
            DtTemp.Columns.Add("ProdLine");
            DtTemp.Columns.Add("TOPID");
            DtTemp.Columns.Add("TOPVALUE");

            for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
            {
                if (sPONO == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() && dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() != "")
                {
                    string PL = dgvDetail2.Rows[lCounter].Cells["g_iDetProdLine1"].Value.ToString();

                    DataView dv = DtTemp.DefaultView;
                    dv.RowFilter = "ProdLine = '" + PL + "'";

                    DataTable dTempDv = dv.ToTable();
                    if (dTempDv.Rows.Count == 0)
                    {
                        if (!string.IsNullOrEmpty(PL))
                        {
                            if (sPayType == "T" || top_by_cust || isDivisionTOP == false)
                            {
                                DtTemp.Rows.Add(PL, sTopDefault, sTopDefaultValue);
                            }
                            else
                            {
                                //strSQL = "select distinct top 1 PM.prm_prd_line_code pcodelinenew,SP.ssp_spggroup_code,SP.ssp_prdline_code plinegrup,SMT.smt_prdline_id plinemap,SMT.smt_top_id, PTC.pptc_no_of_days " +
                                //         "   from IM_PRD_MASTER PM " +
                                //         "   left join TBL_SD_SPGGROUP_PRDLINE SP " +
                                //         "       ON  " +
                                //         "           SP.ssp_prdline_code = PM.prm_prd_line_code " +
                                //         "   left join SO_SPG_GIRL_MAN SPG " +
                                //         "       ON	" +
                                //         "           SPG.sgm_spgm_group = SP.ssp_spggroup_code " +
                                //         "   left join TBL_SD_CUSTCOVER TSC " +
                                //         "       ON " +
                                //         "           TSC.csc_entity = SPG.sgm_entity_id " +
                                //         "           AND TSC.csc_branch = SPG.sgm_branch_id " +
                                //         "           AND TSC.csc_salesman_id = SPG.sgm_spgm_id " +
                                //         "   left join SO_CUST_MASTER CM " +
                                //         "       ON " +
                                //         "           CM.cm_cust_code1 = TSC.csc_cust_code1 " +
                                //         "           and cm_cust_code2 = TSC.csc_cust_code2 " +
                                //         "           and cm_entity = TSC.csc_entity " +
                                //         "           and cm_branch = TSC.csc_branch " +
                                //         "   left Join SO_MAPPING_TOPBYPRDLINE SMT " +
                                //         "       ON " +
                                //         "           SMT.smt_entity_id = CM.cm_entity " +
                                //         "           and SMT.smt_branch_id = CM.cm_branch " +
                                //         "           and SMT.smt_cust_code1 = CM.cm_cust_code1 " +
                                //         "           and SMT.smt_cust_code2 = CM.cm_cust_code2 " +
                                //         "           and SMT.smt_prdline_id = PM.prm_prd_line_code " +
                                //         "   Left Join PO_PAYMENT_TERM_CODES PTC " +
                                //         "       ON " +
                                //         "           PTC.pptc_term_code = SMT.smt_top_id " +
                                //         "   where " +
                                //         "   SPG.sgm_spgm_id = '" + txtSalesman.Text.Trim() + "' " +
                                //         "   and SP.ssp_del_flag = 'N' " +
                                //         "   and SPG.sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' " +
                                //         "   and SPG.sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' " +
                                //         "   and CM.cm_cust_code1 = '" + sCustNo + "' " +
                                //         "   and SMT.smt_prdline_id = '"+ PL +"' " +
                                //         "   and CM.cm_top_by_cust = 'N' ";
                                strSQL = "select distinct top 1 PM.prm_prd_line_code pcodelinenew, SMT.smt_prdline_id  plinemap, SMT.smt_top_id, PTC.pptc_no_of_days ";
                                strSQL += "   from SO_CUST_MASTER CM WITH (NOLOCK) ";
                                strSQL += "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                                strSQL += "       ON  ";
                                strSQL += "           TSC.csc_entity          = CM.cm_entity ";
                                strSQL += "           AND TSC.csc_branch      = CM.cm_branch ";
                                strSQL += "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 ";
                                strSQL += "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 ";
                                strSQL += "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK)  ";
                                strSQL += "       ON	";
                                strSQL += "           SPG.sgm_spgm_id  = TSC.csc_salesman_id ";
                                strSQL += "           AND SPG.sgm_entity_id = TSC.csc_entity ";
                                strSQL += "           AND SPG.sgm_branch_id = TSC.csc_branch ";
                                strSQL += "   left join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK)  ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_entity_id      = CM.cm_entity ";
                                if (!isMultibranch)
                                {
                                    strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                                    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                }
                                else
                                {
                                    strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                                    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                }
                                strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC ";
                                strSQL += "       ON ";
                                strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                                strSQL += "   where ";
                                strSQL += "   SPG.sgm_spgm_id = '" + txtSalesman.Text.Trim() + "' ";
                                strSQL += "   and SPG.sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' ";
                                strSQL += "   and SPG.sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                                strSQL += "   and CM.cm_cust_code1 = '" + sCustNo + "' ";
                                strSQL += "   and SMT.smt_prdline_id = '" + PL + "' ";
                                strSQL += "   and CM.cm_top_by_cust = 'N' ";

                                DataTable dtNewProduct = _clsGlobal.ExecDTTrans(strSQL);
                                if (dtNewProduct.Rows.Count > 0)
                                {
                                    if (DtTemp.Rows.Count == 0)
                                    {
                                        DtTemp.Rows.Add(PL, dtNewProduct.Rows[0]["smt_top_id"].ToString(), dtNewProduct.Rows[0]["pptc_no_of_days"].ToString());
                                    }
                                    else
                                    {
                                        int no_days_exist = Convert.ToInt32(DtTemp.Rows[0][2].ToString());
                                        int no_days_this = Convert.ToInt32(dtNewProduct.Rows[0]["pptc_no_of_days"].ToString());
                                        if (no_days_exist > no_days_this)
                                        {
                                            DtTemp.Rows[0][1] = dtNewProduct.Rows[0]["smt_top_id"].ToString();
                                            DtTemp.Rows[0][2] = dtNewProduct.Rows[0]["pptc_no_of_days"].ToString();
                                        }
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("TOP Mapping untuk product line code = '" + PL + "' belum disetting.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }

                            }
                        }
                    }

                }
            }
            return DtTemp;

        }

        private DataTable InitNefoKAMProsesKAM(string sEntity, string sBranch, string sSalesman, string sPONO, bool top_by_cust)
        {

            string PCode1 = "";
            string PCode2 = "";


            DataTable DtTemp = new DataTable();
            DtTemp.Columns.Add("ProdLine");
            DtTemp.Columns.Add("TOPID");
            DtTemp.Columns.Add("TOPVALUE");

            for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
            {
                if (sPONO == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() && dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() != "")
                {
                    string PL = dgvDetail2.Rows[lCounter].Cells["g_iDetProdLine1"].Value.ToString();

                    DataView dv = DtTemp.DefaultView;
                    dv.RowFilter = "ProdLine = '" + PL + "'";

                    DataTable dTempDv = dv.ToTable();
                    if (dTempDv.Rows.Count == 0)
                    {
                        if (!string.IsNullOrEmpty(PL))
                        {
                            if (sPayType == "T" || top_by_cust || isDivisionTOP == false)
                            {
                                DtTemp.Rows.Add(PL, sTopDefault, sTopDefaultValue);
                            }
                            else
                            {                                
                                strSQL = "select distinct top 1 PM.prm_prd_line_code pcodelinenew, SMT.smt_prdline_id  plinemap, SMT.smt_top_id, PTC.pptc_no_of_days ";
                                strSQL += "   from SO_CUST_MASTER CM WITH (NOLOCK) ";
                                strSQL += "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                                strSQL += "       ON  ";
                                strSQL += "           TSC.csc_entity          = CM.cm_entity ";
                                strSQL += "           AND TSC.csc_branch      = CM.cm_branch ";
                                strSQL += "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 ";
                                strSQL += "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 ";
                                strSQL += "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK) ";
                                strSQL += "       ON	";
                                strSQL += "           SPG.sgm_spgm_id  = TSC.csc_salesman_id ";
                                strSQL += "           AND SPG.sgm_entity_id = TSC.csc_entity ";
                                strSQL += "           AND SPG.sgm_branch_id = TSC.csc_branch ";
                                strSQL += "   left join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_entity_id      = CM.cm_entity ";
                                if (!isMultibranch)
                                {
                                    strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                                    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                }
                                else
                                {
                                    strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                                    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                }
                                strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                                strSQL += "   where ";
                                strSQL += "   SPG.sgm_spgm_id = '" + sSalesman + "' ";
                                strSQL += "   and SPG.sgm_entity_id = '" + sEntity + "' ";
                                strSQL += "   and SPG.sgm_branch_id = '" + sBranch + "' ";
                                strSQL += "   and CM.cm_cust_code1 = '" + sCustNo + "' ";
                                strSQL += "   and SMT.smt_prdline_id = '" + PL + "' ";
                                strSQL += "   and CM.cm_top_by_cust = 'N' ";

                                DataTable dtNewProduct = _clsGlobal.ExecDTTrans(strSQL);
                                if (dtNewProduct.Rows.Count > 0)
                                {
                                    if (DtTemp.Rows.Count == 0)
                                    {
                                        DtTemp.Rows.Add(PL, dtNewProduct.Rows[0]["smt_top_id"].ToString(), dtNewProduct.Rows[0]["pptc_no_of_days"].ToString());
                                    }
                                    else
                                    {
                                        int no_days_exist = Convert.ToInt32(DtTemp.Rows[0][2].ToString());
                                        int no_days_this = Convert.ToInt32(dtNewProduct.Rows[0]["pptc_no_of_days"].ToString());
                                        if (no_days_exist > no_days_this)
                                        {
                                            DtTemp.Rows[0][1] = dtNewProduct.Rows[0]["smt_top_id"].ToString();
                                            DtTemp.Rows[0][2] = dtNewProduct.Rows[0]["pptc_no_of_days"].ToString();
                                        }
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("TOP Mapping untuk product line code = '" + PL + "' belum disetting.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }

                            }
                        }
                    }

                }
            }
            return DtTemp;

        }

        private void cbTipeOrder_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbTipeOrder.SelectedIndex != -1)
            {
                if (fChangeSOType() == false)
                {
                    return;
                }
            }
        }

        private void cbWH_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbTipeOrder.Text != "")
            {
                fChangeWH();
            }
        }

        private void dgvHeader_Click(object sender, EventArgs e)
        {
            if (dgvHeader.Rows.Count > 0)
            {
                if (dgvHeader.CurrentCell.ColumnIndex == 2)
                {
                    return;
                }
                fViewDetail(dgvHeader.CurrentCell.RowIndex);
            }
        }

        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvHeader.Rows)
            {
                if (chkAll.Checked == true)
                {
                    row.Cells["g_iProses"].Value = true;
                }
                else
                {
                    row.Cells["g_iProses"].Value = false;
                }
            }
        }

        private void cbEntityID_SelectedIndexChanged(object sender, EventArgs e)
        {
            BindDetailBranch();
        }

        private void txtSalesman_Leave(object sender, EventArgs e)
        {
            DataTable dtFill1 = new DataTable();
            strSQL = "select sgm_spgm_id as [Sales ID], sgm_spgm_name as [Sales Name], sgm_type_operasi as [Type Operasi], sgm_employee as [Employee] From dbo.SO_SPG_GIRL_MAN WITH (NOLOCK)";
            strSQL = strSQL + "  where sgm_spgm_id = '" + txtSalesman.Text + "'  and sgm_spgm_status = '2' and sgm_entity_id = '" + cbEntityBranch.txtCM.Text.Trim() + "' and sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' ";

            strSQL = strSQL + "  and sgm_active_flag = 'A' ";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                txtSalesmanDesc.Text = dtFill1.Rows[0][1].ToString();
                //txtSalesmanDesc.Text = frm.ArrField[1].Trim();
                lblEmployee.Text = dtFill1.Rows[0][3].ToString();
                sSldID = dtFill1.Rows[0][0].ToString();

                DataTable dtFill2 = new DataTable();
                strSQL = " select msow_so_type code, msow_so_type + ' ~ ' + ot_desc order_type From TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) ";
                strSQL = strSQL + " inner join SO_ORDER_TYPE WITH (NOLOCK) on msow_so_type = ot_order_type";
                strSQL = strSQL + " where msow_sld_id = '" + txtSalesman.Text + "' and msow_entity_id = '" + cbEntityBranch.txtCM.Text + "' and msow_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                strSQL = strSQL + " and ot_transaction_type = 'D'";
                strSQL = strSQL + " Order By msow_so_type";
                dtFill2 = _clsGlobal.ExecDT(strSQL);

                if (dtFill2.Rows.Count > 0)
                {
                    //cbTipeOrder.DataSource = null;
                    //cbTipeOrder.Items.Clear();
                    cbTipeOrder.DataSource = _clsGlobal.ExecDT(strSQL);
                    cbTipeOrder.ValueMember = "code";
                    cbTipeOrder.DisplayMember = "order_type";
                }
            }
            else
            {
                txtSalesman.Text = "";
                txtSalesmanDesc.Text = "";
            }
        }


        #endregion

        #region Function

        private void AccessButton()
        {
            tsb_new.Enabled = false;
            tsb_edit.Enabled = false;
            tsb_delete.Enabled = false;
            tsb_print.Enabled = false;
        }
        private bool fSaveSO(string bExtSono, string sCredLimit, bool bCalcDisc, int iRow)
        {
            bool keyFound = false;

            string sBranch = "";
            string sArea = "";
            string sWilayah = "";
            string sRayon = "";
            string sCustGroup = "";
            string sLineCode = "";
            string sYear = "";
            string sPeriode = "";
            //string sBranch = "";
            string sWeekNo = "";

            string sSOStatus = "";
            long lROrdQty = 0;
            decimal curROrdAmt = 0;
            decimal curCost = 0;
            decimal curTotalCost = 0;
            //clsPCode
            string sCurrCode = "";
            double dRateCurr = 0;
            long lBrs = 0;
            decimal curDisc1 = 0;
            decimal curDisc2 = 0;
            string sErrDesc = "";
            decimal currCredLimit = 0;
            //bool bFlagCr = false;
            string sPriceCode = "";
            string sOpenStatus = "";
            long lOrdQty = 0;
            long lLeadTime = 0;
            decimal curOrdAmt = 0;
            string pdaNO = "";
            string pdaFlag = "";
            string extsap = "";
            string runcext = "";
            string QasirFlag = "";
            string nik = "";

            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cSled = "";
            string cLineCode = "";
            string cGroup = "";
            string cSubGroup = "";
            string cModel = "";
            string cRawMatUsed = "";
            string cWashingCollorCode = "";
            long cConv1 = 0;
            long cConv2 = 0;

            string sCustCode1;
            string sCustCode2;
            string sPONo;
            //string sSoNo;
            Dictionary<string, string> SOPrincipal;
            string qasirflag = "";
            string sFullfilment = "N";


            bool sCancel = false;
            string reason = "";
            // string sValidity = "";

            try
            {
                keyFound = false;

                sPONo = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString();
                sCustCode1 = dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString();
                sCustCode2 = dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value.ToString();
                reason = dgvHeader.Rows[iRow].Cells["g_iReasonRejection"].Value.ToString();
                if (dgvHeader.Rows[iRow].Cells["g_iValid"].Value.ToString() == "1" && dgvHeader.Rows[iRow].Cells["g_iRemark"].Value.ToString().ToLower() != "ok")
                {
                    sCancel = true;
                }

                if (!cbSource.SelectedValue.ToString().CompareC("M"))
                {
                    qasirflag = cbSource.SelectedValue.ToString().Trim();
                }
                else
                {
                    qasirflag = "N";
                }

                if (bCalcDisc == false)
                {
                    if (fGetSONumber() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    sSONo = g_sDefaultSONo;

                    strSQL = " EXEC SP_DELETE_SO_TABLE ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }


                DataTable dtNIK = new DataTable();
                strSQL = "select sah_NIK from tbl_salesman_history WITH (NOLOCK) " +
                         " where" +
                         " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and sah_entity_id = '" + cbEntityBranch.txtCM.Text.Trim() + "'" +
                         " and sah_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'" +
                         " and sah_salesmen_id = '" + txtSalesman.Text + "'";
                dtNIK = _clsGlobal.ExecDTTrans(strSQL);

                if (dtNIK.Rows.Count == 1)
                {
                    nik = dtNIK.Rows[0]["sah_NIK"].ToString();
                }
                else if (dtNIK.Rows.Count > 1)
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }

                    MessageBox.Show("Salesman History Pada Periode Entry Lebih dari 1", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return keyFound;
                }
                else
                {

                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Salesman History Tidak Ditemukan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }



                //'Get branch, area, wilayah, rayon, cust group dan line code
                DataTable dtFill8 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') cm_line_code, cm_lead_time, ISNULL(cm_fullfilment,'N') cm_fullfilment From SO_CUST_MASTER WITH (NOLOCK)";
                strSQL = strSQL + " where cm_cust_code1 = '" + sCustCode1 + "' AND";
                strSQL = strSQL + " cm_cust_code2 = '" + sCustCode2 + "' ";
                dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill8.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill8.Rows[0]["cm_branch"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Area untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill8.Rows[0]["cm_area"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Wilayah untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill8.Rows[0]["cm_wilayah"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Rayon untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill8.Rows[0]["cm_rayon"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Cust group untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill8.Rows[0]["cm_cust_group"].ToString();
                    }

                    sLineCode = dtFill8.Rows[0]["cm_line_code"].ToString();
                    sFullfilment = dtFill8.Rows[0]["cm_fullfilment"].ToString();
                    //if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_lead_time"].ToString()))
                    //{
                    //    MessageBox.Show("Lead Time untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}
                    //else
                    //{
                    //    lLeadTime = Convert.ToInt64(dtFill8.Rows[0]["cm_lead_time"].ToString());
                    //}

                }
                else
                {
                    sBranch = "";
                    sArea = "";
                    sWilayah = "";
                    sRayon = "";
                    sCustGroup = "";
                    sLineCode = "";
                    sFullfilment = "N";
                }

                //'Get Year, Periode, Week untuk Tgl SO
                DataTable dtFill9 = new DataTable();
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK) ";
                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + dtSODate.Value.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '01' "; //entity id harus di lempar
                dtFill9 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill9.Rows.Count > 0)
                {
                    sWeekNo = dtFill9.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill9.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill9.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + dtSODate.Value.ToString("dd MMM yyyy") + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get SO Status
                DataTable dtFill10 = new DataTable();
                strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                strSQL = strSQL + " and gh_sequence_no = '1' ";
                dtFill10 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill10.Rows.Count > 0)
                {
                    sSOStatus = dtFill10.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    MessageBox.Show("Status SO (Open) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


                //Get Price Code
                //DataTable dtFill11 = new DataTable();
                //strSQL = " exec SP_GET_PRICE_CODE '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + sCustCode1 + "','" + sCustCode2 + "'";
                //dtFill11 = _clsGlobal.ExecDTTrans(strSQL);
                //if (dtFill11.Rows.Count > 0)
                //{
                //    if (String.IsNullOrEmpty(dtFill11.Rows[0]["cm_def_price_code"].ToString()))
                //    {
                //        MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        keyFound = false;
                //        return keyFound;
                //    }
                //    else
                //    {
                //        sPriceCode = dtFill11.Rows[0]["cm_def_price_code"].ToString();
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvDetail.Rows.Count; lCounter++)
                {
                    if (sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString() && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != "")
                    {
                        //fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + (int)Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curROrdAmt = curROrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            curROrdAmt = curROrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);
                        }

                        fConvertQty(dgvDetail.Rows[lCounter].Cells["g_iDetQtyOrd"].Value.ToString(), Convert.ToInt16(iConv1), Convert.ToInt16(iConv2));
                        lOrdQty = lOrdQty + Convert.ToInt32(_fConvertQty);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curOrdAmt = curOrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);

                        }
                        else
                        {
                            curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);

                        }

                        if (fGetCost(dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetGrade"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetSize"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        curTotalCost = curTotalCost + _curCost;
                    }
                }

                //if (sSONo.Length == 0)
                //{
                //    MessageBox.Show("nomor SO Blank", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //if (_xFlagDC == "N")
                //{
                // '===> Cek Kredit Limit dihilangkan jika flagDC = Y - Project Improvement TiraSnD
                //'/*-----------------------------------------------------------------------------------------
                //'Credit Limit
                if (fGetCreditLimit(sCustCode1, sCustCode2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (bFlagCr == true)
                {
                    if (_curAvCredLimit < curROrdAmt)
                    {
                        sCredLimit = "";
                    }
                    else
                    {
                        sCredLimit = "";
                    }
                }
                //}

                DataTable dtFill4 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_DETAIL A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsd_so_seq_no = '" + sSONo + "'";
                //strSQL = strSQL + "  AND wsh_branch_id = '" + cbBranchID.SelectedValue.ToString()  + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill5 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_HEADER A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "'";
                strSQL = strSQL + "  AND wsh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill12 = new DataTable();
                #region Query Insert
                strSQL = "Insert into SO_WEB_SALES_HEADER (";
                strSQL = strSQL + " wsh_entity_id ,wsh_area ,wsh_wilayah ,wsh_rayon ";
                strSQL = strSQL + " ,wsh_branch_id ,wsh_cust_group ,wsh_so_type ,wsh_seq_no";
                strSQL = strSQL + " ,wsh_so_date,wsh_cust_code1 ,wsh_cust_code2 ,wsh_loc_id1";
                strSQL = strSQL + " ,wsh_loc_id2,wsh_brand_line ,wsh_so_ref_no,wsh_so_txn_date_from";
                strSQL = strSQL + " ,wsh_so_txn_date_to,wsh_week_no,wsh_spgm_id ,wsh_spgm_group";
                strSQL = strSQL + " ,wsh_spgm_coordinator_id ,wsh_spgm_loc1 ,wsh_spgm_loc2 ,wsh_shift_no";
                strSQL = strSQL + " ,[wsh_participation_disc_%] ,wsh_participation_disc_amount ,wsh_tot_part_line_amount ,wsh_total_line";
                strSQL = strSQL + " , wsh_relevan_pod ,wsh_status_so ,wsh_real_order_qty   ,wsh_real_order_amt";
                strSQL = strSQL + " ,wsh_total_so_qty ,wsh_total_so_amt      ,wsh_include_tax ,[wsh_tax_%]";
                strSQL = strSQL + " ,wsh_total_partisipasi_disc_amount ,wsh_total_margin_disc_amount ,wsh_net_amount,wsh_dpp_amount";
                strSQL = strSQL + " ,wsh_tax_amount,wsh_invoice_process_flag ,wsh_invoice_no ,wsh_invoice_date";
                strSQL = strSQL + " ,wsh_invoice_process_by ,wsh_disc_margin,wsh_total_cost,wsh_year";
                strSQL = strSQL + " ,wsh_period,wsh_cancel_date,wsh_cancel_by,wsh_approval";
                strSQL = strSQL + " ,wsh_approval_by ,wsh_approval_date,wsh_ret_seq_no ,wsh_ret_by ,wsh_remarks";
                strSQL = strSQL + " ,wsh_participation_code ,wsh_entry_date,wsh_user_id ,wsh_from_OB_flag";
                strSQL = strSQL + " ,wsh_gross_net,wsh_gen_flag,wsh_tot_disc_pc,wsh_tot_disc_1";
                strSQL = strSQL + " ,wsh_tot_disc_2,wsh_tot_tpr_amt,wsh_tax_code,wsh_tax_pct";
                strSQL = strSQL + " ,wsh_pct_cash_disc, wsh_cash_disc, wsh_tax_pbm_pct, wsh_tax_pbm_amount, wsh_pay_type";
                strSQL = strSQL + " ,wsh_gross_net_amt, wsh_validate, wsh_po_no ";
                strSQL = strSQL + " ,wsh_flag_so_limit ";
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty,wsh_po_ol_flag ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                strSQL = strSQL + " , wsh_extract_sap ";
                strSQL = strSQL + " , wsh_runcode_ext  ";
                strSQL = strSQL + " , wsh_sch_date  ";
                strSQL = strSQL + " , wsh_expired_date  ";
                strSQL = strSQL + " , wsh_nik  ";
                strSQL = strSQL + " , wsh_qasir_flag  ";
                strSQL = strSQL + " , wsh_beban_cashdisc  ";
                if (isMultiSource)
                {
                    strSQL = strSQL + " , wsh_flag_sloc ";
                }
                strSQL = strSQL + " , wsh_is_fulfillment  ";
                strSQL = strSQL + " , wsh_sled_flag  ";
                string __tglorder = dtSODate.Value.ToString("yyyyMMdd");
                //int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '";  //'wsh_so_type
                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_so_date
                }
                else
                {
                    strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                }
                strSQL = strSQL + sCustCode1 + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + sCustCode2 + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + lblWHLoc1.Text + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + lblWHLoc2.Text + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + txtSalesman.Text + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "null , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "null , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "null , "; //'wsh_tot_part_line_amount
                strSQL = strSQL + "0, "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                strSQL = strSQL + "null , "; //'wsh_include_tax
                strSQL = strSQL + "null , "; //'[wsh_tax_%]
                strSQL = strSQL + "null , "; //'wsh_total_partisipasi_disc_amount
                strSQL = strSQL + "null , "; //'wsh_total_margin_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_net_amount
                strSQL = strSQL + "0 , "; //'wsh_dpp_amount
                strSQL = strSQL + "0 , "; //'wsh_tax_amount
                strSQL = strSQL + "'N' , "; //'wsh_invoice_process_flag
                strSQL = strSQL + "null , "; //'wsh_invoice_no
                strSQL = strSQL + "null , "; //'wsh_invoice_date
                strSQL = strSQL + "null , "; //'wsh_invoice_process_by
                strSQL = strSQL + "null , "; //'wsh_disc_margin
                strSQL = strSQL + curTotalCost + ", '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , "; //'wsh_approval
                strSQL = strSQL + "null , "; //'wsh_approval_by
                strSQL = strSQL + "null , "; //'wsh_approval_date
                strSQL = strSQL + "null , "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                //strSQL = strSQL + "null , "; //'wsh_participation_code
                if (cbSource.SelectedValue.ToString().Trim().CompareC("L") || cbSource.SelectedValue.ToString().Trim().CompareC("S") || cbSource.SelectedValue.ToString().Trim().CompareC("W"))
                {
                    if (string.IsNullOrEmpty(dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value.ToString()))
                    {
                        strSQL = strSQL + "null , "; //'wsh_participation_code
                    }
                    else
                    {
                        strSQL = strSQL + "'" + dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value.ToString() + "' , "; //'wsh_participation_code
                    }
                }
                else
                {
                    strSQL = strSQL + "null , "; //'wsh_participation_code
                }
                strSQL = strSQL + "getdate() , '"; //'wsh_entry_date
                strSQL = strSQL + clsLogin.USERID + "', "; //'wsh_user_id
                strSQL = strSQL + "null , "; //'wsh_from_OB_flag
                strSQL = strSQL + "null , "; //'wsh_gross_net
                strSQL = strSQL + "null , "; //'wsh_gen_flag
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_pc
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_1
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_2
                strSQL = strSQL + "0 , "; //'wsh_tot_tpr_amt
                strSQL = strSQL + "null , "; //'wsh_tax_code
                strSQL = strSQL + "null , "; //'wsh_tax_pct

                //if (dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() == "")
                //{
                //    strSQL = strSQL + "0" + " , ";
                //}
                //else
                //{
                //    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                //}
                bool iscashdiscbeban = false;
                if (dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() == "")
                {
                    strSQL = strSQL + "0" + " , ";
                }
                else
                {
                    if (cbSource.SelectedValue.ToString().Trim().CompareC("L") || cbSource.SelectedValue.ToString().Trim().CompareC("S") || cbSource.SelectedValue.ToString().Trim().CompareC("W"))
                    {
                        decimal cashValue = 0;
                        decimal cashGross = 0;
                        decimal pct = 0;
                        if (cbSource.SelectedValue.ToString().CompareC("L"))
                        {
                            cashValue = Convert.ToDecimal(dgvHeader.Rows[iRow].Cells["g_iCashDisc"].Value.ToString());
                            cashGross = Convert.ToDecimal(curROrdAmt);
                            pct = (cashValue / cashGross) * 100;
                        }
                        if (Convert.ToDecimal(dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString()) > 0)
                        {
                            strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                            //iscashdiscbeban = true;
                        }
                        else if (pct > 0)
                        {
                            //strSQL = strSQL + pct + " , "; //'wsh_pct_cash_disc
                            strSQL = strSQL + "0 , "; //'wsh_pct_cash_disc
                            //iscashdiscbeban = true;
                        }
                        else
                        {
                            strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                        }
                    }
                    else
                    {
                        strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                    }

                }


                strSQL = strSQL + "0 , "; //'wsh_cash_disc
                //strSQL = strSQL + "0 , "; //'wsh_cash_disc
                strSQL = strSQL + "0 , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "0 , '"; //'wsh_tax_pbm_amount
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value.ToString() + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , "; //'wsh_gross_net_amt
                strSQL = strSQL + "null , '"; // Convert.ToDateTime(dtTglValidasi.Value).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + sPONo + "' , '"; //'wsh_po_no
                strSQL = strSQL + sCredLimit + "' , '"; //'wsh_flag_so_limit
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                strSQL = strSQL + "N' , '"; //'wsh_po_ol_flag
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iTOP"].Value.ToString() + "' , '"; //'wsh_top
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iTOPId"].Value.ToString() + "' , '"; //'wsh_top_id
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_po_date
                }
                else
                {
                    strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "','"; //'wsh_po_date
                }
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value.ToString() + "'"; //'wsh_cust_leadtime
                strSQL = strSQL + ",'" + "N" + "' , '"; //'extsap
                strSQL = strSQL + "'"; //'runcext
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + ",'" + dgvHeader.Rows[iRow].Cells["g_iDeliveryDate"].Value.ToString() + "'"; // planning date
                    strSQL = strSQL + ",'" + dgvHeader.Rows[iRow].Cells["g_iExpiredDate"].Value.ToString() + "'"; // expired date
                }
                else
                {
                    strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // planning date
                    strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // expired date
                }
                strSQL = strSQL + ",'" + nik + "'"; // nik
                strSQL = strSQL + ",'" + qasirflag + "'"; // qasirflag 30-06-2021 -- update BL : kalo M kosong kalo yang lain diisi sama dengan cbSource.selectedvalue
                //if (iscashdiscbeban) // kalau ada cashdisc
                //{
                //    strSQL = strSQL + ",1"; //wsh_beban_cashdisc
                //}
                //else // kalau ga ada cashdisc
                //{
                strSQL = strSQL + ",0"; //wsh_beban_cashdisc
                //}
                //Convert.ToDateTime(dtTglOrder.Text).ToString("yyyyMMdd")
                if (isMultiSource)
                {
                    strSQL = strSQL + ",'" + cbFlagSloc.SelectedValue.ToString() + "'";
                }
                strSQL = strSQL + ",'" + sFullfilment.Trim() + "'"; //wsh_is_fulfillment
                if (dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                {
                    strSQL = strSQL + ",'N'"; //wsh_sled_flag
                }
                else
                {
                    strSQL = strSQL + ",'"+ dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().Trim() + "'"; //wsh_sled_flag
                }
                strSQL = strSQL + " ) ";
                dtFill12 = _clsGlobal.ExecDTTrans(strSQL);
                //'END OF Saving SO Header
                //if (dtFill12.Rows.Count > 0)
                //{
                DataTable dtFill13 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK) ";
                strSQL = strSQL + " where cc_currency_type = 'H' ";
                #endregion

                dtFill13 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill13.Rows.Count > 0)
                {
                    sCurrCode = dtFill13.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill13.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }

                lCounter = 0;
                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                {

                    if (sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString())
                    {
                        #region save detail
                        DataTable dtFill14 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK)";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "'";
                        dtFill14 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill14.Rows.Count > 0)
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = dtFill14.Rows[0]["prm_prd_line_code"].ToString();
                            cGroup = dtFill14.Rows[0]["prm_prd_group_code"].ToString();
                            cSubGroup = dtFill14.Rows[0]["prm_prd_sgroup_code"].ToString();
                            cModel = dtFill14.Rows[0]["prm_prd_model_code"].ToString();
                            cRawMatUsed = dtFill14.Rows[0]["prm_raw_mat_used_code"].ToString();
                            cWashingCollorCode = dtFill14.Rows[0]["prm_washing_collor_code"].ToString();
                            cConv1 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_purc"].ToString());
                            cConv2 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_sales"].ToString());
                        }
                        else
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = "";
                            cGroup = "";
                            cSubGroup = "";
                            cModel = "";
                            cRawMatUsed = "";
                            cWashingCollorCode = "";
                            cConv1 = 0;
                            cConv2 = 0;
                        }
                        //Get Price Code
                        DataTable dtFillGrpPrice = new DataTable();
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + sCustCode1 + "','" + sCustCode2 + "','" + __tglorder + "'";
                        dtFillGrpPrice = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFillGrpPrice.Rows.Count > 0)
                        {
                            if (String.IsNullOrEmpty(dtFillGrpPrice.Rows[0]["cp_price_group"].ToString()) == true)
                            {
                                MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                            else
                            {
                                sPriceCode = dtFillGrpPrice.Rows[0]["cp_price_group"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }


                        DataTable dtFill15 = new DataTable();
                        strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                        strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                        strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                        strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                        strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                        strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                        strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                        strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                        strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                        strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_sled)";

                        strSQL = strSQL + " values ( '";
                        strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '"; //'wsd_entity_id
                        strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '"; //'wsd_so_type
                        strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                        int lineno = lCounter + 1;
                        strSQL = strSQL + lineno + "', '"; //'wsd_so_detail_line_no
                        strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','"; //wsd_so_branch_id NEW
                        strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                        strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                        strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                        strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                        strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                        strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                        strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                        strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                        strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', '";
                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";
                        decimal OrderAmt = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString());
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            strSQL = strSQL + Math.Round(OrderAmt, 0, MidpointRounding.AwayFromZero) + "', '"; //'wsd_real_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + OrderAmt + "', '"; //'wsd_real_order_amt
                        }

                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";//'wsd_qty_sales
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', "; //'wsd_total_so_amount
                        strSQL = strSQL + "null, "; //'wsd_shipped_xqty
                        strSQL = strSQL + "null, "; //'wsd_shipped_qty
                        strSQL = strSQL + "null, "; //'wsd_shipped_amount
                        strSQL = strSQL + "null, "; //'wsd_pod_xqty
                        strSQL = strSQL + "null, "; //'wsd_pod_qty
                        strSQL = strSQL + "null, "; //'wsd_pod_amount
                        strSQL = strSQL + "null, "; //'wsd_invoice_xqty
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty
                        strSQL = strSQL + "null, "; //'wsd_invoice_amount
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_%
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_amt
                        strSQL = strSQL + "null, '"; //'wsd_margin

                        if (fGetCost(dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        strSQL = strSQL + _curCost.ToString() + "', '"; //'wsd_unit_cost
                        fConvertQty(dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString(), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value));
                        //Rounding Mekanism
                        decimal ttlcst = _curCost * Convert.ToDecimal(_fConvertQty);
                        if (isRoundingMekanism)
                        {
                            ttlcst = Math.Round(_curCost * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                        }
                        strSQL = strSQL + ttlcst.ToString() + "', "; //'wsd_total_cost
                        strSQL = strSQL + "null, "; //'wsd_remarks
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_flag
                        strSQL = strSQL + "null, "; //'wsd_invoice_no
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty_sales
                        strSQL = strSQL + "null, "; //'wsd_invoice_date
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_by
                        strSQL = strSQL + "getdate() , '"; //'wsd_entry_date
                        strSQL = strSQL + clsLogin.USERID + "', '"; //wsd_user_id
                        strSQL = strSQL + sCurrCode + "', "; //'wsd_curr_code
                        strSQL = strSQL + dRateCurr + ", "; //'wsd_rate
                        Decimal orghet = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                        Decimal orgso = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orgso + ", "; //'wsd_org_total_so_amount
                        strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                        strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                        strSQL = strSQL + cConv2 + ", "; //'wsd_uom_convert_big
                        //fConvertQty(dgvDetail2.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                        strSQL = strSQL + uomorg + ", "; //'wsd_uom_org_qty
                        strSQL = strSQL + "0, '";

                        strSQL = strSQL + "0" + "', '"; //'wsd_tot_disc_1
                        strSQL = strSQL + "0" + "', "; //'wsd_tot_disc_2
                        strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString() + "', '"; //wsd_tax_code
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString() + "',"; //wsd_tax_pct
                        strSQL = strSQL + "0, "; //wsd_tax_amount
                        strSQL = strSQL + "0, "; //wsd_gross_net
                        strSQL = strSQL + "0, "; //wsd_cash_disc
                        strSQL = strSQL + "'N', '"; //wsd_line_type
                        strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "', '";
                        strSQL = strSQL + sPriceCode + "', "; //wsd_price_code

                        if (dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                        {
                            strSQL = strSQL + "NULL"; //wsd_sled
                        }
                        else
                        {
                            strSQL = strSQL + "'"+ dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString().Trim() + "'"; //wsd_sled
                        }

                        strSQL = strSQL + ")"; //wsd_reason
                        dtFill15 = _clsGlobal.ExecDTTrans(strSQL);
                        #endregion
                    }


                }

                if (sCancel)
                {
                    strSQL = " EXEC SP_CANCEL_SO '" + sSONo + "','" + clsLogin.USERID + "','" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'";
                    _clsGlobal.ExecuteTrans(strSQL);

                    strSQL = " EXEC SP_CHECK_BALANCE_SO_MULTIBRANCH '1','" + sSONo + "', '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'";
                    _clsGlobal.ExecuteTrans(strSQL);

                    strSQL = "update A set A.wsd_reason = '" + reason + "' FROM SO_WEB_SALES_DETAIL A WITH (NOLOCK) where wsd_so_seq_no = '" + sSONo + "' and wsd_entity_id='" + cbEntityBranch.txtCM.Text.Trim() + "' and wsd_so_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'";
                    _clsGlobal.ExecuteTrans(strSQL);
                }

                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fSaveSOEcommerce(string bExtSono, string sCredLimit, bool bCalcDisc, int iRow)
        {
            bool keyFound = false;

            string sEntity = "";
            string sBranch = "";
            string sArea = "";
            string sWilayah = "";
            string sRayon = "";
            string sCustGroup = "";
            string sLineCode = "";
            string sYear = "";
            string sPeriode = "";
            //string sBranch = "";
            string sWeekNo = "";
            string orderType = "1101";
            string wh1 = "";
            string wh2 = "";
            string sSalesman = "";

            string sSOStatus = "";
            long lROrdQty = 0;
            decimal curROrdAmt = 0;
            decimal curCost = 0;
            decimal curTotalCost = 0;
            //clsPCode
            string sCurrCode = "";
            double dRateCurr = 0;
            long lBrs = 0;
            decimal curDisc1 = 0;
            decimal curDisc2 = 0;
            string sErrDesc = "";
            decimal currCredLimit = 0;
            //bool bFlagCr = false;
            string sPriceCode = "";
            string sOpenStatus = "";
            long lOrdQty = 0;
            long lLeadTime = 0;
            decimal curOrdAmt = 0;
            string pdaNO = "";
            string pdaFlag = "";
            string extsap = "";
            string runcext = "";
            string QasirFlag = "";
            string nik = "";
            DateTime sTglGudang;

            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cSled = "";
            string cLineCode = "";
            string cGroup = "";
            string cSubGroup = "";
            string cModel = "";
            string cRawMatUsed = "";
            string cWashingCollorCode = "";
            long cConv1 = 0;
            long cConv2 = 0;

            string sCustCode1;
            string sCustCode2;
            string sPONo;
            //string sSoNo;
            Dictionary<string, string> SOPrincipal;
            string qasirflag = "";

            string sCCOrderType = "";
            string sFlagColdChainOT = "N";
            string sFullfilment = "N";



            bool sCancel = false;
            string reason = "";
            // string sValidity = "";

            try
            {
                keyFound = false;

                sPONo = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString();
                sCustCode1 = dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString();
                sCustCode2 = dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value.ToString();
                reason = dgvHeader.Rows[iRow].Cells["g_iReasonRejection"].Value.ToString();
                sSalesman = dgvHeader.Rows[iRow].Cells["g_iSalesmanCode"].Value.ToString();
                sBranch = dgvHeader.Rows[iRow].Cells["g_iBranch"].Value.ToString();
                sEntity = dgvHeader.Rows[iRow].Cells["g_iEntity"].Value.ToString();



                if (dgvHeader.Rows[iRow].Cells["g_iValid"].Value.ToString() == "1" && dgvHeader.Rows[iRow].Cells["g_iRemark"].Value.ToString().ToLower() != "ok")
                {
                    sCancel = true;
                }

                if (!cbSource.SelectedValue.ToString().CompareC("M"))
                {
                    qasirflag = cbSource.SelectedValue.ToString().Trim();
                }
                else
                {
                    qasirflag = "N";
                }

                if (bCalcDisc == false)
                {
                    if (fGetSONumber() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    sSONo = g_sDefaultSONo;

                    strSQL = " EXEC SP_DELETE_SO_TABLE ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }

                /* ticket 27978 : ubah order type ikut dengan gs_gen_hardcoded  */
                strSQL = "";
                strSQL += "SELECT * FROM GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL += "WHERE gh_function_name = 'FLAGCOLDCHAIN_OT'";
                DataTable dtCheckColdChain = _clsGlobal.ExecDTTrans(strSQL);
                if (dtCheckColdChain.Rows.Count > 0)
                {
                    sFlagColdChainOT = dtCheckColdChain.Rows[0]["gh_function_code"].ToString();
                    if (sFlagColdChainOT.Trim().Equals("Y"))
                    {
                        sCCOrderType = dtCheckColdChain.Rows[0]["gh_min_seq_no"].ToString();
                        if (string.IsNullOrEmpty(sCCOrderType.Trim()))
                        {
                            MessageBox.Show("ColdChain Tipe Order : Data tipe order belum diisi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return keyFound;
                        }

                        strSQL = "select * from SO_ORDER_TYPE WITH (NOLOCK) where ot_order_type = '" + sCCOrderType + "' ";
                        DataTable dtCheckTipeOrder = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtCheckTipeOrder.Rows.Count == 0)
                        {
                            MessageBox.Show("ColdChain Tipe Order : Data tipe order tidak ada di master order type mohon di create terlebih dahulu.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return keyFound;

                        }

                        strSQL = "select * from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) where msow_sld_id = '" + sSalesman + "' and msow_entity_id = '" + sEntity + "' and msow_branch_id = '" + sBranch + "' and  msow_so_type = '" + sCCOrderType + "' AND msow_wh_loc1 +'|'+msow_wh_loc2 = (SELECT top 1 gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_function_name  LIKE 'MAINWHLOC' ) ";
                        DataTable dtCheckMAPSLD = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtCheckMAPSLD.Rows.Count == 0)
                        {
                            MessageBox.Show("ColdChain Tipe Order : Data tipe order tidak ada pada mapping salesman ke Tipe Order", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return keyFound;
                        }
                    }
                }

                if (sFlagColdChainOT.Trim().Equals("Y"))
                {
                    orderType = sCCOrderType;
                }


                /* end of ticket 27978 */


                strSQL = "";
                strSQL = "SELECT wh_loc_id1,wh_loc_id2,wh_loc_last_work_date FROM IM_WH_LOC WITH (NOLOCK) ";
                strSQL += "WHERE ";
                strSQL += "wh_loc_entity = '" + sEntity + "' ";
                strSQL += "AND wh_branch_id = '" + sBranch + "' ";
                strSQL += "AND wh_loc_id1 +'|'+wh_loc_id2 = (SELECT top 1 gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_function_name  LIKE 'MAINWHLOC' ) ";
                DataTable dtTglGudang = _clsGlobal.ExecDTTrans(strSQL);
                if (dtTglGudang.Rows.Count > 0)
                {
                    wh1 = dtTglGudang.Rows[0]["wh_loc_id1"].ToString();
                    wh2 = dtTglGudang.Rows[0]["wh_loc_id2"].ToString();
                    sTglGudang = Convert.ToDateTime(dtTglGudang.Rows[0]["wh_loc_last_work_date"].ToString());
                }
                else
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Tanggal Gudang tidak ditemukan!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }

                DataTable dtNIK = new DataTable();
                strSQL = "select sah_NIK from tbl_salesman_history WITH (NOLOCK) " +
                         " where" +
                         " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and sah_entity_id = '" + sEntity + "'" +
                         " and sah_branch_id = '" + sBranch + "'" +
                         " and sah_salesmen_id = '" + sSalesman + "'";
                dtNIK = _clsGlobal.ExecDTTrans(strSQL);

                if (dtNIK.Rows.Count == 1)
                {
                    nik = dtNIK.Rows[0]["sah_NIK"].ToString();
                }
                else if (dtNIK.Rows.Count > 1)
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }

                    MessageBox.Show("Salesman History Pada Periode Entry Lebih dari 1", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return keyFound;
                }
                else
                {

                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Salesman History Tidak Ditemukan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }



                //'Get branch, area, wilayah, rayon, cust group dan line code
                DataTable dtFill8 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') cm_line_code, cm_lead_time, isnull(cm_fullfilment,'N') cm_fullfilment From SO_CUST_MASTER WITH (NOLOCK)";
                strSQL = strSQL + " where cm_cust_code1 = '" + sCustCode1 + "' AND";
                strSQL = strSQL + " cm_cust_code2 = '" + sCustCode2 + "'  and cm_entity = '" + sEntity + "' and cm_branch = '" + sBranch + "'";
                dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill8.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill8.Rows[0]["cm_branch"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Area untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill8.Rows[0]["cm_area"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Wilayah untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill8.Rows[0]["cm_wilayah"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Rayon untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill8.Rows[0]["cm_rayon"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Cust group untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill8.Rows[0]["cm_cust_group"].ToString();
                    }

                    sLineCode = dtFill8.Rows[0]["cm_line_code"].ToString();
                    sFullfilment = dtFill8.Rows[0]["cm_fullfilment"].ToString();

                }
                else
                {
                    sBranch = "";
                    sArea = "";
                    sWilayah = "";
                    sRayon = "";
                    sCustGroup = "";
                    sLineCode = "";
                    sFullfilment = "N";
                }

                //'Get Year, Periode, Week untuk Tgl SO
                DataTable dtFill9 = new DataTable();
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK) ";
                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + Convert.ToDateTime(sTglGudang).ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '" + sEntity + "' and fwd_branch = '" + sBranch + "' "; //entity id harus di lempar
                dtFill9 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill9.Rows.Count > 0)
                {
                    sWeekNo = dtFill9.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill9.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill9.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + dtSODate.Value.ToString("dd MMM yyyy") + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get SO Status
                DataTable dtFill10 = new DataTable();
                strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                strSQL = strSQL + " and gh_sequence_no = '1' ";
                dtFill10 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill10.Rows.Count > 0)
                {
                    sSOStatus = dtFill10.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    MessageBox.Show("Status SO (Open) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvDetail.Rows.Count; lCounter++)
                {
                    if (
                        sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString()
                        && sEntity == dgvDetail.Rows[lCounter].Cells["g_iDetEntity"].Value.ToString()
                        && sBranch == dgvDetail.Rows[lCounter].Cells["g_iDetBranch"].Value.ToString()
                        && sSalesman == dgvDetail.Rows[lCounter].Cells["g_iDetSalesman"].Value.ToString()
                        && sCustCode1 == dgvDetail.Rows[lCounter].Cells["g_iDetCustNo"].Value.ToString()
                        && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != "")
                    {
                        //fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + (int)Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curROrdAmt = curROrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            curROrdAmt = curROrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);
                        }

                        fConvertQty(dgvDetail.Rows[lCounter].Cells["g_iDetQtyOrd"].Value.ToString(), Convert.ToInt16(iConv1), Convert.ToInt16(iConv2));
                        lOrdQty = lOrdQty + Convert.ToInt32(_fConvertQty);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curOrdAmt = curOrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);

                        }
                        else
                        {
                            curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);

                        }

                        if (fGetCostEcommerce(sEntity, sBranch, dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetGrade"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetSize"].Value.ToString(), wh1, wh2) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        curTotalCost = curTotalCost + _curCost;
                    }
                }

                //if (sSONo.Length == 0)
                //{
                //    MessageBox.Show("nomor SO Blank", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //if (_xFlagDC == "N")
                //{
                // '===> Cek Kredit Limit dihilangkan jika flagDC = Y - Project Improvement TiraSnD
                //'/*-----------------------------------------------------------------------------------------
                //'Credit Limit
                if (fGetCreditLimit(sCustCode1, sCustCode2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (bFlagCr == true)
                {
                    if (_curAvCredLimit < curROrdAmt)
                    {
                        sCredLimit = "";
                    }
                    else
                    {
                        sCredLimit = "";
                    }
                }
                //}

                DataTable dtFill4 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_DETAIL A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsd_so_seq_no = '" + sSONo + "'";
                //strSQL = strSQL + "  AND wsh_branch_id = '" + cbBranchID.SelectedValue.ToString()  + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill5 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_HEADER A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "'";
                strSQL = strSQL + "  AND wsh_branch_id = '" + sBranch + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill12 = new DataTable();
                #region Query Insert
                strSQL = "Insert into SO_WEB_SALES_HEADER (";
                strSQL = strSQL + " wsh_entity_id ,wsh_area ,wsh_wilayah ,wsh_rayon ";
                strSQL = strSQL + " ,wsh_branch_id ,wsh_cust_group ,wsh_so_type ,wsh_seq_no";
                strSQL = strSQL + " ,wsh_so_date,wsh_cust_code1 ,wsh_cust_code2 ,wsh_loc_id1";
                strSQL = strSQL + " ,wsh_loc_id2,wsh_brand_line ,wsh_so_ref_no,wsh_so_txn_date_from";
                strSQL = strSQL + " ,wsh_so_txn_date_to,wsh_week_no,wsh_spgm_id ,wsh_spgm_group";
                strSQL = strSQL + " ,wsh_spgm_coordinator_id ,wsh_spgm_loc1 ,wsh_spgm_loc2 ,wsh_shift_no";
                strSQL = strSQL + " ,[wsh_participation_disc_%] ,wsh_participation_disc_amount ,wsh_tot_part_line_amount ,wsh_total_line";
                strSQL = strSQL + " , wsh_relevan_pod ,wsh_status_so ,wsh_real_order_qty   ,wsh_real_order_amt";
                strSQL = strSQL + " ,wsh_total_so_qty ,wsh_total_so_amt      ,wsh_include_tax ,[wsh_tax_%]";
                strSQL = strSQL + " ,wsh_total_partisipasi_disc_amount ,wsh_total_margin_disc_amount ,wsh_net_amount,wsh_dpp_amount";
                strSQL = strSQL + " ,wsh_tax_amount,wsh_invoice_process_flag ,wsh_invoice_no ,wsh_invoice_date";
                strSQL = strSQL + " ,wsh_invoice_process_by ,wsh_disc_margin,wsh_total_cost,wsh_year";
                strSQL = strSQL + " ,wsh_period,wsh_cancel_date,wsh_cancel_by,wsh_approval";
                strSQL = strSQL + " ,wsh_approval_by ,wsh_approval_date,wsh_ret_seq_no ,wsh_ret_by ,wsh_remarks";
                strSQL = strSQL + " ,wsh_participation_code ,wsh_entry_date,wsh_user_id ,wsh_from_OB_flag";
                strSQL = strSQL + " ,wsh_gross_net,wsh_gen_flag,wsh_tot_disc_pc,wsh_tot_disc_1";
                strSQL = strSQL + " ,wsh_tot_disc_2,wsh_tot_tpr_amt,wsh_tax_code,wsh_tax_pct";
                strSQL = strSQL + " ,wsh_pct_cash_disc, wsh_cash_disc, wsh_tax_pbm_pct, wsh_tax_pbm_amount, wsh_pay_type";
                strSQL = strSQL + " ,wsh_gross_net_amt, wsh_validate, wsh_po_no ";
                strSQL = strSQL + " ,wsh_flag_so_limit ";
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty,wsh_po_ol_flag ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                strSQL = strSQL + " , wsh_extract_sap ";
                strSQL = strSQL + " , wsh_runcode_ext  ";
                strSQL = strSQL + " , wsh_sch_date  ";
                strSQL = strSQL + " , wsh_expired_date  ";
                strSQL = strSQL + " , wsh_nik  ";
                strSQL = strSQL + " , wsh_qasir_flag  ";
                strSQL = strSQL + " , wsh_beban_cashdisc  ";
                if (isMultiSource)
                {
                    strSQL = strSQL + " , wsh_flag_sloc ";
                }
                strSQL = strSQL + ", wsh_is_fulfillment ";
                strSQL = strSQL + ", wsh_sled_flag ";
                string __tglorder = sTglGudang.ToString("yyyyMMdd");
                //int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + sEntity + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + sBranch + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                strSQL = strSQL + orderType + "' , '";  //'wsh_so_type
                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_so_date
                }
                else
                {
                    strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                }
                strSQL = strSQL + sCustCode1 + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + sCustCode2 + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + wh1 + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + wh2 + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + sSalesman + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "null , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "null , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "null , "; //'wsh_tot_part_line_amount
                strSQL = strSQL + "0, "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                strSQL = strSQL + "null , "; //'wsh_include_tax
                strSQL = strSQL + "null , "; //'[wsh_tax_%]
                strSQL = strSQL + "null , "; //'wsh_total_partisipasi_disc_amount
                strSQL = strSQL + "null , "; //'wsh_total_margin_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_net_amount
                strSQL = strSQL + "0 , "; //'wsh_dpp_amount
                strSQL = strSQL + "0 , "; //'wsh_tax_amount
                strSQL = strSQL + "'N' , "; //'wsh_invoice_process_flag
                strSQL = strSQL + "null , "; //'wsh_invoice_no
                strSQL = strSQL + "null , "; //'wsh_invoice_date
                strSQL = strSQL + "null , "; //'wsh_invoice_process_by
                strSQL = strSQL + "null , "; //'wsh_disc_margin
                strSQL = strSQL + curTotalCost + ", '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , "; //'wsh_approval
                strSQL = strSQL + "null , "; //'wsh_approval_by
                strSQL = strSQL + "null , "; //'wsh_approval_date
                strSQL = strSQL + "null , "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                //strSQL = strSQL + "null , "; //'wsh_participation_code
                if (cbSource.SelectedValue.ToString().Trim().CompareC("L") || cbSource.SelectedValue.ToString().Trim().CompareC("S") || cbSource.SelectedValue.ToString().Trim().CompareC("W"))
                {
                    if (string.IsNullOrEmpty(dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value.ToString()))
                    {
                        strSQL = strSQL + "null , "; //'wsh_participation_code
                    }
                    else
                    {
                        strSQL = strSQL + "'" + dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value.ToString() + "' , "; //'wsh_participation_code
                    }
                }
                else
                {
                    strSQL = strSQL + "null , "; //'wsh_participation_code
                }
                strSQL = strSQL + "getdate() , '"; //'wsh_entry_date
                strSQL = strSQL + clsLogin.USERID + "', "; //'wsh_user_id
                strSQL = strSQL + "null , "; //'wsh_from_OB_flag
                strSQL = strSQL + "null , "; //'wsh_gross_net
                strSQL = strSQL + "null , "; //'wsh_gen_flag
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_pc
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_1
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_2
                strSQL = strSQL + "0 , "; //'wsh_tot_tpr_amt
                strSQL = strSQL + "null , "; //'wsh_tax_code
                strSQL = strSQL + "null , "; //'wsh_tax_pct
                strSQL = strSQL + "0" + " , ";
                strSQL = strSQL + "0 , "; //'wsh_cash_disc
                strSQL = strSQL + "0 , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "0 , '"; //'wsh_tax_pbm_amount
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value.ToString() + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , "; //'wsh_gross_net_amt
                strSQL = strSQL + "null , '"; // Convert.ToDateTime(dtTglValidasi.Value).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + sPONo + "' , '"; //'wsh_po_no
                strSQL = strSQL + sCredLimit + "' , '"; //'wsh_flag_so_limit
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                strSQL = strSQL + "N' , '"; //'wsh_po_ol_flag
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iTOP"].Value.ToString() + "' , '"; //'wsh_top
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iTOPId"].Value.ToString() + "' , '"; //'wsh_top_id
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_po_date
                }
                else
                {
                    string __poDate = dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString();
                    strSQL = strSQL + __poDate.Substring(6, 4) + __poDate.Substring(3, 2) + __poDate.Substring(0, 2) + "','"; //'wsh_po_date
                }
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value.ToString() + "'"; //'wsh_cust_leadtime
                strSQL = strSQL + ",'" + "N" + "' , '"; //'extsap
                strSQL = strSQL + "'"; //'runcext
                string __delvdate = dgvHeader.Rows[iRow].Cells["g_iDeliveryDate"].Value.ToString();
                string __expdate = dgvHeader.Rows[iRow].Cells["g_iExpiredDate"].Value.ToString();
                strSQL = strSQL + ",'" + __delvdate.Substring(6, 4) + __delvdate.Substring(3, 2) + __delvdate.Substring(0, 2) + "'"; // planning date
                strSQL = strSQL + ",'" + __expdate.Substring(6, 4) + __expdate.Substring(3, 2) + __expdate.Substring(0, 2) + "'"; // expired date
                strSQL = strSQL + ",'" + nik + "'"; // nik
                strSQL = strSQL + ",'" + qasirflag + "'"; // qasirflag 30-06-2021 -- update BL : kalo M kosong kalo yang lain diisi sama dengan cbSource.selectedvalue
                strSQL = strSQL + ",0"; //wsh_beban_cashdisc
                if (isMultiSource)
                {
                    strSQL = strSQL + ",'" + cbFlagSloc.SelectedValue.ToString() + "'";
                }
                strSQL = strSQL + ",'" + sFullfilment + "'"; //wsh_is_fulfillment
                if (dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                {
                    strSQL = strSQL + ",NULL"; // wsd_sled_flag
                }
                else
                {
                    strSQL = strSQL + ",'"+ dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().Trim() + "'"; // wsd_sled_flag
                }


                strSQL = strSQL + " ) ";
                dtFill12 = _clsGlobal.ExecDTTrans(strSQL);
                //'END OF Saving SO Header
                //if (dtFill12.Rows.Count > 0)
                //{
                DataTable dtFill13 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK) ";
                strSQL = strSQL + " where cc_currency_type = 'H' ";
                #endregion

                dtFill13 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill13.Rows.Count > 0)
                {
                    sCurrCode = dtFill13.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill13.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }

                lCounter = 0;
                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                {

                    if (
                        sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString()
                        && sEntity == dgvDetail2.Rows[lCounter].Cells["g_iDetEntity1"].Value.ToString()
                        && sBranch == dgvDetail2.Rows[lCounter].Cells["g_iDetBranch1"].Value.ToString()
                        && sSalesman == dgvDetail2.Rows[lCounter].Cells["g_iDetSalesmanID1"].Value.ToString()
                        && sCustCode1 == dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value.ToString()
                        )
                    {
                        #region save detail
                        DataTable dtFill14 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK)";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "'";
                        dtFill14 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill14.Rows.Count > 0)
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = dtFill14.Rows[0]["prm_prd_line_code"].ToString();
                            cGroup = dtFill14.Rows[0]["prm_prd_group_code"].ToString();
                            cSubGroup = dtFill14.Rows[0]["prm_prd_sgroup_code"].ToString();
                            cModel = dtFill14.Rows[0]["prm_prd_model_code"].ToString();
                            cRawMatUsed = dtFill14.Rows[0]["prm_raw_mat_used_code"].ToString();
                            cWashingCollorCode = dtFill14.Rows[0]["prm_washing_collor_code"].ToString();
                            cConv1 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_purc"].ToString());
                            cConv2 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_sales"].ToString());
                        }
                        else
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = "";
                            cGroup = "";
                            cSubGroup = "";
                            cModel = "";
                            cRawMatUsed = "";
                            cWashingCollorCode = "";
                            cConv1 = 0;
                            cConv2 = 0;
                        }
                        //Get Price Code
                        DataTable dtFillGrpPrice = new DataTable();
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + sEntity + "','" + sBranch + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + sCustCode1 + "','" + sCustCode2 + "','" + __tglorder + "'";
                        dtFillGrpPrice = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFillGrpPrice.Rows.Count > 0)
                        {
                            if (String.IsNullOrEmpty(dtFillGrpPrice.Rows[0]["cp_price_group"].ToString()) == true)
                            {
                                MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                            else
                            {
                                sPriceCode = dtFillGrpPrice.Rows[0]["cp_price_group"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }


                        DataTable dtFill15 = new DataTable();
                        strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                        strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                        strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                        strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                        strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                        strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                        strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                        strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                        strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                        strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_sled)";

                        strSQL = strSQL + " values ( '";
                        strSQL = strSQL + sEntity + "' , '"; //'wsd_entity_id
                        strSQL = strSQL + orderType + "' , '"; //'wsd_so_type
                        strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                        int lineno = lCounter + 1;
                        strSQL = strSQL + lineno + "', '"; //'wsd_so_detail_line_no
                        strSQL = strSQL + sBranch + "','"; //wsd_so_branch_id NEW
                        strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                        strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                        strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                        strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                        strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                        strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                        strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                        strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                        strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', '";
                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";
                        decimal OrderAmt = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString());
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            strSQL = strSQL + Math.Round(OrderAmt, 0, MidpointRounding.AwayFromZero) + "', '"; //'wsd_real_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + OrderAmt + "', '"; //'wsd_real_order_amt
                        }

                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";//'wsd_qty_sales
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', "; //'wsd_total_so_amount
                        strSQL = strSQL + "null, "; //'wsd_shipped_xqty
                        strSQL = strSQL + "null, "; //'wsd_shipped_qty
                        strSQL = strSQL + "null, "; //'wsd_shipped_amount
                        strSQL = strSQL + "null, "; //'wsd_pod_xqty
                        strSQL = strSQL + "null, "; //'wsd_pod_qty
                        strSQL = strSQL + "null, "; //'wsd_pod_amount
                        strSQL = strSQL + "null, "; //'wsd_invoice_xqty
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty
                        strSQL = strSQL + "null, "; //'wsd_invoice_amount
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_%
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_amt
                        strSQL = strSQL + "null, '"; //'wsd_margin

                        if (fGetCostEcommerce(sEntity, sBranch, dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString(), wh1, wh2) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        strSQL = strSQL + _curCost.ToString() + "', '"; //'wsd_unit_cost
                        fConvertQty(dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString(), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value));
                        //Rounding Mekanism
                        decimal ttlcst = _curCost * Convert.ToDecimal(_fConvertQty);
                        if (isRoundingMekanism)
                        {
                            ttlcst = Math.Round(_curCost * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                        }
                        strSQL = strSQL + ttlcst.ToString() + "', "; //'wsd_total_cost
                        strSQL = strSQL + "null, "; //'wsd_remarks
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_flag
                        strSQL = strSQL + "null, "; //'wsd_invoice_no
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty_sales
                        strSQL = strSQL + "null, "; //'wsd_invoice_date
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_by
                        strSQL = strSQL + "getdate() , '"; //'wsd_entry_date
                        strSQL = strSQL + clsLogin.USERID + "', '"; //wsd_user_id
                        strSQL = strSQL + sCurrCode + "', "; //'wsd_curr_code
                        strSQL = strSQL + dRateCurr + ", "; //'wsd_rate
                        Decimal orghet = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                        Decimal orgso = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orgso + ", "; //'wsd_org_total_so_amount
                        strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                        strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                        strSQL = strSQL + cConv2 + ", "; //'wsd_uom_convert_big
                        //fConvertQty(dgvDetail2.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                        strSQL = strSQL + uomorg + ", "; //'wsd_uom_org_qty
                        strSQL = strSQL + "0, '";

                        strSQL = strSQL + "0" + "', '"; //'wsd_tot_disc_1
                        strSQL = strSQL + "0" + "', "; //'wsd_tot_disc_2
                        strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString() + "', '"; //wsd_tax_code
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString() + "',"; //wsd_tax_pct
                        strSQL = strSQL + "0, "; //wsd_tax_amount
                        strSQL = strSQL + "0, "; //wsd_gross_net
                        strSQL = strSQL + "0, "; //wsd_cash_disc
                        strSQL = strSQL + "'N', '"; //wsd_line_type
                        strSQL = strSQL + sTglGudang.ToString("yyyy-MM-dd") + "', '";
                        strSQL = strSQL + sPriceCode + "', "; //wsd_price_code
                        if (dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                        {
                            strSQL = strSQL + "NULL"; //wsd_sled
                        }
                        else
                        {
                            strSQL = strSQL + "'"+ dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString() + "'"; //wsd_sled
                        }
                        
                        strSQL = strSQL + ")"; //wsd_reason
                        dtFill15 = _clsGlobal.ExecDTTrans(strSQL);
                        #endregion
                    }


                }
                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fSaveSO_KAM(string bExtSono, string sCredLimit, bool bCalcDisc, int iRow, string TOP, string ValueTOP, string _sProdLine)
        {
            bool keyFound = false;

            string sBranch = "";
            string sArea = "";
            string sWilayah = "";
            string sRayon = "";
            string sCustGroup = "";
            string sLineCode = "";
            string sYear = "";
            string sPeriode = "";
            //string sBranch = "";
            string sWeekNo = "";

            string sSOStatus = "";
            long lROrdQty = 0;
            decimal curROrdAmt = 0;
            decimal curCost = 0;
            decimal curTotalCost = 0;
            //clsPCode
            string sCurrCode = "";
            double dRateCurr = 0;
            long lBrs = 0;
            decimal curDisc1 = 0;
            decimal curDisc2 = 0;
            string sErrDesc = "";
            decimal currCredLimit = 0;
            //bool bFlagCr = false;
            string sPriceCode = "";
            string sOpenStatus = "";
            long lOrdQty = 0;
            long lLeadTime = 0;
            decimal curOrdAmt = 0;
            string pdaNO = "";
            string pdaFlag = "";
            string extsap = "";
            string runcext = "";
            string QasirFlag = "";
            string nik = "";

            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cSled = "";
            string cLineCode = "";
            string cGroup = "";
            string cSubGroup = "";
            string cModel = "";
            string cRawMatUsed = "";
            string cWashingCollorCode = "";
            long cConv1 = 0;
            long cConv2 = 0;

            string sCustCode1;
            string sCustCode2;
            string sPONo;
            //string sSoNo;
            Dictionary<string, string> SOPrincipal;

            string sFullfilment = "N";

            try
            {
                keyFound = false;

                sPONo = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString();
                sCustCode1 = dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString();
                sCustCode2 = dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value.ToString();

                if (bCalcDisc == false)
                {
                    if (fGetSONumber() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    sSONo = g_sDefaultSONo;

                    strSQL = " EXEC SP_DELETE_SO_TABLE ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }


                DataTable dtNIK = new DataTable();
                strSQL = "select sah_NIK from tbl_salesman_history WITH (NOLOCK) " +
                         " where" +
                         " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and sah_entity_id = '" + cbEntityBranch.txtCM.Text.Trim() + "'" +
                         " and sah_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'" +
                         " and sah_salesmen_id = '" + txtSalesman.Text + "'";
                dtNIK = _clsGlobal.ExecDTTrans(strSQL);

                if (dtNIK.Rows.Count == 1)
                {
                    nik = dtNIK.Rows[0]["sah_NIK"].ToString();
                }
                else if (dtNIK.Rows.Count > 1)
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }

                    MessageBox.Show("Salesman History Pada Periode Entry Lebih dari 1", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return keyFound;
                }
                else
                {

                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Salesman History Tidak Ditemukan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }



                //'Get branch, area, wilayah, rayon, cust group dan line code
                DataTable dtFill8 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') cm_line_code, cm_lead_time, ISNULL(cm_fullfilment,'N') cm_fullfilment From SO_CUST_MASTER WITH (NOLOCK)";
                strSQL = strSQL + " where cm_cust_code1 = '" + sCustCode1 + "' AND";
                strSQL = strSQL + " cm_cust_code2 = '" + sCustCode2 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text.Trim() + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' ";
                dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill8.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill8.Rows[0]["cm_branch"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Area untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill8.Rows[0]["cm_area"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Wilayah untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill8.Rows[0]["cm_wilayah"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Rayon untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill8.Rows[0]["cm_rayon"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Cust group untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill8.Rows[0]["cm_cust_group"].ToString();
                    }

                    sLineCode = dtFill8.Rows[0]["cm_line_code"].ToString();
                    sFullfilment = dtFill8.Rows[0]["cm_fullfilment"].ToString();

                    //if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_lead_time"].ToString()))
                    //{
                    //    MessageBox.Show("Lead Time untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}
                    //else
                    //{
                    //    lLeadTime = Convert.ToInt64(dtFill8.Rows[0]["cm_lead_time"].ToString());
                    //}

                }
                else
                {
                    sBranch = "";
                    sArea = "";
                    sWilayah = "";
                    sRayon = "";
                    sCustGroup = "";
                    sLineCode = "";
                    sFullfilment = "N";
                }

                //'Get Year, Periode, Week untuk Tgl SO
                DataTable dtFill9 = new DataTable();
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK) ";
                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + dtSODate.Value.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '" + cbEntityBranch.txtCM.Text.Trim() + "' and fwd_branch = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'"; //entity id harus di lempar
                dtFill9 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill9.Rows.Count > 0)
                {
                    sWeekNo = dtFill9.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill9.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill9.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + dtSODate.Value.ToString("dd MMM yyyy") + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get SO Status
                DataTable dtFill10 = new DataTable();
                strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                strSQL = strSQL + " and gh_sequence_no = '1' ";
                dtFill10 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill10.Rows.Count > 0)
                {
                    sSOStatus = dtFill10.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    MessageBox.Show("Status SO (Open) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


                //Get Price Code
                //DataTable dtFill11 = new DataTable();
                //strSQL = " exec SP_GET_PRICE_CODE '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + sCustCode1 + "','" + sCustCode2 + "'";
                //dtFill11 = _clsGlobal.ExecDTTrans(strSQL);
                //if (dtFill11.Rows.Count > 0)
                //{
                //    if (String.IsNullOrEmpty(dtFill11.Rows[0]["cm_def_price_code"].ToString()))
                //    {
                //        MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        keyFound = false;
                //        return keyFound;
                //    }
                //    else
                //    {
                //        sPriceCode = dtFill11.Rows[0]["cm_def_price_code"].ToString();
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvDetail.Rows.Count; lCounter++)
                {
                    if (sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString() && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != "" && _sProdLine.Trim() == dgvDetail.Rows[lCounter].Cells["g_iDetProdLine"].Value.ToString().Trim())
                    {
                        //fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + (int)Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curROrdAmt = curROrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            curROrdAmt = curROrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);
                        }

                        fConvertQty(dgvDetail.Rows[lCounter].Cells["g_iDetQtyOrd"].Value.ToString(), Convert.ToInt16(iConv1), Convert.ToInt16(iConv2));
                        lOrdQty = lOrdQty + Convert.ToInt32(_fConvertQty);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curOrdAmt = curOrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);

                        }
                        else
                        {
                            curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);

                        }

                        if (fGetCost(dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetGrade"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetSize"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        curTotalCost = curTotalCost + _curCost;
                    }
                }

                //if (sSONo.Length == 0)
                //{
                //    MessageBox.Show("nomor SO Blank", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //if (_xFlagDC == "N")
                //{
                // '===> Cek Kredit Limit dihilangkan jika flagDC = Y - Project Improvement TiraSnD
                //'/*-----------------------------------------------------------------------------------------
                //'Credit Limit
                if (fGetCreditLimit(sCustCode1, sCustCode2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (bFlagCr == true)
                {
                    if (_curAvCredLimit < curROrdAmt)
                    {
                        sCredLimit = "";
                    }
                    else
                    {
                        sCredLimit = "";
                    }
                }
                //}

                DataTable dtFill4 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_DETAIL A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsd_so_seq_no = '" + sSONo + "'";
                //strSQL = strSQL + "  AND wsh_branch_id = '" + cbBranchID.SelectedValue.ToString()  + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill5 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_HEADER A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "'";
                strSQL = strSQL + "  AND wsh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill12 = new DataTable();
                #region Query Insert
                strSQL = "Insert into SO_WEB_SALES_HEADER (";
                strSQL = strSQL + " wsh_entity_id ,wsh_area ,wsh_wilayah ,wsh_rayon ";
                strSQL = strSQL + " ,wsh_branch_id ,wsh_cust_group ,wsh_so_type ,wsh_seq_no";
                strSQL = strSQL + " ,wsh_so_date,wsh_cust_code1 ,wsh_cust_code2 ,wsh_loc_id1";
                strSQL = strSQL + " ,wsh_loc_id2,wsh_brand_line ,wsh_so_ref_no,wsh_so_txn_date_from";
                strSQL = strSQL + " ,wsh_so_txn_date_to,wsh_week_no,wsh_spgm_id ,wsh_spgm_group";
                strSQL = strSQL + " ,wsh_spgm_coordinator_id ,wsh_spgm_loc1 ,wsh_spgm_loc2 ,wsh_shift_no";
                strSQL = strSQL + " ,[wsh_participation_disc_%] ,wsh_participation_disc_amount ,wsh_tot_part_line_amount ,wsh_total_line";
                strSQL = strSQL + " , wsh_relevan_pod ,wsh_status_so ,wsh_real_order_qty   ,wsh_real_order_amt";
                strSQL = strSQL + " ,wsh_total_so_qty ,wsh_total_so_amt      ,wsh_include_tax ,[wsh_tax_%]";
                strSQL = strSQL + " ,wsh_total_partisipasi_disc_amount ,wsh_total_margin_disc_amount ,wsh_net_amount,wsh_dpp_amount";
                strSQL = strSQL + " ,wsh_tax_amount,wsh_invoice_process_flag ,wsh_invoice_no ,wsh_invoice_date";
                strSQL = strSQL + " ,wsh_invoice_process_by ,wsh_disc_margin,wsh_total_cost,wsh_year";
                strSQL = strSQL + " ,wsh_period,wsh_cancel_date,wsh_cancel_by,wsh_approval";
                strSQL = strSQL + " ,wsh_approval_by ,wsh_approval_date,wsh_ret_seq_no ,wsh_ret_by ,wsh_remarks";
                strSQL = strSQL + " ,wsh_participation_code ,wsh_entry_date,wsh_user_id ,wsh_from_OB_flag";
                strSQL = strSQL + " ,wsh_gross_net,wsh_gen_flag,wsh_tot_disc_pc,wsh_tot_disc_1";
                strSQL = strSQL + " ,wsh_tot_disc_2,wsh_tot_tpr_amt,wsh_tax_code,wsh_tax_pct";
                strSQL = strSQL + " ,wsh_pct_cash_disc, wsh_cash_disc, wsh_tax_pbm_pct, wsh_tax_pbm_amount, wsh_pay_type";
                strSQL = strSQL + " ,wsh_gross_net_amt, wsh_validate, wsh_po_no ";
                strSQL = strSQL + " ,wsh_flag_so_limit ";
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty,wsh_po_ol_flag ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                strSQL = strSQL + " , wsh_extract_sap ";
                strSQL = strSQL + " , wsh_runcode_ext  ";
                strSQL = strSQL + " , wsh_sch_date  ";
                strSQL = strSQL + " , wsh_expired_date  ";
                strSQL = strSQL + " , wsh_nik  ";
                strSQL = strSQL + " , wsh_is_fulfillment  ";
                strSQL = strSQL + " , wsh_sled_flag  ";
                string __tglorder = dtSODate.Value.ToString("yyyyMMdd");
                //int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '";  //'wsh_so_type
                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no
                strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                strSQL = strSQL + sCustCode1 + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + sCustCode2 + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + lblWHLoc1.Text + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + lblWHLoc2.Text + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + txtSalesman.Text + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "null , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "null , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "null , "; //'wsh_tot_part_line_amount
                strSQL = strSQL + "0, "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                strSQL = strSQL + "null , "; //'wsh_include_tax
                strSQL = strSQL + "null , "; //'[wsh_tax_%]
                strSQL = strSQL + "null , "; //'wsh_total_partisipasi_disc_amount
                strSQL = strSQL + "null , "; //'wsh_total_margin_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_net_amount
                strSQL = strSQL + "0 , "; //'wsh_dpp_amount
                strSQL = strSQL + "0 , "; //'wsh_tax_amount
                strSQL = strSQL + "'N' , "; //'wsh_invoice_process_flag
                strSQL = strSQL + "null , "; //'wsh_invoice_no
                strSQL = strSQL + "null , "; //'wsh_invoice_date
                strSQL = strSQL + "null , "; //'wsh_invoice_process_by
                strSQL = strSQL + "null , "; //'wsh_disc_margin
                strSQL = strSQL + curTotalCost + ", '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , "; //'wsh_approval
                strSQL = strSQL + "null , "; //'wsh_approval_by
                strSQL = strSQL + "null , "; //'wsh_approval_date
                strSQL = strSQL + "null , "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                strSQL = strSQL + "null , "; //'wsh_participation_code
                strSQL = strSQL + "getdate() , '"; //'wsh_entry_date
                strSQL = strSQL + clsLogin.USERID + "', "; //'wsh_user_id
                strSQL = strSQL + "null , "; //'wsh_from_OB_flag
                strSQL = strSQL + "null , "; //'wsh_gross_net
                strSQL = strSQL + "null , "; //'wsh_gen_flag
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_pc
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_1
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_2
                strSQL = strSQL + "0 , "; //'wsh_tot_tpr_amt
                strSQL = strSQL + "null , "; //'wsh_tax_code
                strSQL = strSQL + "null , "; //'wsh_tax_pct

                if (dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() == "")
                {
                    strSQL = strSQL + "0" + " , ";
                }
                else
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                }

                strSQL = strSQL + "0 , "; //'wsh_cash_disc
                //strSQL = strSQL + "0 , "; //'wsh_cash_disc
                strSQL = strSQL + "0 , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "0 , '"; //'wsh_tax_pbm_amount
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value.ToString() + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , "; //'wsh_gross_net_amt
                strSQL = strSQL + "null , '"; // Convert.ToDateTime(dtTglValidasi.Value).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + sPONo + "' , '"; //'wsh_po_no
                strSQL = strSQL + sCredLimit + "' , '"; //'wsh_flag_so_limit
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                strSQL = strSQL + "N' , '"; //'wsh_po_ol_flag
                strSQL = strSQL + ValueTOP + "' , '"; //'wsh_top
                strSQL = strSQL + TOP + "' , '"; //'wsh_top_id
                strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "','"; //'wsh_po_date
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value.ToString() + "'"; //'wsh_cust_leadtime
                strSQL = strSQL + ",'" + "N" + "' , '"; //'extsap
                strSQL = strSQL + "'"; //'runcext
                strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // planning date
                strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // expired date
                strSQL = strSQL + ",'" + nik + "'"; // nik
                strSQL = strSQL + ",'" + sFullfilment + "'"; // nik
                if (dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                {
                    strSQL = strSQL + ",'N'"; //'wsh_sled_flag
                }
                else
                {
                    strSQL = strSQL + "'"+ dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString() + "'"; //'wsh_sled_flag
                }           
                //Convert.ToDateTime(dtTglOrder.Text).ToString("yyyyMMdd")
                strSQL = strSQL + " ) ";
                dtFill12 = _clsGlobal.ExecDTTrans(strSQL);
                //'END OF Saving SO Header
                //if (dtFill12.Rows.Count > 0)
                //{
                DataTable dtFill13 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK) ";
                strSQL = strSQL + " where cc_currency_type = 'H' ";
                #endregion

                dtFill13 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill13.Rows.Count > 0)
                {
                    sCurrCode = dtFill13.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill13.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }


                int lineNo = 0;
                lCounter = 0;
                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                {

                    if (sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString())
                    {
                        #region save detail
                        DataTable dtFill14 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK)";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "'";
                        dtFill14 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill14.Rows.Count > 0)
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = dtFill14.Rows[0]["prm_prd_line_code"].ToString();
                            cGroup = dtFill14.Rows[0]["prm_prd_group_code"].ToString();
                            cSubGroup = dtFill14.Rows[0]["prm_prd_sgroup_code"].ToString();
                            cModel = dtFill14.Rows[0]["prm_prd_model_code"].ToString();
                            cRawMatUsed = dtFill14.Rows[0]["prm_raw_mat_used_code"].ToString();
                            cWashingCollorCode = dtFill14.Rows[0]["prm_washing_collor_code"].ToString();
                            cConv1 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_purc"].ToString());
                            cConv2 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_sales"].ToString());
                        }
                        else
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = "";
                            cGroup = "";
                            cSubGroup = "";
                            cModel = "";
                            cRawMatUsed = "";
                            cWashingCollorCode = "";
                            cConv1 = 0;
                            cConv2 = 0;
                        }
                        //Get Price Code
                        DataTable dtFillGrpPrice = new DataTable();
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + sCustCode1 + "','" + sCustCode2 + "','" + __tglorder + "'";
                        dtFillGrpPrice = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFillGrpPrice.Rows.Count > 0)
                        {
                            if (String.IsNullOrEmpty(dtFillGrpPrice.Rows[0]["cp_price_group"].ToString()) == true)
                            {
                                MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                            else
                            {
                                sPriceCode = dtFillGrpPrice.Rows[0]["cp_price_group"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        if (cLineCode == _sProdLine)
                        {


                            DataTable dtFill15 = new DataTable();
                            strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                            strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                            strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                            strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                            strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                            strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                            strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                            strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                            strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                            strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_sled)";

                            strSQL = strSQL + " values ( '";
                            strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '"; //'wsd_entity_id
                            strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '"; //'wsd_so_type
                            strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                            int lineno = lineNo + 1;
                            strSQL = strSQL + lineno + "', '"; //'wsd_so_detail_line_no
                            strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','"; //wsd_so_branch_id NEW
                            strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                            strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                            strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                            strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                            strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                            strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                            strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                            strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                            strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                            strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', '";
                            strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                            strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";
                            decimal OrderAmt = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString());
                            //Rounding Mekanism
                            if (isRoundingMekanism)
                            {
                                strSQL = strSQL + Math.Round(OrderAmt, 0, MidpointRounding.AwayFromZero) + "', '"; //'wsd_real_order_amt
                            }
                            else
                            {
                                strSQL = strSQL + OrderAmt + "', '"; //'wsd_real_order_amt
                            }

                            strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                            strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";//'wsd_qty_sales
                            strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', "; //'wsd_total_so_amount
                            strSQL = strSQL + "null, "; //'wsd_shipped_xqty
                            strSQL = strSQL + "null, "; //'wsd_shipped_qty
                            strSQL = strSQL + "null, "; //'wsd_shipped_amount
                            strSQL = strSQL + "null, "; //'wsd_pod_xqty
                            strSQL = strSQL + "null, "; //'wsd_pod_qty
                            strSQL = strSQL + "null, "; //'wsd_pod_amount
                            strSQL = strSQL + "null, "; //'wsd_invoice_xqty
                            strSQL = strSQL + "null, "; //'wsd_invoice_qty
                            strSQL = strSQL + "null, "; //'wsd_invoice_amount
                            strSQL = strSQL + "null, "; //'wsd_participation_disc_%
                            strSQL = strSQL + "null, "; //'wsd_participation_disc_amt
                            strSQL = strSQL + "null, '"; //'wsd_margin

                            if (fGetCost(dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                            {
                                //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            strSQL = strSQL + _curCost.ToString() + "', '"; //'wsd_unit_cost
                            fConvertQty(dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString(), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value));
                            //Rounding Mekanism
                            decimal ttlcst = _curCost * Convert.ToDecimal(_fConvertQty);
                            if (isRoundingMekanism)
                            {
                                ttlcst = Math.Round(_curCost * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                            }
                            strSQL = strSQL + ttlcst.ToString() + "', "; //'wsd_total_cost
                            strSQL = strSQL + "null, "; //'wsd_remarks
                            strSQL = strSQL + "null, "; //'wsd_invoice_process_flag
                            strSQL = strSQL + "null, "; //'wsd_invoice_no
                            strSQL = strSQL + "null, "; //'wsd_invoice_qty_sales
                            strSQL = strSQL + "null, "; //'wsd_invoice_date
                            strSQL = strSQL + "null, "; //'wsd_invoice_process_by
                            strSQL = strSQL + "getdate() , '"; //'wsd_entry_date
                            strSQL = strSQL + clsLogin.USERID + "', '"; //wsd_user_id
                            strSQL = strSQL + sCurrCode + "', "; //'wsd_curr_code
                            strSQL = strSQL + dRateCurr + ", "; //'wsd_rate
                            Decimal orghet = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                            strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                            Decimal orgso = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                            strSQL = strSQL + orgso + ", "; //'wsd_org_total_so_amount
                            strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                            strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                            strSQL = strSQL + cConv2 + ", "; //'wsd_uom_convert_big
                            //fConvertQty(dgvDetail2.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                            double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                            strSQL = strSQL + uomorg + ", "; //'wsd_uom_org_qty
                            strSQL = strSQL + "0, '";

                            strSQL = strSQL + "0" + "', '"; //'wsd_tot_disc_1
                            strSQL = strSQL + "0" + "', "; //'wsd_tot_disc_2
                            strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                            strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString() + "', '"; //wsd_tax_code
                            strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString() + "',"; //wsd_tax_pct
                            strSQL = strSQL + "0, "; //wsd_tax_amount
                            strSQL = strSQL + "0, "; //wsd_gross_net
                            strSQL = strSQL + "0, "; //wsd_cash_disc
                            strSQL = strSQL + "'N', '"; //wsd_line_type
                            strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "', '";
                            strSQL = strSQL + sPriceCode + "', "; //wsd_price_code
                            if (dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                            {
                                strSQL = strSQL + "NULL"; //wsd_sled
                            }
                            else
                            {
                                strSQL = strSQL + "'"+ dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString().Trim() + "'"; //wsd_sled
                            }
                            strSQL = strSQL + ")"; //wsd_reason
                            dtFill15 = _clsGlobal.ExecDTTrans(strSQL);

                            lineNo++;
                            #endregion
                        }
                    }

                }

                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private string initDateDelvExp(string entity, string branch, string sodate, string custCode1, string custCode2, int executeType)
        {
            DataTable dt = new DataTable();
            string sqlScript = "Select dbo.F_GET_DATE_PLAN ('" + entity + "','" + branch + "','" + sodate + "','" + custCode1 + "','" + custCode2 + "')";
            if (executeType == 1) // exec biasa ga pake tran2an
            {
                dt = _clsGlobal.ExecDT(sqlScript);
            }
            else
            {
                dt = _clsGlobal.ExecDTTrans(sqlScript);
            }

            // yyyyMMdd
            return dt.Rows[0][0].ToString();

        }

        private bool fGetSONumber()
        {
            bool keyFound = true;

            string sBranchID = "";
            string sErrDesc = "";

            try
            {
                if (!isMultibranch)
                {
                    DataTable dtFill1 = new DataTable();
                    strSQL = "if not exists( select * FROM GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'SOSEQNUMBER')";
                    strSQL = strSQL + " Begin";
                    strSQL = strSQL + " insert into GS_GEN_HARDCODED (gh_sys, gh_function_name, gh_sequence_no, gh_function_code, gh_function_desc, ";
                    strSQL = strSQL + " gh_min_seq_no, gh_max_seq_no, gh_last_seq_no ) values ('S', 'SOSEQNUMBER', 1, 'SOSEQNUMBER', 'SOSEQNUMBER', 1, 9999999, 0 ) END";
                    dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

                    //DataTable dtFill2 = new DataTable();
                    //strSQL = " select br_branch_id from GS_BRANCH ";
                    //dtFill2 = _clsGlobal.ExecDT(strSQL);
                    //if (dtFill2.Rows.Count > 0)
                    //{
                    //    sBranchID = dtFill2.Rows[0]["br_branch_id"].ToString();
                    //}
                    //else
                    //{
                    //    MessageBox.Show("Data Branch tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}

                    //DataTable dtFill3 = new DataTable();
                    //strSQL = " select count(br_branch_id) br_branch_id from GS_BRANCH ";
                    //dtFill3 = _clsGlobal.ExecDT(strSQL);

                    //if (dtFill3.Rows.Count > 0)
                    //{
                    //    if (dtFill3.Rows[0]["br_branch_id"].ToString() != "1")
                    //    {
                    //        //MessageBox.Show("Data Branch tidak sama dengan satu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //        //keyFound = false;
                    //        //return keyFound;
                    //    }
                    //}
                    //else
                    //{
                    //    MessageBox.Show("Data Branch tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}

                    sBranchID = cbEntityBranch.txtBranchIdCM.Text.Trim();

                    DataTable dtFill4 = new DataTable();
                    strSQL = "Update GS_GEN_HARDCODED";
                    strSQL = strSQL + " set gh_last_seq_no = gh_last_seq_no + 1";
                    strSQL = strSQL + " where gh_sys = 'S'";
                    strSQL = strSQL + " and gh_function_name = 'SOSEQNUMBER'";
                    dtFill4 = _clsGlobal.ExecDTTrans(strSQL);

                    DataTable dtFill5 = new DataTable();
                    strSQL = "Select gh_last_seq_no from GS_GEN_HARDCODED WITH (NOLOCK)";
                    strSQL = strSQL + " where gh_sys = 'S' ";
                    strSQL = strSQL + " and gh_function_name = 'SOSEQNUMBER'";
                    dtFill5 = _clsGlobal.ExecDTTrans(strSQL);

                    if (dtFill5.Rows.Count > 0)
                    {
                        long i = Convert.ToInt64(dtFill5.Rows[0]["gh_last_seq_no"].ToString());
                        string cNol = "0000000" + i.ToString();
                        sSONo = "S" + sBranchID + cNol.Substring(cNol.Length - 7);
                    }
                    else
                    {
                        MessageBox.Show("Sequence No Order tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    int lengthseq = 10;
                    string prefix = "";
                    strSQL = "if not exists( ";
                    strSQL += " select gh_last_seq_no ";
                    strSQL += " from GS_GEN_HARDCODED WITH (NOLOCK) ";
                    strSQL += " where ";
                    strSQL += " gh_sys = 'S' and  gh_function_name = 'SOSEQNUMBER' ";
                    strSQL += " and gh_function_name <> gh_function_code ";
                    strSQL += " and Isnull(gh_group_menu_id,'') = '' ";
                    strSQL += " AND Isnull(gh_parent_menu,'') = '' ) ";
                    strSQL += " begin ";
                    strSQL += " insert into GS_GEN_HARDCODED (gh_sys, gh_function_name, gh_sequence_no, gh_function_code, gh_function_desc, gh_min_seq_no, gh_max_seq_no, gh_last_seq_no ) ";
                    strSQL += " values ('S', 'SOSEQNUMBER', 1, '', 'SO SEQUENTIAL NATIONAL', 1, 9999999, 0 ) ";
                    strSQL += " end ";

                    _clsGlobal.ExecuteTrans(strSQL);

                    DataTable dtFillCheck = new DataTable();

                    strSQL = " DECLARE @NextID numeric,@Prefix varchar(10) ";
                    strSQL += " UPDATE GH ";
                    strSQL += " SET @NextID = gh_last_seq_no = gh_last_seq_no + 1 ";
                    strSQL += " ,@Prefix = gh_function_code ";
                    strSQL += " FROM GS_GEN_HARDCODED GH WITH (UPDLOCK) ";
                    strSQL += " INNER JOIN ";
                    strSQL += " ( 	SELECT ";
                    strSQL += " 'SOSEQNUMBER' gh_function_name ";
                    strSQL += " ,CASE WHEN C.cnt2 > 0  THEN '" + cbEntityBranch.txtCM.Text.Trim() + "' ELSE '' END [gh_group_menu_id] ";
                    strSQL += " ,CASE WHEN C.cnt2 > 0  THEN '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' ELSE '' END [gh_parent_menu] ";
                    strSQL += "  FROM (SELECT COUNT(gh_function_name) as [cnt2] FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_group_menu_id = '" + cbEntityBranch.txtCM.Text.Trim() + "' AND gh_parent_menu = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' and  gh_function_name = 'SOSEQNUMBER' ) C ";
                    strSQL += " ) GH2 ";
                    strSQL += " ON GH2.gh_function_name = GH.gh_function_name AND ISNULL(GH2.gh_parent_menu,'') = ISNULL(GH.gh_parent_menu,'') AND ISNULL(GH2.gh_group_menu_id,'') = ISNULL(GH.gh_group_menu_id,'') ";
                    strSQL += " WHERE	gh_sys = 'S' ";
                    strSQL += " AND GH.gh_function_name = 'SOSEQNUMBER' ";
                    strSQL += " AND GH.gh_function_code <> GH.gh_function_name ";
                    strSQL += " SELECT @NextID [nextid], @Prefix [prefix] ";

                    dtFillCheck = _clsGlobal.ExecDTTrans(strSQL);

                    if (dtFillCheck.Rows.Count > 0)
                    {
                        int lengthangka = dtFillCheck.Rows[0]["nextid"].ToString().Trim().Length;
                        prefix = dtFillCheck.Rows[0]["prefix"].ToString();
                        int prefixlength = prefix.Length;
                        int lengthjumlahnol = lengthseq - (prefixlength + lengthangka);
                        string cnNol = "";

                        long i = Convert.ToInt64(dtFillCheck.Rows[0]["nextid"].ToString());
                        for (int j = 0; j < lengthjumlahnol; j++)
                        {
                            cnNol = cnNol + "0";
                        }

                        sSONo = prefix + cnNol + i.ToString();

                        DataTable dtSOCheck = new DataTable();
                        dtSOCheck = _clsGlobal.ExecDTTrans("Select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK) where wsh_seq_no = '" + sSONo + "' and wsh_entity_id = '" + cbEntityBranch.txtCM.Text.Trim() + "' and wsh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' ");
                        if (dtSOCheck.Rows.Count > 0)
                        {
                            MessageBox.Show("Sequence No Order sudah terpakai.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Sequence No Order tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                }


            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                keyFound = false;
                return keyFound;
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private string fPC2Qty(double sQty, int iConv1, int iConv2)
        {
            string qtyformat = "";
            DataTable dtFill1 = new DataTable();
            strSQL = " select dbo.SQTY_FORMAT (";
            strSQL = strSQL + sQty + ",";
            strSQL = strSQL + iConv1 + ",";
            strSQL = strSQL + iConv2 + ") as Qty_Format";
            dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

            if (dtFill1.Rows.Count > 0)
            {
                qtyformat = dtFill1.Rows[0]["Qty_Format"].ToString();
            }

            return qtyformat;
        }

        //mdlutility
        private bool fGetAFS(string sPcode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " select dbo.F_GET_AFS('";
            strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "','";
            strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','";
            strSQL = strSQL + sPcode + "','";
            strSQL = strSQL + sGrade + "','";
            strSQL = strSQL + sSize + "','";
            strSQL = strSQL + sWHLoc1 + "','";
            strSQL = strSQL + sWHLoc2 + "') as afs_qty";
            dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

            if (dtFill1.Rows.Count > 0)
            {
                if (dtFill1.Rows[0]["afs_qty"].ToString() == "")//NULL
                {
                    keyFound = false;
                }
                else
                {
                    int afs = Convert.ToInt32(dtFill1.Rows[0]["afs_qty"].ToString());
                    if (afs < 0)
                    {
                        lAFSQty = 0;
                    }
                    else
                    {
                        lAFSQty = afs;
                    }

                    keyFound = true;
                }
            }
            else
            {
                keyFound = false;
            }

            return keyFound;
        }
        private bool fGetAFSeCommerce(string sEntity, string sBranch, string sPcode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " select dbo.F_GET_AFS('";
            strSQL = strSQL + sEntity + "','";
            strSQL = strSQL + sBranch + "','";
            strSQL = strSQL + sPcode + "','";
            strSQL = strSQL + sGrade + "','";
            strSQL = strSQL + sSize + "','";
            strSQL = strSQL + sWHLoc1 + "','";
            strSQL = strSQL + sWHLoc2 + "') as afs_qty";
            dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

            if (dtFill1.Rows.Count > 0)
            {
                if (dtFill1.Rows[0]["afs_qty"].ToString() == "")//NULL
                {
                    keyFound = false;
                }
                else
                {
                    int afs = Convert.ToInt32(dtFill1.Rows[0]["afs_qty"].ToString());
                    if (afs < 0)
                    {
                        lAFSQty = 0;
                    }
                    else
                    {
                        lAFSQty = afs;
                    }

                    keyFound = true;
                }
            }
            else
            {
                keyFound = false;
            }

            return keyFound;
        }

        private bool fCheckBalance(string sPCode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;
            try
            {
                int lRecordsAffected = 0;
                int avgcost = 0;

                DataTable dtFill1 = new DataTable();

                strSQL = " select isnull(isnull(wlb_last_avg_ucost, cs_standard_unit_cost), 0) as wlb_last_avg_ucost from IM_STOCK_BALANCE WITH (NOLOCK) ";
                strSQL = strSQL + " LEFT JOIN IM_COST_STD WITH (NOLOCK) ON wlb_prd_master_code = cs_prd_master_code AND wlb_grade = cs_grade AND wlb_prd_size = cs_size";
                strSQL = strSQL + " where wlb_loc_Id1 = '" + sWHLoc1 + "'";
                strSQL = strSQL + " and wlb_loc_Id2 = '" + sWHLoc2 + "' and wlb_entity_id = '" + cbEntityBranch.txtCM.Text + "' and wlb_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                strSQL = strSQL + " and wlb_prd_master_code = '" + sPCode + "'";
                strSQL = strSQL + " and wlb_grade = '" + sGrade + "'";
                strSQL = strSQL + " and wlb_prd_size = '" + sSize + "'";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {

                }
                else //' JIka di balance tidak ada maka insert dengan qty nol , request Pa Wawan BL Project
                {
                    DataTable dtFill2 = new DataTable();
                    strSQL = "select dbo.GET_COST('" + cbEntityBranch.txtCM.Text + "','" + cbEntityBranch.txtBranchIdCM.Text + "','" + sWHLoc1 + "','" + sWHLoc2 + "','" + sPCode + "','" + sGrade + "','" + sSize + "') as cost ";
                    dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill2.Rows.Count > 0)
                    {
                        avgcost = Convert.ToInt16(dtFill2.Rows[0]["cost"].ToString());
                    }
                    else
                    {
                        avgcost = 0;
                    }

                    DataTable dtFill3 = new DataTable();
                    strSQL = "INSERT INTO IM_STOCK_BALANCE (wlb_entity_id,wlb_branch_id,wlb_loc_Id1, wlb_loc_Id2, wlb_prd_master_code, wlb_grade ,wlb_prd_size,wlb_qty_on_hand,wlb_last_avg_ucost,wlb_creation_date,wlb_first_receipt_date)";
                    strSQL = strSQL + " SELECT '" + cbEntityBranch.txtCM.Text + "','" + cbEntityBranch.txtBranchIdCM.Text + "','" + sWHLoc1 + "','" + sWHLoc2 + "',prm_prd_master_code,prm_grade,prm_prd_size,0," + avgcost + ",GETDATE(),GETDATE() FROM IM_PRD_MASTER WITH (NOLOCK) ";
                    strSQL = strSQL + " where prm_prd_master_code = '" + sPCode + "' and prm_grade = '" + sGrade + "' and prm_prd_size ='" + sSize + "'";
                    dtFill3 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill3.Rows.Count > 0)
                    {
                        lRecordsAffected = 1;
                    }

                    if (lRecordsAffected < 1)
                    {
                        keyFound = false;
                        return keyFound;
                    }


                }

                keyFound = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return keyFound;
        }

        private bool fCheckBalanceEcommerce(string sEntity, string sBranch, string sPCode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;
            try
            {
                int lRecordsAffected = 0;
                int avgcost = 0;

                DataTable dtFill1 = new DataTable();

                strSQL = " select isnull(isnull(wlb_last_avg_ucost, cs_standard_unit_cost), 0) as wlb_last_avg_ucost from IM_STOCK_BALANCE WITH (NOLOCK) ";
                strSQL = strSQL + " LEFT JOIN IM_COST_STD WITH (NOLOCK) ON wlb_prd_master_code = cs_prd_master_code AND wlb_grade = cs_grade AND wlb_prd_size = cs_size";
                strSQL = strSQL + " where wlb_loc_Id1 = '" + sWHLoc1 + "'";
                strSQL = strSQL + " and wlb_loc_Id2 = '" + sWHLoc2 + "' and wlb_entity_id = '" + sEntity + "' and wlb_branch_id = '" + sBranch + "'";
                strSQL = strSQL + " and wlb_prd_master_code = '" + sPCode + "'";
                strSQL = strSQL + " and wlb_grade = '" + sGrade + "'";
                strSQL = strSQL + " and wlb_prd_size = '" + sSize + "'";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {

                }
                else //' JIka di balance tidak ada maka insert dengan qty nol , request Pa Wawan BL Project
                {
                    DataTable dtFill2 = new DataTable();
                    strSQL = "select dbo.GET_COST('" + sEntity + "','" + sBranch + "','" + sWHLoc1 + "','" + sWHLoc2 + "','" + sPCode + "','" + sGrade + "','" + sSize + "') as cost ";
                    dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill2.Rows.Count > 0)
                    {
                        avgcost = Convert.ToInt16(dtFill2.Rows[0]["cost"].ToString());
                    }
                    else
                    {
                        avgcost = 0;
                    }

                    DataTable dtFill3 = new DataTable();
                    strSQL = "INSERT INTO IM_STOCK_BALANCE (wlb_entity_id,wlb_branch_id,wlb_loc_Id1, wlb_loc_Id2, wlb_prd_master_code, wlb_grade ,wlb_prd_size,wlb_qty_on_hand,wlb_last_avg_ucost,wlb_creation_date,wlb_first_receipt_date)";
                    strSQL = strSQL + " SELECT '" + sEntity + "','" + sBranch + "','" + sWHLoc1 + "','" + sWHLoc2 + "',prm_prd_master_code,prm_grade,prm_prd_size,0," + avgcost + ",GETDATE(),GETDATE() FROM IM_PRD_MASTER WITH (NOLOCK) ";
                    strSQL = strSQL + " where prm_prd_master_code = '" + sPCode + "' and prm_grade = '" + sGrade + "' and prm_prd_size ='" + sSize + "'";
                    dtFill3 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill3.Rows.Count > 0)
                    {
                        lRecordsAffected = 1;
                    }

                    if (lRecordsAffected < 1)
                    {
                        keyFound = false;
                        return keyFound;
                    }


                }

                keyFound = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return keyFound;
        }


        private void eExecute()
        {
            try
            {
                if (fCheckMandatory() == false)
                {
                    return;
                }
                //DataTable dtExcel = new DataTable();
                //dtExcel = ReadExcel(filePath, fileExt);//read excel file
                //dgvCommerce.Visible = true;
                //dgvCommerce.DataSource = dtExcel;
                Cursor.Current = Cursors.WaitCursor;
                pClearGrid();
                pEnabled(0);
                if (chkEcommerce.Checked)
                {
                    dgvHeader.Columns["g_iSalesmanCode"].Visible = true;
                    dgvHeader.Columns["g_iSalesmanName"].Visible = true;
                    dgvHeader.Columns["g_iVoucherNo"].Visible = false;
                    dgvHeader.Columns["g_iSOValue"].Visible = false;
                    dgvHeader.Columns["g_iTXTValue"].Visible = false;
                    if (chkmappingproduct.Checked)
                    {
                        dgvDetail.Columns["g_iDetPCodeMap"].Visible = true;
                    }
                    else
                    {
                        dgvDetail.Columns["g_iDetPCodeMap"].Visible = false;
                    }
                }
                else
                {
                    dgvHeader.Columns["g_iSalesmanCode"].Visible = false;
                    dgvHeader.Columns["g_iSalesmanName"].Visible = false;
                    dgvHeader.Columns["g_iVoucherNo"].Visible = true;
                    dgvHeader.Columns["g_iSOValue"].Visible = true;
                    dgvHeader.Columns["g_iTXTValue"].Visible = true;
                    dgvDetail.Columns["g_iDetPCodeMap"].Visible = false;
                }

                if (fLoadData(lblPath.Text) == false)
                {
                    pClearGrid();
                    return;
                }

                if (fProsesData() == false)
                {
                    _clsGlobal.RollbackTrans();
                    pClearGrid();
                    return;
                }
                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void PopulateDataTableFromUploadedFile(string strm)
        {
            System.IO.StreamReader srdr = new System.IO.StreamReader(strm);
            String strLine = String.Empty;
            Int32 iLineCount = 0;
            do
            {
                strLine = srdr.ReadLine();
                if (strLine == null)
                {
                    break;
                }
                if (0 == iLineCount++)
                {
                    dtView = this.CreateDataTableForCSVData(strLine);
                }
                this.AddDataRowToTable(strLine, dtView);
            } while (true);
        }

        private DataTable CreateDataTableForCSVData(String strLine)
        {
            DataTable dt = new DataTable("CSVTable");
            String[] strVals = strLine.Split(new char[] { '|' });
            m_iColumnCount = strVals.Length;
            int idx = 0;
            foreach (String strVal in strVals)
            {
                String strColumnName = String.Format("Column-{0}", idx++);
                dt.Columns.Add(strColumnName, Type.GetType("System.String"));
            }
            return dt;
        }

        private DataRow AddDataRowToTable(String strCSVLine, DataTable dt)
        {
            String[] strVals = strCSVLine.Split(new char[] { '|' });
            Int32 iTotalNumberOfValues = strVals.Length;
            // If number of values in this line are more than the columns
            // currently in table, then we need to add more columns to table.
            if (iTotalNumberOfValues > m_iColumnCount)
            {
                Int32 iDiff = iTotalNumberOfValues - m_iColumnCount;
                for (Int32 i = 0; i < iDiff; i++)
                {
                    String strColumnName = String.Format("Column-{0}", (m_iColumnCount + i));
                    dt.Columns.Add(strColumnName, Type.GetType("System.String"));
                }
                m_iColumnCount = iTotalNumberOfValues;
            }
            int idx = 0;
            DataRow drow = dt.NewRow();
            foreach (String strVal in strVals)
            {
                String strColumnName = String.Format("Column-{0}", idx++);
                drow[strColumnName] = strVal.Trim();
            }
            dt.Rows.Add(drow);
            return drow;
        }
        private void _checkPaymentType()
        {
            try
            {
                // columns 0 -> flag header
                // columns 1 -> cust code
                // columns 2 -> payment type
                if (isUploadMigrasiToNEFOKAM)
                {

                    DataRow[] __r = dtView.Select(string.Format("[{0}]='SOHDR'", dtView.Columns[0].ColumnName));
                    for (int i = 0; i < __r.Length; i++)
                    {
                        if (__r[i] == null) return;
                        __r[i][2] = _clsGlobal.GetFieldValue(
                            string.Format("SELECT TOP 1 cm_payment_type FROM SO_CUST_MASTER WITH(NOLOCK) WHERE cm_entity='{0}' AND cm_branch='{1}' AND cm_cust_code1='{2}'"
                            , cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __r[i][1].ToString()));

                    }

                }
                else
                {
                    DataRow __r = dtView.Select(string.Format("[{0}]='SOHDR'", dtView.Columns[0].ColumnName)).FirstOrDefault();
                    if (__r == null) return;
                    if (!cbSource.SelectedValue.ToString().CompareC("L"))
                        __r[2] = _clsGlobal.GetFieldValue(
                            string.Format("SELECT TOP 1 cm_payment_type FROM SO_CUST_MASTER WITH(NOLOCK) WHERE cm_entity='{0}' AND cm_branch='{1}' AND cm_cust_code1='{2}'"
                            , cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __r[1].ToString()));
                }

            }
            catch (Exception) { }
        }

        private bool fLoadData(string sPath)
        {
            bool keyFound = false;

            bool eFailedDtl = false;
            string vPONo = "";
            String name = "Items";
            string vLastPO = "";
            decimal dTotCust = 0;
            decimal dTotTRS = 0;

            if (!chkEcommerce.Checked)
            {
                PopulateDataTableFromUploadedFile(dataFile);
            }
            else
            {
                PopulateDataTableFromUploadedFileBFIExcel(dataFile, ".XLS");
            }

            if (!chkEcommerce.Checked)
            {
                _checkPaymentType();
            }


            if (dgvdata.Rows.Count > 0)
            {
                dgvdata.DataSource = null;
                dgvdata.Rows.Clear();
                dgvdata.Columns.Clear();
            }


            int _i = 0;


            SetGridData(dgvdata);
            DataTable dt1 = dgvdata.DataSource as DataTable;

            foreach (DataRow drow in dtView.Rows)
            {
                dt1.Rows.Add();
                if (!chkEcommerce.Checked)
                {
                    //dgvdata.Rows.Add();
                    dgvdata.Rows[_i].Cells[0].Value = dtView.Rows[_i][0].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[1].Value = dtView.Rows[_i][1].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[2].Value = dtView.Rows[_i][2].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[3].Value = dtView.Rows[_i][3].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[4].Value = dtView.Rows[_i][4].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[5].Value = dtView.Rows[_i][5].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[6].Value = dtView.Rows[_i][6].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[7].Value = dtView.Rows[_i][7].ToString().Replace("\"", "");
                    //dgvdata.Rows[_i].Cells[8].Value = dtView.Rows[_i][8].ToString();
                    //dgvdata.Rows[_i].Cells[9].Value = dtView.Rows[_i][9].ToString();
                    //dgvdata.Rows[_i].Cells[10].Value = dtView.Rows[_i][10].ToString();
                    //dgvdata.Rows[_i].Cells[11].Value = dtView.Rows[_i][11].ToString();
                    //dgvdata.Rows[_i].Cells[12].Value = dtView.Rows[_i][12].ToString();
                    //dgvdata.Rows[_i].Cells[13].Value = dtView.Rows[_i][13].ToString();
                }
                else
                {
                    dgvdata.Rows[_i].Cells[0].Value = dtView.Rows[_i][0].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[1].Value = dtView.Rows[_i][1].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[2].Value = dtView.Rows[_i][2].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[3].Value = dtView.Rows[_i][3].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[4].Value = dtView.Rows[_i][4].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[5].Value = dtView.Rows[_i][5].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[6].Value = dtView.Rows[_i][6].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[7].Value = dtView.Rows[_i][7].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[8].Value = dtView.Rows[_i][8].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[9].Value = dtView.Rows[_i][9].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[10].Value = dtView.Rows[_i][10].ToString().Replace("\"", "");
                    dgvdata.Rows[_i].Cells[11].Value = dtView.Rows[_i][11].ToString().Replace("\"", "");
                }


                _i++;
            }

            keyFound = true;

            return keyFound;
        }

        private void PopulateDataTableFromUploadedFileBFIExcel(string strm, string ext)
        {
            //System.IO.StreamReader srdr = new System.IO.StreamReader(strm);
            //String strLine = String.Empty;
            //Int32 iLineCount = 0;
            //do
            //{
            //    strLine = srdr.ReadLine();
            //    if (strLine == null)
            //    {
            //        break;
            //    }
            //    if (0 == iLineCount++)
            //    {
            //        dtView = this.CreateDataTableForCSVData(strLine);
            //    }
            //    this.AddDataRowToTable(strLine, dtView);
            //} while (true);
            String constr = null;
            if (ext.ToUpper().ToString().Equals(".XLS"))
            {
                constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strm +
                         ";Extended Properties='Excel 12.0 xml;HDR=YES;';";
            }
            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [Sheet1$]", con);
            con.Open();
            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            DataSet dataSet = new DataSet();
            sda.Fill(dataSet);
            DataTable dtTemp = dataSet.Tables[0];
            con.Close();

            /** sorting data table **/
            DataView dv = dtTemp.DefaultView;
            dv.Sort = "BRANCH,CUST_CODE,SALESMAN,PO_NO ASC";
            dtTemp = dv.Table;
            /** end sorting data table **/


            dtHeader = new DataTable();
            dtHeader.Columns.Add("Entity"); //1
            dtHeader.Columns.Add("Branch"); //2
            dtHeader.Columns.Add("PONO"); //3
            dtHeader.Columns.Add("PODate"); //4
            dtHeader.Columns.Add("kdsales_vendor"); //5
            dtHeader.Columns.Add("kdsales_trs"); //6
            dtHeader.Columns.Add("kdoutlet_vendor"); //7
            dtHeader.Columns.Add("kdoutlet_trs"); //8
            dtHeader.Columns.Add("paytype"); //9
            dtHeader.Columns.Add("DelvDate");//10
            dtHeader.Columns.Add("ExpDate");//11



            dtDetail = new DataTable();
            dtDetail.Columns.Add("Entity");//1
            dtDetail.Columns.Add("Branch"); //2
            dtDetail.Columns.Add("PoNo"); //3
            dtDetail.Columns.Add("kdsales_trs"); //4
            dtDetail.Columns.Add("kdoutlet_trs"); //5
            dtDetail.Columns.Add("sku_map"); //6
            dtDetail.Columns.Add("sku_trs"); //7
            dtDetail.Columns.Add("sku_pl"); //8
            dtDetail.Columns.Add("qty"); //9
            dtDetail.Columns.Add("paytype"); //10
            dtDetail.Columns.Add("Conv1"); //11
            dtDetail.Columns.Add("Conv2"); //12



            //dtStagingAllData.Columns.Add("error"); //11
            //dtStagingAllData.Columns.Add("keterangan"); //12


            bool stillread = true;
            string __poexisting = "";
            string __salesvendorexisting = "";
            string __outletvendorexisting = "";
            string __skuvendorexisting = "";


            string __kdoutletbefore = "";
            string __kdsalestrsbefore = "";
            string __branchbefore = "";
            string __ponobefore = "";

            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {



                string __kdoutlet = dtTemp.Rows[i][0].ToString().Trim();
                string __kdsalestrs = dtTemp.Rows[i][1].ToString().Trim();
                string __branch = dtTemp.Rows[i][2].ToString().Trim();
                string __pono = dtTemp.Rows[i][3].ToString().Trim();
                string __podate = dtTemp.Rows[i][4].ToString().Trim();
                string __expdate = dtTemp.Rows[i][5].ToString().Trim();
                string __delvdate = dtTemp.Rows[i][6].ToString().Trim();
                string __paytype = dtTemp.Rows[i][7].ToString().Trim();
                string __sku = dtTemp.Rows[i][8].ToString().Trim();
                string __qty = dtTemp.Rows[i][9].ToString().Trim();
                string __uom = dtTemp.Rows[i][10].ToString().Trim();
                string __prdline = "";
                string __conv1 = "";
                string __conv2 = "";

                string __skumap = "";
                string __skutrs = "";
                sPcode = "";
                sProdLine = "";
                iConv1 = "";
                iConv2 = "";
                if (chkEcommerce.Checked && chkmappingproduct.Checked)
                {
                    __skumap = dtTemp.Rows[i][8].ToString().Trim();
                    getProductMapping("01", __branch, __kdoutlet, "000", __skumap, 1);
                    __prdline = sProdLine;
                    __skutrs = sPcode;
                    __conv1 = iConv1;
                    __conv2 = iConv2;
                }
                else
                {
                    __skutrs = dtTemp.Rows[i][8].ToString().Trim();
                    getProductMapping("01", __branch, __kdoutlet, "000", __skutrs, 2);
                    __prdline = sProdLine;
                    __conv1 = iConv1;
                    __conv2 = iConv2;
                }

                __paytype = _clsGlobal.GetFieldValue(string.Format("SELECT TOP 1 cm_payment_type FROM SO_CUST_MASTER WITH(NOLOCK) WHERE cm_entity='{0}' AND cm_branch='{1}' AND cm_cust_code1='{2}'"
                       , cbEntityBranch.txtCM.Text.Trim(), __branch, __kdoutlet));
                #region Insert Header
                if (string.IsNullOrEmpty(__kdoutletbefore))
                {
                    dtHeader.Rows.Add
                    (
                        "01"//dtHeader.Columns.Add("Entity"); //1"
                        , __branch//dtHeader.Columns.Add("Branch"); //2
                        , __pono//dtHeader.Columns.Add("PONO"); //3
                        , __podate//dtHeader.Columns.Add("PODate"); //4
                        , ""//dtHeader.Columns.Add("kdsales_vendor"); //5
                        , __kdsalestrs//dtHeader.Columns.Add("kdsales_trs"); //6
                        , ""//dtHeader.Columns.Add("kdoutlet_vendor"); //7
                        , __kdoutlet//dtHeader.Columns.Add("kdoutlet_trs"); //8
                        , __paytype//dtHeader.Columns.Add("paytype"); //9
                        , __delvdate //dtHeader.Columns.Add("Delvdate"); //10
                        , __expdate//dtHeader.Columns.Add("ExpDate");//11
                    );
                    __branchbefore = __branch;
                    __ponobefore = __pono;
                    __kdsalestrsbefore = __kdsalestrs;
                    __kdoutletbefore = __kdoutlet;
                }
                else
                {
                    if (__branch == __branchbefore
                       && __ponobefore == __pono
                       && __kdoutletbefore == __kdoutlet
                       && __kdsalestrsbefore == __kdsalestrs
                      )
                    {
                        // ga usah diapa2in
                        __branchbefore = __branch;
                        __ponobefore = __pono;
                        __kdsalestrsbefore = __kdsalestrs;
                        __kdoutletbefore = __kdoutlet;
                    }
                    else
                    {
                        dtHeader.Rows.Add
                        (
                            "01"//dtHeader.Columns.Add("Entity"); //1"
                            , __branch//dtHeader.Columns.Add("Branch"); //2
                            , __pono//dtHeader.Columns.Add("PONO"); //3
                            , __podate//dtHeader.Columns.Add("PODate"); //4
                            , ""//dtHeader.Columns.Add("kdsales_vendor"); //5
                            , __kdsalestrs//dtHeader.Columns.Add("kdsales_trs"); //6
                            , ""//dtHeader.Columns.Add("kdoutlet_vendor"); //7
                            , __kdoutlet//dtHeader.Columns.Add("kdoutlet_trs"); //8
                            , __paytype//dtHeader.Columns.Add("paytype"); //9
                            , __delvdate //dtHeader.Columns.Add("Delvdate"); //10
                            , __expdate//dtHeader.Columns.Add("ExpDate");//11
                        );
                        __branchbefore = __branch;
                        __ponobefore = __pono;
                        __kdsalestrsbefore = __kdsalestrs;
                        __kdoutletbefore = __kdoutlet;
                    }
                }
                #endregion

                #region Insert Detail
                dtDetail.Rows.Add(
                "01"    //dtDetail.Columns.Add("Entity");//1
                , __branch    //dtDetail.Columns.Add("Branch"); //2
                , __pono    //dtDetail.Columns.Add("PoNo"); //3
                , __kdsalestrs    //dtDetail.Columns.Add("kdsales_trs"); //4
                , __kdoutlet    //dtDetail.Columns.Add("kdoutlet_trs"); //5
                , __skumap    //dtDetail.Columns.Add("sku_map"); //6
                , __skutrs    //dtDetail.Columns.Add("sku_trs"); //7
                , __prdline    //dtDetail.Columns.Add("sku_pl"); //8
                , __qty    //dtDetail.Columns.Add("qty"); //9
                , __paytype    //dtDetail.Columns.Add("paytype"); //10
                , __conv1    //dtDetail.Columns.Add("Conv1"); //11
                , __conv2    //dtDetail.Columns.Add("Conv2"); //12
                );
                #endregion


            }
            dtView = dtDetail;

        }

        private void getProductMapping(string entity, string branch, string custcode1, string custcode2, string skumap, int type)
        {
            /** type 1 = mappingan,  type 2 = real product **/
            if (type == 1)
            {
                strSQL = "SELECT top 1 mp_trs_prd_id,prm_prd_line_code,prm_conversion_purc conv1,prm_conversion_sales conv2, mcps_prd_sled as [prm_prd_sled] FROM TBL_MAP_PRD_B2B WITH (NOLOCK) INNER JOIN IM_PRD_MASTER WITH (NOLOCK) ON prm_prd_master_code = mp_trs_prd_id and prm_grade = mp_trs_grade and prm_prd_size = mp_trs_size ";
                strSQL = " LEFT JOIN TBL_MAP_CUST_PRD_SLED WITH (NOLOCK) mcps_entity_id=mp_entity  and mcps_branch_id=mp_branch_id and mcps_cust_code1=mp_cust_code1 and mcps_cust_code2=mp_cust_code2 and mcps_prd_code=mp_trs_prd_id and mcps_prd_grade=mp_trs_grade and mcps_prd_size=mp_trs_size ";
                strSQL += "where mp_entity = '" + entity + "' and mp_branch_id = '" + branch + "' and mp_cust_prd_id = '" + skumap + "' and mp_cust_code1 = '" + custcode1 + "' ";
                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    sPcode = dt.Rows[0]["mp_trs_prd_id"].ToString().Trim();
                    sProdLine = dt.Rows[0]["prm_prd_line_code"].ToString().Trim();
                    sSled     = dt.Rows[0]["prm_prd_sled"].ToString().Trim();
                    iConv1 = dt.Rows[0]["conv1"].ToString().Trim();
                    iConv2 = dt.Rows[0]["conv2"].ToString().Trim();

                }
            }
            else
            {
                strSQL = "select TOP 1 prm_prd_master_code,prm_prd_line_code,prm_conversion_purc conv1,prm_conversion_sales conv2, mcps_prd_sled prm_prd_sled ";
                strSQL += "from IM_PRD_MASTER WITH (NOLOCK) ";
                strSQL += " LEFT JOIN TBL_MAP_CUST_PRD_SLED WITH (NOLOCK) ON mcps_entity_id='"+entity+ "' and mcps_branch_id='"+branch+"' ";
                strSQL += " AND mcps_cust_code1='"+custcode1+ "' AND mcps_cust_code2='"+custcode2+ "' AND mcps_prd_code=prm_prd_master_code and mcps_prd_grade=prm_grade and mcps_prd_size=prm_prd_size ";
                strSQL += " where prm_prd_master_code = '" + skumap + "' ";
                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    sPcode = dt.Rows[0]["prm_prd_master_code"].ToString().Trim();
                    sProdLine = dt.Rows[0]["prm_prd_line_code"].ToString().Trim();
                    sSled = dt.Rows[0]["prm_prd_sled"].ToString().Trim();
                    iConv1 = dt.Rows[0]["conv1"].ToString().Trim();
                    iConv2 = dt.Rows[0]["conv2"].ToString().Trim();
                }
            }

        }


        private string getOutletVendor(string kodeoutletvendor, string kodesalesman)
        {
            string __outlet = "";
            strSQL = "SELECT cm_cust_code1 FROM VW_OUTLET_VENDOR WITH (NOLOCK) where cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' and cm_account_manager = '" + kodeoutletvendor + "' and csc_salesman_id = '" + kodesalesman + "' ";
            DataTable dt = _clsGlobal.ExecDT(strSQL);
            if (dt.Rows.Count > 0)
            {
                __outlet = dt.Rows[0]["cm_cust_code1"].ToString().Trim();
            }

            return __outlet;
        }

        private bool checkHeaderSO(string PoNo, string salestrs, string outlettrs)
        {
            bool result = false;
            string expression = "PoNo ='" + PoNo + "' AND kdsales_trs ='" + salestrs + "' AND kdoutlet_trs ='" + outlettrs + "' ";
            DataRow[] selectedRows = dtHeader.Select(expression);
            if (selectedRows.Count() > 0)
            {
                result = true;
            }

            return result;
        }

        private bool fProsesData()
        {
            bool keyFound = false;

            int CountDtl = 0;
            if (!chkEcommerce.Checked)
            {
                if (dgvdata.Rows[0].Cells[0].Value.ToString() == "SOHDR")
                {
                    if (dgvdata.ColumnCount == 8)
                    {
                        if (fProsesTextFile() == false)
                        {
                            keyFound = false;
                            return keyFound;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Block Data not valid please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    MessageBox.Show("Row Header Code not valid please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }
            }
            else
            {
                fProsesTextFileEcommerce();
            }


            if (fGetDetailAttr() == false)
            {
                keyFound = false;
                return keyFound;
            }

            int iFail = 0;
            for (int r = 0; r < dgvHeader.Rows.Count; r++)
            {
                iFail = 0;
                CountDtl = 0;
                bool foundSLED = false;


                for (int i = 0; i < dgvDetail2.Rows.Count; i++)
                {
                    if (chkEcommerce.Checked)
                    {
                        if (
                            dgvHeader.Rows[r].Cells["g_iCustPONo"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetPONO1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iEntity"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetEntity1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iBranch"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetBranch1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iSalesmanCode"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetSalesmanID1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iCustCode"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetCustNo1"].Value.ToString()
                            )
                        {
                            if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() != "OK")
                            {
                                iFail = iFail + 1;

                            }
                            CountDtl = CountDtl + 1;

                            if (!string.IsNullOrWhiteSpace(dgvDetail2.Rows[i].Cells["g_iDetSled1"].Value?.ToString()))
                            {
                                foundSLED = true;
                            }
                        }

                       
                        if (
                            dgvHeader.Rows[r].Cells["g_iCustPONo"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetPONO1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iEntity"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetEntity1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iBranch"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetBranch1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iSalesmanCode"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetSalesmanID1"].Value.ToString()
                            && dgvHeader.Rows[r].Cells["g_iCustCode"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetCustNo1"].Value.ToString()
                            )
                        {
                            if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() != "OK")
                            {
                                iFail = iFail + 1;

                            }
                            CountDtl = CountDtl + 1;

                            if (!string.IsNullOrWhiteSpace(dgvDetail2.Rows[i].Cells["g_iDetSled1"].Value?.ToString()))
                            {
                                foundSLED = true;
                            }
                        }
                    }
                    else
                    {
                        if (dgvHeader.Rows[r].Cells["g_iCustPONo"].Value.ToString() == dgvDetail2.Rows[i].Cells["g_iDetPONO1"].Value.ToString())
                        {
                            if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() != "OK")
                            {
                                iFail = iFail + 1;

                            }
                            CountDtl = CountDtl + 1;

                            if (!string.IsNullOrWhiteSpace(dgvDetail2.Rows[i].Cells["g_iDetSled1"].Value?.ToString()))
                            {
                                foundSLED = true;
                            }
                        }
                    }

                }
                if (dgvHeader.Rows[r].Cells["g_iCustName"].Value.ToString() == "")
                {
                    dgvHeader.Rows[r].Cells["g_iRemark"].Value = "Customer Tidak terdaftar";
                }
                else if (CountDtl == 0)
                {
                    dgvHeader.Rows[r].Cells["g_iRemark"].Value = "Error, detail Product not found";
                }

                else if (iFail != 0)
                {
                    dgvHeader.Rows[r].Cells["g_iRemark"].Value = "Error on detail data";
                }

                else if (fIsPOExists(dgvHeader.Rows[r].Cells["g_iCustPONo"].Value.ToString()) != "")
                {
                    dgvHeader.Rows[r].Cells["g_iRemark"].Value = "PO already Exists with SO No : " + fIsPOExists(dgvHeader.Rows[r].Cells["g_iCustPONo"].Value.ToString());
                }
                else
                {
                    dgvHeader.Rows[r].Cells["g_iRemark"].Value = "OK";
                }

                if (dgvHeader.Rows[r].Cells["g_iRemark"].Value.ToString().ToLower() == "ok")
                {
                    dgvHeader.Rows[r].Cells["g_iValid"].Value = "1";
                }
                else
                {
                    dgvHeader.Rows[r].Cells["g_iValid"].Value = "0";
                }

                dgvHeader.Rows[r].Cells["g_iReasonRejection"].Value = "";
                dgvHeader.Rows[r].Cells["g_iSledFlag"].Value = foundSLED ? "Y" : "N";


            }



            if (cbSource.SelectedValue.ToString() == "L" || cbSource.SelectedValue.ToString() == "S" || cbSource.SelectedValue.ToString() == "W")
            {
                if (!isNefoKAM)
                {
                    for (int i = 0; i < dgvHeader.Rows.Count; i++)
                    {
                        if (fViewDetail(i))
                        {
                            calculateSOValue(i);
                        }

                    }
                }

            }

            if (dgvHeader.Rows.Count > 0)
            {
                if (fViewDetail(dgvHeader.CurrentCell.RowIndex) == false)
                {
                    keyFound = false;
                    return keyFound;
                }
            }

            keyFound = true;
            return keyFound;
        }

        private bool fProsesTextFileEcommerce()
        {
            bool keyFound = false;

            try
            {
                string vPONo = "";
                sCustNo = "";
                ConvFail = "";
                string sCustPCode = "";
                //LineNum = 0
                string CCP = "";
                string CTEMP1 = ""; //customer code untuk division
                string CTEMP2 = ""; //customer code2 untuk filter division
                string PONO = "";
                string PCode1 = "";
                string PCode2 = "";
                string sGradeT = "";
                string sSizeT = "";
                string sSled = "";
                string PayType = "";

                string tglgudang = "";
                string wh1 = "";
                string wh2 = "";

                dgvHeader.Rows.Clear();
                dgvDetail.Rows.Clear();
                //dgvdata.Rows.Clear();



                int _i = 0;
                //dgvdata.DataSource = dtView;
                //dgvdata.Columns.Clear();
                //SetGridData(dgvdata);
                //DataTable dt1 = dgvdata.DataSource as DataTable;

                string __strSQL = "";
                __strSQL = "SELECT top 1 wh_loc_id1,wh_loc_id2,wh_loc_last_work_date FROM IM_WH_LOC WITH (NOLOCK) ";
                __strSQL += "WHERE ";
                //__strSQL += "wh_loc_entity = '" + cbEntityBranch.txtCM.Text + "' ";
                //__strSQL += "AND wh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                __strSQL += "wh_loc_id1 +'|'+wh_loc_id2 = (SELECT top 1 gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_function_name  LIKE 'MAINWHLOC' ) ";
                DataTable dtTglGudang = _clsGlobal.ExecDT(__strSQL);
                if (dtTglGudang.Rows.Count > 0)
                {
                    //tglgudang = dtTglGudang.Rows[0]["wh_loc_last_work_date"].ToString();
                    wh1 = dtTglGudang.Rows[0]["wh_loc_id1"].ToString();
                    wh2 = dtTglGudang.Rows[0]["wh_loc_id2"].ToString();
                }
                else
                {
                    MessageBox.Show(" Tanggal Gudang tidak ditemukan!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }




                /** Temp Data untuk check double PO **/
                DataTable dtPO = new DataTable();
                dtPO.Columns.Add("Branch");
                dtPO.Columns.Add("PO");
                dtPO.Columns.Add("Outlet");
                dtPO.Columns.Add("Salesman");
                /** end Temp Data untuk check double PO **/

                #region Validasi

                /* detail */
                for (lCounter = 0; lCounter < dgvdata.Rows.Count; lCounter++)
                {

                    //if (dgvdata.Rows[lCounter].Cells[6].Value.ToString() == "")
                    //{
                    //    MessageBox.Show("Row Header [Line " + lCounter + "] : Invalid PCode For Vendor Pcode " + dgvdata.Rows[lCounter].Cells[5].Value.ToString() + ", please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}

                    string ___entity = dgvdata.Rows[lCounter].Cells[0].Value.ToString();
                    string ___branch = dgvdata.Rows[lCounter].Cells[1].Value.ToString();

                    __strSQL = "select br_branch_id from GS_BRANCH WITH (NOLOCK) where br_gl_entity_initial  = '01' and br_branch_id = '" + ___branch + "'";
                    DataTable dtCheckBranch = _clsGlobal.ExecDT(__strSQL);
                    if (dtCheckBranch.Rows.Count == 0)
                    {
                        MessageBox.Show("Branch " + ___branch + " tidak ditemukan!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }


                    __strSQL = "SELECT top 1 wh_loc_id1,wh_loc_id2,wh_loc_last_work_date FROM IM_WH_LOC WITH (NOLOCK) ";
                    __strSQL += "WHERE ";
                    __strSQL += "wh_loc_entity = '" + ___entity + "' ";
                    __strSQL += "AND wh_branch_id = '" + ___branch + "' ";
                    __strSQL += "AND wh_loc_id1 +'|'+wh_loc_id2 = (SELECT top 1 gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_function_name  LIKE 'MAINWHLOC' ) ";
                    DataTable dtCheckTglGudang = _clsGlobal.ExecDT(__strSQL);
                    if (dtTglGudang.Rows.Count == 0)
                    {
                        MessageBox.Show(" Tanggal Gudang untuk gudang utama branch " + ___branch + " tidak ditemukan!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }

                    if (chkmappingproduct.Checked)
                    {
                        string __outlet = dgvdata.Rows[lCounter].Cells[4].Value.ToString();
                        string __pcodex = dgvdata.Rows[lCounter].Cells[5].Value.ToString();
                        string __pcode = dgvdata.Rows[lCounter].Cells[6].Value.ToString();
                        if (string.IsNullOrEmpty(__pcode))
                        {
                            MessageBox.Show("Row Header [Line " + lCounter + 1 + "] : Invalid mapping product " + __pcodex + " and outlet = " + __outlet + ", please Tcode SDP50A !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }

                    if (IsNumeric(dgvdata.Rows[lCounter].Cells[8].Value.ToString()) == false)
                    {
                        MessageBox.Show("Row Header [Line " + lCounter + 1 + "] : Invalid Conversion , please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }

                    if (IsNumeric(dgvdata.Rows[lCounter].Cells[8].Value.ToString()) == false && dgvdata.Rows[lCounter].Cells[8].Value.ToString() != "" || IsNumeric(dgvdata.Rows[lCounter].Cells[8].Value.ToString()) == false && dgvdata.Rows[lCounter].Cells[8].Value.ToString() != "")
                    {
                        MessageBox.Show("Row Header [Line " + lCounter + 1 + "] : Invalid Qty Value, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }

                    if (chkEcommerce.Checked)
                    {
                        string __branch = dgvdata.Rows[lCounter].Cells[1].Value.ToString().Trim();
                        string __PO = dgvdata.Rows[lCounter].Cells[2].Value.ToString().Trim();
                        string __Salesman = dgvdata.Rows[lCounter].Cells[3].Value.ToString().Trim();
                        string __Outlet = dgvdata.Rows[lCounter].Cells[4].Value.ToString().Trim();

                        DataView dv = new DataView(dtPO);
                        dv.RowFilter = "PO = '" + __PO + "'";
                        DataTable temp = dv.ToTable();
                        if (temp.Rows.Count == 0)
                        {
                            /** insert ke data PO **/
                            dtPO.Rows.Add(__branch, __PO, __Outlet, __Salesman);
                        }
                        else
                        {
                            DataTable dtCheck = dv.ToTable();
                            string __branchX = dtCheck.Rows[0]["Branch"].ToString().Trim();
                            string __SalesmanX = dtCheck.Rows[0]["Salesman"].ToString().Trim();
                            string __OutletX = dtCheck.Rows[0]["Outlet"].ToString().Trim();
                            if (__branchX != __branch || __Outlet != __OutletX || __Salesman != __SalesmanX)
                            {
                                MessageBox.Show("Row Header [Line " + (lCounter + 1).ToString() + "] : Double PO untuk Salesman/Outlet yang berbeda Mohon Cek Kembali Data Upload !!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                        }
                    }


                }
                /*end of detail */

                /* header */
                for (lCounter = 0; lCounter < dtHeader.Rows.Count; lCounter++)
                {

                    if (string.IsNullOrEmpty(dtHeader.Rows[lCounter]["kdsales_trs"].ToString()))
                    {
                        MessageBox.Show(" Salesman Code is empty for Salesman Code = " + dtHeader.Rows[lCounter]["kdsales_trs"].ToString() + ", please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }

                    /** SALESMAN **/
                    string __entity = dtHeader.Rows[lCounter]["Entity"].ToString();
                    string __branch = dtHeader.Rows[lCounter]["Branch"].ToString();
                    string __salesman = dtHeader.Rows[lCounter]["kdsales_trs"].ToString();
                    string __outlet = dtHeader.Rows[lCounter]["kdoutlet_trs"].ToString();

                    DataTable dtCheckSalesman = _clsGlobal.ExecDT("select * from  SO_SPG_GIRL_MAN WITH (NOLOCK) where sgm_entity_id = '" + __entity + "' and sgm_branch_id = '" + __branch + "' and sgm_spgm_id = '" + __salesman + "'");
                    if (dtCheckSalesman.Rows.Count == 0)
                    {

                        MessageBox.Show(" Salesman Code not found for Salesman Code = " + dtHeader.Rows[lCounter]["kdsales_trs"].ToString() + ", please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }



                    /** END SALESMAN **/

                    #region CekValidasi Tanggal PO, EXP, DELIVERY


                    //validation po date, sch date dan expired date
                    string _poDate = dtHeader.Rows[lCounter]["PODate"].ToString();
                    string _delvDate = dtHeader.Rows[lCounter]["DelvDate"].ToString();
                    string _expDate = dtHeader.Rows[lCounter]["ExpDate"].ToString();

                    if (string.IsNullOrEmpty(_poDate))
                    {
                        MessageBox.Show(" Tanggal PO tidak boleh kosong!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        string _poDateSOC = DateTime.ParseExact(_poDate, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                        DateTime _poDateSO = Convert.ToDateTime(_poDateSOC);
                        TimeSpan selisihTanggal = g_dTglGudang - _poDateSO;
                        if (selisihTanggal.Days <= 1)
                        {

                        }
                        else
                        {
                            MessageBox.Show(" Tanggal PO tidak boleh H-1 dari tanggal gudang!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }

                    if (string.IsNullOrEmpty(_delvDate))
                    {
                        MessageBox.Show(" Request Delivery Date tidak boleh kosong!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        string _delvDateSOC = DateTime.ParseExact(_delvDate, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                        DateTime _delvDateSO = Convert.ToDateTime(_delvDateSOC);
                        TimeSpan selisihTanggal = g_dTglGudang - _delvDateSO;
                        if (selisihTanggal.Days <= 0)
                        {

                        }
                        else
                        {
                            MessageBox.Show("Request Delivery Date tidak boleh lebih kecil dari tanggal gudang!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }


                    if (string.IsNullOrEmpty(_expDate))
                    {
                        MessageBox.Show(" Expired Date tidak boleh kosong!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        string _expDateSOC = DateTime.ParseExact(_expDate, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                        DateTime _expDateSO = Convert.ToDateTime(_expDateSOC);
                        TimeSpan selisihTanggal = g_dTglGudang - _expDateSO;
                        if (selisihTanggal.Days <= 0)
                        {

                        }
                        else
                        {
                            MessageBox.Show("Expired Date tidak boleh lebih kecil dari tanggal gudang!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }

                    #endregion

                    if (string.IsNullOrEmpty(dtHeader.Rows[lCounter]["kdoutlet_trs"].ToString()))
                    {
                        MessageBox.Show(" Customer Code is empty for Customer Code = " + dtHeader.Rows[lCounter]["kdoutlet_trs"].ToString() + ", please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }

                    if (string.IsNullOrEmpty(dtHeader.Rows[lCounter]["PONO"].ToString().Trim()))
                    {
                        MessageBox.Show("Po Number is empty, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    if (isNefoKAM)
                    {
                        if (dtHeader.Rows[lCounter]["PONO"].ToString().Trim().Length > 35)
                        {
                            MessageBox.Show("Row Header [Line " + lCounter + "] : Po Number - " + dgvdata.Rows[lCounter].Cells[3].Value.ToString() + " is can't be more than 35 characters, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;

                        }

                    }

                    /** COVER CUSTOMER **/
                    DataTable dtCheckCustomerCover = _clsGlobal.ExecDT("SELECT cm_cust_code1 FROM SO_CUST_MASTER WITH(NOLOCK) inner join TBL_SD_CUSTCOVER WITH(NOLOCK) ON cm_cust_code1 = csc_cust_code1 and cm_cust_code2 = csc_cust_code2 and cm_entity = csc_entity and cm_branch = csc_branch where csc_salesman_id = '" + __salesman + "' and csc_cust_code1 = '" + __outlet + "' and csc_entity = '" + __entity + "' and csc_branch = '" + __branch + "' ");
                    if (dtCheckCustomerCover.Rows.Count == 0)
                    {
                        MessageBox.Show(" Customer Code " + __outlet + " not covered by Salesman " + __salesman + ", please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }



                }
                /* end of header */





                #endregion

                #region InsertToGrid

                //sCustNo = dgvdata.Rows[lCounter].Cells[1].Value.ToString();
                //vPONo = dgvdata.Rows[lCounter].Cells[3].Value.ToString();

                //if (fGetCustNO() == false)
                //{
                //    dgvdata.DataSource = null;
                //    dgvdata.Rows.Clear();
                //    keyFound = false;
                //    return keyFound;
                //}

                int iRow = 0;
                int __No = 1;
                for (lCounter = 0; lCounter < dtHeader.Rows.Count; lCounter++)
                {
                    string __entity = dtHeader.Rows[lCounter]["Entity"].ToString();
                    string __branch = dtHeader.Rows[lCounter]["Branch"].ToString();
                    string __poDate = dtHeader.Rows[lCounter]["PODate"].ToString();
                    string __delvDate = dtHeader.Rows[lCounter]["DelvDate"].ToString();
                    string __expDate = dtHeader.Rows[lCounter]["ExpDate"].ToString();


                    sCustNo = dtHeader.Rows[lCounter]["kdoutlet_trs"].ToString();
                    vPONo = dtHeader.Rows[lCounter]["PONO"].ToString();

                    if (fGetCustNOPerSalesman(dtHeader.Rows[lCounter]["kdsales_trs"].ToString().Trim(), __entity, __branch) == false)
                    {
                        dgvdata.DataSource = null;
                        dgvdata.Rows.Clear();
                        keyFound = false;
                        return keyFound;
                    }

                    if (dgvdata.Rows.Count > 0)
                    {
                        dgvHeader.Rows.Add();
                        dgvHeader.Rows[iRow].Cells["g_iNo"].Value = __No;
                        dgvHeader.Rows[iRow].Cells["g_iProses"].Value = false;
                        if (!string.IsNullOrEmpty(sCust1))
                        {
                            dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value = sCust1;
                        }
                        else
                        {
                            dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value = sCustNo;
                        }
                        if (!string.IsNullOrEmpty(sCust2))
                        {
                            dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value = sCust2;
                        }
                        else
                        {
                            dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value = "000";
                        }
                        dgvHeader.Rows[iRow].Cells["g_iEntity"].Value = __entity;
                        dgvHeader.Rows[iRow].Cells["g_iBranch"].Value = __branch;
                        dgvHeader.Rows[iRow].Cells["g_iCustName"].Value = sCustName;
                        dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value = vPONo;
                        //dgvHeader.Rows[iRow].Cells["g_iSalesmanVendor"].Value = dtHeader.Rows[lCounter]["kdsales_vendor"].ToString();
                        DataTable dtCheckSalesman = _clsGlobal.ExecDT("select * from  SO_SPG_GIRL_MAN WITH (NOLOCK) where sgm_entity_id = '" + __entity + "' and sgm_branch_id = '" + __branch + "' and sgm_spgm_id = '" + dtHeader.Rows[lCounter]["kdsales_trs"].ToString() + "'");
                        if (dtCheckSalesman.Rows.Count > 0)
                        {
                            dgvHeader.Rows[iRow].Cells["g_iSalesmanName"].Value = dtCheckSalesman.Rows[0]["sgm_spgm_name"].ToString();
                        }

                        dgvHeader.Rows[iRow].Cells["g_iSalesmanCode"].Value = dtHeader.Rows[lCounter]["kdsales_trs"].ToString();
                        DataTable dtFill1 = new DataTable();
                        strSQL = " select isnull(soh_so_no,'') so_no From TBL_SD_UPLOADSO_HDR WITH (NOLOCK) where soh_po_no='" + vPONo + "' and soh_cust_code1='" + sCust1 + "' and soh_cust_code2='" + sCust2 + "'";
                        dtFill1 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill1.Rows.Count > 0)
                        {
                            dgvHeader.Rows[iRow].Cells["g_iSONo"].Value = dtFill1.Rows[0]["so_no"].ToString();
                            //dgvHeader.Rows[iRow].Cells["g_iRemark"].Value = "PO Already Uploaded";
                        }
                        else
                        {
                            dgvHeader.Rows[iRow].Cells["g_iSONo"].Value = "";
                            //dgvHeader.Rows[iRow].Cells["g_iRemark"].Value = "";
                        }
                        string sPODate = __poDate.Substring(6, 2) + "-" + __poDate.Substring(4, 2) + "-" + __poDate.Substring(0, 4); ;
                        string sDeliveryDate = __delvDate.Substring(6, 2) + "-" + __delvDate.Substring(4, 2) + "-" + __delvDate.Substring(0, 4);
                        string sExpiredDate = __expDate.Substring(6, 2) + "-" + __expDate.Substring(4, 2) + "-" + __expDate.Substring(0, 4);

                        //DateTime dtime = Convert.ToDateTime(sPODate);
                        //sPODate = dtSODate.Value.ToString("yyyy-MM-dd");
                        if (dtHeader.Rows[lCounter]["paytype"].ToString() == "T" || dtHeader.Rows[lCounter]["paytype"].ToString() == "D" || dtHeader.Rows[lCounter]["paytype"].ToString() == "S")
                        {
                            sTopId = "P0";
                            sTOPVal = "0";
                        }
                        dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value = sPODate;
                        dgvHeader.Rows[iRow].Cells["g_iTOPId"].Value = sTopId;
                        dgvHeader.Rows[iRow].Cells["g_iTOP"].Value = sTOPVal;
                        dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value = dtHeader.Rows[lCounter]["paytype"].ToString();
                        dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value = sLeadtime;
                        dgvHeader.Rows[iRow].Cells["g_iDeliveryDate"].Value = sDeliveryDate;
                        dgvHeader.Rows[iRow].Cells["g_iExpiredDate"].Value = sExpiredDate;

                        iRow++;
                        __No++;

                    }
                }

                iRow = 0;

                if (dgvDetail2.Rows.Count > 0)
                {
                    dgvDetail2.Rows.Clear();
                }
                for (lCounter = 0; lCounter < dgvdata.Rows.Count; lCounter++)
                {
                    sCustNo = dgvdata.Rows[lCounter].Cells[4].Value.ToString();
                    vPONo = dgvdata.Rows[lCounter].Cells[0].Value.ToString();
                    string __entity = dgvdata.Rows[lCounter].Cells[0].Value.ToString();
                    string __branch = dgvdata.Rows[lCounter].Cells[1].Value.ToString();
                    string __salesman = dgvdata.Rows[lCounter].Cells[4].Value.ToString();
                    string __outlet = dgvdata.Rows[lCounter].Cells[5].Value.ToString();


                    string __cust_code = dgvdata.Rows[lCounter].Cells[4].Value.ToString();

                    if (dgvdata.Rows.Count > 0)
                    {
                        string sPOCodex = dgvdata.Rows[lCounter].Cells[6].Value.ToString();

                        dCustPcode = sPOCodex;

                        iConv1 = "0";
                        iConv2 = "0";

                        if (chkEcommerce.Checked)
                        {
                            if (fGetProdAttrPerWHEcommerce(wh1, wh2, __entity, __branch, __cust_code) == false)
                            {
                                dtView.Rows.Clear();
                                keyFound = false;
                                return keyFound;
                            }
                        }
                        else
                        {
                            if (fGetProdAttrPerWH(wh1, wh2, __cust_code) == false)
                            {
                                dtView.Rows.Clear();
                                keyFound = false;
                                return keyFound;
                            }
                        }


                        if (iConv1 == "")
                        {
                            iConv1 = "0";
                        }

                        if (iConv2 == "")
                        {
                            iConv2 = "0";
                        }

                        bExists = false;
                        //if (fPcodeExistsBFI(vPONo, sCustNo) == false)
                        //{
                        //    dtView.Rows.Clear();
                        //    keyFound = false;
                        //    return keyFound;
                        //}

                        //if (dgvdata.Rows[lCounter].Cells[5].Value.ToString() == "")
                        //{
                        //dgvdata.Rows[lCounter].Cells[5].Value = "0";
                        //}

                        //if (dgvdata.Rows[lCounter].Cells[6].Value.ToString() == "")
                        //{
                        //    dgvdata.Rows[lCounter].Cells[6].Value = "0";
                        //}

                        if (dgvdata.Rows[lCounter].Cells[8].Value.ToString() == "")
                        {
                            dgvdata.Rows[lCounter].Cells[8].Value = "0";
                        }

                        int sData5 = 0;
                        int sData6 = 0;
                        int sData7 = Convert.ToInt16(dgvdata.Rows[lCounter].Cells[8].Value);

                        if (sData5 > 0 || sData6 > 0 || sData7 > 0)
                        {
                            CCP = sData5.ToString() + "." + sData6.ToString() + "." + sData7.ToString();
                            fQtyFormat(CCP, Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));
                            CCP = _fQtyFormat;
                            dgvDetail2.Rows.Add();
                            dgvDetail2.Rows[iRow].Cells["g_iDetEntity1"].Value = dgvdata.Rows[lCounter].Cells[0].Value.ToString().Replace("\"", "");
                            dgvDetail2.Rows[iRow].Cells["g_iDetBranch1"].Value = dgvdata.Rows[lCounter].Cells[1].Value.ToString().Replace("\"", "");
                            dgvDetail2.Rows[iRow].Cells["g_iDetPONO1"].Value = dgvdata.Rows[lCounter].Cells[2].Value.ToString().Replace("\"", "");
                            dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value = dgvdata.Rows[lCounter].Cells[4].Value.ToString().Replace("\"", "");
                            dgvDetail2.Rows[iRow].Cells["g_iDetSalesmanID1"].Value = dgvdata.Rows[lCounter].Cells[3].Value.ToString().Replace("\"", "");
                            if (chkmappingproduct.Checked)
                            {
                                dgvDetail2.Rows[iRow].Cells["g_iDetPCode1"].Value = sPcode;
                                dgvDetail2.Rows[iRow].Cells["g_iProductMap1"].Value = dgvdata.Rows[lCounter].Cells[5].Value.ToString();
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(sPcode))
                                {
                                    dgvDetail2.Rows[iRow].Cells["g_iDetPCode1"].Value = dgvdata.Rows[lCounter].Cells[6].Value.ToString();
                                }
                                else
                                {
                                    dgvDetail2.Rows[iRow].Cells["g_iDetPCode1"].Value = sPcode;
                                }
                                dgvDetail2.Rows[iRow].Cells["g_iProductMap1"].Value = "";
                            }
                            //if (sPcode != "")
                            //{

                            //    //dgvDetail2.Rows[iRow].Cells["g_iDetPCode1"].Value = dCustPcode;
                            //}
                            //else
                            //{
                            //    dgvDetail2.Rows[iRow].Cells["g_iDetPCode1"].Value = dCustPcode;
                            //}
                            dgvDetail2.Rows[iRow].Cells["g_iDetPDesc1"].Value = sDesc;
                            dgvDetail2.Rows[iRow].Cells["g_iDetGrade1"].Value = sGrade;
                            dgvDetail2.Rows[iRow].Cells["g_iDetSize1"].Value = sSize;
                            dgvDetail2.Rows[iRow].Cells["g_iDetSled1"].Value = sPrdSLED;
                            dgvDetail2.Rows[iRow].Cells["g_iDetConv11"].Value = iConv1;
                            dgvDetail2.Rows[iRow].Cells["g_iDetConv21"].Value = iConv2;
                            dgvDetail2.Rows[iRow].Cells["g_iDetValue1"].Value = "0";
                            fConvertQty(CCP, Convert.ToInt16(dgvDetail2.Rows[iRow].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[iRow].Cells["g_iDetConv21"].Value));
                            dgvDetail2.Rows[iRow].Cells["g_iDetQty1"].Value = _fConvertQty;
                            dgvDetail2.Rows[iRow].Cells["g_iDetQtyOrd1"].Value = CCP;
                            dgvDetail2.Rows[iRow].Cells["g_iDetTaxCode1"].Value = sTaxCode;
                            dgvDetail2.Rows[iRow].Cells["g_iDetProdLine1"].Value = sProdLine;
                            dgvDetail2.Rows[iRow].Cells["g_iDetTaxPct1"].Value = siTaxPct;
                            dgvDetail2.Rows[iRow].Cells["g_iDetPrice1"].Value = "0";
                            // dgvDetail2.Rows[lCounter].Cells["g_iDetTRSPrice1"].Value = dTotTRS;
                            dgvDetail2.Rows[iRow].Cells["g_iDetStatus1"].Value = sStatus;

                            if (sPcode == "")
                            {
                                dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "Master Product not found";
                            }
                            else if (sStatus != "A")
                            {
                                if (sStatus != "D")
                                {
                                    dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "Product is not Active";
                                }
                                else
                                {
                                    dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "OK";
                                }
                            }
                            else
                            {
                                /** cover salesman produk ga **/
                                dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "OK";

                            }
                            
                            iRow++;
                            //dgvDetail2.Rows[lCounter].Cells["g_iDetTRSUnitPrice1"].Value = dTRSPrice;
                        }


                    }
                }

                // jika di kam dan divisionnya di checklist cari top nya apabila Division Y dan customer TOP N //
                if (isNefoKAM)
                {
                    if (isDivisionTOP)
                    {
                        for (int i = 0; i < dgvDetail2.Rows.Count; i++)
                        {
                            string __entity = "";
                            string __branch = "";
                            string __salesman = "";
                            if (chkEcommerce.Checked)
                            {
                                __entity = dgvDetail2.Rows[i].Cells["g_iDetEntity1"].Value.ToString();
                                __branch = dgvDetail2.Rows[i].Cells["g_iDetBranch1"].Value.ToString();
                                __salesman = dgvDetail2.Rows[i].Cells["g_iDetSalesmanID1"].Value.ToString();
                            }

                            PONO = dgvDetail2.Rows[i].Cells["g_iDetPONO1"].Value.ToString();
                            PCode1 = dgvDetail2.Rows[i].Cells["g_iDetPCode1"].Value.ToString();
                            PCode2 = "000";
                            sGradeT = dgvDetail2.Rows[i].Cells["g_iDetGrade1"].Value.ToString();
                            sSizeT = dgvDetail2.Rows[i].Cells["g_iDetSize1"].Value.ToString();
                            sSled = dgvDetail2.Rows[i].Cells["g_iDetSled1"].Value.ToString();

                            DataView dvTemp = dtHeader.DefaultView;
                            if (chkEcommerce.Checked)
                            {
                                dvTemp.RowFilter = "PONO = '" + PONO + "'";
                            }
                            else
                            {
                                dvTemp.RowFilter = "PO = '" + PONO + "'";
                            }


                            CTEMP1 = "";
                            CTEMP2 = "";
                            DataTable dtTempHeader = dvTemp.ToTable();
                            if (dtTempHeader.Rows.Count > 0)
                            {
                                if (chkEcommerce.Checked)
                                {
                                    CTEMP1 = dtTempHeader.Rows[0]["kdoutlet_trs"].ToString().Trim();
                                    CTEMP2 = "000";
                                    PayType = dtTempHeader.Rows[0]["PayType"].ToString().Trim();
                                }
                                else
                                {
                                    CTEMP1 = dtTempHeader.Rows[0]["CustCode1"].ToString().Trim();
                                    CTEMP2 = dtTempHeader.Rows[0]["CustCode2"].ToString().Trim();
                                    PayType = dtTempHeader.Rows[0]["PayType"].ToString().Trim();
                                }


                                if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() == "OK")
                                {
                                    strSQL = " Select smt_prdline_id ";
                                    strSQL += " FROM IM_PRD_MASTER WITH (NOLOCK) ";
                                    strSQL += " INNER JOIN SO_MAPPING_TOPBYPRDLINE WITH (NOLOCK) ";
                                    strSQL += " ON smt_prdline_id = prm_prd_line_code ";
                                    strSQL += " WHERE  ";
                                    if (chkEcommerce.Checked)
                                    {
                                        strSQL += " smt_entity_id = '" + __entity + "' ";
                                        strSQL += " and smt_branch_id = '" + __branch + "' ";
                                    }
                                    else
                                    {
                                        strSQL += " smt_entity_id = '" + cbEntityBranch.txtCM.Text + "' ";
                                        strSQL += " and smt_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                                    }

                                    strSQL += " and smt_cust_code1 = '" + CTEMP1 + "' ";
                                    strSQL += " and smt_cust_code2 = '" + CTEMP2 + "' ";
                                    strSQL += " and prm_prd_master_code = '" + PCode1 + "' ";

                                    DataTable dtPrdlineTOP = _clsGlobal.ExecDT(strSQL);
                                    if (dtPrdlineTOP.Rows.Count == 0)
                                    {
                                        dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value = "Product Line Belum di Mapping ke Customer";
                                    }

                                }

                                if (chkEcommerce.Checked)
                                {
                                    strSQL = "select cm_top_by_cust from SO_CUST_MASTER WITH (NOLOCK) where cm_entity = '" + __entity + "' and cm_branch = '" + __branch + "' and cm_cust_code1 = '" + CTEMP1 + "' and cm_cust_code2 = '" + CTEMP2 + "'";
                                }
                                else
                                {
                                    strSQL = "select cm_top_by_cust from SO_CUST_MASTER WITH (NOLOCK) where cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' and cm_cust_code1 = '" + CTEMP1 + "' and cm_cust_code2 = '" + CTEMP2 + "'";
                                }

                                DataTable dtCust = _clsGlobal.ExecDT(strSQL);
                                if (dtCust.Rows.Count > 0 && PayType != "T")
                                {
                                    if (dtCust.Rows[0]["cm_top_by_cust"].ToString() != "Y")
                                    {

                                        strSQL = "select SMT.smt_top_id ";
                                        strSQL += " from SO_CUST_MASTER CM WITH (NOLOCK) ";
                                        strSQL += "   LEFT JOIN TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                                        strSQL += "       ON  ";
                                        strSQL += "           TSC.csc_entity          = cm_entity ";
                                        strSQL += "           AND TSC.csc_branch      = CM.cm_branch ";
                                        strSQL += "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 ";
                                        strSQL += "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 ";
                                        strSQL += "   left Join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK) ";
                                        strSQL += "       ON ";
                                        strSQL += "           SMT.smt_entity_id      = CM.cm_entity ";
                                        if (!isMultibranch)
                                        {
                                            strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                        }
                                        else
                                        {
                                            strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                        }
                                        strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                        strSQL += "       ON ";
                                        strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                        strSQL += "           where ";
                                        strSQL += "   prm_prd_master_code = '" + PCode1 + "' ";
                                        strSQL += "   and prm_grade = '" + sGradeT + "' ";
                                        strSQL += "   and prm_prd_size = '" + sSizeT + "' ";
                                        if (chkEcommerce.Checked)
                                        {
                                            strSQL += "   and TSC.csc_salesman_id = '" + __salesman + "' ";
                                            strSQL += "   and cm_entity = '" + __entity + "' ";
                                            strSQL += "   and cm_branch = '" + __branch + "' ";
                                        }
                                        else
                                        {
                                            strSQL += "   and TSC.csc_salesman_id = '" + txtSalesman.Text.Trim() + "' ";
                                            strSQL += "   and cm_entity = '" + cbEntityBranch.txtCM.Text + "' ";
                                            strSQL += "   and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                                        }
                                        strSQL += "   and CM.cm_cust_code1 = '" + CTEMP1 + "' ";
                                        strSQL += "   and CM.cm_cust_code2 = '" + CTEMP2 + "'";
                                        DataTable dtNewProduct = _clsGlobal.ExecDT(strSQL);
                                        if (dtNewProduct.Rows.Count > 0)
                                        {
                                            if (string.IsNullOrEmpty(dtNewProduct.Rows[0]["smt_top_id"].ToString().Trim()))
                                            {
                                                //if (dtNewProduct.Rows[0]["pcodelinenew"].ToString().Trim() != dtNewProduct.Rows[0]["plinemap"].ToString().Trim())
                                                //{
                                                if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() == "OK")
                                                {
                                                    dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value = "Mapping TOP Belum Disetting.";
                                                }
                                                //}
                                            }
                                            else
                                            {
                                                dgvDetail2.Rows[i].Cells["g_iDetTOPDiv1"].Value = dtNewProduct.Rows[0]["smt_top_id"].ToString().Trim();
                                            }
                                            //else
                                            //{
                                            //    if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() == "OK")
                                            //    {
                                            //        dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value = "Mapping TOP Belum Disetting.";
                                            //    }
                                            //}
                                        }
                                        else
                                        {
                                            if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() == "OK")
                                            {
                                                dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value = "Mapping TOP Belum Disetting.";
                                            }
                                        }
                                    }
                                }
                                else if (PayType == "T")
                                {
                                    //jika tunai
                                    dgvDetail2.Rows[i].Cells["g_iDetTOPDiv1"].Value = "P0";

                                }

                            }



                        }
                    }
                }


                #endregion

                DataTable dt = dgvDetail2.DataSource as DataTable;

                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }
        private bool fGetProdAttrPerWH(string wh1, string wh2, string cust_code)
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " SELECT  prm_prd_line_code,ISNULL(pl_prd_line_desc, '') AS pl_prd_line_desc, prm_prd_master_code,prm_grade,prm_prd_size,ISNULL(prm_prd_desc, '') AS prm_prd_desc,prm_conversion_purc,";
            strSQL = strSQL + " prm_conversion_sales, ";
            if (isNefoKAM)
            {
                strSQL = strSQL + " case when cust_tax_code='PPN0' THEN cust_tax_code ELSE  stc_tax_code END prm_tax_code, CASE WHEN cust_tax_code='PPN0' THEN isnull(cust_tax_amount,0) ELSE  isnull(stc_tax_amount, 0) END stc_tax_amount, ";
            }
            else
            {
                strSQL = strSQL + " prm_tax_code,case when isnull(excludeTax,'N') = 'Y' then 0 else isnull(stc_tax_amount, 0) end stc_tax_amount, ";
            }

            strSQL = strSQL + " isnull(prm_status, '') prm_status,isnull(wlb_loc_Id1,'X') wlb_loc_id, prm_prd_line_code,mcps_prd_sled AS [prm_prd_sled] From IM_PRD_MASTER WITH (NOLOCK) ";
            strSQL = strSQL + " LEFT OUTER JOIN IM_STOCK_BALANCE WITH (NOLOCK) ON ";
            strSQL = strSQL + " wlb_loc_Id1 = '" + wh1 + "' AND wlb_loc_Id2 = '" + wh2 + "' AND wlb_entity_id='" + cbEntityBranch.txtCM.Text + "' AND wlb_branch_id ='" + cbEntityBranch.txtBranchIdCM.Text + "' AND wlb_prd_master_code = prm_prd_master_code AND wlb_grade = prm_grade ";
            strSQL = strSQL + " AND wlb_prd_size = prm_prd_size AND isnull(wlb_stock_sts, 'N') = 'N' LEFT JOIN IM_PRD_LINE WITH (NOLOCK) ON prm_prd_line_code = pl_prd_line_code ";
            strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
            strSQL = strSQL + " LEFT JOIN TBL_MAP_CUST_PRD_SLED WITH (NOLOCK) ON mcps_entity_id='" + cbEntityBranch.txtCM.Text + "' AND mcps_branch_id='" + cbEntityBranch.txtBranchIdCM.Text + "' AND mcps_cust_code1='" + cust_code + "' AND mcps_prd_code=prm_prd_master_code AND mcps_prd_grade=prm_grade";
            strSQL = strSQL + " AND mcps_prd_size = prm_prd_size ";
            //if (isNefoKAM)
            //{
            //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON  CAST(CONVERT(DATE, '" + dtSODate.Value.ToString("yyyyMMdd") + "', 112) AS DATETIME) >=cast(stc_tax_valid_from as datetime) ";
            //    strSQL = strSQL + "  AND CAST(CONVERT(DATE,'" + dtSODate.Value.ToString("yyyyMMdd") + "', 112) AS DATETIME) <=cast(stc_tax_valid_to as datetime)  ";
            //}
            //else
            //{
            //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
            //}


            if (isNefoKAM)
            {
                /** customer PPN **/
                strSQL = strSQL + " LEFT JOIN  (SELECT cm_entity,cm_branch, cm_ship_tax_id AS cust_tax_code , stc_tax_amount as cust_tax_amount FROM SO_CUST_MASTER WITH (NOLOCK)  ";
                strSQL = strSQL + "  INNER JOIN dbo.SO_TAX_CODE  WITH (NOLOCK) ON cm_ship_tax_id = stc_tax_code  ";
                strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + cust_code + "'";
                //strSQL = strSQL + "  AND cm_cust_code2 = '" + cust_code2 + "'";
                strSQL = strSQL + "  AND cm_entity = '" + cbEntityBranch.txtCM.Text + "'";
                strSQL = strSQL + "  AND cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                strSQL = strSQL + " )CUST ON cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                /** End customer PPN **/
            }
            else
            {
                /** exclude tax **/
                strSQL = strSQL + " LEFT JOIN ( SELECT gh_group_menu_id as __entity,gh_parent_menu __branch,gh_function_code excludeTax FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE GH_FUNCTION_NAME = 'Exclude-TAX-KAM' AND GH_SYS = 'S' ) FLAGTAX ON __entity = '" + cbEntityBranch.txtCM.Text + "' and __branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                /** end exclude tax **/
            }

            strSQL = strSQL + " WHERE  prm_prd_master_code ='" + dCustPcode + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                if (isNefoKAM)
                {
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["prm_tax_code"].ToString()))
                    {
                        MessageBox.Show("Parameter untuk so_tax_code valid date '" + dtSODate.Value.ToString("yyyyMMdd") + "' tidak ditemukan!.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }
                    else
                    {
                        keyFound = true;
                    }
                }
                sLineCode = dtFill1.Rows[0]["prm_prd_line_code"].ToString().Trim();
                sLineDesc = dtFill1.Rows[0]["pl_prd_line_desc"].ToString().Trim();
                sPcode = dtFill1.Rows[0]["prm_prd_master_code"].ToString().Trim();
                sGrade = dtFill1.Rows[0]["prm_grade"].ToString().Trim();
                sSize = dtFill1.Rows[0]["prm_prd_size"].ToString().Trim();
                sSled = dtFill1.Rows[0]["prm_prd_sled"].ToString().Trim();  
                sDesc = dtFill1.Rows[0]["prm_prd_desc"].ToString().Trim();
                iConv1 = dtFill1.Rows[0]["prm_conversion_purc"].ToString().Trim();
                iConv2 = dtFill1.Rows[0]["prm_conversion_sales"].ToString().Trim();
                sTaxCode = dtFill1.Rows[0]["prm_tax_code"].ToString().Trim();
                siTaxPct = dtFill1.Rows[0]["stc_tax_amount"].ToString().Trim();
                sStatus = dtFill1.Rows[0]["prm_status"].ToString().Trim();
                sWHLoc = dtFill1.Rows[0]["wlb_loc_id"].ToString().Trim();
                sProdLine = dtFill1.Rows[0]["prm_prd_line_code"].ToString().Trim();

            }
            else
            {
                sLineCode = "";
                sLineDesc = "";
                sPcode = "";
                sGrade = "";
                sSize = "";
                sSled = "";
                sDesc = "";
                iConv1 = "";
                iConv2 = "";
                sTaxCode = "";
                siTaxPct = "";
                sStatus = "";
                sWHLoc = "";
                sProdLine = "";

            }

            keyFound = true;
            return keyFound;

        }
        private bool fGetProdAttrPerWHEcommerce(string wh1, string wh2, string entity, string branch, string cust_code)
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " SELECT  prm_prd_line_code,ISNULL(pl_prd_line_desc, '') AS pl_prd_line_desc, prm_prd_master_code,prm_grade,prm_prd_size,ISNULL(prm_prd_desc, '') AS prm_prd_desc,prm_conversion_purc,";
            strSQL = strSQL + " prm_conversion_sales, ";
            if (isNefoKAM)
            {
                strSQL = strSQL + " case when cust_tax_code='PPN0' THEN cust_tax_code ELSE  stc_tax_code END prm_tax_code, CASE WHEN cust_tax_code='PPN0' THEN isnull(cust_tax_amount,0) ELSE  isnull(stc_tax_amount, 0) END stc_tax_amount, ";
            }
            else
            {
                strSQL = strSQL + " prm_tax_code,case when isnull(excludeTax,'N') = 'Y' then 0 else isnull(stc_tax_amount, 0) end stc_tax_amount, ";
            }


            strSQL = strSQL + " isnull(prm_status, '') prm_status,isnull(wlb_loc_Id1,'X') wlb_loc_id, prm_prd_line_code,mcps_prd_sled AS [prm_prd_sled] From IM_PRD_MASTER WITH (NOLOCK) ";
            strSQL = strSQL + " LEFT OUTER JOIN IM_STOCK_BALANCE WITH (NOLOCK) ON ";
            strSQL = strSQL + " wlb_loc_Id1 = '" + wh1 + "' AND wlb_loc_Id2 = '" + wh2 + "' AND wlb_entity_id='" + entity + "' AND wlb_branch_id ='" + branch + "' AND wlb_prd_master_code = prm_prd_master_code AND wlb_grade = prm_grade ";
            strSQL = strSQL + " AND wlb_prd_size = prm_prd_size AND isnull(wlb_stock_sts, 'N') = 'N' LEFT JOIN IM_PRD_LINE WITH (NOLOCK) ON prm_prd_line_code = pl_prd_line_code ";
            strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
            strSQL = strSQL + " LEFT JOIN TBL_MAP_CUST_PRD_SLED WITH (NOLOCK) ON mcps_entity_id='" + entity + "' AND mcps_branch_id='" + branch + "' AND mcps_cust_code1='" + cust_code + "' AND mcps_prd_code=prm_prd_master_code AND mcps_prd_grade=prm_grade";
            strSQL = strSQL + " AND mcps_prd_size = prm_prd_size ";
            //if (isNefoKAM)
            //{
            //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON  CAST(CONVERT(DATE, '" + dtSODate.Value.ToString("yyyyMMdd") + "', 112) AS DATETIME) >=cast(stc_tax_valid_from as datetime) ";
            //    strSQL = strSQL + "  AND CAST(CONVERT(DATE,'"+ dtSODate.Value.ToString("yyyyMMdd") + "', 112) AS DATETIME) <=cast(stc_tax_valid_to as datetime)  ";
            //}
            //else
            //{
            //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
            //}   


            if (isNefoKAM)
            {
                /** customer PPN **/
                strSQL = strSQL + " LEFT JOIN  (SELECT cm_entity,cm_branch, cm_ship_tax_id AS cust_tax_code , stc_tax_amount as cust_tax_amount FROM SO_CUST_MASTER WITH (NOLOCK)  ";
                strSQL = strSQL + "  INNER JOIN dbo.SO_TAX_CODE  WITH (NOLOCK) ON cm_ship_tax_id = stc_tax_code  ";
                strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + cust_code + "'";
                // strSQL = strSQL + "  AND cm_cust_code2 = '" + cust_code_2 + "'";
                strSQL = strSQL + "  AND cm_entity = '" + entity + "'";
                strSQL = strSQL + "  AND cm_branch = '" + branch + "'";
                strSQL = strSQL + " )CUST ON cm_entity = '" + entity + "' and cm_branch = '" + branch + "' ";
                /** End customer PPN **/
            }
            else
            {
                /** exclude tax **/
                strSQL = strSQL + " LEFT JOIN ( SELECT gh_group_menu_id as __entity,gh_parent_menu __branch,gh_function_code excludeTax FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE GH_FUNCTION_NAME = 'Exclude-TAX-KAM' AND GH_SYS = 'S' ) FLAGTAX ON __entity = '" + entity + "' and __branch = '" + branch + "' ";
                /** end exclude tax **/
            }

            strSQL = strSQL + " WHERE  prm_prd_master_code ='" + dCustPcode + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                if (isNefoKAM)
                {
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["prm_tax_code"].ToString()))
                    {
                        MessageBox.Show("Parameter untuk so_tax_code valid date '" + dtSODate.Value.ToString("yyyyMMdd") + "' tidak ditemukan!.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }
                    else
                    {
                        keyFound = true;
                    }
                }


                sLineCode = dtFill1.Rows[0]["prm_prd_line_code"].ToString().Trim();
                sLineDesc = dtFill1.Rows[0]["pl_prd_line_desc"].ToString().Trim();
                sPcode = dtFill1.Rows[0]["prm_prd_master_code"].ToString().Trim();
                sGrade = dtFill1.Rows[0]["prm_grade"].ToString().Trim();
                sSize = dtFill1.Rows[0]["prm_prd_size"].ToString().Trim();
                sDesc = dtFill1.Rows[0]["prm_prd_desc"].ToString().Trim();
                iConv1 = dtFill1.Rows[0]["prm_conversion_purc"].ToString().Trim();
                iConv2 = dtFill1.Rows[0]["prm_conversion_sales"].ToString().Trim();
                sTaxCode = dtFill1.Rows[0]["prm_tax_code"].ToString().Trim();
                siTaxPct = dtFill1.Rows[0]["stc_tax_amount"].ToString().Trim();
                sStatus = dtFill1.Rows[0]["prm_status"].ToString().Trim();
                sWHLoc = dtFill1.Rows[0]["wlb_loc_id"].ToString().Trim();
                sProdLine = dtFill1.Rows[0]["prm_prd_line_code"].ToString().Trim();
                sPrdSLED = dtFill1.Rows[0]["prm_prd_sled"].ToString().Trim();
            }
            else
            {
                sLineCode = "";
                sLineDesc = "";
                sPcode = "";
                sGrade = "";
                sSize = "";
                sDesc = "";
                iConv1 = "";
                iConv2 = "";
                sTaxCode = "";
                siTaxPct = "";
                sStatus = "";
                sWHLoc = "";
                sProdLine = "";
                sPrdSLED = "";

            }

            keyFound = true;
            return keyFound;

        }
        private bool fGetCustNOPerSalesman(string salesman)
        {
            bool keyFound = false;
            DataTable dtFill1 = new DataTable();

            strSQL = "SELECT cm_cust_code1, cm_cust_code2, cm_cust_name, isnull(csc_salesman_id,'X') csc_salesman_id,isnull(pptc_term_code,'') top_id,isnull(cm_top,'') cm_top,isnull(cm_lead_time,0) cm_lead_time,";
            strSQL = strSQL + " isnull(cm_payment_type, '') as cm_payment_type,sgm_spgm_name From SO_CUST_MASTER WITH (NOLOCK) LEFT OUTER JOIN TBL_SD_CUSTCOVER WITH (NOLOCK) ON  cm_cust_code1 = csc_cust_code1 AND cm_cust_code2 = csc_cust_code2 and cm_entity = csc_entity and cm_branch = csc_branch INNER JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON  pptc_no_of_days = cm_top ";
            strSQL = strSQL + " LEFT JOIN SO_SPG_GIRL_MAN WITH (NOLOCK) ON csc_entity = sgm_entity_id and csc_branch = sgm_branch_id and csc_salesman_id = sgm_spgm_id ";
            strSQL = strSQL + " WHERE csc_salesman_id = '" + salesman + "' AND cm_active_flag <> 'D' AND cm_cust_code1 = '" + sCustNo + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                sCust1 = dtFill1.Rows[0]["cm_cust_code1"].ToString();
                sCust2 = dtFill1.Rows[0]["cm_cust_code2"].ToString();
                sCustName = dtFill1.Rows[0]["cm_cust_name"].ToString();
                sPayType = dtFill1.Rows[0]["cm_payment_type"].ToString();
                sTopId = dtFill1.Rows[0]["top_id"].ToString();
                sTOPVal = dtFill1.Rows[0]["cm_top"].ToString();
                sLeadtime = dtFill1.Rows[0]["cm_lead_time"].ToString();
                sSalesmanCode = salesman;
                sSalesmanName = dtFill1.Rows[0]["sgm_spgm_name"].ToString();
            }
            else
            {
                sCust1 = "";
                sCust2 = "";
                sCustName = "";
                sPayType = "";
                sTopId = "";
                sTOPVal = "";
                sLeadtime = "";
            }

            DataTable dtFill2 = new DataTable();
            strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            DataTable dtFillDivision = new DataTable();
            strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name='TOPBYDIVISION' and gh_sys='H' and gh_function_code='Y' ";
            dtFillDivision = _clsGlobal.ExecDT(strSQL);
            #region TOP TEAM SALESMAN
            if (dtFill2.Rows.Count > 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = "select coalesce(cm_top_by_cust,'N') as cm_top_by_cust,cm_top_id,cm_payment_type, coalesce(cg_cust_top,'') as topnya,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment ";
                strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) inner join SO_CUST_GROUP WITH (NOLOCK) on cm_cust_group = cg_cust_group left join PO_PAYMENT_TERM_CODES on cg_cust_top=pptc_term_code ";
                strSQL = strSQL + " where cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text.ToString() + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                dtFill3 = _clsGlobal.ExecDT(strSQL);
                if (dtFill3.Rows.Count > 0)
                {

                    if (dtFill3.Rows[0]["cm_payment_type"].ToString() == "T") //'langsung TOP 0
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select  pptc_term_code,pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES WITH (NOLOCK) where pptc_term_code='P0'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else if (dtFill3.Rows[0]["cm_top_by_cust"].ToString() == "Y")
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select pptc_term_code,pptc_no_of_days ";
                        strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON cm_top_id=pptc_term_code";
                        strSQL = strSQL + " where cm_cust_code1='" + sCust1 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else
                    {
                        if (dtFill3.Rows[0]["topnya"].ToString() != "")
                        {
                            sPayType = dtFill3.Rows[0]["top_payment"].ToString();
                        }
                        else
                        {
                            DataTable dtFill6 = new DataTable();
                            strSQL = "select pptc_no_of_days,sgl_group_top,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from SO_SPG_GIRL_MAN WITH (NOLOCK) ";
                            strSQL = strSQL + " left join SO_SPG_GROUP WITH (NOLOCK) on sgm_spgm_group = sgl_group_code inner join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on sgl_group_top=pptc_term_code ";
                            strSQL = strSQL + " where sgm_spgm_id='" + txtSalesman.Text + "' and sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' and sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                            dtFill6 = _clsGlobal.ExecDT(strSQL);
                            if (dtFill6.Rows.Count > 0) //ikut team salesman
                            {
                                sTopId = dtFill6.Rows[0]["sgl_group_top"].ToString();
                                sTOPVal = dtFill6.Rows[0]["pptc_no_of_days"].ToString();
                            }
                        }
                    }
                }
            }
            #endregion

            #region TOP BY DIVISION
            else if (dtFillDivision.Rows.Count > 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = "select coalesce(cm_top_by_cust,'N') as cm_top_by_cust,cm_top_id,cm_payment_type, coalesce(cg_cust_top,'') as topnya,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment ";
                strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) inner join SO_CUST_GROUP WITH (NOLOCK) on cm_cust_group = cg_cust_group left join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on cg_cust_top=pptc_term_code ";
                strSQL = strSQL + " where cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                dtFill3 = _clsGlobal.ExecDT(strSQL);
                if (dtFill3.Rows.Count > 0)
                {

                    if (dtFill3.Rows[0]["cm_payment_type"].ToString() == "T") //'langsung TOP 0
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select  pptc_term_code,pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES WITH (NOLOCK) where pptc_term_code='P0'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else if (dtFill3.Rows[0]["cm_top_by_cust"].ToString() == "Y")
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select pptc_term_code,pptc_no_of_days ";
                        strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON cm_top_id=pptc_term_code";
                        strSQL = strSQL + " where cm_cust_code1='" + sCust1 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else
                    {
                        strSQL = "select distinct top 1 PM.prm_prd_line_code pcodelinenew,SP.ssp_spggroup_code,SP.ssp_prdline_code plinegrup,SMT.smt_prdline_id plinemap,SMT.smt_top_id, PTC.pptc_no_of_days " +
                                 "   from IM_PRD_MASTER PM WITH (NOLOCK) " +
                                 "   left join TBL_SD_SPGGROUP_PRDLINE SP WITH (NOLOCK) " +
                                 "       ON  " +
                                 "           SP.ssp_prdline_code = PM.prm_prd_line_code " +
                                 "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK) " +
                                 "       ON	" +
                                 "           SPG.sgm_spgm_group = SP.ssp_spggroup_code " +
                                 "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) " +
                                 "       ON " +
                                 "           TSC.csc_entity = SPG.sgm_entity_id " +
                                 "           AND TSC.csc_branch = SPG.sgm_branch_id " +
                                 "           AND TSC.csc_salesman_id = SPG.sgm_spgm_id " +
                                 "   left join SO_CUST_MASTER CM WITH (NOLOCK) " +
                                 "       ON " +
                                 "           CM.cm_cust_code1 = TSC.csc_cust_code1 " +
                                 "           and cm_cust_code2 = TSC.csc_cust_code2 " +
                                 "           and cm_entity = TSC.csc_entity " +
                                 "           and cm_branch = TSC.csc_branch " +
                                 "   left Join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK) " +
                                 "       ON " +
                                 "           SMT.smt_entity_id = CM.cm_entity " +
                                 "           and SMT.smt_branch_id = CM.cm_branch " +
                                 "           and SMT.smt_cust_code1 = CM.cm_cust_code1 " +
                                 "           and SMT.smt_cust_code2 = CM.cm_cust_code2 " +
                                 "           and SMT.smt_prdline_id = PM.prm_prd_line_code " +
                                 "   Left Join PO_PAYMENT_TERM_CODES PTC " +
                                 "       ON " +
                                 "           PTC.pptc_term_code = SMT.smt_top_id " +
                                 "   where " +
                                 "   SPG.sgm_spgm_id = '" + txtSalesman.Text.Trim() + "' " +
                                 "   and SP.ssp_del_flag = 'N' " +
                                 "   and SPG.sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' " +
                                 "   and SPG.sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' " +
                                 "   and CM.cm_cust_code1 = '" + sCust1 + "' " +
                                 "   and CM.cm_top_by_cust = 'N' " +
                                 "   and CM.cm_cust_code2 = '" + sCust2 + "' and smt_top_id is not null ";
                        DataTable dtNewProduct = _clsGlobal.ExecDT(strSQL);
                        if (dtNewProduct.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dtNewProduct.Rows[0]["smt_top_id"].ToString()))
                            {
                                sTopId = dtNewProduct.Rows[0]["smt_top_id"].ToString();
                                sTOPVal = dtNewProduct.Rows[0]["pptc_no_of_days"].ToString();
                            }
                            else
                            {
                                MessageBox.Show("Mapping TOP Belum Disetting untuk Customer '" + sCust1 + "'.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Mapping TOP Belum Disetting untuk Customer '" + sCust1 + "'.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }
                }
            }
            #endregion

            keyFound = true;
            return keyFound;
        }

        private bool fGetCustNOPerSalesman(string salesman, string entity, string branch)
        {
            bool keyFound = false;
            DataTable dtFill1 = new DataTable();

            strSQL = "SELECT cm_cust_code1, cm_cust_code2, cm_cust_name, isnull(csc_salesman_id,'X') csc_salesman_id,isnull(pptc_term_code,'') top_id,isnull(cm_top,'') cm_top,isnull(cm_lead_time,0) cm_lead_time,";
            strSQL = strSQL + " isnull(cm_payment_type, '') as cm_payment_type,sgm_spgm_name From SO_CUST_MASTER WITH (NOLOCK) LEFT OUTER JOIN TBL_SD_CUSTCOVER WITH (NOLOCK) ON  cm_cust_code1 = csc_cust_code1 AND cm_cust_code2 = csc_cust_code2 and cm_entity = csc_entity and cm_branch = csc_branch INNER JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON  pptc_no_of_days = cm_top ";
            strSQL = strSQL + " LEFT JOIN SO_SPG_GIRL_MAN WITH (NOLOCK) ON csc_entity = sgm_entity_id and csc_branch = sgm_branch_id and csc_salesman_id = sgm_spgm_id ";
            strSQL = strSQL + " WHERE csc_salesman_id = '" + salesman + "' AND cm_active_flag <> 'D' AND cm_cust_code1 = '" + sCustNo + "' and cm_entity = '" + entity + "' and cm_branch = '" + branch + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                sCust1 = dtFill1.Rows[0]["cm_cust_code1"].ToString();
                sCust2 = dtFill1.Rows[0]["cm_cust_code2"].ToString();
                sCustName = dtFill1.Rows[0]["cm_cust_name"].ToString();
                sPayType = dtFill1.Rows[0]["cm_payment_type"].ToString();
                sTopId = dtFill1.Rows[0]["top_id"].ToString();
                sTOPVal = dtFill1.Rows[0]["cm_top"].ToString();
                sLeadtime = dtFill1.Rows[0]["cm_lead_time"].ToString();
                sSalesmanCode = salesman;
                sSalesmanName = dtFill1.Rows[0]["sgm_spgm_name"].ToString();
            }
            else
            {
                sCust1 = "";
                sCust2 = "";
                sCustName = "";
                sPayType = "";
                sTopId = "";
                sTOPVal = "";
                sLeadtime = "";
            }

            DataTable dtFill2 = new DataTable();
            strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            DataTable dtFillDivision = new DataTable();
            strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name='TOPBYDIVISION' and gh_sys='H' and gh_function_code='Y' ";
            dtFillDivision = _clsGlobal.ExecDT(strSQL);
            #region TOP TEAM SALESMAN
            if (dtFill2.Rows.Count > 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = "select coalesce(cm_top_by_cust,'N') as cm_top_by_cust,cm_top_id,cm_payment_type, coalesce(cg_cust_top,'') as topnya,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment ";
                strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) inner join SO_CUST_GROUP WITH (NOLOCK) on cm_cust_group = cg_cust_group left join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on cg_cust_top=pptc_term_code ";
                strSQL = strSQL + " where cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "' and cm_entity = '" + entity + "' and cm_branch = '" + branch + "'";
                dtFill3 = _clsGlobal.ExecDT(strSQL);
                if (dtFill3.Rows.Count > 0)
                {

                    if (dtFill3.Rows[0]["cm_payment_type"].ToString() == "T") //'langsung TOP 0
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select  pptc_term_code,pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES WITH (NOLOCK) where pptc_term_code='P0'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else if (dtFill3.Rows[0]["cm_top_by_cust"].ToString() == "Y")
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select pptc_term_code,pptc_no_of_days ";
                        strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON cm_top_id=pptc_term_code";
                        strSQL = strSQL + " where cm_cust_code1='" + sCust1 + "' and cm_entity = '" + entity + "' and cm_branch = '" + branch + "'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else
                    {
                        if (dtFill3.Rows[0]["topnya"].ToString() != "")
                        {
                            sPayType = dtFill3.Rows[0]["top_payment"].ToString();
                        }
                        else
                        {
                            DataTable dtFill6 = new DataTable();
                            strSQL = "select pptc_no_of_days,sgl_group_top,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from SO_SPG_GIRL_MAN WITH (NOLOCK) ";
                            strSQL = strSQL + " left join SO_SPG_GROUP WITH (NOLOCK) on sgm_spgm_group = sgl_group_code inner join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on sgl_group_top=pptc_term_code ";
                            strSQL = strSQL + " where sgm_spgm_id='" + salesman + "' and sgm_entity_id = '" + entity + "' and sgm_branch_id = '" + branch + "'";
                            dtFill6 = _clsGlobal.ExecDT(strSQL);
                            if (dtFill6.Rows.Count > 0) //ikut team salesman
                            {
                                sTopId = dtFill6.Rows[0]["sgl_group_top"].ToString();
                                sTOPVal = dtFill6.Rows[0]["pptc_no_of_days"].ToString();
                            }
                        }
                    }
                }
            }
            #endregion

            #region TOP BY DIVISION
            else if (dtFillDivision.Rows.Count > 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = "select coalesce(cm_top_by_cust,'N') as cm_top_by_cust,cm_top_id,cm_payment_type, coalesce(cg_cust_top,'') as topnya,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment ";
                strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) inner join SO_CUST_GROUP WITH (NOLOCK)  on cm_cust_group = cg_cust_group left join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on cg_cust_top=pptc_term_code ";
                strSQL = strSQL + " where cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "' and cm_entity = '" + entity + "' and cm_branch = '" + branch + "' ";
                dtFill3 = _clsGlobal.ExecDT(strSQL);
                if (dtFill3.Rows.Count > 0)
                {

                    if (dtFill3.Rows[0]["cm_payment_type"].ToString() == "T") //'langsung TOP 0
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select  pptc_term_code,pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES WITH (NOLOCK) where pptc_term_code='P0'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else if (dtFill3.Rows[0]["cm_top_by_cust"].ToString() == "Y")
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select pptc_term_code,pptc_no_of_days ";
                        strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK)";
                        strSQL = strSQL + " JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON cm_top_id=pptc_term_code";
                        strSQL = strSQL + " where cm_cust_code1='" + sCust1 + "' and cm_entity = '" + entity + "' and cm_branch = '" + branch + "'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else
                    {
                        strSQL = "select distinct top 1 PM.prm_prd_line_code pcodelinenew,SP.ssp_spggroup_code,SP.ssp_prdline_code plinegrup,SMT.smt_prdline_id plinemap,SMT.smt_top_id, PTC.pptc_no_of_days ";
                        strSQL += "   from IM_PRD_MASTER PM WITH (NOLOCK) ";
                        strSQL += "   left join TBL_SD_SPGGROUP_PRDLINE SP WITH (NOLOCK) ";
                        strSQL += "       ON  ";
                        strSQL += "           SP.ssp_prdline_code = PM.prm_prd_line_code ";
                        strSQL += "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK) ";
                        strSQL += "       ON	";
                        strSQL += "           SPG.sgm_spgm_group = SP.ssp_spggroup_code ";
                        strSQL += "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           TSC.csc_entity = SPG.sgm_entity_id ";
                        strSQL += "           AND TSC.csc_branch = SPG.sgm_branch_id ";
                        strSQL += "           AND TSC.csc_salesman_id = SPG.sgm_spgm_id ";
                        strSQL += "   left join SO_CUST_MASTER CM WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           CM.cm_cust_code1 = TSC.csc_cust_code1 ";
                        strSQL += "           and cm_cust_code2 = TSC.csc_cust_code2 ";
                        strSQL += "           and cm_entity = TSC.csc_entity ";
                        strSQL += "           and cm_branch = TSC.csc_branch ";
                        strSQL += "   left Join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           SMT.smt_entity_id = CM.cm_entity ";
                        //strSQL += "           and SMT.smt_branch_id = CM.cm_branch ";
                        //strSQL += "           and SMT.smt_cust_code1 = CM.cm_cust_code1 ";
                        //strSQL += "           and SMT.smt_cust_code2 = CM.cm_cust_code2 ";
                        //strSQL += "           and SMT.smt_prdline_id = PM.prm_prd_line_code ";
                        if (!isMultibranch)
                        {
                            strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                        }
                        else
                        {
                            strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                        }
                        strSQL += "   Left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                        strSQL += "   where ";
                        strSQL += "   SPG.sgm_spgm_id = '" + salesman + "' ";
                        strSQL += "   and SP.ssp_del_flag = 'N' ";
                        strSQL += "   and SPG.sgm_entity_id = '" + entity + "' ";
                        strSQL += "   and SPG.sgm_branch_id = '" + branch + "' ";
                        strSQL += "   and CM.cm_cust_code1 = '" + sCust1 + "' ";
                        strSQL += "   and CM.cm_top_by_cust = 'N' ";
                        strSQL += "   and CM.cm_cust_code2 = '" + sCust2 + "' and smt_top_id is not null ";
                        DataTable dtNewProduct = _clsGlobal.ExecDT(strSQL);
                        if (dtNewProduct.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dtNewProduct.Rows[0]["smt_top_id"].ToString()))
                            {
                                sTopId = dtNewProduct.Rows[0]["smt_top_id"].ToString();
                                sTOPVal = dtNewProduct.Rows[0]["pptc_no_of_days"].ToString();
                            }
                            else
                            {
                                MessageBox.Show("Mapping TOP Belum Disetting untuk Customer '" + sCust1 + "'.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Mapping TOP Belum Disetting untuk Customer '" + sCust1 + "'.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }
                }
            }
            #endregion

            keyFound = true;
            return keyFound;
        }

        private bool fViewDetail(int iRow)
        {
            bool keyFound = false;
            string sPono = "";
            string sCustNo = "";
            string sPayType = "";
            string sEntity = "";
            string sBranch = "";
            string sSalesman = "";
            string sOutlet = "";

            sPono = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString().Replace("\"", "");
            sCustNo = dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString().Replace("\"", "");
            sPayType = dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value.ToString().Replace("\"", "");

            if (chkEcommerce.Checked)
            {
                sEntity = dgvHeader.Rows[iRow].Cells["g_iEntity"].Value.ToString().Replace("\"", "");
                sBranch = dgvHeader.Rows[iRow].Cells["g_iBranch"].Value.ToString().Replace("\"", "");
                sSalesman = dgvHeader.Rows[iRow].Cells["g_iSalesmanCode"].Value.ToString().Replace("\"", "");
            }


            dgvDetail.Rows.Clear();
            int RowD = 0;
            int __DetNo = 1;
            for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
            {
                if (chkEcommerce.Checked)
                {
                    if (
                        dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() == sPono
                        && dgvDetail2.Rows[lCounter].Cells["g_iDetEntity1"].Value.ToString() == sEntity
                        && dgvDetail2.Rows[lCounter].Cells["g_iDetBranch1"].Value.ToString() == sBranch
                        && dgvDetail2.Rows[lCounter].Cells["g_iDetSalesmanID1"].Value.ToString() == sSalesman
                        && dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value.ToString() == sCustNo
                        )
                    {

                        dgvDetail.Rows.Add();
                        if (chkEcommerce.Checked)
                        {
                            dgvDetail.Rows[RowD].Cells["g_iDetEntity"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetEntity1"].Value.ToString().Replace("\"", "");
                            dgvDetail.Rows[RowD].Cells["g_iDetBranch"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetBranch1"].Value.ToString().Replace("\"", "");
                            dgvDetail.Rows[RowD].Cells["g_iDetSalesman"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetSalesmanID1"].Value.ToString().Replace("\"", "");
                            dgvDetail.Rows[RowD].Cells["g_iDetCustNo"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value.ToString().Replace("\"", "");
                            if (chkmappingproduct.Checked)
                            {
                                dgvDetail.Rows[RowD].Cells["g_iDetPCodeMap"].Value = dgvDetail2.Rows[lCounter].Cells["g_iProductMap1"].Value.ToString().Replace("\"", "");
                            }
                        }
                        dgvDetail.Rows[RowD].Cells["g_iDetNo"].Value = __DetNo;
                        dgvDetail.Rows[RowD].Cells["g_iDetPONO"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString().Replace("\"", "");
                        //di VB6 di tutup//dgvDetail.Rows[RowD].Cells["g_iDetCustNo"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value.ToString().Replace("\"", "");
                        dgvDetail.Rows[RowD].Cells["g_iDetPCode"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetPDesc"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetPDesc1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetGrade"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetSize"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetSled"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetConv1"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetConv2"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetValue"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetValue1"].Value.ToString()).ToString("#,##0.00");
                        dgvDetail.Rows[RowD].Cells["g_iDetValue"].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dgvDetail.Rows[RowD].Cells["g_iDetQty"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString()).ToString("#,##0");
                        dgvDetail.Rows[RowD].Cells["g_iDetQty"].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dgvDetail.Rows[RowD].Cells["g_iDetQtyOrd"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetTaxCode"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetTaxPct"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString();
                        if (isNefoKAM)
                        {
                            dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()).ToString("#,##0.0000");
                        }
                        else
                        {
                            dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()).ToString("#,##0.00");
                        }
                        //dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()).ToString("#,##0.00");
                        dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        //dgvDetail.Rows[RowD].Cells["g_iDetTRSPrice"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTRSPrice1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetStatus"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetStatus1"].Value.ToString();
                        //dgvDetail.Rows[RowD].Cells["g_iDetTRSUnitPrice"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTRSUnitPrice1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetRemark"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetRemark1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetProdLine"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetProdLine1"].Value.ToString();
                        if (isNefoKAM)
                        {
                            if (sPayType == "T")
                            {
                                dgvDetail.Rows[RowD].Cells["g_iDetTOPDiv"].Value = "P0";
                            }
                            else
                            {
                                try
                                {
                                    if (!string.IsNullOrEmpty(dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString()))
                                    {
                                        dgvDetail.Rows[RowD].Cells["g_iDetTOPDiv"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString();
                                    }
                                }
                                catch
                                {

                                }


                            }

                        }

                        RowD++;
                        __DetNo++;
                    }
                }
                else
                {
                    if (dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() == sPono)
                    {

                        dgvDetail.Rows.Add();
                        if (chkEcommerce.Checked)
                        {
                            dgvDetail.Rows[RowD].Cells["g_iDetEntity"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetEntity1"].Value.ToString().Replace("\"", "");
                            dgvDetail.Rows[RowD].Cells["g_iDetBranch"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetBranch1"].Value.ToString().Replace("\"", "");
                        }
                        dgvDetail.Rows[RowD].Cells["g_iDetPONO"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString().Replace("\"", "");
                        //di VB6 di tutup//dgvDetail.Rows[RowD].Cells["g_iDetCustNo"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value.ToString().Replace("\"", "");
                        dgvDetail.Rows[RowD].Cells["g_iDetPCode"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetPDesc"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetPDesc1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetGrade"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetSize"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetSled"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetConv1"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetConv2"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetValue"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetValue1"].Value.ToString()).ToString("#,##0.00");
                        dgvDetail.Rows[RowD].Cells["g_iDetValue"].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dgvDetail.Rows[RowD].Cells["g_iDetQty"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString()).ToString("#,##0");
                        dgvDetail.Rows[RowD].Cells["g_iDetQty"].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        dgvDetail.Rows[RowD].Cells["g_iDetQtyOrd"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetTaxCode"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetTaxPct"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString();
                        if (isNefoKAM)
                        {
                            dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()).ToString("#,##0.0000");
                        }
                        else
                        {
                            dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()).ToString("#,##0.00");
                        }
                        //dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Value = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()).ToString("#,##0.00");
                        dgvDetail.Rows[RowD].Cells["g_iDetPrice"].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                        //dgvDetail.Rows[RowD].Cells["g_iDetTRSPrice"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTRSPrice1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetStatus"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetStatus1"].Value.ToString();
                        //dgvDetail.Rows[RowD].Cells["g_iDetTRSUnitPrice"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTRSUnitPrice1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetRemark"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetRemark1"].Value.ToString();
                        dgvDetail.Rows[RowD].Cells["g_iDetProdLine"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetProdLine1"].Value.ToString();
                        if (isNefoKAM)
                        {
                            if (sPayType == "T")
                            {
                                dgvDetail.Rows[RowD].Cells["g_iDetTOPDiv"].Value = "P0";
                            }
                            else
                            {
                                try
                                {
                                    if (!string.IsNullOrEmpty(dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString()))
                                    {
                                        dgvDetail.Rows[RowD].Cells["g_iDetTOPDiv"].Value = dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString();
                                    }
                                }
                                catch
                                {

                                }


                            }

                        }

                        RowD++;
                    }
                }

            }

            keyFound = true;
            return keyFound;

        }

        private string fIsPOExists(string sPO)
        {
            string fIsPOExists = "";

            DataTable dtFill1 = new DataTable();
            strSQL = " select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK) WHERE wsh_po_no= '" + sPO + "' and wsh_entity_id = '" + cbEntityBranch.txtCM.Text + "' and wsh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                fIsPOExists = dtFill1.Rows[0]["wsh_seq_no"].ToString();
            }

            return fIsPOExists;
        }

        private bool fGetDetailAttr()
        {
            bool keyFound = false;
            decimal currPrice = 0;

            for (int r = 0; r < dgvDetail2.Rows.Count; r++)
            {

                if (chkEcommerce.Checked)
                {
                    if (fGetCustCodeByPO(
                        dgvDetail2.Rows[r].Cells["g_iDetPONO1"].Value.ToString()
                        , dgvDetail2.Rows[r].Cells["g_iDetEntity1"].Value.ToString()
                        , dgvDetail2.Rows[r].Cells["g_iDetBranch1"].Value.ToString()
                        , dgvDetail2.Rows[r].Cells["g_iDetCustNo1"].Value.ToString()
                        , dgvDetail2.Rows[r].Cells["g_iDetSalesmanID1"].Value.ToString()
                        ) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    if (fGetCustCodeByPO(dgvDetail2.Rows[r].Cells["g_iDetPONO1"].Value.ToString()) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }

                dTRSPrice = 0;
                sPcode = dgvDetail2.Rows[r].Cells["g_iDetPCode1"].Value.ToString();
                sGrade = dgvDetail2.Rows[r].Cells["g_iDetGrade1"].Value.ToString();
                sSize = dgvDetail2.Rows[r].Cells["g_iDetSize1"].Value.ToString();
                sSled = dgvDetail2.Rows[r].Cells["g_iDetSled1"].Value.ToString();
                string sEntity = "";
                string sBranch = "";
                if (chkEcommerce.Checked)
                {
                    sEntity = dgvDetail2.Rows[r].Cells["g_iDetEntity1"].Value.ToString();
                    sBranch = dgvDetail2.Rows[r].Cells["g_iDetBranch1"].Value.ToString();
                }

                if (chkEcommerce.Checked)
                {

                    string __strSQL = "";
                    __strSQL = "SELECT top 1 wh_loc_id1,wh_loc_id2,wh_loc_last_work_date FROM IM_WH_LOC WITH (NOLOCK) ";
                    __strSQL += "WHERE ";
                    __strSQL += "wh_loc_entity = '" + sEntity + "' ";
                    __strSQL += "AND wh_branch_id = '" + sBranch + "' ";
                    __strSQL += "AND wh_loc_id1 +'|'+wh_loc_id2 = (SELECT top 1 gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_function_name  LIKE 'MAINWHLOC' ) ";
                    DataTable dtTglGudang = _clsGlobal.ExecDT(__strSQL);
                    if (dtTglGudang.Rows.Count > 0)
                    {
                        DateTime tgl = Convert.ToDateTime(dtTglGudang.Rows[0]["wh_loc_last_work_date"].ToString());
                        if (fPricePCode(sEntity, sBranch, tgl) == false)
                        {
                            keyFound = false;
                            return keyFound;
                        }
                    }
                    else
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    if (fPricePCode() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }


                dgvDetail2.Rows[r].Cells["g_iDetPrice1"].Value = dTRSPrice;
                decimal vlu = dTRSPrice * Convert.ToInt16(dgvDetail2.Rows[r].Cells["g_iDetQty1"].Value);
                dgvDetail2.Rows[r].Cells["g_iDetValue1"].Value = vlu;
            }

            keyFound = true;
            return keyFound;
        }

        private bool fPricePCode()
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " IP_PRICEPCODE_MBRANCH '" + sPcode + "','";
            strSQL = strSQL + sGrade + "','";
            strSQL = strSQL + sSize + "','";
            strSQL = strSQL + sCust1 + "','";
            strSQL = strSQL + sCust2 + "','";
            strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "','";
            strSQL = strSQL + cbEntityBranch.txtCM.Text + "','";
            strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                if (string.IsNullOrEmpty(dtFill1.Rows[0]["opl_het"].ToString()))
                {
                    keyFound = false;
                }
                else
                {
                    dTRSPrice = Convert.ToDecimal(dtFill1.Rows[0]["opl_het"].ToString());
                }
            }

            keyFound = true;
            return keyFound;
        }

        private bool fPricePCode(string entity, string branch, DateTime tglgudang)
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " IP_PRICEPCODE_MBRANCH '" + sPcode + "','";
            strSQL = strSQL + sGrade + "','";
            strSQL = strSQL + sSize + "','";
            strSQL = strSQL + sCust1 + "','";
            strSQL = strSQL + sCust2 + "','";
            strSQL = strSQL + tglgudang.ToString("yyyy-MM-dd") + "','";
            strSQL = strSQL + entity + "','";
            strSQL = strSQL + branch + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                if (string.IsNullOrEmpty(dtFill1.Rows[0]["opl_het"].ToString()))
                {
                    keyFound = false;
                }
                else
                {
                    dTRSPrice = Convert.ToDecimal(dtFill1.Rows[0]["opl_het"].ToString());
                }
            }

            keyFound = true;
            return keyFound;
        }

        private bool fGetCustCodeByPO(string sPONO)
        {
            bool keyFound = false;

            for (int x = 0; x < dgvHeader.Rows.Count; x++)
            {
                if (sPONO == dgvHeader.Rows[x].Cells["g_iCustPONo"].Value.ToString())
                {
                    sCust1 = dgvHeader.Rows[x].Cells["g_iCustCode"].Value.ToString();
                    sCust2 = dgvHeader.Rows[x].Cells["g_iCustCode2"].Value.ToString();
                    break;
                }
            }

            keyFound = true;
            return keyFound;
        }

        private bool fGetCustCodeByPO(string sPONO, string sEntity, string sBranch, string sCust, string sSalesman)
        {
            bool keyFound = false;

            for (int x = 0; x < dgvHeader.Rows.Count; x++)
            {
                if (sPONO == dgvHeader.Rows[x].Cells["g_iCustPONo"].Value.ToString()
                    && sEntity == dgvHeader.Rows[x].Cells["g_iEntity"].Value.ToString()
                    && sBranch == dgvHeader.Rows[x].Cells["g_iBranch"].Value.ToString()
                    && sSalesman == dgvHeader.Rows[x].Cells["g_iSalesmanCode"].Value.ToString()
                    && sCust == dgvHeader.Rows[x].Cells["g_iCustCode"].Value.ToString()
                    )
                {
                    sCust1 = dgvHeader.Rows[x].Cells["g_iCustCode"].Value.ToString();
                    sCust2 = dgvHeader.Rows[x].Cells["g_iCustCode2"].Value.ToString();
                    break;
                }
            }

            keyFound = true;
            return keyFound;
        }

        //validasi
        private bool fProsesTextFile()
        {
            bool keyFound = false;

            try
            {
                string vPONo = "";
                sCustNo = "";
                ConvFail = "";
                string sCustPCode = "";
                //LineNum = 0
                string CCP = "";
                string CTEMP1 = ""; //customer code untuk division
                string CTEMP2 = ""; //customer code2 untuk filter division
                string PONO = "";
                string PCode1 = "";
                string PCode2 = "";
                string sGradeT = "";
                string sSizeT = "";
                string PayType = "";


                DataTable dHeader = new DataTable();
                dHeader.Columns.Add("PO");
                dHeader.Columns.Add("CustCode1");
                dHeader.Columns.Add("CustCode2");
                dHeader.Columns.Add("PayType");


                dgvHeader.Rows.Clear();
                dgvDetail.Rows.Clear();
                //dgvdata.Rows.Clear();

                int _i = 0;
                //dgvdata.DataSource = dtView;
                //dgvdata.Columns.Clear();
                //SetGridData(dgvdata);
                //DataTable dt1 = dgvdata.DataSource as DataTable;

                #region Validasi

                for (lCounter = 0; lCounter < dgvdata.Rows.Count; lCounter++)
                {
                    if (dgvdata.Rows.Count > 0)
                    {
                        if (dgvdata.Rows[lCounter].Cells[0].Value.ToString() == "SOHDR")
                        {
                            if (dgvdata.Rows[lCounter].Cells[1].Value.ToString() == "")
                            {
                                MessageBox.Show("Row Header [Line " + lCounter + "] : Customer No is empty, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            if (dgvdata.Rows[lCounter].Cells[3].Value.ToString() == "")
                            {
                                MessageBox.Show("Row Header [Line " + lCounter + "] : Po Number is empty, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            if (isNefoKAM)
                            {
                                if (dgvdata.Rows[lCounter].Cells[3].Value.ToString().Length > 35)
                                {
                                    MessageBox.Show("Row Header [Line " + lCounter + "] : Po Number - " + dgvdata.Rows[lCounter].Cells[3].Value.ToString() + " is can't be more than 35 characters, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    keyFound = false;
                                    return keyFound;
                                }
                            }

                            if (dgvdata.Rows[lCounter].Cells[4].Value.ToString() == "")
                            {
                                MessageBox.Show("Row Header [Line " + lCounter + "] : Failed Value Disc Header, character comma found please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            if (IsNumeric(dgvdata.Rows[lCounter].Cells[4].Value.ToString()) == false)
                            {
                                MessageBox.Show("Row Header [Line " + lCounter + "] : Failed Value Disc Header, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            if (sNoPO == dgvdata.Rows[lCounter].Cells[3].Value.ToString())
                            {
                                MessageBox.Show("Terdapat No Po yang Double : " + sNoPO, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            //sCustNo = dgvdata.Rows[lCounter].Cells[1].Value.ToString();
                            //sNoPO = dgvdata.Rows[lCounter].Cells[3].Value.ToString();

                            //if (fGetCustNO() == false)
                            //{
                            //    dgvdata.DataSource = null;
                            //    dgvdata.Rows.Clear();
                            //    keyFound = false;
                            //    return keyFound;
                            //}

                            //dgvHeader.Rows.Add();
                            //dgvHeader

                        }
                        else if (dgvdata.Rows[lCounter].Cells[0].Value.ToString() == "SODTL")
                        {
                            if (dgvdata.Rows[lCounter].Cells[3].Value.ToString() == "")
                            {
                                MessageBox.Show("Row Detail [Line" + lCounter + "] : Invalid PCode, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            if (IsNumeric(dgvdata.Rows[lCounter].Cells[4].Value.ToString()) == false)
                            {
                                MessageBox.Show("Row Detail [Line " + lCounter + "] : Invalid Conversion, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            if (IsNumeric(dgvdata.Rows[lCounter].Cells[5].Value.ToString()) == false && dgvdata.Rows[lCounter].Cells[5].Value.ToString() != "" || IsNumeric(dgvdata.Rows[lCounter].Cells[6].Value.ToString()) == false && dgvdata.Rows[lCounter].Cells[6].Value.ToString() != "")
                            {
                                MessageBox.Show("Row Detail [Line " + lCounter + "] : Invalid Qty Value, please check textfile !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                        }
                    }
                }

                #endregion

                #region InsertToGrid

                //sCustNo = dgvdata.Rows[lCounter].Cells[1].Value.ToString();
                //vPONo = dgvdata.Rows[lCounter].Cells[3].Value.ToString();

                //if (fGetCustNO() == false)
                //{
                //    dgvdata.DataSource = null;
                //    dgvdata.Rows.Clear();
                //    keyFound = false;
                //    return keyFound;
                //}

                int iRow = 0;
                for (lCounter = 0; lCounter < dgvdata.Rows.Count; lCounter++)
                {
                    sCustNo = dgvdata.Rows[lCounter].Cells[1].Value.ToString();
                    vPONo = dgvdata.Rows[lCounter].Cells[3].Value.ToString();

                    if (fGetCustNO() == false)
                    {
                        dgvdata.DataSource = null;
                        dgvdata.Rows.Clear();
                        keyFound = false;
                        return keyFound;
                    }

                    if (dgvdata.Rows.Count > 0)
                    {
                        if (dgvdata.Rows[lCounter].Cells[0].Value.ToString() == "SOHDR")
                        {

                            dgvHeader.Rows.Add();
                            dgvHeader.Rows[iRow].Cells["g_iProses"].Value = false;
                            if (sCust1 != "")
                            {
                                dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value = sCust1;
                            }
                            else
                            {
                                dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value = sCustNo;
                            }
                            if (sCust2 != "")
                            {
                                dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value = sCust2;
                            }
                            else
                            {
                                dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value = "000";
                            }
                            dgvHeader.Rows[iRow].Cells["g_iCustName"].Value = sCustName;
                            dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value = vPONo;

                            DataTable dtFill1 = new DataTable();
                            strSQL = " select isnull(soh_so_no,'') so_no From TBL_SD_UPLOADSO_HDR WITH (NOLOCK) where soh_po_no='" + vPONo + "' and soh_cust_code1='" + sCust1 + "' and soh_cust_code2='" + sCust2 + "'";
                            dtFill1 = _clsGlobal.ExecDT(strSQL);
                            if (dtFill1.Rows.Count > 0)
                            {
                                dgvHeader.Rows[iRow].Cells["g_iSONo"].Value = dtFill1.Rows[0]["so_no"].ToString();
                                //dgvHeader.Rows[iRow].Cells["g_iRemark"].Value = "PO Already Uploaded";
                            }
                            else
                            {
                                dgvHeader.Rows[iRow].Cells["g_iSONo"].Value = "";
                                //dgvHeader.Rows[iRow].Cells["g_iRemark"].Value = "";
                            }
                            string sPODate = dgvdata.Rows[lCounter].Cells[3].Value.ToString();

                            if (isUploadMigrasiToNEFOKAM)
                            {
                                sPODate = dgvdata.Rows[lCounter].Cells[5].Value.ToString();
                                sPODate = sPODate.Substring(sPODate.Length - 4) + "-" + sPODate.Substring(2, 2) + "-" + sPODate.Substring(0, 2);
                            }
                            else
                            {
                                sPODate = sPODate.Substring(0, 10);
                                sPODate = sPODate.Substring(2, 8);
                                sPODate = "20" + sPODate.Substring(sPODate.Length - 2) + "-" + sPODate.Substring(2, 2) + "-" + sPODate.Substring(0, 2);
                            }
                            //DateTime dtime = Convert.ToDateTime(sPODate);
                            //sPODate = dtSODate.Value.ToString("yyyy-MM-dd");
                            if (dgvdata.Rows[lCounter].Cells[2].Value.ToString() == "T" || dgvdata.Rows[lCounter].Cells[2].Value.ToString() == "D" || dgvdata.Rows[lCounter].Cells[2].Value.ToString() == "S")
                            {
                                sTopId = "P0";
                                sTOPVal = "0";
                            }
                            dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value = sPODate;
                            dgvHeader.Rows[iRow].Cells["g_iTOPId"].Value = sTopId;
                            dgvHeader.Rows[iRow].Cells["g_iTOP"].Value = sTOPVal;
                            dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value = dgvdata.Rows[lCounter].Cells[2].Value.ToString();
                            dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value = dgvdata.Rows[lCounter].Cells[4].Value.ToString().Replace(",", ".");
                            dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value = sLeadtime;
                            if (isUploadMigrasiToNEFOKAM)
                            {
                                string sDeliveryDate = dgvdata.Rows[lCounter].Cells[6].Value.ToString();
                                string sExpiredDate = dgvdata.Rows[lCounter].Cells[7].Value.ToString();

                                sDeliveryDate = sDeliveryDate.Substring(sDeliveryDate.Length - 4) + "-" + sDeliveryDate.Substring(2, 2) + "-" + sDeliveryDate.Substring(0, 2);
                                sExpiredDate = sExpiredDate.Substring(sExpiredDate.Length - 4) + "-" + sExpiredDate.Substring(2, 2) + "-" + sExpiredDate.Substring(0, 2);

                                dgvHeader.Rows[iRow].Cells["g_iDeliveryDate"].Value = sDeliveryDate;
                                dgvHeader.Rows[iRow].Cells["g_iExpiredDate"].Value = sExpiredDate;

                            }

                            if (cbSource.SelectedValue.ToString().Trim() == "L" || cbSource.SelectedValue.ToString().Trim() == "S" || cbSource.SelectedValue.ToString().Trim() == "W")
                            {
                                dgvHeader.Rows[iRow].Cells["g_iTXTValue"].Value = dgvdata.Rows[lCounter].Cells[7].Value.ToString();
                                dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value = dgvdata.Rows[lCounter].Cells[6].Value.ToString();
                                dgvHeader.Rows[iRow].Cells["g_iCashDisc"].Value = dgvdata.Rows[lCounter].Cells[5].Value.ToString();
                            }
                            //else
                            //{
                            //    dgvHeader.Rows[iRow].Cells["g_iTXTValue"].Value = "0";
                            //    dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value = "0";
                            //    dgvHeader.Rows[iRow].Cells["g_iCashDisc"].Value = "0";
                            //}

                            dHeader.Rows.Add(vPONo, dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value, dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value, dgvdata.Rows[lCounter].Cells[2].Value.ToString());

                            iRow++;
                        }
                    }
                }

                iRow = 0;
                if (dgvDetail2.Rows.Count > 0)
                {
                    dgvDetail2.Rows.Clear();
                }
                for (lCounter = 0; lCounter < dgvdata.Rows.Count; lCounter++)
                {
                    //sCustNo = dgvdata.Rows[lCounter].Cells[1].Value.ToString();
                    vPONo = dgvdata.Rows[lCounter].Cells[3].Value.ToString();

                    if (dgvdata.Rows.Count > 0)
                    {
                        if (dgvdata.Rows[lCounter].Cells[0].Value.ToString() == "SODTL")
                        {
                            string sPOCodex = dgvdata.Rows[lCounter].Cells[3].Value.ToString();
                            dCustPcode = sPOCodex.Substring(0, 6);

                            iConv1 = "0";
                            iConv2 = "0";

                            if (fGetProdAttr() == false)
                            {
                                dtView.Rows.Clear();
                                keyFound = false;
                                return keyFound;
                            }

                            if (iConv1 == "")
                            {
                                iConv1 = "0";
                            }

                            if (iConv2 == "")
                            {
                                iConv2 = "0";
                            }

                            bExists = false;
                            if (fPcodeExists(vPONo) == false)
                            {
                                dtView.Rows.Clear();
                                keyFound = false;
                                return keyFound;
                            }

                            if (dgvdata.Rows[lCounter].Cells[5].Value.ToString() == "")
                            {
                                dgvdata.Rows[lCounter].Cells[5].Value = "0";
                            }

                            if (dgvdata.Rows[lCounter].Cells[6].Value.ToString() == "")
                            {
                                dgvdata.Rows[lCounter].Cells[6].Value = "0";
                            }

                            if (dgvdata.Rows[lCounter].Cells[7].Value.ToString() == "")
                            {
                                dgvdata.Rows[lCounter].Cells[7].Value = "0";
                            }

                            int sData5 = Convert.ToInt16(dgvdata.Rows[lCounter].Cells[5].Value);
                            int sData6 = Convert.ToInt16(dgvdata.Rows[lCounter].Cells[6].Value);
                            int sData7 = Convert.ToInt16(dgvdata.Rows[lCounter].Cells[7].Value);

                            if (sData5 > 0 || sData6 > 0 || sData7 > 0)
                            {
                                CCP = sData5.ToString() + "." + sData6.ToString() + "." + sData7.ToString();
                                fQtyFormat(CCP, Convert.ToInt32(iConv1), Convert.ToInt32(iConv2));
                                CCP = _fQtyFormat;
                                dgvDetail2.Rows.Add();
                                dgvDetail2.Rows[iRow].Cells["g_iDetPONO1"].Value = dgvdata.Rows[lCounter].Cells[1].Value.ToString().Replace("\"", "");
                                //dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value = dgvdata.Rows[lCounter].Cells[0].Value.ToString().Replace("\"", "");
                                if (sPcode != "")
                                {
                                    dgvDetail2.Rows[iRow].Cells["g_iDetPCode1"].Value = sPcode;
                                }
                                else
                                {
                                    dgvDetail2.Rows[iRow].Cells["g_iDetPCode1"].Value = dCustPcode;
                                }
                                dgvDetail2.Rows[iRow].Cells["g_iDetPDesc1"].Value = sDesc;
                                dgvDetail2.Rows[iRow].Cells["g_iDetGrade1"].Value = sGrade;
                                dgvDetail2.Rows[iRow].Cells["g_iDetSize1"].Value = sSize;
                                dgvDetail2.Rows[iRow].Cells["g_iDetSled1"].Value = sSled;
                                dgvDetail2.Rows[iRow].Cells["g_iDetConv11"].Value = iConv1;
                                dgvDetail2.Rows[iRow].Cells["g_iDetConv21"].Value = iConv2;
                                dgvDetail2.Rows[iRow].Cells["g_iDetValue1"].Value = "0";
                                fConvertQty(CCP, Convert.ToInt16(dgvDetail2.Rows[iRow].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[iRow].Cells["g_iDetConv21"].Value));
                                dgvDetail2.Rows[iRow].Cells["g_iDetQty1"].Value = _fConvertQty;
                                dgvDetail2.Rows[iRow].Cells["g_iDetQtyOrd1"].Value = CCP;
                                dgvDetail2.Rows[iRow].Cells["g_iDetTaxCode1"].Value = sTaxCode;
                                dgvDetail2.Rows[iRow].Cells["g_iDetProdLine1"].Value = sProdLine;
                                dgvDetail2.Rows[iRow].Cells["g_iDetTaxPct1"].Value = siTaxPct;
                                dgvDetail2.Rows[iRow].Cells["g_iDetPrice1"].Value = "0";
                                // dgvDetail2.Rows[lCounter].Cells["g_iDetTRSPrice1"].Value = dTotTRS;
                                dgvDetail2.Rows[iRow].Cells["g_iDetStatus1"].Value = sStatus;

                                if (sPcode == "")
                                {
                                    dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "Master Product not found";
                                }
                                else if (sStatus != "A")
                                {
                                    if (sStatus != "D")
                                    {
                                        dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "Product is not Active";
                                    }
                                    else
                                    {
                                        dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "OK";
                                    }
                                }
                                else
                                {

                                    dgvDetail2.Rows[iRow].Cells["g_iDetRemark1"].Value = "OK";
                                }
                                iRow++;
                                //dgvDetail2.Rows[lCounter].Cells["g_iDetTRSUnitPrice1"].Value = dTRSPrice;
                            }
                        }

                    }
                }

                // jika di kam dan divisionnya di checklist cari top nya apabila Division Y dan customer TOP N //
                if (isNefoKAM)
                {
                    if (isDivisionTOP)
                    {
                        for (int i = 0; i < dgvDetail2.Rows.Count; i++)
                        {
                            PONO = dgvDetail2.Rows[i].Cells["g_iDetPONO1"].Value.ToString();
                            PCode1 = dgvDetail2.Rows[i].Cells["g_iDetPCode1"].Value.ToString();
                            PCode2 = "000";
                            sGradeT = dgvDetail2.Rows[i].Cells["g_iDetGrade1"].Value.ToString();
                            sSizeT = dgvDetail2.Rows[i].Cells["g_iDetSize1"].Value.ToString();
                            sSled = dgvDetail2.Rows[i].Cells["g_iDetSled1"].Value.ToString();

                            DataView dvTemp = dHeader.DefaultView;
                            dvTemp.RowFilter = "PO = '" + PONO + "'";

                            CTEMP1 = "";
                            CTEMP2 = "";
                            DataTable dtTempHeader = dvTemp.ToTable();
                            if (dtTempHeader.Rows.Count > 0)
                            {
                                CTEMP1 = dtTempHeader.Rows[0]["CustCode1"].ToString().Trim();
                                CTEMP2 = dtTempHeader.Rows[0]["CustCode2"].ToString().Trim();
                                PayType = dtTempHeader.Rows[0]["PayType"].ToString().Trim();

                                strSQL = "select cm_top_by_cust from SO_CUST_MASTER WITH (NOLOCK) where cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' and cm_cust_code1 = '" + CTEMP1 + "' and cm_cust_code2 = '" + CTEMP2 + "'";
                                DataTable dtCust = _clsGlobal.ExecDT(strSQL);
                                if (dtCust.Rows.Count > 0 && PayType != "T")
                                {
                                    if (dtCust.Rows[0]["cm_top_by_cust"].ToString() != "Y")
                                    {
                                        strSQL = "select SMT.smt_top_id ";
                                        strSQL += "   from SO_CUST_MASTER CM WITH (NOLOCK) ";
                                        strSQL += "   LEFT JOIN TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                                        strSQL += "       ON  ";
                                        strSQL += "           TSC.csc_entity          = cm_entity ";
                                        strSQL += "           AND TSC.csc_branch      = CM.cm_branch ";
                                        strSQL += "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 ";
                                        strSQL += "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 ";
                                        strSQL += "   left Join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK) ";
                                        strSQL += "       ON ";
                                        strSQL += "           SMT.smt_entity_id      = CM.cm_entity ";
                                        if (!isMultibranch)
                                        {
                                            strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                        }
                                        else
                                        {
                                            strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                        }

                                        strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                        strSQL += "       ON ";
                                        strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                        strSQL += "   where ";
                                        strSQL += "   prm_prd_master_code = '" + PCode1 + "' ";
                                        strSQL += "   and prm_grade = '" + sGradeT + "' ";
                                        strSQL += "   and prm_prd_size = '" + sSizeT + "' ";
                                        strSQL += "   and TSC.csc_salesman_id = '" + txtSalesman.Text.Trim() + "' ";
                                        strSQL += "   and cm_entity = '" + cbEntityBranch.txtCM.Text + "' ";
                                        strSQL += "   and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                                        strSQL += "   and CM.cm_cust_code1 = '" + CTEMP1 + "' ";
                                        strSQL += "   and CM.cm_cust_code2 = '" + CTEMP2 + "'";
                                        DataTable dtNewProduct = _clsGlobal.ExecDT(strSQL);
                                        if (dtNewProduct.Rows.Count > 0)
                                        {
                                            if (string.IsNullOrEmpty(dtNewProduct.Rows[0]["smt_top_id"].ToString().Trim()))
                                            {
                                                if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() == "OK")
                                                {
                                                    dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value = "Mapping TOP Belum Disetting.";
                                                }
                                            }
                                            else
                                            {
                                                dgvDetail2.Rows[i].Cells["g_iDetTOPDiv1"].Value = dtNewProduct.Rows[0]["smt_top_id"].ToString().Trim();
                                            }
                                        }
                                        else
                                        {
                                            if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() == "OK")
                                            {
                                                dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value = "Mapping TOP Belum Disetting.";
                                            }
                                        }
                                    }
                                }
                                else if (PayType == "T")
                                {
                                    dgvDetail2.Rows[i].Cells["g_iDetTOPDiv1"].Value = "P0";
                                }

                                /** Validasi Tambahan untuk TOP PER PRODUCT LINE **/
                                if (dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value.ToString() == "OK")
                                {
                                    strSQL = " Select smt_prdline_id ";
                                    strSQL += " FROM IM_PRD_MASTER WITH (NOLOCK) ";
                                    strSQL += " INNER JOIN SO_MAPPING_TOPBYPRDLINE WITH (NOLOCK) ";
                                    strSQL += " ON smt_prdline_id = prm_prd_line_code ";
                                    strSQL += " WHERE  ";
                                    strSQL += " smt_entity_id = '" + cbEntityBranch.txtCM.Text + "' ";
                                    strSQL += " and smt_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                                    strSQL += " and smt_cust_code1 = '" + CTEMP1 + "' ";
                                    strSQL += " and smt_cust_code2 = '" + CTEMP2 + "' ";
                                    strSQL += " and prm_prd_master_code = '" + PCode1 + "' ";

                                    DataTable dtPrdlineTOP = _clsGlobal.ExecDT(strSQL);
                                    if (dtPrdlineTOP.Rows.Count == 0)
                                    {
                                        dgvDetail2.Rows[i].Cells["g_iDetRemark1"].Value = "Product Line Belum di Mapping ke Customer";
                                    }

                                }
                                /** end validasi Tambahan untuk TOP Per Product LINE **/

                            }



                        }
                    }
                }


                #endregion
                DataTable dt = dgvDetail2.DataSource as DataTable;

                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fPcodeExists(string PONOx)
        {
            bool keyFound = false;

            //for (int x = 0; x < dgvdata.Rows.Count; x++)
            //{
            //    if (dgvdata.Rows.Count > 0)
            //    {
            //        if (dgvdata.Rows[lCounter].Cells[0].Value.ToString() == "SODTL")
            //        {
            if (PONOx == dgvdata.Rows[lCounter].Cells[1].Value.ToString())
            {
                if (sPcode == dgvdata.Rows[lCounter].Cells[3].Value.ToString())
                {
                    bExists = true;
                    keyFound = true;
                    return keyFound;
                }
            }
            //        }

            //    }
            //}

            keyFound = true;
            return keyFound;
        }

        private bool IsNumeric(string s)
        {
            float output;
            return float.TryParse(s, out output);
        }

        private void pEnabled(int iType)
        {
            if (iType == 0)
            {
                txtSalesman.Enabled = false;
                btnPopUpSalesman.Enabled = false;
                btnOk.Enabled = false;
                btnBrowse.Enabled = false;
                cbTipeOrder.Enabled = false;
                cbWH.Enabled = false;
                //btnPopUpAccount.Enabled = false;
                btnCancel.Enabled = true;
                btnProses.Enabled = true;
                chkAll.Enabled = true;
                cbSource.Enabled = false;
                cbFlagSloc.Enabled = false;
                chkEcommerce.Enabled = false;
                chkmappingproduct.Enabled = false;
            }
            else
            {
                txtSalesman.Enabled = true;
                btnPopUpSalesman.Enabled = true;
                btnOk.Enabled = true;
                btnBrowse.Enabled = true;
                cbTipeOrder.Enabled = true;
                cbWH.Enabled = true;
                //btnPopUpAccount.Enabled = true;
                btnCancel.Enabled = false;
                btnProses.Enabled = false;
                chkAll.Enabled = false;
                chkAll.Checked = false;
                cbSource.Enabled = true;
                cbFlagSloc.Enabled = true;
                chkEcommerce.Enabled = true;
                chkmappingproduct.Enabled = true;
            }

        }

        private void BindDetailEntity()
        {
            strSQL = "";
            strSQL = "SELECT '0' as id, '' as descc ";
            strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "select distinct e.ge_entity_id  as id,  e.ge_entity_id + ' - ' + e.ge_entity as descc from GS_ENTITY e WITH (NOLOCK) INNER JOIN GS_USERS_SECURITY u WITH (NOLOCK) on u.gu_entity = e.ge_entity_id WHERE gu_user_id = '" + clsLogin.USERID + "' order by id ";


            //cbEntityID.DataSource = _clsGlobal.ExecDT(strSQL);
            //cbEntityID.ValueMember = "id";
            //cbEntityID.DisplayMember = "descc";
            //cbSOType.SelectedIndex = 0;

        }

        private void BindDetailBranch()
        {
            strSQL = "";
            //strSQL = "SELECT '0' as id, 'ALL' as descc ";
            //strSQL = strSQL + " union ALL ";
            strSQL = strSQL + "Select distinct br_branch_id as id, br_branch_id + ' - ' + br_branch_desc as descc from GS_BRANCH WITH (NOLOCK) inner join GS_ENTITY WITH (NOLOCK) on br_gl_entity_initial = ge_entity_id INNER JOIN GS_USERS_SECURITY u WITH (NOLOCK) on u.gu_entity = ge_entity_id where gu_user_id = '" + clsLogin.USERID + "' and ge_entity_id = '" + cbEntityBranch.txtCM.Text.Trim() + "' ";
            //strSQL = "distinct br_branch_id as id, br_branch_id + ' - ' + br_branch_desc as descc  frpw"
            //cbBranchID.DataSource = _clsGlobal.ExecDT(strSQL);
            //cbBranchID.ValueMember = "id";
            //cbBranchID.DisplayMember = "descc";
            //cbSOType.SelectedIndex = 0;
        }

        private bool fGetProdAttr()
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " SELECT  prm_prd_line_code,ISNULL(pl_prd_line_desc, '') AS pl_prd_line_desc, prm_prd_master_code,prm_grade,prm_prd_size,ISNULL(prm_prd_desc, '') AS prm_prd_desc,prm_conversion_purc,";
            strSQL = strSQL + " prm_conversion_sales, ";
            if (isNefoKAM)
            {
                strSQL = strSQL + " case when cust_tax_code='PPN0' THEN cust_tax_code ELSE  stc_tax_code END prm_tax_code, CASE WHEN cust_tax_code='PPN0' THEN isnull(cust_tax_amount,0) ELSE  isnull(stc_tax_amount, 0) END stc_tax_amount";
            }
            else
            {
                strSQL = strSQL + " prm_tax_code,case when isnull(excludeTax,'N') = 'Y' then 0 else isnull(stc_tax_amount, 0) end stc_tax_amount, ";

            }

            strSQL = strSQL + " isnull(prm_status, '') prm_status,isnull(wlb_loc_Id1,'X') wlb_loc_id, prm_prd_line_code,mcps_prd_sled AS [prm_prd_sled] From IM_PRD_MASTER WITH (NOLOCK) ";
            strSQL = strSQL + " LEFT OUTER JOIN IM_STOCK_BALANCE WITH (NOLOCK) ON ";
            strSQL = strSQL + " wlb_loc_Id1 = '" + lblWHLoc1.Text + "' AND wlb_loc_Id2 = '" + lblWHLoc2.Text + "' AND wlb_entity_id='" + cbEntityBranch.txtCM.Text + "' AND wlb_branch_id ='" + cbEntityBranch.txtBranchIdCM.Text + "' AND wlb_prd_master_code = prm_prd_master_code AND wlb_grade = prm_grade ";
            strSQL = strSQL + " AND wlb_prd_size = prm_prd_size AND isnull(wlb_stock_sts, 'N') = 'N' LEFT JOIN IM_PRD_LINE WITH (NOLOCK) ON prm_prd_line_code = pl_prd_line_code ";
            strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
            strSQL = strSQL + " LEFT JOIN TBL_MAP_CUST_PRD_SLED WITH (NOLOCK) ON mcps_entity_id='" + cbEntityBranch.txtCM.Text + "' AND mcps_branch_id='" + cbEntityBranch.txtBranchIdCM.Text + "' AND mcps_cust_code1='" + sCust1 + "' AND mcps_cust_code1='" + sCust2 + "' AND mcps_prd_code=prm_prd_master_code AND mcps_prd_grade=prm_grade";
            strSQL = strSQL + " AND mcps_prd_size = prm_prd_size ";
            //if (isNefoKAM)
            //{
            //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON  CAST(CONVERT(DATE, '" + dtSODate.Value.ToString("yyyyMMdd") + "', 112) AS DATETIME) >=cast(stc_tax_valid_from as datetime) ";
            //    strSQL = strSQL + "  AND CAST(CONVERT(DATE,'" + dtSODate.Value.ToString("yyyyMMdd") + "', 112) AS DATETIME) <=cast(stc_tax_valid_to as datetime)  ";
            //}
            //else
            //{
            //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
            //}

            if (isNefoKAM)
            {
                /** customer PPN **/
                strSQL = strSQL + " LEFT JOIN  (SELECT cm_entity,cm_branch, cm_ship_tax_id AS cust_tax_code , stc_tax_amount as cust_tax_amount FROM SO_CUST_MASTER WITH (NOLOCK)  ";
                strSQL = strSQL + "  INNER JOIN dbo.SO_TAX_CODE  WITH (NOLOCK) ON cm_ship_tax_id = stc_tax_code  ";
                strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + sCust1 + "'";
                strSQL = strSQL + "  AND cm_cust_code2 = '" + sCust2 + "'";
                strSQL = strSQL + "  AND cm_entity = '" + cbEntityBranch.txtCM.Text + "'";
                strSQL = strSQL + "  AND cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                strSQL = strSQL + " )CUST ON cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                /** End customer PPN **/
            }
            else
            {
                strSQL = strSQL + " LEFT JOIN ( SELECT gh_group_menu_id as __entity,gh_parent_menu __branch,gh_function_code excludeTax FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE GH_FUNCTION_NAME = 'Exclude-TAX-KAM' AND GH_SYS = 'S' ) FLAGTAX ON __entity = '" + cbEntityBranch.txtCM.Text + "' and __branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
            }


            strSQL = strSQL + " WHERE  prm_prd_master_code ='" + dCustPcode + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                if (isNefoKAM)
                {
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["prm_tax_code"].ToString()))
                    {
                        MessageBox.Show("Parameter untuk so_tax_code valid date '" + dtSODate.Value.ToString("yyyyMMdd") + "' tidak ditemukan!.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }
                    else
                    {
                        keyFound = true;
                    }
                }
                sLineCode = dtFill1.Rows[0]["prm_prd_line_code"].ToString().Trim();
                sLineDesc = dtFill1.Rows[0]["pl_prd_line_desc"].ToString().Trim();
                sPcode = dtFill1.Rows[0]["prm_prd_master_code"].ToString().Trim();
                sGrade = dtFill1.Rows[0]["prm_grade"].ToString().Trim();
                sSize = dtFill1.Rows[0]["prm_prd_size"].ToString().Trim();
                sSled = dtFill1.Rows[0]["prm_prd_sled"].ToString().Trim();
                sDesc = dtFill1.Rows[0]["prm_prd_desc"].ToString().Trim();
                iConv1 = dtFill1.Rows[0]["prm_conversion_purc"].ToString().Trim();
                iConv2 = dtFill1.Rows[0]["prm_conversion_sales"].ToString().Trim();
                sTaxCode = dtFill1.Rows[0]["prm_tax_code"].ToString().Trim();
                siTaxPct = dtFill1.Rows[0]["stc_tax_amount"].ToString().Trim();
                sStatus = dtFill1.Rows[0]["prm_status"].ToString().Trim();
                sWHLoc = dtFill1.Rows[0]["wlb_loc_id"].ToString().Trim();
                sProdLine = dtFill1.Rows[0]["prm_prd_line_code"].ToString().Trim();
            }
            else
            {
                sLineCode = "";
                sLineDesc = "";
                sPcode = "";
                sGrade = "";
                sSize = "";
                sSled = "";
                sDesc = "";
                iConv1 = "";
                iConv2 = "";
                sTaxCode = "";
                siTaxPct = "";
                sStatus = "";
                sWHLoc = "";
                sProdLine = "";
                keyFound = true;
            }


            return keyFound;
        }

        private bool fGetCustNO()
        {
            bool keyFound = false;
            DataTable dtFill1 = new DataTable();

            strSQL = "SELECT cm_cust_code1, cm_cust_code2, cm_cust_name, isnull(csc_salesman_id,'X') csc_salesman_id,isnull(pptc_term_code,'') top_id,isnull(cm_top,'') cm_top,isnull(cm_lead_time,0) cm_lead_time,";
            strSQL = strSQL + " isnull(cm_payment_type, '') as cm_payment_type From SO_CUST_MASTER WITH (NOLOCK) LEFT OUTER JOIN TBL_SD_CUSTCOVER WITH (NOLOCK) ON  cm_cust_code1 = csc_cust_code1 AND cm_cust_code2 = csc_cust_code2 and cm_entity = csc_entity and cm_branch = csc_branch INNER JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON  pptc_no_of_days = cm_top ";
            strSQL = strSQL + " WHERE csc_salesman_id = '" + txtSalesman.Text + "' AND cm_active_flag <> 'D' AND cm_cust_code1 = '" + sCustNo + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                sCust1 = dtFill1.Rows[0]["cm_cust_code1"].ToString();
                sCust2 = dtFill1.Rows[0]["cm_cust_code2"].ToString();
                sCustName = dtFill1.Rows[0]["cm_cust_name"].ToString();
                sPayType = dtFill1.Rows[0]["cm_payment_type"].ToString();
                sTopId = dtFill1.Rows[0]["top_id"].ToString();
                sTOPVal = dtFill1.Rows[0]["cm_top"].ToString();
                sLeadtime = dtFill1.Rows[0]["cm_lead_time"].ToString();
            }
            else
            {
                sCust1 = "";
                sCust2 = "";
                sCustName = "";
                sPayType = "";
                sTopId = "";
                sTOPVal = "";
                sLeadtime = "";
            }

            DataTable dtFill2 = new DataTable();
            strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            DataTable dtFillDivision = new DataTable();
            strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name='TOPBYDIVISION' and gh_sys='H' and gh_function_code='Y' ";
            dtFillDivision = _clsGlobal.ExecDT(strSQL);
            #region TOP TEAM SALESMAN
            if (dtFill2.Rows.Count > 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = "select coalesce(cm_top_by_cust,'N') as cm_top_by_cust,cm_top_id,cm_payment_type, coalesce(cg_cust_top,'') as topnya,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment ";
                strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) inner join SO_CUST_GROUP WITH (NOLOCK)  on cm_cust_group = cg_cust_group left join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on cg_cust_top=pptc_term_code ";
                //strSQL = strSQL + " where cg_cust_top IS NOT NULL AND cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "'";
                strSQL = strSQL + " where cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text.ToString() + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                dtFill3 = _clsGlobal.ExecDT(strSQL);
                if (dtFill3.Rows.Count > 0)
                {

                    if (dtFill3.Rows[0]["cm_payment_type"].ToString() == "T") //'langsung TOP 0
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select  pptc_term_code,pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES WITH (NOLOCK) where pptc_term_code='P0'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else if (dtFill3.Rows[0]["cm_top_by_cust"].ToString() == "Y")
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select pptc_term_code,pptc_no_of_days ";
                        strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON cm_top_id=pptc_term_code";
                        strSQL = strSQL + " where cm_cust_code1='" + sCust1 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                        //sTopId = dtFill3.Rows[0]["cm_top_id"].ToString();
                        //DataTable dtFill5 = new DataTable();
                        //strSQL = "select  pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES where pptc_term_code='" + sTopId + "'";
                        //dtFill5 = _clsGlobal.ExecDT(strSQL);
                        //if (dtFill5.Rows.Count > 0)
                        //{
                        //    sTOPVal = dtFill5.Rows[0]["pptc_no_of_days"].ToString();
                        //}
                    }
                    else
                    {
                        if (dtFill3.Rows[0]["topnya"].ToString() != "")
                        {
                            sPayType = dtFill3.Rows[0]["top_payment"].ToString();
                        }
                        else
                        {
                            DataTable dtFill6 = new DataTable();
                            strSQL = "select pptc_no_of_days,sgl_group_top,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from SO_SPG_GIRL_MAN WITH (NOLOCK) ";
                            strSQL = strSQL + " left join SO_SPG_GROUP WITH (NOLOCK) on sgm_spgm_group = sgl_group_code inner join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on sgl_group_top=pptc_term_code ";
                            strSQL = strSQL + " where sgm_spgm_id='" + txtSalesman.Text + "' and sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' and sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                            dtFill6 = _clsGlobal.ExecDT(strSQL);
                            if (dtFill6.Rows.Count > 0) //ikut team salesman
                            {
                                sTopId = dtFill6.Rows[0]["sgl_group_top"].ToString();
                                sTOPVal = dtFill6.Rows[0]["pptc_no_of_days"].ToString();
                            }
                        }
                    }
                }
            }
            #endregion

            #region TOP BY DIVISION
            else if (dtFillDivision.Rows.Count > 0)
            {
                DataTable dtFill3 = new DataTable();
                strSQL = "select coalesce(cm_top_by_cust,'N') as cm_top_by_cust,cm_top_id,cm_payment_type, coalesce(cg_cust_top,'') as topnya,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment ";
                strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) inner join SO_CUST_GROUP WITH (NOLOCK)  on cm_cust_group = cg_cust_group left join PO_PAYMENT_TERM_CODES WITH (NOLOCK) on cg_cust_top=pptc_term_code ";
                //strSQL = strSQL + " where cg_cust_top IS NOT NULL AND cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "'";
                strSQL = strSQL + " where cm_cust_code1 ='" + sCust1 + "' and cm_cust_code2='" + sCust2 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                dtFill3 = _clsGlobal.ExecDT(strSQL);
                if (dtFill3.Rows.Count > 0)
                {

                    if (dtFill3.Rows[0]["cm_payment_type"].ToString() == "T") //'langsung TOP 0
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select  pptc_term_code,pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES WITH (NOLOCK) where pptc_term_code='P0'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                    }
                    else if (dtFill3.Rows[0]["cm_top_by_cust"].ToString() == "Y")
                    {
                        DataTable dtFill4 = new DataTable();
                        strSQL = "select pptc_term_code,pptc_no_of_days ";
                        strSQL = strSQL + " from SO_CUST_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON cm_top_id=pptc_term_code";
                        strSQL = strSQL + " where cm_cust_code1='" + sCust1 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill4.Rows.Count > 0)
                        {
                            sTopId = dtFill4.Rows[0]["pptc_term_code"].ToString();
                            sTOPVal = dtFill4.Rows[0]["pptc_no_of_days"].ToString();
                        }
                        //sTopId = dtFill3.Rows[0]["cm_top_id"].ToString();
                        //DataTable dtFill5 = new DataTable();
                        //strSQL = "select  pptc_no_of_days,convert(varchar,pptc_no_of_days) + ' ~ ' + pptc_term_desc top_payment from PO_PAYMENT_TERM_CODES where pptc_term_code='" + sTopId + "'";
                        //dtFill5 = _clsGlobal.ExecDT(strSQL);
                        //if (dtFill5.Rows.Count > 0)
                        //{
                        //    sTOPVal = dtFill5.Rows[0]["pptc_no_of_days"].ToString();
                        //}
                    }
                    else
                    {
                        strSQL = "select distinct top 1 PM.prm_prd_line_code pcodelinenew,SP.ssp_spggroup_code,SP.ssp_prdline_code plinegrup,SMT.smt_prdline_id plinemap,SMT.smt_top_id, PTC.pptc_no_of_days ";
                        strSQL += "   from IM_PRD_MASTER PM WITH (NOLOCK) ";
                        strSQL += "   left join TBL_SD_SPGGROUP_PRDLINE SP WITH (NOLOCK) ";
                        strSQL += "       ON  ";
                        strSQL += "           SP.ssp_prdline_code = PM.prm_prd_line_code ";
                        strSQL += "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK) ";
                        strSQL += "       ON	";
                        strSQL += "           SPG.sgm_spgm_group = SP.ssp_spggroup_code ";
                        strSQL += "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           TSC.csc_entity = SPG.sgm_entity_id ";
                        strSQL += "           AND TSC.csc_branch = SPG.sgm_branch_id ";
                        strSQL += "           AND TSC.csc_salesman_id = SPG.sgm_spgm_id ";
                        strSQL += "   left join SO_CUST_MASTER CM WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           CM.cm_cust_code1 = TSC.csc_cust_code1 ";
                        strSQL += "           and cm_cust_code2 = TSC.csc_cust_code2 ";
                        strSQL += "           and cm_entity = TSC.csc_entity ";
                        strSQL += "           and cm_branch = TSC.csc_branch ";
                        strSQL += "   left Join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           SMT.smt_entity_id = CM.cm_entity ";
                        //"           and SMT.smt_branch_id = CM.cm_branch " +
                        //"           and SMT.smt_cust_code1 = CM.cm_cust_code1 " +
                        //"           and SMT.smt_cust_code2 = CM.cm_cust_code2 " +
                        //"           and SMT.smt_prdline_id = PM.prm_prd_line_code " +
                        if (!isMultibranch)
                        {
                            strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                        }
                        else
                        {
                            strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                            strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                            strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                        }
                        strSQL += "   Left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                        strSQL += "       ON ";
                        strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                        strSQL += "   where ";
                        strSQL += "   SPG.sgm_spgm_id = '" + txtSalesman.Text.Trim() + "' ";
                        strSQL += "   and SP.ssp_del_flag = 'N' ";
                        strSQL += "   and SPG.sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' ";
                        strSQL += "   and SPG.sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                        strSQL += "   and CM.cm_cust_code1 = '" + sCust1 + "' ";
                        strSQL += "   and CM.cm_top_by_cust = 'N' ";
                        strSQL += "   and CM.cm_cust_code2 = '" + sCust2 + "' and smt_top_id is not null ";
                        DataTable dtNewProduct = _clsGlobal.ExecDT(strSQL);
                        if (dtNewProduct.Rows.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(dtNewProduct.Rows[0]["smt_top_id"].ToString()))
                            {
                                sTopId = dtNewProduct.Rows[0]["smt_top_id"].ToString();
                                sTOPVal = dtNewProduct.Rows[0]["pptc_no_of_days"].ToString();
                            }
                            else
                            {
                                MessageBox.Show("Mapping TOP Belum Disetting untuk Customer '" + sCust1 + "'.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Mapping TOP Belum Disetting untuk Customer '" + sCust1 + "'.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }
                }
            }
            #endregion

            keyFound = true;
            return keyFound;
        }

        private bool fGetCost(string sPcode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;
            try
            {
                DataTable dtFill1 = new DataTable();
                string strSQLL = "select isnull(isnull(wlb_last_avg_ucost, cs_standard_unit_cost), 0) as wlb_last_avg_ucost from IM_STOCK_BALANCE WITH (NOLOCK) LEFT JOIN IM_COST_STD WITH (NOLOCK) ON wlb_prd_master_code = cs_prd_master_code ";
                strSQLL = strSQLL + " AND wlb_grade = cs_grade AND wlb_prd_size = cs_size";
                strSQLL = strSQLL + " where wlb_loc_Id1 =  '" + sWHLoc1 + "' ";
                strSQLL = strSQLL + " and wlb_loc_Id2 =  '" + sWHLoc2 + "' and wlb_entity_id = '" + cbEntityBranch.txtCM.Text + "' and wlb_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                strSQLL = strSQLL + " and wlb_prd_master_code =  '" + sPcode + "' ";
                strSQLL = strSQLL + " and wlb_grade =  '" + sGrade + "' ";
                strSQLL = strSQLL + " and wlb_prd_size =  '" + sSize + "' ";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQLL);
                if (dtFill1.Rows.Count > 0)
                {
                    _curCost = Convert.ToDecimal(dtFill1.Rows[0]["wlb_last_avg_ucost"].ToString());
                    keyFound = true;
                    return keyFound;
                }
                else
                {
                    MessageBox.Show("Cost untuk produk " + sPcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fGetCostEcommerce(string sEntity, string sBranch, string sPcode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;
            try
            {
                DataTable dtFill1 = new DataTable();
                string strSQLL = "select isnull(isnull(wlb_last_avg_ucost, cs_standard_unit_cost), 0) as wlb_last_avg_ucost from IM_STOCK_BALANCE  LEFT JOIN IM_COST_STD WITH (NOLOCK) ON wlb_prd_master_code = cs_prd_master_code ";
                strSQLL = strSQLL + " AND wlb_grade = cs_grade AND wlb_prd_size = cs_size";
                strSQLL = strSQLL + " where wlb_loc_Id1 =  '" + sWHLoc1 + "' ";
                strSQLL = strSQLL + " and wlb_loc_Id2 =  '" + sWHLoc2 + "' and wlb_entity_id = '" + sEntity + "' and wlb_branch_id = '" + sBranch + "' ";
                strSQLL = strSQLL + " and wlb_prd_master_code =  '" + sPcode + "' ";
                strSQLL = strSQLL + " and wlb_grade =  '" + sGrade + "' ";
                strSQLL = strSQLL + " and wlb_prd_size =  '" + sSize + "' ";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQLL);
                if (dtFill1.Rows.Count > 0)
                {
                    _curCost = Convert.ToDecimal(dtFill1.Rows[0]["wlb_last_avg_ucost"].ToString());
                    keyFound = true;
                    return keyFound;
                }
                else
                {
                    MessageBox.Show("Cost untuk produk " + sPcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fGetCreditLimit(string sCust1, string sCust2)
        {
            bool keyFound = false;
            try
            {
                decimal currCredLimit = 0;

                _curAvCredLimit = 0;
                DataTable dtFill1 = new DataTable();
                strSQL = "select cm_status_cr, isnull(cm_ship_credit_limit, 0) -  isnull(cm_ar_saldo, 0) - isnull(cm_so_comitted, 0) as cred_limit  from SO_CUST_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + "  where cm_cust_code1 = '" + sCust1 + "' ";
                strSQL = strSQL + " and cm_cust_code2 = '" + sCust2 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    _curAvCredLimit = Convert.ToDecimal(dtFill1.Rows[0]["cred_limit"].ToString());

                    if (dtFill1.Rows[0]["cred_limit"].ToString() == "N")
                    {
                        bFlagCr = false;
                        keyFound = false;
                    }
                    else
                    {
                        bFlagCr = true;
                        keyFound = true;
                    }
                }
                else
                {
                    _curAvCredLimit = 0;
                    MessageBox.Show("Data Credit Limit tidak ada untuk Customer : " + sCust1, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                return keyFound;
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                keyFound = false;
                return keyFound;

            }
            return keyFound;
        }

        private bool fGetCreditLimit(string sEntity, string sBranch, string sCust1, string sCust2)
        {
            bool keyFound = false;
            try
            {
                decimal currCredLimit = 0;

                _curAvCredLimit = 0;
                DataTable dtFill1 = new DataTable();
                strSQL = "select cm_status_cr, isnull(cm_ship_credit_limit, 0) -  isnull(cm_ar_saldo, 0) - isnull(cm_so_comitted, 0) as cred_limit  from SO_CUST_MASTER ";
                strSQL = strSQL + "  where cm_cust_code1 = '" + sCust1 + "' ";
                strSQL = strSQL + " and cm_cust_code2 = '" + sCust2 + "' and cm_entity = '" + sEntity + "' and cm_branch = '" + sBranch + "' ";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    _curAvCredLimit = Convert.ToDecimal(dtFill1.Rows[0]["cred_limit"].ToString());

                    if (dtFill1.Rows[0]["cred_limit"].ToString() == "N")
                    {
                        bFlagCr = false;
                        keyFound = false;
                    }
                    else
                    {
                        bFlagCr = true;
                        keyFound = true;
                    }
                }
                else
                {
                    _curAvCredLimit = 0;
                    MessageBox.Show("Data Credit Limit tidak ada untuk Customer : " + sCust1, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                return keyFound;
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                keyFound = false;
                return keyFound;

            }
            return keyFound;
        }


        private bool fChangeSOType()
        {
            bool keyFound = false;

            //DataRow selectedDataRow = ((DataRowView)cbTipeOrder.SelectedItem).Row;
            //int idO = Convert.ToInt32(selectedDataRow["Code"]);
            //string NameO = selectedDataRow["order_type"].ToString();

            //lblOrderType.Text = idO.ToString();

            DataTable dtFill = new DataTable();
            //int key = Convert.ToInt32(cbTipeOrder.SelectedValue);
            strSQL = " select msow_wh_loc1 as id,msow_wh_loc1+msow_wh_loc2 + ' ~ ' + wh_loc_name as wh_loc from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) INNER JOIN IM_WH_LOC WITH (NOLOCK) on wh_loc_id1=msow_wh_loc1 and wh_loc_id2=msow_wh_loc2 and wh_loc_entity = msow_entity_id  and wh_branch_id = msow_branch_id where msow_so_type = '" + cbTipeOrder.SelectedValue.ToString() + "'  and msow_sld_id = '" + txtSalesman.Text + "' and wh_loc_entity = '" + cbEntityBranch.txtCM.Text + "' and wh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                cbWH.DataSource = _clsGlobal.ExecDT(strSQL);
                cbWH.ValueMember = "id";
                cbWH.DisplayMember = "wh_loc";
                cbWH.SelectedIndex = -1;

                keyFound = true;
            }
            return keyFound;
        }

        private bool fChangeWH()
        {
            bool keyFound = false;

            DataRow selectedDataRow = ((DataRowView)cbTipeOrder.SelectedItem).Row;
            int idO = Convert.ToInt32(selectedDataRow["Code"]);
            string NameO = selectedDataRow["order_type"].ToString();

            //lblOrderType.Text = idO.ToString();

            DataTable dtFill = new DataTable();
            //int key = Convert.ToInt32(cbTipeOrder.SelectedValue);
            strSQL = " select msow_wh_loc1, msow_wh_loc2 from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK)  where msow_so_type = '" + idO + "'  and msow_sld_id = '" + txtSalesman.Text + "' and msow_entity_id = '" + cbEntityBranch.txtCM.Text + "' and msow_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                lblWHLoc1.Text = dtFill.Rows[0]["msow_wh_loc1"].ToString();
                lblWHLoc2.Text = dtFill.Rows[0]["msow_wh_loc2"].ToString();

                keyFound = true;
            }
            else
            {
                lblWHLoc1.Text = "";
                lblWHLoc2.Text = "";
                //lblOrderType.Text = "";
                MessageBox.Show("Gudang pada Mapping Order type : " + cbTipeOrder.Text + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                keyFound = false;
            }

            DataTable dtFill2 = new DataTable();
            strSQL = " select cast(convert(varchar(10), wh_loc_last_work_date, 120) as datetime) as wh_loc_last_work_date  from IM_WH_LOC WITH (NOLOCK) ";
            strSQL = strSQL + " where wh_loc_id1 = '" + lblWHLoc1.Text + "'";
            strSQL = strSQL + " AND wh_loc_id2 = '" + lblWHLoc2.Text + "' and wh_loc_entity = '" + cbEntityBranch.txtCM.Text + "' and wh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'";
            dtFill2 = _clsGlobal.ExecDT(strSQL);

            if (dtFill2.Rows.Count > 0)
            {
                dtSODate.Text = Convert.ToDateTime(dtFill2.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                //dtTglPO.Text = Convert.ToDateTime(dtFill2.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                keyFound = true;
            }
            else
            {
                lblWHLoc1.Text = "";
                lblWHLoc2.Text = "";
                //lblOrderType.Text = "";
                MessageBox.Show("Tanggal Gudang : " + lblWHLoc1.Text + " - " + lblWHLoc2.Text + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                keyFound = false;
            }


            return keyFound;
        }

        private bool fLoadSourceData()
        {
            bool keyFound = false;

            DataTable dtFill = new DataTable();
            strSQL = " select gh_function_code,gh_function_desc from gs_gen_hardcoded  WITH (NOLOCK) where gh_function_name = 'SOURCEDATA' and gh_sys = 'H'";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                cbSource.DataSource = _clsGlobal.ExecDT(strSQL);
                cbSource.ValueMember = "gh_function_code";
                cbSource.DisplayMember = "gh_function_desc";
                cbSource.SelectedValue = "M";

                keyFound = true;
            }
            return keyFound;
        }

        private bool fLoadFlagSloc()
        {
            bool keyFound = false;

            DataTable dtFill = new DataTable();
            strSQL = " select gh_function_code,gh_function_desc from gs_gen_hardcoded WITH (NOLOCK) where gh_function_name = 'MULTISOURCE' and gh_sys = 'H'";
            dtFill = _clsGlobal.ExecDT(strSQL);

            if (dtFill.Rows.Count > 0)
            {
                cbFlagSloc.DataSource = _clsGlobal.ExecDT(strSQL);
                cbFlagSloc.ValueMember = "gh_function_code";
                cbFlagSloc.DisplayMember = "gh_function_desc";
                cbFlagSloc.SelectedValue = "";

                keyFound = true;
            }
            return keyFound;
        }

        private void pClearGrid()
        {
            dgvHeader.DataSource = null;
            dgvDetail.DataSource = null;
            dgvdata.DataSource = null;
            dgvHeader.Rows.Clear();
            dgvDetail.Rows.Clear();
            //dgvDetail2.Rows.Clear();
            dgvdata.Rows.Clear();
        }

        private void pWHDate()
        {
            DataTable dtFill1 = new DataTable();
            strSQL = "select isnull(isnull(wlb_last_avg_ucost, cs_standard_unit_cost), 0) as wlb_last_avg_ucost from IM_STOCK_BALANCE WITH (NOLOCK)  LEFT JOIN IM_COST_STD WITH (NOLOCK)  ON wlb_prd_master_code = cs_prd_master_code where wlb_entity_id = '" + cbEntityBranch.txtCM.Text + "' and wlb_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
            dtFill1 = _clsGlobal.ExecDT(strSQL);
            if (dtFill1.Rows.Count > 0)
            {
                dtSODate.Text = dtFill1.Rows[0]["wh_loc_last_work_date"].ToString();
            }
            else
            {
                dtSODate.Value = DateTime.Now;
                btnOk.Enabled = false;
                MessageBox.Show("Tanggal gudang tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void SetGridData(DataGridView dgv)
        {
            dtGridData = new DataTable();
            if (!chkEcommerce.Checked)
            {
                dtGridData.Columns.Add("C0", typeof(string));
                dtGridData.Columns.Add("C1", typeof(string));
                dtGridData.Columns.Add("C2", typeof(string));
                dtGridData.Columns.Add("C3", typeof(string));
                dtGridData.Columns.Add("C4", typeof(string));
                dtGridData.Columns.Add("C5", typeof(string));
                dtGridData.Columns.Add("C6", typeof(string));
                dtGridData.Columns.Add("C7", typeof(string));
                //dtGridData.Columns.Add("C8", typeof(string));
                //dtGridData.Columns.Add("C9", typeof(string));
                //dtGridData.Columns.Add("C10", typeof(string));
                //dtGridData.Columns.Add("C11", typeof(string));
                //dtGridData.Columns.Add("C12", typeof(string));
                //dtGridData.Columns.Add("C13", typeof(string));
                //dtGridData.Columns.Add("C14", typeof(string));
            }
            else
            {
                dtGridData.Columns.Add("C0", typeof(string));
                dtGridData.Columns.Add("C1", typeof(string));
                dtGridData.Columns.Add("C2", typeof(string));
                dtGridData.Columns.Add("C3", typeof(string));
                dtGridData.Columns.Add("C4", typeof(string));
                dtGridData.Columns.Add("C5", typeof(string));
                dtGridData.Columns.Add("C6", typeof(string));
                dtGridData.Columns.Add("C7", typeof(string));
                dtGridData.Columns.Add("C8", typeof(string));
                dtGridData.Columns.Add("C9", typeof(string));
                dtGridData.Columns.Add("C10", typeof(string));
                dtGridData.Columns.Add("C11", typeof(string));
                //dtGridData.Columns.Add("C12", typeof(string));
                //dtGridData.Columns.Add("C13", typeof(string));
                //dtGridData.Columns.Add("C14", typeof(string));
            }

            dgv.DataSource = dtGridData;
            //dgv.Rows.Clear();
        }

        //mdlutility
        private void fQtyFormat(string sQty, int iConv1, int iConv2)
        {
            string sFormatQty = "";
            int iSplit = 0;
            string[] sQuantity = sQty.Split('.');
            iSplit = sQuantity.GetUpperBound(0);
            string cNol = "";
            string cSatu = "";
            string cDua = "";

            string __fNol = "00000";
            string __fSatu = "000";
            string __fDua = "0000";
            if (isNefoKAM)
            {
                __fDua = "00000";
            }

            if (iSplit == 0)
            {
                cNol = __fNol + sQuantity[0].ToString();
                sFormatQty = cNol.Substring(cNol.Length - __fNol.Length) + "." + __fSatu + "." + __fDua;
            }
            else if (iSplit == 1)
            {
                cNol = __fNol + sQuantity[0].ToString();
                cSatu = __fSatu + sQuantity[1].ToString();
                cDua = __fDua + sQuantity[2].ToString();
                sFormatQty = cNol.Substring(cNol.Length - __fNol.Length) + "." + cSatu.Substring(cSatu.Length - __fSatu.Length) + "." + __fDua;
            }
            else if (iSplit == 2)
            {
                cNol = __fNol + sQuantity[0].ToString();
                cSatu = __fSatu + sQuantity[1].ToString();
                cDua = __fDua + sQuantity[2].ToString();
                sFormatQty = cNol.Substring(cNol.Length - __fNol.Length) + "." + cSatu.Substring(cSatu.Length - __fSatu.Length) + "." + cDua.Substring(cDua.Length - __fDua.Length);
            }
            else if (iSplit > 2)
            {
                cNol = __fNol + sQuantity[0].ToString();
                cSatu = __fSatu + sQuantity[1].ToString();
                cDua = __fDua + sQuantity[2].ToString();
                sFormatQty = cNol.Substring(cNol.Length - __fNol.Length) + "." + cSatu.Substring(cSatu.Length - __fSatu.Length) + "." + cDua.Substring(cDua.Length - __fDua.Length);
            }

            if (iConv1 == 0 && iConv2 == 0)
            {
                _fQtyFormat = sFormatQty;
            }
            else
            {
                fConvertQty(sFormatQty, iConv1, iConv2);
                DataTable dtFill1 = new DataTable();
                strSQL = "select dbo.SQTY_FORMAT (" + _fConvertQty + ", " + iConv1 + "," + iConv2 + ") as Qty_Format";
                if (_clsGlobal.tr == null)
                {
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                }
                else if (_clsGlobal.tr.Connection == null)
                {
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                }
                else
                {
                    dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                }

                if (dtFill1.Rows.Count > 0)
                {
                    _fQtyFormat = dtFill1.Rows[0]["Qty_Format"].ToString();
                }
            }

        }

        //mdlutility
        private void fConvertQty(string sQty, int iConv1, int iConv2)
        {
            //double keyFound = 0;
            _fConvertQty = "0";
            int iSplit = 0;
            string[] sQuantity = sQty.Split('.');
            iSplit = sQuantity.Length;

            if (iSplit == 0)
            {
                if (sQuantity[0].ToString() == "")
                {
                    sQuantity[0] = "0";
                }
                _fConvertQty = sQuantity[0].ToString();
            }
            else if (iSplit == 1)
            {
                if (sQuantity[0].ToString() == "")
                {
                    sQuantity[0] = "0";
                }
                if (sQuantity[1].ToString() == "")
                {
                    sQuantity[1] = "0";
                }

                int qtyconv = (iConv1 * Convert.ToInt32(sQuantity[0])) + (iConv2 * Convert.ToInt32(sQuantity[1]));

                _fConvertQty = qtyconv.ToString();

            }
            else if (iSplit == 2)
            {
                if (sQuantity[0] == "")
                {
                    sQuantity[0] = "0";
                }
                if (sQuantity[1] == "")
                {
                    sQuantity[1] = "0";
                }
                if (sQuantity[2] == "")
                {
                    sQuantity[2] = "0";
                }
                int qtyconv = (iConv1 * Convert.ToInt32(sQuantity[0])) + (iConv2 * Convert.ToInt32(sQuantity[1])) + Convert.ToInt32(sQuantity[2]);
                _fConvertQty = qtyconv.ToString();
            }
            else if (iSplit > 2)
            {
                if (sQuantity[0] == "")
                {
                    sQuantity[0] = "0";
                }
                if (sQuantity[1] == "")
                {
                    sQuantity[1] = "0";
                }
                if (sQuantity[2] == "")
                {
                    sQuantity[2] = "0";
                }

                int qtyconv = (iConv1 * Convert.ToInt32(sQuantity[0])) + (iConv2 * Convert.ToInt32(sQuantity[1])) + Convert.ToInt32(sQuantity[2]);
                _fConvertQty = qtyconv.ToString();
            }

            //return keyFound;
        }

        private bool fCekMandatoryProses(int iRow)
        {
            bool keyFound = false;
            string sPONo;

            sPONo = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString();

            if (dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString() == "")
            {
                MessageBox.Show("Outlet masih kosong!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                _clsGlobal.RollbackTrans();
                keyFound = false;
                return keyFound;
            }

            if (!chkEcommerce.Checked)
            {
                if (cbTipeOrder.Text == "")
                {
                    MessageBox.Show("Order Type masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    _clsGlobal.RollbackTrans();
                    keyFound = false;
                    return keyFound;
                }
            }


            for (int lCounterDetail = 0; lCounterDetail < dgvDetail.Rows.Count; lCounterDetail++)
            {
                if (dgvDetail.Rows[lCounterDetail].Cells["g_iDetPONO"].Value.ToString().CompareC(sPONo))
                {
                    if (dgvDetail.Rows[lCounterDetail].Cells["g_iDetPCode"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                    {
                        MessageBox.Show("Product masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        _clsGlobal.RollbackTrans();
                        keyFound = false;
                        return keyFound;
                    }

                    if (dgvDetail.Rows[lCounterDetail].Cells["g_iDetQtyOrd"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                    {
                        MessageBox.Show("Order Qty masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        _clsGlobal.RollbackTrans();
                        keyFound = false;
                        return keyFound;
                    }
                }
            }

            if (!chkEcommerce.Checked)
            {
                //'Cek Week dan Month berjalan
                DataTable dtFill1 = new DataTable();
                strSQL = "select fwd_entity From IM_WH_LOC WITH (NOLOCK) inner join TBL_GS_WORKING_DAY WITH (NOLOCK) on convert(varchar(10), fwd_work_day, 120) = convert(varchar(10), wh_loc_last_work_date, 120)" +
                    "  and fwd_week_no = wh_loc_week_closing and fwd_periode = wh_loc_month_closing and fwd_year = wh_loc_year_closing and fwd_entity = wh_loc_entity and fwd_branch = wh_branch_id" +
                    " where wh_loc_id1 ='" + lblWHLoc1.Text + "'  and wh_loc_id2 ='" + lblWHLoc2.Text + "' and wh_loc_entity = '" + cbEntityBranch.txtCM.Text + "' and wh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "'  ";
                dtFill1 = _clsGlobal.ExecDT(strSQL);

                if (dtFill1.Rows.Count > 0)
                {
                }
                else
                {
                    MessageBox.Show("Tanggal transaksi Gudang tidak di pekan/periode/year berjalan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }
            }
            keyFound = true;

            return keyFound;
        }

        private bool fCheckMandatory()
        {
            bool keyFound = false;

            if (cbEntityBranch.txtCM.Text == "")
            {
                MessageBox.Show("Entity belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                keyFound = false;
                return keyFound;
            }

            if (cbEntityBranch.txtBranchIdCM.Text == "")
            {
                MessageBox.Show("Branch belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                keyFound = false;
                return keyFound;
            }

            //if (txtAccountDesc.Text == "")
            //{
            //    MessageBox.Show("Account E-Commerce belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    keyFound = false;
            //    return keyFound;
            //}

            if (!chkEcommerce.Checked)
            {
                if (txtSalesman.Text == "")
                {
                    MessageBox.Show("Salesman belum di isi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }
            }


            if (lblPath.Text == "")
            {
                MessageBox.Show("File belum di isi ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                keyFound = false;
                return keyFound;
            }
            if (!chkEcommerce.Checked)
            {
                if (cbTipeOrder.Text == "")
                {
                    MessageBox.Show("Order Type belum di isi ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }
            }


            if (!chkEcommerce.Checked)
            {
                if (lblWHLoc1.Text == "")
                {
                    MessageBox.Show("Kode Gudang masih kosong, pastikan sudah di setting pada parameter Salesman", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }
            }


            //if (txtAccountDesc.Text == "JD.ID" && txtSheet.Text == "")
            //{
            //    MessageBox.Show("Sheet name belum di isi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    keyFound = false;
            //    return keyFound;
            //}

            DataTable dtFill14 = new DataTable();
            strSQL = " select CAST(isnull(gh_last_seq_no,0) AS MONEY) as param_selisih_amt FROM GS_GEN_HARDCODED WITH (NOLOCK) ";
            strSQL = strSQL + " WHERE gh_sys = 'H' AND gh_function_name = 'SO_ECOMMERCE' AND gh_function_code = '9999'";
            dtFill14 = _clsGlobal.ExecDT(strSQL);
            if (dtFill14.Rows.Count > 0)
            {
                eParSelisih = Convert.ToDecimal(dtFill14.Rows[0]["param_selisih_amt"]);
            }
            else
            {
                MessageBox.Show("Parameter selisih Price belum disetting", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                keyFound = false;
                return keyFound;
            }


            if (isMultiSource)
            {
                dtFill14 = new DataTable();
                string __source = cbSource.SelectedValue.ToString();
                string __flagsloc = cbFlagSloc.SelectedValue.ToString();
                strSQL = " select * from GS_PRM_MULTI_SOURCE WITH (NOLOCK) ";
                strSQL += " where pms_entity_id = '" + cbEntityBranch.txtCM.Text + "' and pms_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                strSQL += " and pms_flag_source = '" + __source.ToUpper().Replace('M', 'N').ToString() + "' and pms_flag_sloc = '" + __flagsloc + "' ";
                dtFill14 = _clsGlobal.ExecDT(strSQL);
                if (dtFill14.Rows.Count == 0)
                {
                    MessageBox.Show("Multisource : Kombinasi Source " + __source + " dan FlagSloc " + __flagsloc + " tidak ada pada table master!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }
            }

            keyFound = true;

            return keyFound;
        }

        #endregion

        private void tsb_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void fGetWHDate()
        {
            //bool keyFound = false;

            try
            {
                DataTable dtFill1 = new DataTable();
                strSQL = " exec SP_CURRENT_BOOK_MONTH '" + cbEntityBranch.txtCM.Text + "','" + cbEntityBranch.txtBranchIdCM.Text + "'";
                dtFill1 = _clsGlobal.ExecDT(strSQL);
                if (dtFill1.Rows.Count > 0 && dtFill1.Rows[0]["wh_loc_last_work_date"].ToString().Trim() != "")
                {
                    dtSODate.Value = Convert.ToDateTime(dtFill1.Rows[0]["wh_loc_last_work_date"].ToString());
                    g_dTglGudang = Convert.ToDateTime(dtFill1.Rows[0]["wh_loc_last_work_date"].ToString());
                    //keyFound = true;
                }
                else
                {
                    //MessageBox.Show("Tanggal untuk Gudang tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //keyFound = false;
                    //return keyFound;
                }
            }
            catch (Exception ex)
            {
                //_clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //return keyFound;
        }

        protected void BindDefaultBranch(TextBox txtEntity, TextBox txtBranch)
        {
            try
            {
                DataTable dt = new DataTable();
                strSQL = "select distinct ge_entity_id [Entity], ge_entity [Desc]  from GS_ENTITY WITH (NOLOCK) INNER JOIN GS_USERS_SECURITY u WITH (NOLOCK) on u.gu_entity = ge_entity_id where gu_user_id = '" + clsLogin.USERID + "' order by ge_entity_id ";
                dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    txtEntity.Text = dt.Rows[0]["Entity"].ToString();
                }
                DataTable dtBranch = new DataTable();

                strSQL = "select gu_branch [Branch], br_branch_desc [Desc] FROM VW_USERS_SECURITY WITH (NOLOCK)  where gu_entity = '" + txtEntity.Text.Trim() + "' and gu_user_id = '" + clsLogin.USERID + "'";
                dtBranch = _clsGlobal.ExecDT(strSQL);
                if (dtBranch.Rows.Count > 0)
                {
                    txtBranch.Text = dtBranch.Rows[0]["Branch"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void dgvRowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            try
            {
                int __index = e.RowIndex;
                DataGridView __target = (DataGridView)sender;
                if (__index % 2 == 1)
                    __target.Rows[__index].DefaultCellStyle.BackColor = __COLOR;
            }
            catch (Exception) { }
        }
        private DataTable InitNefoKAMProsesMigrasiSAPKAM_HEADER(string sPONO, bool top_by_cust)
        {

            string PCode1 = "";
            string PCode2 = "";


            DataTable DtTemp = new DataTable();
            DtTemp.Columns.Add("TOPID");
            DtTemp.Columns.Add("TOPVALUE");

            for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
            {
                if (sPONO == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString() && dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() != "")
                {
                    string PL = dgvDetail2.Rows[lCounter].Cells["g_iDetProdLine1"].Value.ToString();

                    string __TOP = "P0";
                    try
                    {
                        if (!string.IsNullOrEmpty(dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString()))
                        {
                            __TOP = dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString();
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                    DataView dv = DtTemp.DefaultView;
                    dv.RowFilter = "TOPID = '" + __TOP + "'";

                    DataTable dTempDv = dv.ToTable();
                    if (dTempDv.Rows.Count == 0)
                    {
                        if (!string.IsNullOrEmpty(PL))
                        {
                            if (sPayType == "T" || top_by_cust || isDivisionTOP == false)
                            {
                                DtTemp.Rows.Add(sTopDefault, sTopDefaultValue);
                            }
                            else
                            {
                                strSQL = "select distinct top 1 PM.prm_prd_line_code pcodelinenew, SMT.smt_prdline_id  plinemap, SMT.smt_top_id, PTC.pptc_no_of_days ";
                                strSQL += "   from SO_CUST_MASTER CM WITH (NOLOCK) ";
                                strSQL += "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                                strSQL += "       ON  ";
                                strSQL += "           TSC.csc_entity          = CM.cm_entity ";
                                strSQL += "           AND TSC.csc_branch      = CM.cm_branch ";
                                strSQL += "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 ";
                                strSQL += "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 ";
                                strSQL += "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK) ";
                                strSQL += "       ON	";
                                strSQL += "           SPG.sgm_spgm_id  = TSC.csc_salesman_id ";
                                strSQL += "           AND SPG.sgm_entity_id = TSC.csc_entity ";
                                strSQL += "           AND SPG.sgm_branch_id = TSC.csc_branch ";
                                strSQL += "   left join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_entity_id      = CM.cm_entity ";
                                //strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                                //strSQL += "           and SMT.smt_cust_code1 = CM.cm_cust_code1 ";
                                //strSQL += "           and SMT.smt_cust_code2 = CM.cm_cust_code2 ";
                                if (!isMultibranch)
                                {
                                    strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                                    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                }
                                else
                                {
                                    strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                                    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                }
                                strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                                strSQL += "   where ";
                                strSQL += "   SPG.sgm_spgm_id = '" + txtSalesman.Text.Trim() + "' ";
                                strSQL += "   and SPG.sgm_entity_id = '" + cbEntityBranch.txtCM.Text + "' ";
                                strSQL += "   and SPG.sgm_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text + "' ";
                                strSQL += "   and CM.cm_cust_code1 = '" + sCustNo + "' ";
                                strSQL += "   and SMT.smt_prdline_id = '" + PL + "' ";
                                strSQL += "   and CM.cm_top_by_cust = 'N' ";

                                DataTable dtNewProduct = _clsGlobal.ExecDTTrans(strSQL);
                                if (dtNewProduct.Rows.Count > 0)
                                {
                                    DataView dv2 = DtTemp.DefaultView;
                                    dv2.RowFilter = "TOPID = '" + dtNewProduct.Rows[0]["smt_top_id"].ToString() + "'";
                                    DataTable dTempDv2 = dv2.ToTable();
                                    if (dTempDv2.Rows.Count == 0)
                                    {
                                        DtTemp.Rows.Add(dtNewProduct.Rows[0]["smt_top_id"].ToString(), dtNewProduct.Rows[0]["pptc_no_of_days"].ToString());
                                    }

                                    //dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value = dtNewProduct.Rows[0]["smt_top_id"].ToString();



                                }
                                else
                                {
                                    MessageBox.Show("TOP Mapping untuk product line code = '" + PL + "' belum disetting.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }

                            }
                        }
                    }

                }
            }
            return DtTemp;

        }

        private bool fSaveSOMigrasiSAP_KAM(string bExtSono, string sCredLimit, bool bCalcDisc, int iRow, string TOP, string ValueTOP)
        {
            bool keyFound = false;

            string sBranch = "";
            string sArea = "";
            string sWilayah = "";
            string sRayon = "";
            string sCustGroup = "";
            string sLineCode = "";
            string sYear = "";
            string sPeriode = "";
            //string sBranch = "";
            string sWeekNo = "";

            string sSOStatus = "";
            long lROrdQty = 0;
            decimal curROrdAmt = 0;
            decimal curCost = 0;
            decimal curTotalCost = 0;
            //clsPCode
            string sCurrCode = "";
            double dRateCurr = 0;
            long lBrs = 0;
            decimal curDisc1 = 0;
            decimal curDisc2 = 0;
            string sErrDesc = "";
            decimal currCredLimit = 0;
            //bool bFlagCr = false;
            string sPriceCode = "";
            string sOpenStatus = "";
            long lOrdQty = 0;
            long lLeadTime = 0;
            decimal curOrdAmt = 0;
            string pdaNO = "";
            string pdaFlag = "";
            string extsap = "";
            string runcext = "";
            string QasirFlag = "";
            string nik = "";

            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cLineCode = "";
            string cGroup = "";
            string cSubGroup = "";
            string cModel = "";
            string cRawMatUsed = "";
            string cWashingCollorCode = "";
            long cConv1 = 0;
            long cConv2 = 0;

            string sCustCode1;
            string sCustCode2;
            string sPONo;
            //string sSoNo;
            Dictionary<string, string> SOPrincipal;
            string qasirflag = "";
            string sFullfilment = "N";

            try
            {
                keyFound = false;

                sPONo = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString();
                sCustCode1 = dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString();
                sCustCode2 = dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value.ToString();


                if (!cbSource.SelectedValue.ToString().CompareC("M"))
                {
                    qasirflag = cbSource.SelectedValue.ToString().Trim();
                }
                else
                {
                    qasirflag = "N";
                }
                if (bCalcDisc == false)
                {
                    if (fGetSONumber() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    sSONo = g_sDefaultSONo;

                    strSQL = " EXEC SP_DELETE_SO_TABLE ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }


                DataTable dtNIK = new DataTable();
                strSQL = "select sah_NIK from tbl_salesman_history WITH (NOLOCK)" +
                         " where" +
                         " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and sah_entity_id = '" + cbEntityBranch.txtCM.Text.Trim() + "'" +
                         " and sah_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'" +
                         " and sah_salesmen_id = '" + txtSalesman.Text + "'";
                dtNIK = _clsGlobal.ExecDTTrans(strSQL);

                if (dtNIK.Rows.Count == 1)
                {
                    nik = dtNIK.Rows[0]["sah_NIK"].ToString();
                }
                else if (dtNIK.Rows.Count > 1)
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }

                    MessageBox.Show("Salesman History Pada Periode Entry Lebih dari 1", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return keyFound;
                }
                else
                {

                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Salesman History Tidak Ditemukan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }



                //'Get branch, area, wilayah, rayon, cust group dan line code
                DataTable dtFill8 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') cm_line_code, cm_lead_time ,ISNULL(cm_fullfilment,'N') cm_fullfilment From SO_CUST_MASTER WITH (NOLOCK)";
                strSQL = strSQL + " where cm_cust_code1 = '" + sCustCode1 + "' AND";
                strSQL = strSQL + " cm_cust_code2 = '" + sCustCode2 + "' and cm_entity = '" + cbEntityBranch.txtCM.Text.Trim() + "' and cm_branch = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' ";
                dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill8.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill8.Rows[0]["cm_branch"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Area untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill8.Rows[0]["cm_area"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Wilayah untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill8.Rows[0]["cm_wilayah"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Rayon untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill8.Rows[0]["cm_rayon"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Cust group untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill8.Rows[0]["cm_cust_group"].ToString();
                    }

                    sLineCode = dtFill8.Rows[0]["cm_line_code"].ToString();
                    sFullfilment = dtFill8.Rows[0]["cm_fullfilment"].ToString();

                    //if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_lead_time"].ToString()))
                    //{
                    //    MessageBox.Show("Lead Time untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}
                    //else
                    //{
                    //    lLeadTime = Convert.ToInt64(dtFill8.Rows[0]["cm_lead_time"].ToString());
                    //}

                }
                else
                {
                    sBranch = "";
                    sArea = "";
                    sWilayah = "";
                    sRayon = "";
                    sCustGroup = "";
                    sLineCode = "";
                    sFullfilment = "N";
                }

                //'Get Year, Periode, Week untuk Tgl SO
                DataTable dtFill9 = new DataTable();
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK) ";
                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + dtSODate.Value.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '" + cbEntityBranch.txtCM.Text.Trim() + "' and fwd_branch = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'"; //entity id harus di lempar
                dtFill9 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill9.Rows.Count > 0)
                {
                    sWeekNo = dtFill9.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill9.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill9.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + dtSODate.Value.ToString("dd MMM yyyy") + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get SO Status
                DataTable dtFill10 = new DataTable();
                strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                strSQL = strSQL + " and gh_sequence_no = '1' ";
                dtFill10 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill10.Rows.Count > 0)
                {
                    sSOStatus = dtFill10.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    MessageBox.Show("Status SO (Open) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


                //Get Price Code
                //DataTable dtFill11 = new DataTable();
                //strSQL = " exec SP_GET_PRICE_CODE '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + sCustCode1 + "','" + sCustCode2 + "'";
                //dtFill11 = _clsGlobal.ExecDTTrans(strSQL);
                //if (dtFill11.Rows.Count > 0)
                //{
                //    if (String.IsNullOrEmpty(dtFill11.Rows[0]["cm_def_price_code"].ToString()))
                //    {
                //        MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        keyFound = false;
                //        return keyFound;
                //    }
                //    else
                //    {
                //        sPriceCode = dtFill11.Rows[0]["cm_def_price_code"].ToString();
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvDetail.Rows.Count; lCounter++)
                {
                    //if (sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString() && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != "" && TOP == dgvDetail.Rows[lCounter].Cells["g_iDetTOPDiv"].Value.ToString().Trim())
                    if (sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString() && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != "")
                    {
                        //fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + (int)Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curROrdAmt = curROrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            curROrdAmt = curROrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);
                        }

                        fConvertQty(dgvDetail.Rows[lCounter].Cells["g_iDetQtyOrd"].Value.ToString(), Convert.ToInt16(iConv1), Convert.ToInt16(iConv2));
                        lOrdQty = lOrdQty + Convert.ToInt32(_fConvertQty);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curOrdAmt = curOrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);

                        }
                        else
                        {
                            curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);

                        }

                        if (fGetCost(dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetGrade"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetSize"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        curTotalCost = curTotalCost + _curCost;
                    }
                }

                //if (sSONo.Length == 0)
                //{
                //    MessageBox.Show("nomor SO Blank", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //if (_xFlagDC == "N")
                //{
                // '===> Cek Kredit Limit dihilangkan jika flagDC = Y - Project Improvement TiraSnD
                //'/*-----------------------------------------------------------------------------------------
                //'Credit Limit
                if (fGetCreditLimit(sCustCode1, sCustCode2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (bFlagCr == true)
                {
                    if (_curAvCredLimit < curROrdAmt)
                    {
                        sCredLimit = "";
                    }
                    else
                    {
                        sCredLimit = "";
                    }
                }
                //}

                DataTable dtFill4 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_DETAIL A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsd_so_seq_no = '" + sSONo + "'";
                //strSQL = strSQL + "  AND wsh_branch_id = '" + cbBranchID.SelectedValue.ToString()  + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill5 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_HEADER A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "'";
                strSQL = strSQL + "  AND wsh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill12 = new DataTable();
                #region Query Insert
                strSQL = "Insert into SO_WEB_SALES_HEADER (";
                strSQL = strSQL + " wsh_entity_id ,wsh_area ,wsh_wilayah ,wsh_rayon ";
                strSQL = strSQL + " ,wsh_branch_id ,wsh_cust_group ,wsh_so_type ,wsh_seq_no";
                strSQL = strSQL + " ,wsh_so_date,wsh_cust_code1 ,wsh_cust_code2 ,wsh_loc_id1";
                strSQL = strSQL + " ,wsh_loc_id2,wsh_brand_line ,wsh_so_ref_no,wsh_so_txn_date_from";
                strSQL = strSQL + " ,wsh_so_txn_date_to,wsh_week_no,wsh_spgm_id ,wsh_spgm_group";
                strSQL = strSQL + " ,wsh_spgm_coordinator_id ,wsh_spgm_loc1 ,wsh_spgm_loc2 ,wsh_shift_no";
                strSQL = strSQL + " ,[wsh_participation_disc_%] ,wsh_participation_disc_amount ,wsh_tot_part_line_amount ,wsh_total_line";
                strSQL = strSQL + " , wsh_relevan_pod ,wsh_status_so ,wsh_real_order_qty   ,wsh_real_order_amt";
                strSQL = strSQL + " ,wsh_total_so_qty ,wsh_total_so_amt      ,wsh_include_tax ,[wsh_tax_%]";
                strSQL = strSQL + " ,wsh_total_partisipasi_disc_amount ,wsh_total_margin_disc_amount ,wsh_net_amount,wsh_dpp_amount";
                strSQL = strSQL + " ,wsh_tax_amount,wsh_invoice_process_flag ,wsh_invoice_no ,wsh_invoice_date";
                strSQL = strSQL + " ,wsh_invoice_process_by ,wsh_disc_margin,wsh_total_cost,wsh_year";
                strSQL = strSQL + " ,wsh_period,wsh_cancel_date,wsh_cancel_by,wsh_approval";
                strSQL = strSQL + " ,wsh_approval_by ,wsh_approval_date,wsh_ret_seq_no ,wsh_ret_by ,wsh_remarks";
                strSQL = strSQL + " ,wsh_participation_code ,wsh_entry_date,wsh_user_id ,wsh_from_OB_flag";
                strSQL = strSQL + " ,wsh_gross_net,wsh_gen_flag,wsh_tot_disc_pc,wsh_tot_disc_1";
                strSQL = strSQL + " ,wsh_tot_disc_2,wsh_tot_tpr_amt,wsh_tax_code,wsh_tax_pct";
                strSQL = strSQL + " ,wsh_pct_cash_disc, wsh_cash_disc, wsh_tax_pbm_pct, wsh_tax_pbm_amount, wsh_pay_type";
                strSQL = strSQL + " ,wsh_gross_net_amt, wsh_validate, wsh_po_no ";
                strSQL = strSQL + " ,wsh_flag_so_limit ";
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty,wsh_po_ol_flag ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                strSQL = strSQL + " , wsh_extract_sap ";
                strSQL = strSQL + " , wsh_runcode_ext  ";
                strSQL = strSQL + " , wsh_sch_date  ";
                strSQL = strSQL + " , wsh_expired_date  ";
                strSQL = strSQL + " , wsh_nik  ";
                strSQL = strSQL + " , wsh_qasir_flag  ";
                strSQL = strSQL + " , wsh_beban_cashdisc  ";
                if (isMultiSource)
                {
                    strSQL = strSQL + " , wsh_flag_sloc ";
                }
                strSQL = strSQL + " , wsh_is_fulfillment ";
                strSQL = strSQL + " , wsh_sled_flag ";

                string __tglorder = dtSODate.Value.ToString("yyyyMMdd");
                //int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '";  //'wsh_so_type
                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no
                //strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_so_date
                }
                else
                {
                    strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                }
                strSQL = strSQL + sCustCode1 + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + sCustCode2 + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + lblWHLoc1.Text + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + lblWHLoc2.Text + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + txtSalesman.Text + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "null , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "null , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "null , "; //'wsh_tot_part_line_amount
                strSQL = strSQL + "0, "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                strSQL = strSQL + "null , "; //'wsh_include_tax
                strSQL = strSQL + "null , "; //'[wsh_tax_%]
                strSQL = strSQL + "null , "; //'wsh_total_partisipasi_disc_amount
                strSQL = strSQL + "null , "; //'wsh_total_margin_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_net_amount
                strSQL = strSQL + "0 , "; //'wsh_dpp_amount
                strSQL = strSQL + "0 , "; //'wsh_tax_amount
                strSQL = strSQL + "'N' , "; //'wsh_invoice_process_flag
                strSQL = strSQL + "null , "; //'wsh_invoice_no
                strSQL = strSQL + "null , "; //'wsh_invoice_date
                strSQL = strSQL + "null , "; //'wsh_invoice_process_by
                strSQL = strSQL + "null , "; //'wsh_disc_margin
                strSQL = strSQL + curTotalCost + ", '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , "; //'wsh_approval
                strSQL = strSQL + "null , "; //'wsh_approval_by
                strSQL = strSQL + "null , "; //'wsh_approval_date
                strSQL = strSQL + "null , "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                strSQL = strSQL + "null , "; //'wsh_participation_code
                strSQL = strSQL + "getdate() , '"; //'wsh_entry_date
                strSQL = strSQL + clsLogin.USERID + "', "; //'wsh_user_id
                strSQL = strSQL + "null , "; //'wsh_from_OB_flag
                strSQL = strSQL + "null , "; //'wsh_gross_net
                strSQL = strSQL + "null , "; //'wsh_gen_flag
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_pc
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_1
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_2
                strSQL = strSQL + "0 , "; //'wsh_tot_tpr_amt
                strSQL = strSQL + "null , "; //'wsh_tax_code
                strSQL = strSQL + "null , "; //'wsh_tax_pct

                if (dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() == "")
                {
                    strSQL = strSQL + "0" + " , ";
                }
                else
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                }

                strSQL = strSQL + "0 , "; //'wsh_cash_disc
                //strSQL = strSQL + "0 , "; //'wsh_cash_disc
                strSQL = strSQL + "0 , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "0 , '"; //'wsh_tax_pbm_amount
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value.ToString() + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , "; //'wsh_gross_net_amt
                strSQL = strSQL + "null , '"; // Convert.ToDateTime(dtTglValidasi.Value).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + sPONo + "' , '"; //'wsh_po_no
                strSQL = strSQL + sCredLimit + "' , '"; //'wsh_flag_so_limit
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                strSQL = strSQL + "N' , '"; //'wsh_po_ol_flag
                strSQL = strSQL + ValueTOP + "' , '"; //'wsh_top
                strSQL = strSQL + TOP + "' , '"; //'wsh_top_id
                //strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "','"; //'wsh_po_date
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_po_date
                }
                else
                {
                    strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "','"; //'wsh_po_date
                }
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value.ToString() + "'"; //'wsh_cust_leadtime
                strSQL = strSQL + ",'" + "N" + "' , '"; //'extsap
                strSQL = strSQL + "'"; //'runcext
                //strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // planning date
                //strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // expired date
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + ",'" + dgvHeader.Rows[iRow].Cells["g_iDeliveryDate"].Value.ToString() + "'"; // planning date
                    strSQL = strSQL + ",'" + dgvHeader.Rows[iRow].Cells["g_iExpiredDate"].Value.ToString() + "'"; // expired date
                }
                else
                {
                    strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // planning date
                    strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // expired date
                }
                strSQL = strSQL + ",'" + nik + "'"; // nik
                //Convert.ToDateTime(dtTglOrder.Text).ToString("yyyyMMdd")
                strSQL = strSQL + ",'" + qasirflag + "'"; // qasirflag 30-06-2021 -- update BL : kalo M kosong kalo yang lain diisi sama dengan cbSource.selectedvalue
                //if (iscashdiscbeban) // kalau ada cashdisc
                //{
                //    strSQL = strSQL + ",1"; //wsh_beban_cashdisc
                //}
                //else // kalau ga ada cashdisc
                //{
                strSQL = strSQL + ",0"; //wsh_beban_cashdisc
                //}
                //Convert.ToDateTime(dtTglOrder.Text).ToString("yyyyMMdd")
                if (isMultiSource)
                {
                    strSQL = strSQL + ",'" + cbFlagSloc.SelectedValue.ToString() + "'";
                }
                strSQL = strSQL + ",'" + sFullfilment.Trim() + "'"; //wsh_is_fulfillment

                if (dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                {
                    strSQL = strSQL + ",'N'";
                }
                else
                {
                    strSQL = strSQL + ",'"+ dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().Trim() + "'";
                }
                

                strSQL = strSQL + " ) ";
                dtFill12 = _clsGlobal.ExecDTTrans(strSQL);
                //'END OF Saving SO Header
                //if (dtFill12.Rows.Count > 0)
                //{
                DataTable dtFill13 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK) ";
                strSQL = strSQL + " where cc_currency_type = 'H' ";
                #endregion

                dtFill13 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill13.Rows.Count > 0)
                {
                    sCurrCode = dtFill13.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill13.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }


                int lineNo = 0;
                lCounter = 0;
                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                {

                    if (sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString())
                    {
                        #region save detail
                        DataTable dtFill14 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK)";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "'";
                        dtFill14 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill14.Rows.Count > 0)
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            sSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = dtFill14.Rows[0]["prm_prd_line_code"].ToString();
                            cGroup = dtFill14.Rows[0]["prm_prd_group_code"].ToString();
                            cSubGroup = dtFill14.Rows[0]["prm_prd_sgroup_code"].ToString();
                            cModel = dtFill14.Rows[0]["prm_prd_model_code"].ToString();
                            cRawMatUsed = dtFill14.Rows[0]["prm_raw_mat_used_code"].ToString();
                            cWashingCollorCode = dtFill14.Rows[0]["prm_washing_collor_code"].ToString();
                            cConv1 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_purc"].ToString());
                            cConv2 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_sales"].ToString());
                        }
                        else
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            sSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = "";
                            cGroup = "";
                            cSubGroup = "";
                            cModel = "";
                            cRawMatUsed = "";
                            cWashingCollorCode = "";
                            cConv1 = 0;
                            cConv2 = 0;
                        }
                        //Get Price Code
                        DataTable dtFillGrpPrice = new DataTable();
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + sCustCode1 + "','" + sCustCode2 + "','" + __tglorder + "'";
                        dtFillGrpPrice = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFillGrpPrice.Rows.Count > 0)
                        {
                            if (String.IsNullOrEmpty(dtFillGrpPrice.Rows[0]["cp_price_group"].ToString()) == true)
                            {
                                MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                            else
                            {
                                sPriceCode = dtFillGrpPrice.Rows[0]["cp_price_group"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        //if (dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString() == TOP)
                        //{


                        DataTable dtFill15 = new DataTable();
                        strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                        strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                        strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                        strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                        strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                        strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                        strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                        strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                        strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                        strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_sled)";
                        strSQL = strSQL + " values ( '";
                        strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '"; //'wsd_entity_id
                        strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '"; //'wsd_so_type
                        strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                        int lineno = lineNo + 1;
                        strSQL = strSQL + lineno + "', '"; //'wsd_so_detail_line_no
                        strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','"; //wsd_so_branch_id NEW
                        strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                        strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                        strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                        strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                        strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                        strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                        strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                        strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                        strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', '";
                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";
                        decimal OrderAmt = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString());
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            strSQL = strSQL + Math.Round(OrderAmt, 0, MidpointRounding.AwayFromZero) + "', '"; //'wsd_real_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + OrderAmt + "', '"; //'wsd_real_order_amt
                        }

                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";//'wsd_qty_sales
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', "; //'wsd_total_so_amount
                        strSQL = strSQL + "null, "; //'wsd_shipped_xqty
                        strSQL = strSQL + "null, "; //'wsd_shipped_qty
                        strSQL = strSQL + "null, "; //'wsd_shipped_amount
                        strSQL = strSQL + "null, "; //'wsd_pod_xqty
                        strSQL = strSQL + "null, "; //'wsd_pod_qty
                        strSQL = strSQL + "null, "; //'wsd_pod_amount
                        strSQL = strSQL + "null, "; //'wsd_invoice_xqty
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty
                        strSQL = strSQL + "null, "; //'wsd_invoice_amount
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_%
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_amt
                        strSQL = strSQL + "null, '"; //'wsd_margin

                        if (fGetCost(dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        strSQL = strSQL + _curCost.ToString() + "', '"; //'wsd_unit_cost
                        fConvertQty(dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString(), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value));
                        //Rounding Mekanism
                        decimal ttlcst = _curCost * Convert.ToDecimal(_fConvertQty);
                        if (isRoundingMekanism)
                        {
                            ttlcst = Math.Round(_curCost * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                        }
                        strSQL = strSQL + ttlcst.ToString() + "', "; //'wsd_total_cost
                        strSQL = strSQL + "null, "; //'wsd_remarks
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_flag
                        strSQL = strSQL + "null, "; //'wsd_invoice_no
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty_sales
                        strSQL = strSQL + "null, "; //'wsd_invoice_date
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_by
                        strSQL = strSQL + "getdate() , '"; //'wsd_entry_date
                        strSQL = strSQL + clsLogin.USERID + "', '"; //wsd_user_id
                        strSQL = strSQL + sCurrCode + "', "; //'wsd_curr_code
                        strSQL = strSQL + dRateCurr + ", "; //'wsd_rate
                        Decimal orghet = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                        Decimal orgso = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orgso + ", "; //'wsd_org_total_so_amount
                        strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                        strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                        strSQL = strSQL + cConv2 + ", "; //'wsd_uom_convert_big
                        //fConvertQty(dgvDetail2.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                        strSQL = strSQL + uomorg + ", "; //'wsd_uom_org_qty
                        strSQL = strSQL + "0, '";

                        strSQL = strSQL + "0" + "', '"; //'wsd_tot_disc_1
                        strSQL = strSQL + "0" + "', "; //'wsd_tot_disc_2
                        strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString() + "', '"; //wsd_tax_code
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString() + "',"; //wsd_tax_pct
                        strSQL = strSQL + "0, "; //wsd_tax_amount
                        strSQL = strSQL + "0, "; //wsd_gross_net
                        strSQL = strSQL + "0, "; //wsd_cash_disc
                        strSQL = strSQL + "'N', '"; //wsd_line_type
                        strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "', '";
                        strSQL = strSQL + sPriceCode + "', "; //wsd_price_code
                        if (dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                        {
                            strSQL = strSQL + "NULL"; //wsd_sled
                        }
                        else
                        {
                            strSQL = strSQL + "'" + dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString() + "'"; //wsd_sled
                        }
                        strSQL = strSQL + ")"; //wsd_reason
                        dtFill15 = _clsGlobal.ExecDTTrans(strSQL);

                        lineNo++;
                        #endregion
                    }
                    //}

                }

                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fSaveSOeCom_KAM(string bExtSono, string sCredLimit, bool bCalcDisc, int iRow, string TOP, string ValueTOP)
        {
            bool keyFound = false;

            string sBranch = "";
            string sArea = "";
            string sWilayah = "";
            string sRayon = "";
            string sCustGroup = "";
            string sLineCode = "";
            string sYear = "";
            string sPeriode = "";
            //string sBranch = "";
            string sWeekNo = "";
            string sOrderType = "1101";

            string sSOStatus = "";
            long lROrdQty = 0;
            decimal curROrdAmt = 0;
            decimal curCost = 0;
            decimal curTotalCost = 0;
            //clsPCode
            string sCurrCode = "";
            double dRateCurr = 0;
            long lBrs = 0;
            decimal curDisc1 = 0;
            decimal curDisc2 = 0;
            string sErrDesc = "";
            decimal currCredLimit = 0;
            //bool bFlagCr = false;
            string sPriceCode = "";
            string sOpenStatus = "";
            long lOrdQty = 0;
            long lLeadTime = 0;
            decimal curOrdAmt = 0;
            string pdaNO = "";
            string pdaFlag = "";
            string extsap = "";
            string runcext = "";
            string QasirFlag = "";
            string nik = "";

            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cSled = "";
            string cLineCode = "";
            string cGroup = "";
            string cSubGroup = "";
            string cModel = "";
            string cRawMatUsed = "";
            string cWashingCollorCode = "";
            long cConv1 = 0;
            long cConv2 = 0;

            string sCustCode1;
            string sCustCode2;
            string sEntity;
            string sPONo;
            string sSalesman;
            //string sSoNo;
            Dictionary<string, string> SOPrincipal;
            string qasirflag = "";
            string sFullfilment = "N";


            DateTime sTglGudang;
            string wh1 = "";
            string wh2 = "";
            try
            {
                keyFound = false;

                sPONo = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString();
                sCustCode1 = dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString();
                sCustCode2 = dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value.ToString();
                sEntity = dgvHeader.Rows[iRow].Cells["g_iEntity"].Value.ToString();
                sBranch = dgvHeader.Rows[iRow].Cells["g_iBranch"].Value.ToString();
                sSalesman = dgvHeader.Rows[iRow].Cells["g_iSalesmanCode"].Value.ToString();

                if (!cbSource.SelectedValue.ToString().CompareC("M"))
                {
                    qasirflag = cbSource.SelectedValue.ToString().Trim();
                }
                else
                {
                    qasirflag = "N";
                }
                if (bCalcDisc == false)
                {
                    if (fGetSONumber() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    sSONo = g_sDefaultSONo;

                    strSQL = " EXEC SP_DELETE_SO_TABLE ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }


                strSQL = "";
                strSQL = "SELECT wh_loc_id1,wh_loc_id2,wh_loc_last_work_date FROM IM_WH_LOC WITH (NOLOCK) ";
                strSQL += "WHERE ";
                strSQL += "wh_loc_entity = '" + sEntity + "' ";
                strSQL += "AND wh_branch_id = '" + sBranch + "' ";
                strSQL += "AND wh_loc_id1 +'|'+wh_loc_id2 = (SELECT top 1 gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_function_name  LIKE 'MAINWHLOC' ) ";
                DataTable dtTglGudang = _clsGlobal.ExecDTTrans(strSQL);
                if (dtTglGudang.Rows.Count > 0)
                {
                    wh1 = dtTglGudang.Rows[0]["wh_loc_id1"].ToString();
                    wh2 = dtTglGudang.Rows[0]["wh_loc_id2"].ToString();
                    sTglGudang = Convert.ToDateTime(dtTglGudang.Rows[0]["wh_loc_last_work_date"].ToString());
                }
                else
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Tanggal Gudang tidak ditemukan!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }

                DataTable dtNIK = new DataTable();
                strSQL = "select sah_NIK from tbl_salesman_history WITH (NOLOCK)" +
                         " where" +
                         " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and sah_entity_id = '" + sEntity + "'" +
                         " and sah_branch_id = '" + sBranch + "'" +
                         " and sah_salesmen_id = '" + sSalesman + "'";
                dtNIK = _clsGlobal.ExecDTTrans(strSQL);

                if (dtNIK.Rows.Count == 1)
                {
                    nik = dtNIK.Rows[0]["sah_NIK"].ToString();
                }
                else if (dtNIK.Rows.Count > 1)
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }

                    MessageBox.Show("Salesman History Pada Periode Entry Lebih dari 1", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return keyFound;
                }
                else
                {

                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Salesman History Tidak Ditemukan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }



                //'Get branch, area, wilayah, rayon, cust group dan line code
                DataTable dtFill8 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') cm_line_code, cm_lead_time, ISNULL(cm_fullfilment,'N') cm_fullfilment From SO_CUST_MASTER WITH (NOLOCK)";
                strSQL = strSQL + " where cm_cust_code1 = '" + sCustCode1 + "' AND";
                strSQL = strSQL + " cm_cust_code2 = '" + sCustCode2 + "' and cm_entity = '" + sEntity + "' and cm_branch = '" + sBranch + "' ";
                dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill8.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill8.Rows[0]["cm_branch"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Area untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill8.Rows[0]["cm_area"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Wilayah untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill8.Rows[0]["cm_wilayah"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Rayon untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill8.Rows[0]["cm_rayon"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Cust group untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill8.Rows[0]["cm_cust_group"].ToString();
                    }

                    sLineCode = dtFill8.Rows[0]["cm_line_code"].ToString();
                    sFullfilment = dtFill8.Rows[0]["cm_fullfilment"].ToString();
                    //if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_lead_time"].ToString()))
                    //{
                    //    MessageBox.Show("Lead Time untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}
                    //else
                    //{
                    //    lLeadTime = Convert.ToInt64(dtFill8.Rows[0]["cm_lead_time"].ToString());
                    //}

                }
                else
                {
                    sBranch = "";
                    sArea = "";
                    sWilayah = "";
                    sRayon = "";
                    sCustGroup = "";
                    sLineCode = "";
                    sFullfilment = "N";
                }

                //'Get Year, Periode, Week untuk Tgl SO
                DataTable dtFill9 = new DataTable();
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK) ";
                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + dtSODate.Value.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '" + sEntity + "' and fwd_branch = '" + sBranch + "'"; //entity id harus di lempar
                dtFill9 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill9.Rows.Count > 0)
                {
                    sWeekNo = dtFill9.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill9.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill9.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + dtSODate.Value.ToString("dd MMM yyyy") + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get SO Status
                DataTable dtFill10 = new DataTable();
                strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                strSQL = strSQL + " and gh_sequence_no = '1' ";
                dtFill10 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill10.Rows.Count > 0)
                {
                    sSOStatus = dtFill10.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    MessageBox.Show("Status SO (Open) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvDetail.Rows.Count; lCounter++)
                {
                    //if (sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString() && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != "" && TOP == dgvDetail.Rows[lCounter].Cells["g_iDetTOPDiv"].Value.ToString().Trim())
                    if (
                        sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString()
                        && sEntity == dgvDetail.Rows[lCounter].Cells["g_iDetEntity"].Value.ToString()
                        && sBranch == dgvDetail.Rows[lCounter].Cells["g_iDetBranch"].Value.ToString()
                        && sSalesman == dgvDetail.Rows[lCounter].Cells["g_iDetSalesman"].Value.ToString()
                        && sCustCode1 == dgvDetail.Rows[lCounter].Cells["g_iDetCustNo"].Value.ToString()
                        && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != ""
                    )
                    {
                        //fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + (int)Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curROrdAmt = curROrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            curROrdAmt = curROrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);
                        }

                        fConvertQty(dgvDetail.Rows[lCounter].Cells["g_iDetQtyOrd"].Value.ToString(), Convert.ToInt16(iConv1), Convert.ToInt16(iConv2));
                        lOrdQty = lOrdQty + Convert.ToInt32(_fConvertQty);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curOrdAmt = curOrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);

                        }
                        else
                        {
                            curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);

                        }

                        if (fGetCostEcommerce(sEntity, sBranch, dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetGrade"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetSize"].Value.ToString(), wh1, wh2) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        curTotalCost = curTotalCost + _curCost;
                    }
                }
                
                //'Credit Limit
                if (fGetCreditLimit(sEntity, sBranch, sCustCode1, sCustCode2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (bFlagCr == true)
                {
                    if (_curAvCredLimit < curROrdAmt)
                    {
                        sCredLimit = "";
                    }
                    else
                    {
                        sCredLimit = "";
                    }
                }
                //}

                DataTable dtFill4 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_DETAIL A WITH (NOLOCK)";
                strSQL = strSQL + " WHERE wsd_so_seq_no = '" + sSONo + "'";
                //strSQL = strSQL + "  AND wsh_branch_id = '" + cbBranchID.SelectedValue.ToString()  + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill5 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_HEADER A WITH (NOLOCK)";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "'";
                strSQL = strSQL + "  AND wsh_branch_id = '" + sBranch + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill12 = new DataTable();
                #region Query Insert
                strSQL = "Insert into SO_WEB_SALES_HEADER (";
                strSQL = strSQL + " wsh_entity_id ,wsh_area ,wsh_wilayah ,wsh_rayon ";
                strSQL = strSQL + " ,wsh_branch_id ,wsh_cust_group ,wsh_so_type ,wsh_seq_no";
                strSQL = strSQL + " ,wsh_so_date,wsh_cust_code1 ,wsh_cust_code2 ,wsh_loc_id1";
                strSQL = strSQL + " ,wsh_loc_id2,wsh_brand_line ,wsh_so_ref_no,wsh_so_txn_date_from";
                strSQL = strSQL + " ,wsh_so_txn_date_to,wsh_week_no,wsh_spgm_id ,wsh_spgm_group";
                strSQL = strSQL + " ,wsh_spgm_coordinator_id ,wsh_spgm_loc1 ,wsh_spgm_loc2 ,wsh_shift_no";
                strSQL = strSQL + " ,[wsh_participation_disc_%] ,wsh_participation_disc_amount ,wsh_tot_part_line_amount ,wsh_total_line";
                strSQL = strSQL + " , wsh_relevan_pod ,wsh_status_so ,wsh_real_order_qty   ,wsh_real_order_amt";
                strSQL = strSQL + " ,wsh_total_so_qty ,wsh_total_so_amt      ,wsh_include_tax ,[wsh_tax_%]";
                strSQL = strSQL + " ,wsh_total_partisipasi_disc_amount ,wsh_total_margin_disc_amount ,wsh_net_amount,wsh_dpp_amount";
                strSQL = strSQL + " ,wsh_tax_amount,wsh_invoice_process_flag ,wsh_invoice_no ,wsh_invoice_date";
                strSQL = strSQL + " ,wsh_invoice_process_by ,wsh_disc_margin,wsh_total_cost,wsh_year";
                strSQL = strSQL + " ,wsh_period,wsh_cancel_date,wsh_cancel_by,wsh_approval";
                strSQL = strSQL + " ,wsh_approval_by ,wsh_approval_date,wsh_ret_seq_no ,wsh_ret_by ,wsh_remarks";
                strSQL = strSQL + " ,wsh_participation_code ,wsh_entry_date,wsh_user_id ,wsh_from_OB_flag";
                strSQL = strSQL + " ,wsh_gross_net,wsh_gen_flag,wsh_tot_disc_pc,wsh_tot_disc_1";
                strSQL = strSQL + " ,wsh_tot_disc_2,wsh_tot_tpr_amt,wsh_tax_code,wsh_tax_pct";
                strSQL = strSQL + " ,wsh_pct_cash_disc, wsh_cash_disc, wsh_tax_pbm_pct, wsh_tax_pbm_amount, wsh_pay_type";
                strSQL = strSQL + " ,wsh_gross_net_amt, wsh_validate, wsh_po_no ";
                strSQL = strSQL + " ,wsh_flag_so_limit ";
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty,wsh_po_ol_flag ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                strSQL = strSQL + " , wsh_extract_sap ";
                strSQL = strSQL + " , wsh_runcode_ext  ";
                strSQL = strSQL + " , wsh_sch_date  ";
                strSQL = strSQL + " , wsh_expired_date  ";
                strSQL = strSQL + " , wsh_nik  ";
                strSQL = strSQL + " , wsh_qasir_flag  ";
                strSQL = strSQL + " , wsh_beban_cashdisc  ";
                if (isMultiSource)
                {
                    strSQL = strSQL + " , wsh_flag_sloc ";
                }
                strSQL = strSQL + " , wsh_is_fulfillment ";
                strSQL = strSQL + " , wsh_sled_flag ";

                string __tglorder = sTglGudang.ToString("yyyyMMdd");
                //int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + sEntity + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + sBranch + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                strSQL = strSQL + sOrderType + "' , '";  //'wsh_so_type
                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no                                                   
                strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date              
                strSQL = strSQL + sCustCode1 + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + sCustCode2 + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + wh1 + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + wh2 + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + sSalesman + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "null , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "null , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "null , "; //'wsh_tot_part_line_amount
                strSQL = strSQL + "0, "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                strSQL = strSQL + "null , "; //'wsh_include_tax
                strSQL = strSQL + "null , "; //'[wsh_tax_%]
                strSQL = strSQL + "null , "; //'wsh_total_partisipasi_disc_amount
                strSQL = strSQL + "null , "; //'wsh_total_margin_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_net_amount
                strSQL = strSQL + "0 , "; //'wsh_dpp_amount
                strSQL = strSQL + "0 , "; //'wsh_tax_amount
                strSQL = strSQL + "'N' , "; //'wsh_invoice_process_flag
                strSQL = strSQL + "null , "; //'wsh_invoice_no
                strSQL = strSQL + "null , "; //'wsh_invoice_date
                strSQL = strSQL + "null , "; //'wsh_invoice_process_by
                strSQL = strSQL + "null , "; //'wsh_disc_margin
                strSQL = strSQL + curTotalCost + ", '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , "; //'wsh_approval
                strSQL = strSQL + "null , "; //'wsh_approval_by
                strSQL = strSQL + "null , "; //'wsh_approval_date
                strSQL = strSQL + "null , "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                strSQL = strSQL + "null , "; //'wsh_participation_code
                strSQL = strSQL + "getdate() , '"; //'wsh_entry_date
                strSQL = strSQL + clsLogin.USERID + "', "; //'wsh_user_id
                strSQL = strSQL + "null , "; //'wsh_from_OB_flag
                strSQL = strSQL + "null , "; //'wsh_gross_net
                strSQL = strSQL + "null , "; //'wsh_gen_flag
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_pc
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_1
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_2
                strSQL = strSQL + "0 , "; //'wsh_tot_tpr_amt
                strSQL = strSQL + "null , "; //'wsh_tax_code
                strSQL = strSQL + "null , "; //'wsh_tax_pct
                strSQL = strSQL + "0" + " , ";
                strSQL = strSQL + "0 , "; //'wsh_cash_disc             
                strSQL = strSQL + "0 , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "0 , '"; //'wsh_tax_pbm_amount
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value.ToString() + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , "; //'wsh_gross_net_amt
                strSQL = strSQL + "null , '"; // Convert.ToDateTime(dtTglValidasi.Value).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + sPONo + "' , '"; //'wsh_po_no
                strSQL = strSQL + sCredLimit + "' , '"; //'wsh_flag_so_limit
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                strSQL = strSQL + "N' , '"; //'wsh_po_ol_flag
                strSQL = strSQL + ValueTOP + "' , '"; //'wsh_top
                strSQL = strSQL + TOP + "' , '"; //'wsh_top_id
                string __poDate = dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString();
                strSQL = strSQL + __poDate.Substring(6, 4) + __poDate.Substring(3, 2) + __poDate.Substring(0, 2) + "','"; //'wsh_po_date

                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value.ToString() + "'"; //'wsh_cust_leadtime
                strSQL = strSQL + ",'" + "N" + "' , '"; //'extsap
                strSQL = strSQL + "'"; //'runcext
                string __delvdate = dgvHeader.Rows[iRow].Cells["g_iDeliveryDate"].Value.ToString();
                string __expdate = dgvHeader.Rows[iRow].Cells["g_iExpiredDate"].Value.ToString();
                strSQL = strSQL + ",'" + __delvdate.Substring(6, 4) + __delvdate.Substring(3, 2) + __delvdate.Substring(0, 2) + "'"; // planning date
                strSQL = strSQL + ",'" + __expdate.Substring(6, 4) + __expdate.Substring(3, 2) + __expdate.Substring(0, 2) + "'"; // expired date

                strSQL = strSQL + ",'" + nik + "'"; // nik                
                strSQL = strSQL + ",'" + qasirflag + "'"; // qasirflag 30-06-2021 -- update BL : kalo M kosong kalo yang lain diisi sama dengan cbSource.selectedvalue               
                strSQL = strSQL + ",0"; //wsh_beban_cashdisc               
                if (isMultiSource)
                {
                    strSQL = strSQL + ",'" + cbFlagSloc.SelectedValue.ToString() + "'";
                }
                strSQL = strSQL + ",'" + sFullfilment.Trim() + "'"; //wsh_is_fulfillment
                if (dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                {
                    strSQL = strSQL + ",'N'"; //wsh_sled_flag
                }
                else
                {
                    strSQL = strSQL + ",'"+ dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().Trim() + "'"; //wsh_sled_flag
                }
                strSQL = strSQL + " ) ";
                dtFill12 = _clsGlobal.ExecDTTrans(strSQL);
                //'END OF Saving SO Header
               
                DataTable dtFill13 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK) ";
                strSQL = strSQL + " where cc_currency_type = 'H' ";
                #endregion

                dtFill13 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill13.Rows.Count > 0)
                {
                    sCurrCode = dtFill13.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill13.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }


                int lineNo = 0;
                lCounter = 0;
                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                {

                    if (
                        sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString()
                        && sEntity == dgvDetail2.Rows[lCounter].Cells["g_iDetEntity1"].Value.ToString()
                        && sBranch == dgvDetail2.Rows[lCounter].Cells["g_iDetBranch1"].Value.ToString()
                        && sSalesman == dgvDetail2.Rows[lCounter].Cells["g_iDetSalesmanID1"].Value.ToString()
                        && sCustCode1 == dgvDetail2.Rows[lCounter].Cells["g_iDetCustNo1"].Value.ToString()
                        )
                    {
                        #region save detail
                        DataTable dtFill14 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK)";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "'";
                        dtFill14 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill14.Rows.Count > 0)
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = dtFill14.Rows[0]["prm_prd_line_code"].ToString();
                            cGroup = dtFill14.Rows[0]["prm_prd_group_code"].ToString();
                            cSubGroup = dtFill14.Rows[0]["prm_prd_sgroup_code"].ToString();
                            cModel = dtFill14.Rows[0]["prm_prd_model_code"].ToString();
                            cRawMatUsed = dtFill14.Rows[0]["prm_raw_mat_used_code"].ToString();
                            cWashingCollorCode = dtFill14.Rows[0]["prm_washing_collor_code"].ToString();
                            cConv1 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_purc"].ToString());
                            cConv2 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_sales"].ToString());
                        }
                        else
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = "";
                            cGroup = "";
                            cSubGroup = "";
                            cModel = "";
                            cRawMatUsed = "";
                            cWashingCollorCode = "";
                            cConv1 = 0;
                            cConv2 = 0;
                        }
                        //Get Price Code
                        DataTable dtFillGrpPrice = new DataTable();
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + sEntity + "','" + sBranch + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + sCustCode1 + "','" + sCustCode2 + "','" + __tglorder + "'";
                        dtFillGrpPrice = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFillGrpPrice.Rows.Count > 0)
                        {
                            if (String.IsNullOrEmpty(dtFillGrpPrice.Rows[0]["cp_price_group"].ToString()) == true)
                            {
                                MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                            else
                            {
                                sPriceCode = dtFillGrpPrice.Rows[0]["cp_price_group"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        //if (dgvDetail2.Rows[lCounter].Cells["g_iDetTOPDiv1"].Value.ToString() == TOP)
                        //{


                        DataTable dtFill15 = new DataTable();
                        strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                        strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                        strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                        strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                        strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                        strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                        strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                        strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                        strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                        strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_sled)";

                        strSQL = strSQL + " values ( '";
                        strSQL = strSQL + sEntity + "' , '"; //'wsd_entity_id
                        strSQL = strSQL + sOrderType + "' , '"; //'wsd_so_type
                        strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                        int lineno = lineNo + 1;
                        strSQL = strSQL + lineno + "', '"; //'wsd_so_detail_line_no
                        strSQL = strSQL + sBranch + "','"; //wsd_so_branch_id NEW
                        strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                        strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                        strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                        strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                        strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                        strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                        strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                        strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                        strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', '";
                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";
                        decimal OrderAmt = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString());
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            strSQL = strSQL + Math.Round(OrderAmt, 0, MidpointRounding.AwayFromZero) + "', '"; //'wsd_real_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + OrderAmt + "', '"; //'wsd_real_order_amt
                        }

                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";//'wsd_qty_sales
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', "; //'wsd_total_so_amount
                        strSQL = strSQL + "null, "; //'wsd_shipped_xqty
                        strSQL = strSQL + "null, "; //'wsd_shipped_qty
                        strSQL = strSQL + "null, "; //'wsd_shipped_amount
                        strSQL = strSQL + "null, "; //'wsd_pod_xqty
                        strSQL = strSQL + "null, "; //'wsd_pod_qty
                        strSQL = strSQL + "null, "; //'wsd_pod_amount
                        strSQL = strSQL + "null, "; //'wsd_invoice_xqty
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty
                        strSQL = strSQL + "null, "; //'wsd_invoice_amount
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_%
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_amt
                        strSQL = strSQL + "null, '"; //'wsd_margin

                        if (fGetCostEcommerce(sEntity, sBranch, dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString(), wh1, wh2) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        strSQL = strSQL + _curCost.ToString() + "', '"; //'wsd_unit_cost
                        fConvertQty(dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString(), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value));
                        //Rounding Mekanism
                        decimal ttlcst = _curCost * Convert.ToDecimal(_fConvertQty);
                        if (isRoundingMekanism)
                        {
                            ttlcst = Math.Round(_curCost * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                        }
                        strSQL = strSQL + ttlcst.ToString() + "', "; //'wsd_total_cost
                        strSQL = strSQL + "null, "; //'wsd_remarks
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_flag
                        strSQL = strSQL + "null, "; //'wsd_invoice_no
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty_sales
                        strSQL = strSQL + "null, "; //'wsd_invoice_date
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_by
                        strSQL = strSQL + "getdate() , '"; //'wsd_entry_date
                        strSQL = strSQL + clsLogin.USERID + "', '"; //wsd_user_id
                        strSQL = strSQL + sCurrCode + "', "; //'wsd_curr_code
                        strSQL = strSQL + dRateCurr + ", "; //'wsd_rate
                        Decimal orghet = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                        Decimal orgso = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orgso + ", "; //'wsd_org_total_so_amount
                        strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                        strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                        strSQL = strSQL + cConv2 + ", "; //'wsd_uom_convert_big
                        //fConvertQty(dgvDetail2.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                        strSQL = strSQL + uomorg + ", "; //'wsd_uom_org_qty
                        strSQL = strSQL + "0, '";

                        strSQL = strSQL + "0" + "', '"; //'wsd_tot_disc_1
                        strSQL = strSQL + "0" + "', "; //'wsd_tot_disc_2
                        strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString() + "', '"; //wsd_tax_code
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString() + "',"; //wsd_tax_pct
                        strSQL = strSQL + "0, "; //wsd_tax_amount
                        strSQL = strSQL + "0, "; //wsd_gross_net
                        strSQL = strSQL + "0, "; //wsd_cash_disc
                        strSQL = strSQL + "'N', '"; //wsd_line_type
                        strSQL = strSQL + sTglGudang.ToString("yyyy-MM-dd") + "', '";
                        strSQL = strSQL + sPriceCode + "' "; //wsd_price_code
                        if (dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                        {
                            strSQL = strSQL + ",NULL"; //wsd_sled
                        }
                        else
                        {
                            strSQL = strSQL + ",'"+ dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString() +"'"; //wsd_sled
                        }
                       
                        strSQL = strSQL + ")"; //wsd_reason
                        dtFill15 = _clsGlobal.ExecDTTrans(strSQL);

                        lineNo++;
                        #endregion
                    }
                    //}

                }

                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }


        private void calculateSOValue(int row)
        {
            _clsGlobal.BeginTrans();

            bool saveSO = fSaveSOCalc("", "", true, row);
            if (saveSO)
            {
                _clsGlobal.ExecuteTrans("exec SP_PROSES_SO_ROVI " + "'1','" + "S" + cbEntityBranch.txtBranchIdCM.Text.ToString() + "XXXXXXX" + "','" + clsLogin.USERID + "','" + cbEntityBranch.txtBranchIdCM.Text.ToString() + "'");
                DataTable dtH = _clsGlobal.ExecDTTrans("Select * from SO_WEB_SALES_HEADER WITH (NOLOCK) where wsh_seq_no = '" + "S" + cbEntityBranch.txtBranchIdCM.Text.ToString() + "XXXXXXX" + "'");
                DataTable dtD = _clsGlobal.ExecDTTrans("Select * from SO_WEB_SALES_DETAIL WITH (NOLOCK) where wsd_so_seq_no = '" + "S" + cbEntityBranch.txtBranchIdCM.Text.ToString() + "XXXXXXX" + "'");
                if (dtH.Rows.Count > 0)
                {

                    dgvHeader.Rows[row].Cells["g_iSOValue"].Value = dtH.Rows[0]["wsh_net_amount"].ToString();
                    decimal diff = Convert.ToDecimal(dgvHeader.Rows[row].Cells["g_iSOValue"].Value) - Convert.ToDecimal(dgvHeader.Rows[row].Cells["g_iTXTValue"].Value);
                    if (diff < 0)
                    {
                        diff = diff * -1;
                    }
                    if (diff > 100)
                    {
                        dgvHeader.Rows[row].Cells["g_iRemark"].Value = "ada selisih 100 rupiah keatas";
                    }
                }
            }

            _clsGlobal.RollbackTrans();
        }
        void _bindCombo()
        {
            try
            {
                strSQL = "";
                strSQL = "select '' rm_reason_code,'' rm_reason_desc Union All ";
                strSQL = "select rm_reason_code, rm_reason_desc from TBL_GS_REASON_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + " where rm_sub_modul_id = 'SALES_CANCEL'";
                strSQL = strSQL + " and rm_auto_flag = 'N'";
                strSQL = strSQL + " order by rm_reason_code";

                DataTable dtCmb = _clsGlobal.ExecDT(strSQL);
                DataGridViewComboBoxColumn c = (DataGridViewComboBoxColumn)dgvHeader.Columns["g_iReasonRejection"];
                c.DataSource = dtCmb;
                c.ValueMember = "rm_reason_code";
                c.DisplayMember = "rm_reason_desc";
            }
            catch (Exception) { }
        }
        private bool fSaveSOCalc(string bExtSono, string sCredLimit, bool bCalcDisc, int iRow)
        {
            bool keyFound = false;

            string sBranch = "";
            string sArea = "";
            string sWilayah = "";
            string sRayon = "";
            string sCustGroup = "";
            string sLineCode = "";
            string sYear = "";
            string sPeriode = "";
            //string sBranch = "";
            string sWeekNo = "";

            string sSOStatus = "";
            long lROrdQty = 0;
            decimal curROrdAmt = 0;
            decimal curCost = 0;
            decimal curTotalCost = 0;
            //clsPCode
            string sCurrCode = "";
            double dRateCurr = 0;
            long lBrs = 0;
            decimal curDisc1 = 0;
            decimal curDisc2 = 0;
            string sErrDesc = "";
            decimal currCredLimit = 0;
            //bool bFlagCr = false;
            string sPriceCode = "";
            string sOpenStatus = "";
            long lOrdQty = 0;
            long lLeadTime = 0;
            decimal curOrdAmt = 0;
            string pdaNO = "";
            string pdaFlag = "";
            string extsap = "";
            string runcext = "";
            string QasirFlag = "";
            string nik = "";

            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cSled = "";
            string cLineCode = "";
            string cGroup = "";
            string cSubGroup = "";
            string cModel = "";
            string cRawMatUsed = "";
            string cWashingCollorCode = "";
            long cConv1 = 0;
            long cConv2 = 0;

            string sCustCode1;
            string sCustCode2;
            string sPONo;
            //string sSoNo;
            Dictionary<string, string> SOPrincipal;
            string qasirflag = "";


            try
            {
                keyFound = false;

                sPONo = dgvHeader.Rows[iRow].Cells["g_iCustPONo"].Value.ToString();
                sCustCode1 = dgvHeader.Rows[iRow].Cells["g_iCustCode"].Value.ToString();
                sCustCode2 = dgvHeader.Rows[iRow].Cells["g_iCustCode2"].Value.ToString();

                if (!cbSource.SelectedValue.ToString().CompareC("M"))
                {
                    qasirflag = cbSource.SelectedValue.ToString().Trim();
                }
                else
                {
                    qasirflag = "N";
                }

                if (bCalcDisc == false)
                {
                    if (fGetSONumber() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
                else
                {
                    sSONo = "S" + cbEntityBranch.txtBranchIdCM.Text.ToString() + "XXXXXXX";

                    //strSQL = " EXEC SP_DELETE_SO_TABLE '"+ sSONo +"' ";
                    //_clsGlobal.ExecuteTrans(strSQL);
                }


                DataTable dtNIK = new DataTable();
                strSQL = "select sah_NIK from tbl_salesman_history WITH (NOLOCK) " +
                         " where" +
                         " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and sah_entity_id = '" + cbEntityBranch.txtCM.Text.Trim() + "'" +
                         " and sah_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'" +
                         " and sah_salesmen_id = '" + txtSalesman.Text + "'";
                dtNIK = _clsGlobal.ExecDTTrans(strSQL);

                if (dtNIK.Rows.Count == 1)
                {
                    nik = dtNIK.Rows[0]["sah_NIK"].ToString();
                }
                else if (dtNIK.Rows.Count > 1)
                {
                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }

                    MessageBox.Show("Salesman History Pada Periode Entry Lebih dari 1", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return keyFound;
                }
                else
                {

                    keyFound = false;
                    if (_clsGlobal.tr != null)
                    {
                        _clsGlobal.RollbackTrans();
                    }
                    MessageBox.Show("Salesman History Tidak Ditemukan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return keyFound;
                }



                //'Get branch, area, wilayah, rayon, cust group dan line code
                DataTable dtFill8 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') cm_line_code, cm_lead_time From SO_CUST_MASTER WITH (NOLOCK)";
                strSQL = strSQL + " where cm_cust_code1 = '" + sCustCode1 + "' AND";
                strSQL = strSQL + " cm_cust_code2 = '" + sCustCode2 + "' ";
                dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill8.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill8.Rows[0]["cm_branch"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Area untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill8.Rows[0]["cm_area"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Wilayah untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill8.Rows[0]["cm_wilayah"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Rayon untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill8.Rows[0]["cm_rayon"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Cust group untuk Customer (" + sCustCode1 + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill8.Rows[0]["cm_cust_group"].ToString();
                    }

                    sLineCode = dtFill8.Rows[0]["cm_line_code"].ToString();

                    //if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_lead_time"].ToString()))
                    //{
                    //    MessageBox.Show("Lead Time untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    keyFound = false;
                    //    return keyFound;
                    //}
                    //else
                    //{
                    //    lLeadTime = Convert.ToInt64(dtFill8.Rows[0]["cm_lead_time"].ToString());
                    //}

                }
                else
                {
                    sBranch = "";
                    sArea = "";
                    sWilayah = "";
                    sRayon = "";
                    sCustGroup = "";
                    sLineCode = "";
                }

                //'Get Year, Periode, Week untuk Tgl SO
                DataTable dtFill9 = new DataTable();
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK) ";
                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + dtSODate.Value.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '01' "; //entity id harus di lempar
                dtFill9 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill9.Rows.Count > 0)
                {
                    sWeekNo = dtFill9.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill9.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill9.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + dtSODate.Value.ToString("dd MMM yyyy") + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get SO Status
                DataTable dtFill10 = new DataTable();
                strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                strSQL = strSQL + " and gh_sequence_no = '1' ";
                dtFill10 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill10.Rows.Count > 0)
                {
                    sSOStatus = dtFill10.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    MessageBox.Show("Status SO (Open) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


                //Get Price Code
                //DataTable dtFill11 = new DataTable();
                //strSQL = " exec SP_GET_PRICE_CODE '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + sCustCode1 + "','" + sCustCode2 + "'";
                //dtFill11 = _clsGlobal.ExecDTTrans(strSQL);
                //if (dtFill11.Rows.Count > 0)
                //{
                //    if (String.IsNullOrEmpty(dtFill11.Rows[0]["cm_def_price_code"].ToString()))
                //    {
                //        MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        keyFound = false;
                //        return keyFound;
                //    }
                //    else
                //    {
                //        sPriceCode = dtFill11.Rows[0]["cm_def_price_code"].ToString();
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvDetail.Rows.Count; lCounter++)
                {
                    if (sPONo == dgvDetail.Rows[lCounter].Cells["g_iDetPONO"].Value.ToString() && dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString() != "")
                    {
                        //fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + (int)Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curROrdAmt = curROrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            curROrdAmt = curROrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value == null ? 0 : dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);
                        }

                        fConvertQty(dgvDetail.Rows[lCounter].Cells["g_iDetQtyOrd"].Value.ToString(), Convert.ToInt16(iConv1), Convert.ToInt16(iConv2));
                        lOrdQty = lOrdQty + Convert.ToInt32(_fConvertQty);
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            curOrdAmt = curOrdAmt + Math.Round(Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value), 0, MidpointRounding.AwayFromZero);

                        }
                        else
                        {
                            curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetQty"].Value) * Convert.ToDecimal(dgvDetail.Rows[lCounter].Cells["g_iDetPrice"].Value);

                        }

                        if (fGetCost(dgvDetail.Rows[lCounter].Cells["g_iDetPCode"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetGrade"].Value.ToString(), dgvDetail.Rows[lCounter].Cells["g_iDetSize"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        curTotalCost = curTotalCost + _curCost;
                    }
                }

                //if (sSONo.Length == 0)
                //{
                //    MessageBox.Show("nomor SO Blank", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //if (_xFlagDC == "N")
                //{
                // '===> Cek Kredit Limit dihilangkan jika flagDC = Y - Project Improvement TiraSnD
                //'/*-----------------------------------------------------------------------------------------
                //'Credit Limit
                if (fGetCreditLimit(sCustCode1, sCustCode2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (bFlagCr == true)
                {
                    if (_curAvCredLimit < curROrdAmt)
                    {
                        sCredLimit = "";
                    }
                    else
                    {
                        sCredLimit = "";
                    }
                }
                //}

                DataTable dtFill4 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_DETAIL A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsd_so_seq_no = '" + sSONo + "'";
                //strSQL = strSQL + "  AND wsh_branch_id = '" + cbBranchID.SelectedValue.ToString()  + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill5 = new DataTable();
                strSQL = "delete A FROM SO_WEB_SALES_HEADER A WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "'";
                strSQL = strSQL + "  AND wsh_branch_id = '" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                DataTable dtFill12 = new DataTable();
                #region Query Insert
                strSQL = "Insert into SO_WEB_SALES_HEADER (";
                strSQL = strSQL + " wsh_entity_id ,wsh_area ,wsh_wilayah ,wsh_rayon ";
                strSQL = strSQL + " ,wsh_branch_id ,wsh_cust_group ,wsh_so_type ,wsh_seq_no";
                strSQL = strSQL + " ,wsh_so_date,wsh_cust_code1 ,wsh_cust_code2 ,wsh_loc_id1";
                strSQL = strSQL + " ,wsh_loc_id2,wsh_brand_line ,wsh_so_ref_no,wsh_so_txn_date_from";
                strSQL = strSQL + " ,wsh_so_txn_date_to,wsh_week_no,wsh_spgm_id ,wsh_spgm_group";
                strSQL = strSQL + " ,wsh_spgm_coordinator_id ,wsh_spgm_loc1 ,wsh_spgm_loc2 ,wsh_shift_no";
                strSQL = strSQL + " ,[wsh_participation_disc_%] ,wsh_participation_disc_amount ,wsh_tot_part_line_amount ,wsh_total_line";
                strSQL = strSQL + " , wsh_relevan_pod ,wsh_status_so ,wsh_real_order_qty   ,wsh_real_order_amt";
                strSQL = strSQL + " ,wsh_total_so_qty ,wsh_total_so_amt      ,wsh_include_tax ,[wsh_tax_%]";
                strSQL = strSQL + " ,wsh_total_partisipasi_disc_amount ,wsh_total_margin_disc_amount ,wsh_net_amount,wsh_dpp_amount";
                strSQL = strSQL + " ,wsh_tax_amount,wsh_invoice_process_flag ,wsh_invoice_no ,wsh_invoice_date";
                strSQL = strSQL + " ,wsh_invoice_process_by ,wsh_disc_margin,wsh_total_cost,wsh_year";
                strSQL = strSQL + " ,wsh_period,wsh_cancel_date,wsh_cancel_by,wsh_approval";
                strSQL = strSQL + " ,wsh_approval_by ,wsh_approval_date,wsh_ret_seq_no ,wsh_ret_by ,wsh_remarks";
                strSQL = strSQL + " ,wsh_participation_code ,wsh_entry_date,wsh_user_id ,wsh_from_OB_flag";
                strSQL = strSQL + " ,wsh_gross_net,wsh_gen_flag,wsh_tot_disc_pc,wsh_tot_disc_1";
                strSQL = strSQL + " ,wsh_tot_disc_2,wsh_tot_tpr_amt,wsh_tax_code,wsh_tax_pct";
                strSQL = strSQL + " ,wsh_pct_cash_disc, wsh_cash_disc, wsh_tax_pbm_pct, wsh_tax_pbm_amount, wsh_pay_type";
                strSQL = strSQL + " ,wsh_gross_net_amt, wsh_validate, wsh_po_no ";
                strSQL = strSQL + " ,wsh_flag_so_limit ";
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty,wsh_po_ol_flag ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                strSQL = strSQL + " , wsh_extract_sap ";
                strSQL = strSQL + " , wsh_runcode_ext  ";
                strSQL = strSQL + " , wsh_sch_date  ";
                strSQL = strSQL + " , wsh_expired_date  ";
                strSQL = strSQL + " , wsh_nik  ";
                strSQL = strSQL + " , wsh_qasir_flag  ";
                strSQL = strSQL + " , wsh_sled_flag  ";

                string __tglorder = dtSODate.Value.ToString("yyyyMMdd");
                //int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '";  //'wsh_so_type
                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_so_date
                }
                else
                {
                    strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                }
                strSQL = strSQL + sCustCode1 + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + sCustCode2 + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + lblWHLoc1.Text + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + lblWHLoc2.Text + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + txtSalesman.Text + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "null , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "null , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "null , "; //'wsh_tot_part_line_amount
                strSQL = strSQL + "0, "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                strSQL = strSQL + "null , "; //'wsh_include_tax
                strSQL = strSQL + "null , "; //'[wsh_tax_%]
                strSQL = strSQL + "null , "; //'wsh_total_partisipasi_disc_amount
                strSQL = strSQL + "null , "; //'wsh_total_margin_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_net_amount
                strSQL = strSQL + "0 , "; //'wsh_dpp_amount
                strSQL = strSQL + "0 , "; //'wsh_tax_amount
                strSQL = strSQL + "'N' , "; //'wsh_invoice_process_flag
                strSQL = strSQL + "null , "; //'wsh_invoice_no
                strSQL = strSQL + "null , "; //'wsh_invoice_date
                strSQL = strSQL + "null , "; //'wsh_invoice_process_by
                strSQL = strSQL + "null , "; //'wsh_disc_margin
                strSQL = strSQL + curTotalCost + ", '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , "; //'wsh_approval
                strSQL = strSQL + "null , "; //'wsh_approval_by
                strSQL = strSQL + "null , "; //'wsh_approval_date
                strSQL = strSQL + "null , "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                if (string.IsNullOrEmpty(dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value.ToString()))
                {
                    strSQL = strSQL + "null , "; //'wsh_participation_code
                }
                else
                {
                    strSQL = strSQL + "'" + dgvHeader.Rows[iRow].Cells["g_iVoucherNo"].Value.ToString() + "' , "; //'wsh_participation_code
                }

                strSQL = strSQL + "getdate() , '"; //'wsh_entry_date
                strSQL = strSQL + clsLogin.USERID + "', "; //'wsh_user_id
                strSQL = strSQL + "null , "; //'wsh_from_OB_flag
                strSQL = strSQL + "null , "; //'wsh_gross_net
                strSQL = strSQL + "null , "; //'wsh_gen_flag
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_pc
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_1
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_2
                strSQL = strSQL + "0 , "; //'wsh_tot_tpr_amt
                strSQL = strSQL + "null , "; //'wsh_tax_code
                strSQL = strSQL + "null , "; //'wsh_tax_pct

                if (dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() == "")
                {
                    strSQL = strSQL + "0" + " , ";
                }
                else
                {

                    decimal cashValue = 0;
                    decimal cashGross = 0;
                    decimal pct = 0;
                    if (cbSource.SelectedValue.ToString() == "L" || cbSource.SelectedValue.ToString() == "S" || cbSource.SelectedValue.ToString() == "W")
                    {
                        cashValue = Convert.ToDecimal(dgvHeader.Rows[iRow].Cells["g_iCashDisc"].Value.ToString());
                        cashGross = Convert.ToDecimal(curROrdAmt);
                        pct = (cashValue / cashGross) * 100;
                    }
                    if (Convert.ToDecimal(dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString()) > 0)
                    {
                        strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                    }
                    else if (pct > 0)
                    {
                        strSQL = strSQL + pct + " , "; //'wsh_pct_cash_disc
                    }
                    else
                    {
                        strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCashHdr"].Value.ToString() + " , "; //'wsh_pct_cash_disc
                    }
                }

                strSQL = strSQL + "0 , "; //'wsh_cash_disc
                //strSQL = strSQL + "0 , "; //'wsh_cash_disc
                strSQL = strSQL + "0 , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "0 , '"; //'wsh_tax_pbm_amount
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iPaymentType"].Value.ToString() + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , "; //'wsh_gross_net_amt
                strSQL = strSQL + "null , '"; // Convert.ToDateTime(dtTglValidasi.Value).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + sPONo + "' , '"; //'wsh_po_no
                strSQL = strSQL + sCredLimit + "' , '"; //'wsh_flag_so_limit
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                strSQL = strSQL + "N' , '"; //'wsh_po_ol_flag
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iTOP"].Value.ToString() + "' , '"; //'wsh_top
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iTOPId"].Value.ToString() + "' , '"; //'wsh_top_id
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iCustPODate"].Value.ToString() + "','"; //'wsh_po_date
                }
                else
                {
                    strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "','"; //'wsh_po_date
                }
                strSQL = strSQL + dgvHeader.Rows[iRow].Cells["g_iLeadtime"].Value.ToString() + "'"; //'wsh_cust_leadtime
                strSQL = strSQL + ",'" + "N" + "' , '"; //'extsap
                strSQL = strSQL + "'"; //'runcext
                if (isUploadMigrasiToNEFOKAM)
                {
                    strSQL = strSQL + ",'" + dgvHeader.Rows[iRow].Cells["g_iDeliveryDate"].Value.ToString() + "'"; // planning date
                    strSQL = strSQL + ",'" + dgvHeader.Rows[iRow].Cells["g_iExpiredDate"].Value.ToString() + "'"; // expired date
                }
                else
                {
                    strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // planning date
                    strSQL = strSQL + ",'" + initDateDelvExp(cbEntityBranch.txtCM.Text.Trim(), cbEntityBranch.txtBranchIdCM.Text.Trim(), __tglorder, sCustCode1, sCustCode2, 2) + "'"; // expired date
                }
                strSQL = strSQL + ",'" + nik + "'"; // nik
                strSQL = strSQL + ",'" + qasirflag + "'"; // qasirflag 30-06-2021 -- update BL : kalo M kosong kalo yang lain diisi sama dengan cbSource.selectedvalue
                if (dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                {
                    strSQL = strSQL + ",'N'";
                }
                else
                {
                    strSQL = strSQL + ",'"+ dgvHeader.Rows[iRow].Cells["g_iSledFlag"].Value.ToString().Trim() + "'";
                }

                //Convert.ToDateTime(dtTglOrder.Text).ToString("yyyyMMdd")
                strSQL = strSQL + " ) ";
                dtFill12 = _clsGlobal.ExecDTTrans(strSQL);
                //'END OF Saving SO Header
                //if (dtFill12.Rows.Count > 0)
                //{
                DataTable dtFill13 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK) ";
                strSQL = strSQL + " where cc_currency_type = 'H' ";
                #endregion

                dtFill13 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill13.Rows.Count > 0)
                {
                    sCurrCode = dtFill13.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill13.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }

                lCounter = 0;
                for (lCounter = 0; lCounter < dgvDetail2.Rows.Count; lCounter++)
                {

                    if (sPONo == dgvDetail2.Rows[lCounter].Cells["g_iDetPONO1"].Value.ToString())
                    {
                        #region save detail
                        DataTable dtFill14 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK)";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString() + "'";
                        dtFill14 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill14.Rows.Count > 0)
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = dtFill14.Rows[0]["prm_prd_line_code"].ToString();
                            cGroup = dtFill14.Rows[0]["prm_prd_group_code"].ToString();
                            cSubGroup = dtFill14.Rows[0]["prm_prd_sgroup_code"].ToString();
                            cModel = dtFill14.Rows[0]["prm_prd_model_code"].ToString();
                            cRawMatUsed = dtFill14.Rows[0]["prm_raw_mat_used_code"].ToString();
                            cWashingCollorCode = dtFill14.Rows[0]["prm_washing_collor_code"].ToString();
                            cConv1 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_purc"].ToString());
                            cConv2 = Convert.ToInt16(dtFill14.Rows[0]["prm_conversion_sales"].ToString());
                        }
                        else
                        {
                            cPcode = dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString();
                            cGrade = dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString();
                            cSize = dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString();
                            cSled = dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString();
                            cLineCode = "";
                            cGroup = "";
                            cSubGroup = "";
                            cModel = "";
                            cRawMatUsed = "";
                            cWashingCollorCode = "";
                            cConv1 = 0;
                            cConv2 = 0;
                        }
                        //Get Price Code
                        DataTable dtFillGrpPrice = new DataTable();
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + cbEntityBranch.txtCM.Text.Trim() + "','" + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + sCustCode1 + "','" + sCustCode2 + "','" + __tglorder + "'";
                        dtFillGrpPrice = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFillGrpPrice.Rows.Count > 0)
                        {
                            if (String.IsNullOrEmpty(dtFillGrpPrice.Rows[0]["cp_price_group"].ToString()) == true)
                            {
                                MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                            else
                            {
                                sPriceCode = dtFillGrpPrice.Rows[0]["cp_price_group"].ToString();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Default Price Code untuk Customer " + sCustCode1 + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }


                        DataTable dtFill15 = new DataTable();
                        strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                        strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                        strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                        strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                        strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                        strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                        strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                        strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                        strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                        strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_sled)";

                        strSQL = strSQL + " values ( '";
                        strSQL = strSQL + cbEntityBranch.txtCM.Text.Trim() + "' , '"; //'wsd_entity_id
                        strSQL = strSQL + cbTipeOrder.SelectedValue.ToString() + "' , '"; //'wsd_so_type
                        strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                        int lineno = lCounter + 1;
                        strSQL = strSQL + lineno + "', '"; //'wsd_so_detail_line_no
                        strSQL = strSQL + cbEntityBranch.txtBranchIdCM.Text.Trim() + "','"; //wsd_so_branch_id NEW
                        strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                        strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                        strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                        strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                        strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                        strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                        strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                        strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                        strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', '";
                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";
                        decimal OrderAmt = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value == null ? "0" : dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString());
                        //Rounding Mekanism
                        if (isRoundingMekanism)
                        {
                            strSQL = strSQL + Math.Round(OrderAmt, 0, MidpointRounding.AwayFromZero) + "', '"; //'wsd_real_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + OrderAmt + "', '"; //'wsd_real_order_amt
                        }

                        strSQL = strSQL + fPC2Qty(Convert.ToDouble(dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value)) + "', '";
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetQty1"].Value.ToString() + "', '";//'wsd_qty_sales
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString() + "', "; //'wsd_total_so_amount
                        strSQL = strSQL + "null, "; //'wsd_shipped_xqty
                        strSQL = strSQL + "null, "; //'wsd_shipped_qty
                        strSQL = strSQL + "null, "; //'wsd_shipped_amount
                        strSQL = strSQL + "null, "; //'wsd_pod_xqty
                        strSQL = strSQL + "null, "; //'wsd_pod_qty
                        strSQL = strSQL + "null, "; //'wsd_pod_amount
                        strSQL = strSQL + "null, "; //'wsd_invoice_xqty
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty
                        strSQL = strSQL + "null, "; //'wsd_invoice_amount
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_%
                        strSQL = strSQL + "null, "; //'wsd_participation_disc_amt
                        strSQL = strSQL + "null, '"; //'wsd_margin

                        if (fGetCost(dgvDetail2.Rows[lCounter].Cells["g_iDetPCode1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetGrade1"].Value.ToString(), dgvDetail2.Rows[lCounter].Cells["g_iDetSize1"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        strSQL = strSQL + _curCost.ToString() + "', '"; //'wsd_unit_cost
                        fConvertQty(dgvDetail2.Rows[lCounter].Cells["g_iDetQtyOrd1"].Value.ToString(), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv11"].Value), Convert.ToInt16(dgvDetail2.Rows[lCounter].Cells["g_iDetConv21"].Value));
                        //Rounding Mekanism
                        decimal ttlcst = _curCost * Convert.ToDecimal(_fConvertQty);
                        if (isRoundingMekanism)
                        {
                            ttlcst = Math.Round(_curCost * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                        }
                        strSQL = strSQL + ttlcst.ToString() + "', "; //'wsd_total_cost
                        strSQL = strSQL + "null, "; //'wsd_remarks
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_flag
                        strSQL = strSQL + "null, "; //'wsd_invoice_no
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty_sales
                        strSQL = strSQL + "null, "; //'wsd_invoice_date
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_by
                        strSQL = strSQL + "getdate() , '"; //'wsd_entry_date
                        strSQL = strSQL + clsLogin.USERID + "', '"; //wsd_user_id
                        strSQL = strSQL + sCurrCode + "', "; //'wsd_curr_code
                        strSQL = strSQL + dRateCurr + ", "; //'wsd_rate
                        Decimal orghet = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                        Decimal orgso = Convert.ToDecimal(dgvDetail2.Rows[lCounter].Cells["g_iDetPrice1"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orgso + ", "; //'wsd_org_total_so_amount
                        strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                        strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                        strSQL = strSQL + cConv2 + ", "; //'wsd_uom_convert_big
                        //fConvertQty(dgvDetail2.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                        strSQL = strSQL + uomorg + ", "; //'wsd_uom_org_qty
                        strSQL = strSQL + "0, '";

                        strSQL = strSQL + "0" + "', '"; //'wsd_tot_disc_1
                        strSQL = strSQL + "0" + "', "; //'wsd_tot_disc_2
                        strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxCode1"].Value.ToString() + "', '"; //wsd_tax_code
                        strSQL = strSQL + dgvDetail2.Rows[lCounter].Cells["g_iDetTaxPct1"].Value.ToString() + "',"; //wsd_tax_pct
                        strSQL = strSQL + "0, "; //wsd_tax_amount
                        strSQL = strSQL + "0, "; //wsd_gross_net
                        strSQL = strSQL + "0, "; //wsd_cash_disc
                        strSQL = strSQL + "'N', '"; //wsd_line_type
                        strSQL = strSQL + dtSODate.Value.ToString("yyyy-MM-dd") + "', '";
                        strSQL = strSQL + sPriceCode + "', "; //wsd_price_code
                        if (dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                        {
                            strSQL = strSQL + "NULL"; //wsd_sled
                        }
                        else
                        {
                            strSQL = strSQL + "'"+ dgvDetail2.Rows[lCounter].Cells["g_iDetSled1"].Value.ToString() + "'"; //wsd_sled
                        }
                        strSQL = strSQL + ")"; //wsd_reason
                        dtFill15 = _clsGlobal.ExecDTTrans(strSQL);
                        #endregion
                    }


                }

                keyFound = true;
            }
            catch (Exception ex)
            {
                keyFound = false;
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }





    }
}

