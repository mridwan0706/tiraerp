using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Security.Cryptography;

using System.Collections;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.DXErrorProvider;
namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{

    public partial class frmSOManualEntryEditor : DevExpress.XtraEditors.XtraForm
    {
        // helper: jumlah baris data-source LookUpEdit (pengganti ComboBox.Items.Count)
        private int LookupRowCount(DevExpress.XtraEditors.LookUpEdit lue)
        {
            object ds = lue.Properties.DataSource;
            System.Data.DataView dv = ds as System.Data.DataView;
            if (dv != null) return dv.Count;
            System.Data.DataTable dt = ds as System.Data.DataTable;
            if (dt != null) return dt.Rows.Count;
            System.Collections.ICollection col = ds as System.Collections.ICollection;
            if (col != null) return col.Count;
            return 0;
        }

        // Ambil nilai ValueMember pada baris ke-index langsung dari DataSource.
        // Deterministik: tidak bergantung pada ItemIndex->EditValue LookUpEdit yang
        // tidak sinkron saat editor belum ter-render (beda perilaku dgn WinForms ComboBox).
        private object LookupValueAt(DevExpress.XtraEditors.LookUpEdit lue, int index)
        {
            if (lue == null || index < 0) return null;
            object ds = lue.Properties.DataSource;
            string vm = lue.Properties.ValueMember;
            System.Data.DataView dv = ds as System.Data.DataView;
            System.Data.DataTable dt = ds as System.Data.DataTable;
            if (dt == null && dv != null) dt = dv.Table;
            if (dt == null || string.IsNullOrEmpty(vm) || !dt.Columns.Contains(vm)) return null;
            if (dv != null) { if (index < dv.Count) return dv[index][vm]; return null; }
            if (index < dt.Rows.Count) return dt.Rows[index][vm];
            return null;
        }

        // Grid field constants: gunakan nama field/DataColumn, bukan index angka.
        // Ini membuat logic tetap aman meskipun urutan/tampilan kolom berubah.
        private const string COL_PCODE = "wsd_prd_master_code";
        private const string COL_REAL_ORDER_QTY = "wsd_real_order_xqty";
        private const string COL_ORDER_QTY = "wsd_order_xqty";
        private const string COL_QTY_REQ = "wsd_req_order_xqty";
        private const string COL_REASON = "reason";
        private const string COL_TGL_PRICE = "tgl_price";
        private const string COL_IS_SUBS = "wsd_is_subs";
        private const string COL_SUBS_LOOKUP = "wsd_subs_lookup";

        // DevExpress validation provider untuk mandatory field.
        // MessageBox lama tetap dipertahankan; ini hanya menambahkan tanda merah di form.
        private DXErrorProvider dxErrorProvider1;

        private void InitValidationProvider()
        {
            if (dxErrorProvider1 == null)
            {
                dxErrorProvider1 = new DXErrorProvider();
                dxErrorProvider1.ContainerControl = this;
            }
        }

        private void ClearFormValidationErrors()
        {
            InitValidationProvider();
            dxErrorProvider1.ClearErrors();
        }

        private void SetControlError(Control control, string message)
        {
            InitValidationProvider();
            if (control != null)
                dxErrorProvider1.SetError(control, message);
        }

        private bool IsEmptyControl(Control control)
        {
            return control == null || string.IsNullOrWhiteSpace(Convert.ToString(control.Text));
        }

        private string SafeCellValue(DataGridViewRow row, string columnName)
        {
            try
            {
                if (row == null || row.Cells == null || string.IsNullOrWhiteSpace(columnName))
                    return string.Empty;

                return Convert.ToString(row.Cells[columnName].Value);
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetProductCodeForMessage(DataGridViewRow row)
        {
            string productCode = SafeCellValue(row, COL_PCODE);
            return string.IsNullOrWhiteSpace(productCode) ? "-" : productCode;
        }

        private string GetFirstInvalidProductMessage()
        {
            if (dgvSalesDetail == null || dgvSalesDetail.Rows == null || dgvSalesDetail.Rows.Count <= 0)
                return "Product masih kosong";

            for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
            {
                DataGridViewRow row = dgvSalesDetail.Rows[i];
                if (row == null || row.IsNewRow)
                    continue;

                if (string.IsNullOrWhiteSpace(SafeCellValue(row, COL_PCODE)))
                    return "Product masih kosong";
            }

            return string.Empty;
        }

        private string GetFirstInvalidReasonMessage()
        {
            if (dgvSalesDetail == null || dgvSalesDetail.Rows == null)
                return string.Empty;

            for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
            {
                DataGridViewRow row = dgvSalesDetail.Rows[i];
                if (row == null || row.IsNewRow)
                    continue;

                double qtyRealOrder = 0;
                double qtyReqOrder = 0;

                try
                {
                    fQtyFormat(SafeCellValue(row, COL_REAL_ORDER_QTY),
                        Convert.ToInt16(SafeCellValue(row, "wsd_uom_convert_big")),
                        Convert.ToInt16(SafeCellValue(row, "wsd_uom_convert_mid")));
                    qtyRealOrder = Convert.ToDouble(_fConvertQty);

                    fQtyFormat(SafeCellValue(row, COL_QTY_REQ),
                        Convert.ToInt16(SafeCellValue(row, "wsd_uom_convert_big")),
                        Convert.ToInt16(SafeCellValue(row, "wsd_uom_convert_mid")));
                    qtyReqOrder = Convert.ToDouble(_fConvertQty);
                }
                catch
                {
                    continue;
                }

                if (qtyRealOrder != qtyReqOrder && string.IsNullOrWhiteSpace(SafeCellValue(row, COL_REASON)))
                    return "Reason masih kosong [PCODE: " + GetProductCodeForMessage(row) + "]";
            }

            return string.Empty;
        }

        private void MarkMandatoryValidationErrors()
        {
            ClearFormValidationErrors();

            if (IsEmptyControl(txtKdOutlet1))
                SetControlError(txtKdOutlet1, "Outlet masih kosong");

            if (IsEmptyControl(cbPembyFaktur))
                SetControlError(cbPembyFaktur, "Pembayaran Faktur masih kosong");

            if (IsEmptyControl(txtNoPO))
                SetControlError(txtNoPO, "No PO masih kosong");
            else if (isNefoKAM && txtNoPO.Text.Length > 35)
                SetControlError(txtNoPO, "No PO tidak boleh lebih dari 35 karakter");

            if (IsEmptyControl(cbTOP))
                SetControlError(cbTOP, "TOP masih kosong");

            if (IsEmptyControl(cbTipeOrder))
                SetControlError(cbTipeOrder, "Order Type masih kosong");

            string productMessage = GetFirstInvalidProductMessage();
            if (!string.IsNullOrWhiteSpace(productMessage))
                SetControlError(dgvSalesDetail, productMessage);

            string reasonMessage = GetFirstInvalidReasonMessage();
            if (!string.IsNullOrWhiteSpace(reasonMessage))
                SetControlError(dgvSalesDetail, reasonMessage);

            if (rbCashDiscP.Checked && !string.IsNullOrWhiteSpace(txtCashDiscP.Text))
            {
                decimal cashDiscPercent;
                if (decimal.TryParse(txtCashDiscP.Text.Trim(), out cashDiscPercent) && cashDiscPercent > 0)
                {
                    if (cmbBebanDisc.EditValue != null && Convert.ToString(cmbBebanDisc.EditValue) == "0")
                        SetControlError(cmbBebanDisc, "Apabila Cash Disc % diisi Beban Cash Disc harus diisi");
                }
            }

            if (rbCashDiscU.Checked && !string.IsNullOrWhiteSpace(txtCashDiscU.Text))
            {
                decimal cashDiscAmount;
                if (decimal.TryParse(txtCashDiscU.Text.Trim(), out cashDiscAmount) && cashDiscAmount > 0)
                {
                    if (cmbBebanDisc.EditValue != null && Convert.ToString(cmbBebanDisc.EditValue) == "0")
                        SetControlError(cmbBebanDisc, "Apabila Rp Cash Disc diisi Beban Cash Disc harus diisi");
                }
            }
        }

        private string GetFirstMandatoryPopupMessage()
        {
            if (IsEmptyControl(txtKdOutlet1))
                return "Outlet masih kosong";

            if (IsEmptyControl(cbPembyFaktur))
                return "Pembayaran Faktur masih kosong";

            if (IsEmptyControl(txtNoPO))
                return "No PO masih kosong";

            if (isNefoKAM && txtNoPO.Text.Length > 35)
                return "No PO Tidak Boleh Dari 35 Karakter!!";

            if (IsEmptyControl(cbTOP))
                return "TOP masih kosong";

            if (IsEmptyControl(cbTipeOrder))
                return "Order Type masih kosong";

            string productMessage = GetFirstInvalidProductMessage();
            if (!string.IsNullOrWhiteSpace(productMessage))
                return productMessage;

            string reasonMessage = GetFirstInvalidReasonMessage();
            if (!string.IsNullOrWhiteSpace(reasonMessage))
                return reasonMessage;

            return string.Empty;
        }

        private string GetMainValidationMessage()
        {
            string message = GetFirstMandatoryPopupMessage();
            return string.IsNullOrWhiteSpace(message) ? "Mohon lengkapi data yang wajib diisi" : message;
        }

        private string GetGridColumnName(DataGridView grid)
        {
            if (grid == null || grid.CurrentCell == null)
                return string.Empty;

            int columnIndex = grid.CurrentCell.ColumnIndex;
            if (columnIndex < 0)
                return string.Empty;

            try
            {
                DataTable dt = grid.DataSource as DataTable;
                if (dt != null && columnIndex < dt.Columns.Count)
                    return dt.Columns[columnIndex].ColumnName;
            }
            catch { }

            try
            {
                if (columnIndex < grid.Columns.Count)
                    return grid.Columns[columnIndex].Name;
            }
            catch { }

            return string.Empty;
        }

        private bool IsGridColumn(DataGridView grid, string columnName)
        {
            return string.Equals(GetGridColumnName(grid), columnName, StringComparison.OrdinalIgnoreCase);
        }


        private string ExtractTopDaysSafe(string topText)
        {
            string tops = Convert.ToString(topText).Trim();
            if (string.IsNullOrEmpty(tops))
                return "0";

            int tildeIndex = tops.IndexOf("~");
            if (tildeIndex >= 0)
            {
                int start = tildeIndex + 1;
                int dashIndex = tops.IndexOf("-", start);
                string middle = dashIndex > start ? tops.Substring(start, dashIndex - start) : tops.Substring(start);
                middle = Regex.Replace(middle, @"[^0-9]", "").Trim();
                if (!string.IsNullOrEmpty(middle))
                    return middle;
            }

            Match match = Regex.Match(tops, @"\d+");
            return match.Success ? match.Value : "0";
        }
        #region Variable

        private clsGlobal _clsGlobal = new clsGlobal();
        private string strSQL;
        private DataTable dtGridSODetail = new DataTable();
        private DataTable dtGridTPRB = new DataTable();
        private DataTable dtGridDisc = new DataTable();
        private DataTable dtGridTPRU = new DataTable();
        private DataTable dtGridTax = new DataTable();
        private DataTable dtGridPromoTax = new DataTable();
        private DataGridViewComboBoxCell cbCellReason;
        private RepositoryItemSearchLookUpEdit _repoReasonSearchLookUp;
        private RepositoryItemSearchLookUpEdit _repoSubsSearchLookUp;
        private DataTable _dtReasonSearchLookUp;
        private bool _isReasonSearchLookUpHooked = false;
        private bool _isSubsSearchLookUpHooked = false;
        private bool _ignoreSLEDCheckEvent = false;
        private bool _updatingHeaderParty = false;

        //private int _xHaveClick;
        private string sSONo;
        private string strFlagLimit;
        private string strRoviUpdUser;
        private string strRoviUpdDate;
        private string _xSONo;
        private string _xSalesID;
        private string _xSalesDesc;
        private string _xTipeSales;
        private string _xWHLoc1;
        private string _xWHLoc2;
        private string _xOrderType;
        private string _xEmployee;
        private DateTime _xTglTransaksi;
        private DateTime _xTglOrder;
        private string _xEntityID;
        private string _xBranchID;
        private string _xFlagDC;

        private DateTime TglDeliveryDate;
        private DateTime TglExpiredDate;

        private double lAFSQty = 0;

        private string sSummaryNo;

        private decimal curCost;
        private decimal curPrice;
        private string _fQtyFormat;
        private string _fConvertQty;
        private decimal _curCost;

        private bool xSaveOOS;
        private int xCountOOS;
        private bool bolOOS;
        private bool bolCloseSO;

        private string g_sOutlet;

        private bool bFlagCr;
        private bool g_bProses;
        private bool g_bReqProses;
        private bool g_bReqSummary;
        private bool g_bReqShipment;
        private bool g_bReqPOD;
        private bool g_bReqInvoice;
        private bool g_bSOType;
        private bool g_bTOPFlag;
        private bool g_bChangeOutlet;
        private bool g_bEdit;
        private bool g_bActiveEdit;
        private int lCounter;
        private string g_sDefaultSONo = "XXXXXXXXXX";
        private int lRow = 0;
        private string _CurrTOP;
        private string _CurrPembayaran;

        private decimal _curAvCredLimit;
        private bool isMsgGudangShown = false;
        private bool isMsgOrderTypeShown = false;
        private bool isPcodeGridExists = false;
        private int editedgvRowid = 0;
        private string[] arr_BebanDisc;

        private bool isNefoKAM = false;
        private bool IsNefoForDC = false;
        private bool isCrossSite = false;
        private bool isRoundingMekanism = false;
        private string isTOPDivision = "N";
        private string MsgErrorDivision = "";
        private string TOPDivision = "";
        private bool APIBlockingSKU = false;

        private bool isMultisource = false;


        private bool isMultibranch = false;
        private DateTime g_dTglGudang;
        private DateTime g_CreateAtSAP;

        /** deadlock SeqCalc **/
        private bool isCalcNoLock = false;

        /** deadlock SeqCalc **/
        private string NO_HP_Outlet = "";
        private string NO_HP_Salesman = "";

        public string xSONo
        {
            get { return _xSONo; }
            set { _xSONo = value; }
        }

        public string xEntityID
        {
            get { return _xEntityID; }
            set { _xEntityID = value; }
        }

        public string xBranchID
        {
            get { return _xBranchID; }
            set { _xBranchID = value; }
        }

        public string xFlagDC
        {
            get { return _xFlagDC; }
            set { _xFlagDC = value; }
        }

        public string xSalesID
        {
            get { return _xSalesID; }
            set { _xSalesID = value; }
        }

        public string xSalesDesc
        {
            get { return _xSalesDesc; }
            set { _xSalesDesc = value; }
        }

        public string xTipeSales
        {
            get { return _xTipeSales; }
            set { _xTipeSales = value; }
        }

        public string xWHLoc1
        {
            get { return _xWHLoc1; }
            set { _xWHLoc1 = value; }
        }

        public string xWHLoc2
        {
            get { return _xWHLoc2; }
            set { _xWHLoc2 = value; }
        }

        public string xOrderType
        {
            get { return _xOrderType; }
            set { _xOrderType = value; }
        }

        public string xEmployee
        {
            get { return _xEmployee; }
            set { _xEmployee = value; }
        }

        public DateTime xTglTransaksi
        {
            get { return _xTglTransaksi; }
            set { _xTglTransaksi = value; }
        }

        public DateTime xTglOrder
        {
            get { return _xTglOrder; }
            set { _xTglOrder = value; }
        }

        #endregion

        public frmSOManualEntryEditor()
        {
            InitializeComponent();
            InitValidationProvider();
            txtTtlInvoice.TextChanged += txtHeaderSource_TextChanged;
        }

        #region WinForm

        private void frmSOManualEntryEditor_Load(object sender, EventArgs e)
        {

            this.Text = clsLogin.MENUID + " - " + clsLogin.MENUNAME + " Editor";
            WorkDateAccess();
            dtTglOrder.Text = DateTime.Now.ToString("dd MMM yyyy");
            dtTglPO.Text = DateTime.Now.ToString("dd MMM yyyy");
            txtNoOrder.Text = "XXXXXXXXXX";
            txtSalesID.Text = xSalesID;
            txtSalesDesc.Text = xSalesDesc;
            lblTipeSales.Text = xTipeSales;
            txtTglTrans.Text = xTglTransaksi.ToString("dd MMM yyyy");
            txtCreateDate.Text = "";
            lblWHLoc1.Text = xWHLoc1;
            lblWHLoc2.Text = xWHLoc2;
            isRoundingMekanism = IsRoundingMekanism();
            isMultibranch = IsMultibranch();
            /** deadlock SeqCalc **/
            if (isMultibranch && IsNoLockMultibranch())
            {
                isCalcNoLock = true;
                string orderCalcNoLocking = createSeqCalc(xBranchID, "S");
                txtNoOrder.Text = orderCalcNoLocking;
                g_sDefaultSONo = orderCalcNoLocking;

            }
            txtUser.Text = clsLogin.USERID;
            rbCashDiscP.Checked = true;
            rbCashDiscU.Checked = false;
            txtCashDiscP.Enabled = true;
            txtCashDiscU.Enabled = false;
            if (rbCashDiscP.Checked == true)
            {
                txtCashDiscU.Text = "0";
            }
            this.ActiveControl = txtKdOutlet1;
            BindHeaderEntityBranchLookup();
            BindOutletLookup();

            isNefoKAM = _clsGlobal.IsNefoForKam;
            IsNefoForDC = _clsGlobal.IsNefoForDC;
            isCrossSite = _clsGlobal.IsCrossSite;
            isMultisource = _clsGlobal.IsMultiSource;

            //if (chkSLED.Checked)
            //{
            //    txtSLED.ReadOnly = false;
            //}
            //else
            //{
            //    txtSLED.ReadOnly = true;
            //}


            if (isMultisource)
            {
                //cbSource.Enabled = true;
                cbFlagSloc.Enabled = true;
                BindSource();
                BindSloc();

                if (isCrossSite)
                {
                    label33.Visible = true;
                    txtShippingPlant.Visible = true;
                    btnPopUpShippingPlant.Visible = true;
                    //BindShippingPlant();
                }
                else
                {
                    label33.Visible = false;
                    txtShippingPlant.Visible = false;
                    btnPopUpShippingPlant.Visible = false;
                }

            }
            else
            {
                //cbSource.Enabled = false;
                cbFlagSloc.Enabled = false;
                label33.Visible = false;
                txtShippingPlant.Visible = false;
                btnPopUpShippingPlant.Visible = false;
            }

            cbResetReq.Checked = false;
            cbResetReq.Enabled = false;
            BindBebanDisc();

            BindPembyFaktur();
            g_bActiveEdit = false;
            SetGridSODetail(dgvSalesDetail);
            EnsureSubstitutionSearchLookUpEditor();
            SetGridTPRB(dgvPromosiQty);
            SetGridTPRU(dgvPromosiRp);
            SetGridDisc(dgvDisc);

            txtSLED.Visible = false;

            InitialTOPDivision();

            if (_xFlagDC == "Y")
            {
                bool tflag = fGetTOPFlag();
                if (tflag == false)
                {
                    btnProses.Enabled = false;
                    btnOk.Enabled = false;
                    btnCalcDisc.Enabled = false;
                    btnPopUpOutlet.Enabled = false;
                    dgvSalesDetail.Enabled = false;
                    cmdSales.Enabled = false;
                    cbTipeOrder.Enabled = false;
                    cbPembyFaktur.Enabled = false;
                    cbTOP.Enabled = false;

                }
            }
            else
            {
                g_bTOPFlag = false;
            }

            txtUser.Text = clsLogin.USERID;


            BindTOP();
            cbTOP.EditValueChanged -= cbTOP_SelectedIndexChanged;
            bindTipeOrder();

            BindTOPFrmSalesman();
            cbTOP.EditValueChanged += cbTOP_SelectedIndexChanged;
            //getReason();
            if (clsGlobal.MODE_TRX == 1)
            {
                dtTglDelv.Properties.MinValue = g_dTglGudang;
                dtTglExp.Properties.MinValue = g_dTglGudang;
            }









            if (clsGlobal.MODE_TRX == 2)
            {
                //FillData();
                ShowMe(txtNoOrder.Text);
                IsBlockingAPI();

                if (APIBlockingSKU)
                {
                    //dgvSalesDetail.Columns["wsd_prd_master_code"].ReadOnly = true;
                    //dgvSalesDetail.Columns["wsd_real_order_xqty"].ReadOnly = true;
                    //dgvSalesDetail.Columns["tgl_price"].ReadOnly = true;
                    //dgvSalesDetail.ReadOnly = true;

                    dgvSalesDetail.Enabled = false;
                    cbSource.Enabled = false;
                }
            }


            if (clsGlobal.MODE_TRX == 1)
            {
                initDateDelvExp(Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd"), "", "", 1);


            }
        }

        private bool IsRoundingMekanism()
        {

            DataTable dtGS = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'FLAG_ROUNDING_MEKANISM' and gh_sys = 'H'");
            if (dtGS.Rows.Count > 0)
            {
                if (dtGS.Rows[0][0].ToString().Trim() == "Y")
                {
                    return true;
                }
            }


            return false;
        }

        private bool IsMultibranch()
        {

            DataTable dtGS = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'FLAGMULTIBRANCH' and gh_sys = 'H'");
            if (dtGS.Rows.Count > 0)
            {
                if (dtGS.Rows[0][0].ToString().Trim() == "Y")
                {
                    return true;
                }
            }


            return false;
        }

        private void initDateDelvExp(string sodate, string custCode1, string custCode2, int executeType)
        {
            if (clsGlobal.MODE_TRX == 1)
            {
                DataTable dt = new DataTable();
                string sqlScript = "Select dbo.F_GET_DATE_PLAN ('" + xEntityID + "','" + xBranchID + "','" + sodate + "','" + custCode1 + "','" + custCode2 + "')";
                if (executeType == 1) // exec biasa ga pake tran2an
                {
                    dt = _clsGlobal.ExecDT(sqlScript);
                }
                else
                {
                    dt = _clsGlobal.ExecDTTrans(sqlScript);
                }

                if (dt.Rows.Count > 0)
                {
                    dtTglDelv.Text = DateTime.ParseExact(dt.Rows[0][0].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                    dtTglExp.Text = DateTime.ParseExact(dt.Rows[0][0].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");

                    TglDeliveryDate = DateTime.ParseExact(dt.Rows[0][0].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    TglExpiredDate = DateTime.ParseExact(dt.Rows[0][0].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                }
            }

        }

        private void bindTipeOrder()
        {
            DataTable dtFill1 = new DataTable();


            strSQL = "SELECT distinct CASE WHEN sgm_type_operasi='C' THEN 1102 ELSE 1101 end Code";
            strSQL = strSQL + " ,CASE WHEN sgm_type_operasi='C' THEN '1102 ~ CV-SALES' ELSE '1101 ~ TO-SALES' END order_type";
            strSQL = strSQL + " ,CASE WHEN sgm_type_operasi='C' THEN '1102' ELSE '1101' END order_type";
            strSQL = strSQL + " FROM SO_SPG_GIRL_MAN WITH(NOLOCK)";
            if (!string.IsNullOrWhiteSpace(txtSalesID.Text))
            {
                strSQL = strSQL + " WHERE sgm_spgm_id=" + _clsGlobal.FmtStr(txtSalesID.Text) + "and sgm_entity_id = '" + _xEntityID + "' and sgm_branch_id = '" + _xBranchID + "' ";
            }

            strSQL = strSQL + " UNION";
            strSQL = strSQL + " select distinct msow_so_type code, msow_so_type + ' ~ ' + ot_desc order_type,msow_so_type ";
            strSQL = strSQL + " From TBL_MAP_SLD_OPRT_WH WITH(NOLOCK)";
            strSQL = strSQL + " inner join SO_ORDER_TYPE WITH(NOLOCK) on msow_so_type = ot_order_type where 1=1";
            if (!string.IsNullOrWhiteSpace(txtSalesID.Text))
            {
                strSQL = strSQL + " and msow_sld_id =" + _clsGlobal.FmtStr(txtSalesID.Text) + " and msow_entity_id = '" + _xEntityID + "' and msow_branch_id = '" + _xBranchID + "' ";
            }
            strSQL = strSQL + " and ot_transaction_type = 'D'";
            strSQL = strSQL + " order by code desc";



            //strSQL = " select msow_so_type code, msow_so_type + ' ~ ' + ot_desc order_type From TBL_MAP_SLD_OPRT_WH ";
            //strSQL = strSQL + " inner join SO_ORDER_TYPE on msow_so_type = ot_order_type";
            //strSQL = strSQL + " where 1=1 ";
            //if (!string.IsNullOrWhiteSpace(txtSalesID.Text))
            //{
            //    strSQL = strSQL + " and msow_sld_id = '" + txtSalesID.Text + "'";
            //}
            //strSQL = strSQL + " and ot_transaction_type = 'D'";
            //strSQL = strSQL + " Order By msow_so_type desc";


            if (_clsGlobal.Connect.State == ConnectionState.Closed)
            {
                cbTipeOrder.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            }
            else
            {
                cbTipeOrder.Properties.DataSource = _clsGlobal.ExecDTTrans(strSQL);
            }
            cbTipeOrder.Properties.ValueMember = "code";
            cbTipeOrder.Properties.DisplayMember = "order_type";
            cbTipeOrder.Properties.ForceInitialize();

            for (lCounter = 0; lCounter < LookupRowCount(cbTipeOrder); lCounter++)
            {
                g_bSOType = false;
                cbTipeOrder.EditValue = LookupValueAt(cbTipeOrder, lCounter);
                g_bSOType = true;
                //DataRow selectedDataRow = ((DataRowView)cbTipeOrder.GetSelectedDataRow()).Row;
                string idO = cbTipeOrder.EditValue == null ? "" : Convert.ToString(cbTipeOrder.EditValue);
                string NameO = cbTipeOrder.Text;
                string tipesales = string.Empty;
                string a = _xTipeSales.Trim().ToLower().Replace(" ", "");
                if (_xTipeSales.Trim().ToLower().Replace(" ", "") == "takingorder")
                {
                    tipesales = "1101";
                }
                else
                {
                    tipesales = "1102";
                }
                if (idO == tipesales)
                {
                    return;
                    //cbTipeOrder.ItemIndex = lCounter;
                }
                //if (idO == lblOrderType.Text)
                //{
                //    return;
                //}

                if (lCounter == LookupRowCount(cbTipeOrder))
                {
                    g_bSOType = false;
                    cbTipeOrder.EditValue = LookupValueAt(cbTipeOrder, 0);
                    g_bSOType = true;
                }
            }
        }

        private void BindSource()
        {
            strSQL = "";
            strSQL = "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'SOURCE_CLNT'";

            cbSource.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cbSource.Properties.ValueMember = "fcode";
            cbSource.Properties.DisplayMember = "fdesc";
            cbSource.Properties.ForceInitialize();
            cbSource.EditValue = "N";
        }

        private void BindSloc()
        {
            strSQL = "";
            strSQL = "select gh_function_code as fcode, gh_function_desc as fdesc from GS_GEN_HARDCODED WITH (NOLOCK)   where gh_sys = 'H' and gh_function_name = 'MULTISOURCE' ORDER BY fcode ASC";

            cbFlagSloc.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cbFlagSloc.Properties.ValueMember = "fcode";
            cbFlagSloc.Properties.DisplayMember = "fdesc";
            cbFlagSloc.Properties.ForceInitialize();
            cbFlagSloc.EditValue = LookupValueAt(cbFlagSloc, 0);
        }

        private void frmSOManualEntryEditor_Activated(object sender, EventArgs e)
        {

            try
            {

                g_bProses = false;
                g_bSOType = true;
                //g_bChangeOutlet = false;

                if (g_bActiveEdit == false)
                {

                    DataTable dtFill1 = new DataTable();


                    strSQL = "SELECT distinct CASE WHEN sgm_type_operasi='C' THEN 1102 ELSE 1101 end Code";
                    strSQL = strSQL + " ,CASE WHEN sgm_type_operasi='C' THEN '1102 ~ CV-SALES' ELSE '1101 ~ TO-SALES' END order_type";
                    strSQL = strSQL + " ,CASE WHEN sgm_type_operasi='C' THEN '1102' ELSE '1101' END order_type";
                    strSQL = strSQL + " FROM SO_SPG_GIRL_MAN WITH(NOLOCK)";
                    if (!string.IsNullOrWhiteSpace(txtSalesID.Text))
                    {
                        strSQL = strSQL + " WHERE sgm_spgm_id=" + _clsGlobal.FmtStr(txtSalesID.Text) + "AND sgm_entity_id = '" + xEntityID + "' AND sgm_branch_id = '" + xBranchID + "' ";
                    }

                    strSQL = strSQL + " UNION";
                    strSQL = strSQL + " select distinct msow_so_type code, msow_so_type + ' ~ ' + ot_desc order_type,msow_so_type ";
                    strSQL = strSQL + " From TBL_MAP_SLD_OPRT_WH WITH(NOLOCK)";
                    strSQL = strSQL + " inner join SO_ORDER_TYPE WITH(NOLOCK) on msow_so_type = ot_order_type where 1=1";
                    if (!string.IsNullOrWhiteSpace(txtSalesID.Text))
                    {
                        strSQL = strSQL + " and msow_sld_id =" + _clsGlobal.FmtStr(txtSalesID.Text) + " and  msow_entity_id = '" + xEntityID + "' and  msow_branch_id = '" + xBranchID + "' ";
                    }
                    strSQL = strSQL + " and ot_transaction_type = 'D'";
                    strSQL = strSQL + " order by code desc";

                    strSQL = " select msow_so_type code, msow_so_type + ' ~ ' + ot_desc order_type From TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) ";
                    strSQL = strSQL + " inner join SO_ORDER_TYPE WITH (NOLOCK) on msow_so_type = ot_order_type";
                    strSQL = strSQL + " where msow_sld_id = '" + txtSalesID.Text + "' and msow_entity_id = '" + xEntityID + "' and  msow_branch_id = '" + xBranchID + "' ";
                    strSQL = strSQL + " and ot_transaction_type = 'D'";
                    strSQL = strSQL + " Order By msow_so_type desc";

                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    {
                        dtFill1 = _clsGlobal.ExecDT(strSQL);
                    }
                    else
                    {
                        dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                    }
                    if (dtFill1.Rows.Count > 0)
                    {
                        cbTipeOrder.Properties.DataSource = null;

                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        {
                            cbTipeOrder.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                        }
                        else
                        {
                            cbTipeOrder.Properties.DataSource = _clsGlobal.ExecDTTrans(strSQL);
                        }
                        cbTipeOrder.Properties.ValueMember = "code";
                        cbTipeOrder.Properties.DisplayMember = "order_type";
                        cbTipeOrder.Properties.ForceInitialize();
                    }


                    //coment dulu
                    for (lCounter = 0; lCounter < LookupRowCount(cbTipeOrder); lCounter++)
                    {
                        g_bSOType = false;
                        cbTipeOrder.EditValue = LookupValueAt(cbTipeOrder, lCounter);
                        g_bSOType = true;
                        string idO = Convert.ToString(cbTipeOrder.EditValue);
                        string NameO = cbTipeOrder.Text;
                        string tipesales = string.Empty;
                        string a = _xTipeSales.Trim().ToLower().Replace(" ", "");
                        if (_xTipeSales.Trim().ToLower().Replace(" ", "") == "takingorder")
                        {
                            tipesales = "1101";
                        }
                        else
                        {
                            tipesales = "1102";
                        }
                        if (idO == tipesales)
                        {
                            return;

                        }


                        if (lCounter == LookupRowCount(cbTipeOrder))
                        {
                            g_bSOType = false;
                            cbTipeOrder.EditValue = LookupValueAt(cbTipeOrder, 0);
                            g_bSOType = true;
                        }
                    }

                }

                DataTable dtFill2 = new DataTable();
                strSQL = "select isnull(ot_process_req, '') as  ot_process_req ";
                strSQL = strSQL + " ,isnull(ot_summary_req, '') as ot_summary_req ";
                strSQL = strSQL + " ,isnull(ot_shipment_req, '') as ot_shipment_req ";
                strSQL = strSQL + " ,isnull(ot_pod_req, '') as ot_pod_req ";
                strSQL = strSQL + " ,isnull(ot_inv_req, '') as ot_inv_req   From SO_ORDER_TYPE WITH (NOLOCK) ";
                strSQL = strSQL + " where ot_order_type = '" + lblOrderType.Text + "'";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                {
                    dtFill2 = _clsGlobal.ExecDT(strSQL);
                }
                else
                {
                    dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                }
                if (dtFill2.Rows.Count > 0)
                {
                    if (dtFill2.Rows[0]["ot_process_req"].ToString() == "Y")
                    {
                        g_bReqProses = true;
                        //btnProses.Visible = true;
                    }
                    else
                    {
                        g_bReqProses = false;
                        //btnProses.Visible = false;
                    }

                    if (dtFill2.Rows[0]["ot_summary_req"].ToString() == "Y")
                    {
                        g_bReqSummary = true;
                        //btnProses.Visible = true;
                    }
                    else
                    {
                        g_bReqSummary = false;
                        //btnProses.Visible = false;
                    }

                    if (dtFill2.Rows[0]["ot_shipment_req"].ToString() == "Y")
                    {
                        g_bReqShipment = true;
                        //btnProses.Visible = true;
                    }
                    else
                    {
                        g_bReqShipment = false;
                        //btnProses.Visible = false;
                    }

                    if (dtFill2.Rows[0]["ot_pod_req"].ToString() == "Y")
                    {
                        g_bReqPOD = true;
                        //btnProses.Visible = true;
                    }
                    else
                    {
                        g_bReqPOD = false;
                        //btnProses.Visible = false;
                    }

                    if (dtFill2.Rows[0]["ot_inv_req"].ToString() == "Y")
                    {
                        g_bReqInvoice = true;
                        //btnProses.Visible = true;
                    }
                    else
                    {
                        g_bReqInvoice = false;
                        //btnProses.Visible = false;
                    }
                }
                else
                {
                    if (isMsgOrderTypeShown == false)
                    {
                        MessageBox.Show("Data order type : " + cbTipeOrder.Text + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        isMsgOrderTypeShown = true;
                    }

                }


            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void BindHeaderEntityBranchLookup()
        {
            try
            {
                string sqlEntity = "select ge_entity_id [Code], ge_entity [Description] from GS_ENTITY WITH(NOLOCK) inner join GL_USER_ENTITY WITH(NOLOCK) on ge_entity_id = ue_entity_id where ue_user_id = '" + clsLogin.USERID + "' order by ge_entity_id";
                DataTable dtEntity = (_clsGlobal.Connect.State == ConnectionState.Closed) ? _clsGlobal.ExecDT(sqlEntity) : _clsGlobal.ExecDTTrans(sqlEntity);
                txtHeaderEntity.Properties.DataSource = dtEntity;
                txtHeaderEntity.Properties.DisplayMember = "Code";
                txtHeaderEntity.Properties.ValueMember = "Code";
                txtHeaderEntityView.Columns.Clear();
                txtHeaderEntityView.Columns.AddVisible("Code", "Entity");
                txtHeaderEntityView.Columns.AddVisible("Description", "Description");
                txtHeaderEntityView.BestFitColumns();
                txtHeaderEntity.EditValue = xEntityID;
                SetHeaderLookupDescription(txtHeaderEntity, txtHeaderEntityDesc);

                BindHeaderBranchLookup();
                txtHeaderBranch.EditValue = xBranchID;
                SetHeaderLookupDescription(txtHeaderBranch, txtHeaderBranchDesc);
            }
            catch { }
        }

        private void BindHeaderBranchLookup()
        {
            try
            {
                string entity = Convert.ToString(txtHeaderEntity.EditValue);
                if (string.IsNullOrWhiteSpace(entity)) entity = xEntityID;
                string sqlBranch = "select gu_branch [Code], br_branch_desc [Description] from VW_USERS_SECURITY WITH(NOLOCK) where gu_entity = '" + entity.Replace("'", "''") + "' and gu_user_id = '" + clsLogin.USERID + "' order by gu_branch";
                DataTable dtBranch = (_clsGlobal.Connect.State == ConnectionState.Closed) ? _clsGlobal.ExecDT(sqlBranch) : _clsGlobal.ExecDTTrans(sqlBranch);
                txtHeaderBranch.Properties.DataSource = dtBranch;
                txtHeaderBranch.Properties.DisplayMember = "Code";
                txtHeaderBranch.Properties.ValueMember = "Code";
                txtHeaderBranchView.Columns.Clear();
                txtHeaderBranchView.Columns.AddVisible("Code", "Branch");
                txtHeaderBranchView.Columns.AddVisible("Description", "Description");
                txtHeaderBranchView.BestFitColumns();
            }
            catch { }
        }

        private void SetHeaderLookupDescription(SearchLookUpEdit lookup, TextEdit desc)
        {
            try
            {
                desc.Text = string.Empty;
                if (lookup == null || desc == null || lookup.EditValue == null || !(lookup.Properties.DataSource is DataTable dt)) return;
                string val = Convert.ToString(lookup.EditValue).Replace("'", "''");
                DataRow[] rows = dt.Select("[Code] = '" + val + "'");
                if (rows.Length > 0)
                    desc.Text = Convert.ToString(rows[0]["Description"]).Trim();
            }
            catch { }
        }

        private void txtHeaderEntity_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            if (e.Button.Kind == ButtonPredefines.Search)
                txtHeaderEntity.ShowPopup();
        }

        private void txtHeaderBranch_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            if (e.Button.Kind == ButtonPredefines.Search)
                txtHeaderBranch.ShowPopup();
        }

        private void txtHeaderEntity_EditValueChanged(object sender, EventArgs e)
        {
            SetHeaderLookupDescription(txtHeaderEntity, txtHeaderEntityDesc);
            _xEntityID = Convert.ToString(txtHeaderEntity.EditValue);
            BindHeaderBranchLookup();
            if (!string.Equals(Convert.ToString(txtHeaderBranch.EditValue), xBranchID, StringComparison.OrdinalIgnoreCase))
                txtHeaderBranch.EditValue = xBranchID;
            BindOutletLookup();
            RefreshAllSubstitutionFlags();
        }

        private void txtHeaderBranch_EditValueChanged(object sender, EventArgs e)
        {
            SetHeaderLookupDescription(txtHeaderBranch, txtHeaderBranchDesc);
            _xBranchID = Convert.ToString(txtHeaderBranch.EditValue);
            BindOutletLookup();
            RefreshAllSubstitutionFlags();
        }

        private void BindOutletLookup()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtSalesID.Text) || string.IsNullOrWhiteSpace(xEntityID) || string.IsNullOrWhiteSpace(xBranchID)) return;
                string sqlOutlet = "select cm_cust_code1 [Cust Code1], cm_cust_code2 [Cust Code2], cm_cust_name [Outlet Name], " +
                    "ISNULL(RTRIM(LTRIM(cm_bill_adrress1)), '') + ' ' + ISNULL(RTRIM(LTRIM(cm_bill_adrress2)), '') + ' ' + ISNULL(RTRIM(LTRIM(cm_bill_adrress3)), '') + ' ' + ISNULL(RTRIM(LTRIM(cm_bill_adrress4)), '') AS Address, " +
                    "gh_function_desc [Outlet Status], isnull(cm_payment_type, '') [Payment Type], isnull(cm_top_id, '') [TOP], isnull(cm_flag_blacklist, 'N') [BLACKLIST], isnull(cm_top_by_cust,'N') [TOPBYCUST], isnull(cm_fullfilment,'N') [FULLFILMENT], ISNULL(cm_delv_phone,'') [outlet_phone] " +
                    "from SO_CUST_MASTER WITH(NOLOCK) inner join TBL_SD_CUSTCOVER WITH(NOLOCK) on cm_cust_code1 = csc_cust_code1 and cm_cust_code2 = csc_cust_code2 and cm_entity = csc_entity and cm_branch = csc_branch " +
                    "left join (select gh_function_code, gh_function_desc from GS_GEN_HARDCODED WITH(NOLOCK) where gh_sys = 'H' and gh_function_name = 'OUTLETSTATUS') GS on cm_active_flag = gh_function_code " +
                    "where csc_salesman_id = '" + txtSalesID.Text + "' and cm_active_flag <> 'D' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' order by cm_cust_code1, cm_cust_code2";
                DataTable dtOutlet = (_clsGlobal.Connect.State == ConnectionState.Closed) ? _clsGlobal.ExecDT(sqlOutlet) : _clsGlobal.ExecDTTrans(sqlOutlet);
                txtKdOutlet1.Properties.DataSource = dtOutlet;
                txtKdOutlet1.Properties.DisplayMember = "Cust Code1";
                txtKdOutlet1.Properties.ValueMember = "Cust Code1";
                txtKdOutlet1View.Columns.Clear();
                txtKdOutlet1View.Columns.AddVisible("Cust Code1", "No Outlet");
                txtKdOutlet1View.Columns.AddVisible("Cust Code2", "Outlet 2");
                txtKdOutlet1View.Columns.AddVisible("Outlet Name", "Outlet Name");
                txtKdOutlet1View.Columns.AddVisible("Address", "Address");
                txtKdOutlet1View.Columns.AddVisible("Outlet Status", "Status");
                txtKdOutlet1View.BestFitColumns();

                txtHeaderShipTo.Properties.DataSource = dtOutlet;
                txtHeaderShipTo.Properties.DisplayMember = "Cust Code1";
                txtHeaderShipTo.Properties.ValueMember = "Cust Code1";
                txtHeaderShipToView.Columns.Clear();
                txtHeaderShipToView.Columns.AddVisible("Cust Code1", "No Outlet");
                txtHeaderShipToView.Columns.AddVisible("Cust Code2", "Outlet 2");
                txtHeaderShipToView.Columns.AddVisible("Outlet Name", "Outlet Name");
                txtHeaderShipToView.Columns.AddVisible("Address", "Address");
                txtHeaderShipToView.Columns.AddVisible("Outlet Status", "Status");
                txtHeaderShipToView.BestFitColumns();
                UpdateHeaderPartyFields();
            }
            catch { }
        }

        private void UpdateHeaderPartyFields()
        {
            _updatingHeaderParty = true;
            try
            {
                txtHeaderOrderNo.Text = txtNoOrder.Text;
                txtHeaderSoldTo.Text = txtKdOutlet1.Text;
                txtHeaderSoldToDesc.Text = lblOutletDesc.Text;
                txtHeaderShipTo.EditValue = string.IsNullOrWhiteSpace(txtKdOutlet1.Text) ? null : (object)txtKdOutlet1.Text;
                txtHeaderShipToDesc.Text = lblOutletDesc.Text;
                txtHeaderNetValue.Text = txtTtlInvoice.Text;
            }
            finally
            {
                _updatingHeaderParty = false;
            }
        }

        private void txtHeaderSource_TextChanged(object sender, EventArgs e)
        {
            UpdateHeaderPartyFields();
        }

        private void txtHeaderShipTo_ButtonClick(object sender, ButtonPressedEventArgs e)
        {
            if (e.Button.Kind == ButtonPredefines.Search)
                txtHeaderShipTo.ShowPopup();
        }

        private void txtHeaderShipTo_EditValueChanged(object sender, EventArgs e)
        {
            if (_updatingHeaderParty) return;
            try
            {
                DataRowView rowView = txtHeaderShipTo.Properties.View.GetFocusedRow() as DataRowView;
                DataRow row = rowView != null ? rowView.Row : null;

                if (row == null && txtHeaderShipTo.EditValue != null && txtHeaderShipTo.Properties.DataSource is DataTable dt)
                {
                    string val = txtHeaderShipTo.EditValue.ToString().Replace("'", "''");
                    DataRow[] rows = dt.Select("[Cust Code1] = '" + val + "'");
                    if (rows.Length > 0) row = rows[0];
                }

                if (row == null)
                {
                    txtHeaderShipToDesc.Text = string.Empty;
                    if (txtHeaderShipTo.EditValue == null || string.IsNullOrWhiteSpace(txtHeaderShipTo.Text))
                    {
                        txtKdOutlet1.EditValue = null;
                        txtKdOutlet1.Text = string.Empty;
                    }
                    return;
                }

                txtHeaderShipToDesc.Text = Convert.ToString(row["Outlet Name"]).Trim();
                txtKdOutlet1.EditValue = Convert.ToString(row["Cust Code1"]).Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void txtKdOutlet1_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                DataRowView rowView = txtKdOutlet1.Properties.View.GetFocusedRow() as DataRowView;
                DataRow row = rowView != null ? rowView.Row : null;

                if (row == null && txtKdOutlet1.EditValue != null && txtKdOutlet1.Properties.DataSource is DataTable dt)
                {
                    string val = txtKdOutlet1.EditValue.ToString().Replace("'", "''");
                    DataRow[] rows = dt.Select("[Cust Code1] = '" + val + "'");
                    if (rows.Length > 0) row = rows[0];
                }

                if (row == null)
                {
                    if (txtKdOutlet1.EditValue == null || string.IsNullOrWhiteSpace(txtKdOutlet1.Text))
                    {
                        lblOutletDesc.Text = string.Empty;
                        lblKdOutlet2.Text = string.Empty;
                        txtStatusOutlet.Text = string.Empty;
                        NO_HP_Outlet = string.Empty;
                        UpdateHeaderPartyFields();
                    }
                    return;
                }
                string oldOutlet = g_sOutlet;
                lblKdOutlet2.Text = Convert.ToString(row["Cust Code2"]).Trim();
                lblOutletDesc.Text = Convert.ToString(row["Outlet Name"]).Trim();
                txtStatusOutlet.Text = Convert.ToString(row["Outlet Status"]).Trim();
                NO_HP_Outlet = Convert.ToString(row["outlet_phone"]).Trim();
                cbfullfilment.Checked = Convert.ToString(row["FULLFILMENT"]).Trim().Equals("Y", StringComparison.OrdinalIgnoreCase);
                if (Convert.ToString(row["BLACKLIST"]).Trim().Equals("Y", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Outlet : " + txtKdOutlet1.Text + " ber Status BlackList" + "Proses tidak bisa dilanjutkan !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    pRefresh(true, false);
                    return;
                }
                if (string.IsNullOrWhiteSpace(Convert.ToString(row["Payment Type"])))
                    cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, 0);
                else
                    cbPembyFaktur.Text = Convert.ToString(row["Payment Type"]);
                if (!string.IsNullOrWhiteSpace(Convert.ToString(row["TOP"])))
                    cbTOP.Text = Convert.ToString(row["TOP"]);
                btnOk.Enabled = true;
                if (!string.IsNullOrWhiteSpace(oldOutlet) && oldOutlet != txtKdOutlet1.Text)
                {
                    pRefresh(true, false);
                    APIBlockingSKU = false;
                    if (dgvSalesDetail.Columns.Contains("wsd_prd_master_code")) dgvSalesDetail.Columns["wsd_prd_master_code"].ReadOnly = false;
                    if (dgvSalesDetail.Columns.Contains("wsd_real_order_xqty")) dgvSalesDetail.Columns["wsd_real_order_xqty"].ReadOnly = false;
                }
                initDateDelvExp(Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd"), txtKdOutlet1.Text, lblKdOutlet2.Text, 1);
                UpdateHeaderPartyFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnPopUpOutlet_Click(object sender, EventArgs e)
        {
            bool statusSAP = false;
            bool statusBlackList = false;
            g_bActiveEdit = true;
            string sOutlet = "";
            cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, 0);

            DataTable dtFill1 = new DataTable();
            strSQL = "select gh_function_name,gh_function_code from GS_GEN_HARDCODED WITH(NOLOCK) where gh_function_name = 'TIRASND_SAP'";
            if (_clsGlobal.Connect.State == ConnectionState.Closed)
            {
                dtFill1 = _clsGlobal.ExecDT(strSQL);
            }
            else
            {
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
            }

            if (dtFill1.Rows.Count > 0)
            {
                if (dtFill1.Rows[0]["gh_function_code"].ToString().Trim() == "Y")
                {
                    statusSAP = true;
                }
                else
                {
                    statusSAP = false;
                    //MessageBox.Show("TIRASND_SAP gh_function_code 'N'", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }
            }
            else
            {
                MessageBox.Show("Please define parameter for TIRASND SAP", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            sOutlet = txtKdOutlet1.Text;
            strSQL = "select cm_cust_code1 [Cust Code1], cm_cust_code2 [Cust Code2], ";
            strSQL = strSQL + " cm_cust_name [Outlet Name], ";
            strSQL = strSQL + " ISNULL(RTRIM(LTRIM(cm_bill_adrress1)), '') + ' ' + ISNULL(RTRIM(LTRIM(cm_bill_adrress2)), '') + ' ' + ISNULL(RTRIM(LTRIM(cm_bill_adrress3)), '') + ' ' + ISNULL(RTRIM(LTRIM(cm_bill_adrress4)), '') AS Address, ";
            strSQL = strSQL + " gh_function_desc [Outlet Status], ";
            strSQL = strSQL + " isnull(cm_payment_type, '') [Payment Type], ";
            strSQL = strSQL + " isnull(cm_top_id, '') [TOP],";
            strSQL = strSQL + " isnull(cm_flag_blacklist, 'N') [BLACKLIST], ";
            strSQL = strSQL + " isnull(cm_top_by_cust,'N') [TOPBYCUST], ";
            strSQL = strSQL + " isnull(cm_fullfilment,'N') [FULLFILMENT], ";
            strSQL = strSQL + " ISNULL(cm_delv_phone,'') [outlet_phone] ";
            strSQL = strSQL + " from SO_CUST_MASTER WITH(NOLOCK) ";
            strSQL = strSQL + " inner join TBL_SD_CUSTCOVER WITH(NOLOCK)  on ";
            strSQL = strSQL + " cm_cust_code1 = csc_cust_code1 ";
            strSQL = strSQL + " and cm_cust_code2 = csc_cust_code2 and cm_entity = csc_entity and cm_branch = csc_branch";
            strSQL = strSQL + " left join ( select gh_function_code, gh_function_desc";
            strSQL = strSQL + " from GS_GEN_HARDCODED WITH(NOLOCK) where gh_sys = 'H'";
            strSQL = strSQL + " and gh_function_name = 'OUTLETSTATUS') GS on ";
            strSQL = strSQL + " cm_active_flag = gh_function_code";
            strSQL = strSQL + " where csc_salesman_id = '" + txtSalesID.Text + "'";
            strSQL = strSQL + " and cm_active_flag <> 'D' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' ";
            strSQL = strSQL + " order by cm_cust_code1, cm_cust_code2";


            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Outlet";
            frm.Query = strSQL;
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtKdOutlet1.Text = frm.ArrField[0].Trim();
                lblOutletDesc.Text = frm.ArrField[2].Trim();
                lblKdOutlet2.Text = frm.ArrField[1].Trim();
                txtStatusOutlet.Text = frm.ArrField[3].Trim();
                cbfullfilment.Checked = false;
                NO_HP_Outlet = frm.ArrField[10].Trim();
                if (frm.ArrField[9].Trim().ToUpper().Equals("Y"))
                {
                    cbfullfilment.Checked = true;
                }

                if (sOutlet != txtKdOutlet1.Text && sOutlet != "")
                {
                    pRefresh(true, false);
                    APIBlockingSKU = false;
                    dgvSalesDetail.Columns["wsd_prd_master_code"].ReadOnly = false;
                    dgvSalesDetail.Columns["wsd_real_order_xqty"].ReadOnly = false;
                    //SetGridSODetail(dgvSalesDetail);
                    //SetGridTPRB(dgvPromosiQty);
                    //SetGridTPRU(dgvPromosiRp);
                    //SetGridDisc(dgvDisc);
                    //SetGridTax(dgvTax);
                    //SetGridPromoTax(dgvPromoTax);
                }

                txtKdOutlet1.Text = frm.ArrField[0].Trim();
                lblOutletDesc.Text = frm.ArrField[2].Trim();
                lblKdOutlet2.Text = frm.ArrField[1].Trim();
                txtStatusOutlet.Text = frm.ArrField[4].Trim();

                if (frm.ArrField[7].ToString().Trim() == "Y")
                {
                    MessageBox.Show("Outlet : " + frm.ArrField[0].Trim() + " ber Status BlackList" + "Proses tidak bisa dilanjutkan !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    pRefresh(true, false);
                    //SetGridSODetail(dgvSalesDetail);
                    //SetGridTPRB(dgvPromosiQty);
                    //SetGridTPRU(dgvPromosiRp);
                    //SetGridDisc(dgvDisc);
                    //SetGridTax(dgvTax);
                    //SetGridPromoTax(dgvPromoTax);
                }

                if (frm.ArrField[5].ToString().Trim() == "")
                {
                    cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, 0);
                }
                else
                {
                    g_bEdit = false;
                    for (lCounter = 0; lCounter < LookupRowCount(cbPembyFaktur); lCounter++)
                    {
                        cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, lCounter);
                        string cdPF = frm.ArrField[5].Trim();
                        string cbPemFak = Convert.ToString(cbPembyFaktur.EditValue);
                        if (cbPemFak == cdPF)
                        {
                            g_bEdit = true;

                            if (statusSAP == true)
                            {
                                cbPembyFaktur.Enabled = false;
                            }
                            break;
                        }
                    }

                    //cbPembyFaktur.ItemIndex = 0;

                    if (g_bEdit == false)
                    {
                        MessageBox.Show("Default tipe Pembayaran Faktur tidak ada untuk customer " + frm.ArrField[0].Trim(), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                }
                g_bChangeOutlet = true;
                #region DC
                if (_xFlagDC == "Y")
                {
                    DataTable dtFill2 = new DataTable();
                    strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH(NOLOCK) where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
                    dtFill2 = _clsGlobal.ExecDT(strSQL);

                    DataTable dtFillcheckTopByDivision = new DataTable();
                    strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH(NOLOCK) where gh_function_name='TOPBYDIVISION' and gh_sys='H' and gh_function_code='Y' ";
                    dtFillcheckTopByDivision = _clsGlobal.ExecDT(strSQL);

                    #region TOP BY GROUP SALESMAN
                    if (dtFill2.Rows.Count > 0)
                    {
                        DataTable dtFill3 = new DataTable();
                        strSQL = "select sgl_group_top from SO_SPG_GIRL_MAN WITH(NOLOCK) inner join SO_SPG_GROUP WITH(NOLOCK) on sgm_spgm_group = sgl_group_code ";
                        strSQL = strSQL + " where sgm_spgm_id='" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "' ";
                        dtFill3 = _clsGlobal.ExecDT(strSQL);

                        if (dtFill3.Rows.Count > 0)
                        {
                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string cdPF = dtFill3.Rows[0]["sgl_group_top"].ToString().Trim();
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (Convert.ToString(cbTOP.EditValue) == cdPF)
                                {
                                    cbTOP.Enabled = false;
                                    break;
                                }
                            }
                        }

                        string vPF = Convert.ToString(cbPembyFaktur.EditValue);
                        //jika tipe pembayaran T, TOP tidak berubah, 21-03-2019
                        if (vPF == "T")
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                        }

                        DataTable dtFill4 = new DataTable();
                        strSQL = "select cg_cust_top from SO_CUST_MASTER WITH (NOLOCK) inner join SO_CUST_GROUP WITH (NOLOCK)  on cm_cust_group = cg_cust_group ";
                        strSQL = strSQL + " where cg_cust_top is not null and cm_cust_code1='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'";
                        dtFill4 = _clsGlobal.ExecDT(strSQL);

                        //if (dtFill4.Rows.Count > 0)
                        //{
                        //    for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                        //    {
                        //        cbTOP.ItemIndex = lCounter;
                        //        string cdPF = dtFill4.Rows[0]["cg_cust_top"].ToString().Trim();
                        //        string vTOP = Convert.ToString(cbTOP.EditValue);
                        //        if (vTOP == cdPF)
                        //        {
                        //            break;
                        //        }
                        //    }
                        //}
                        if (dtFill4.Rows.Count > 0)
                        {
                            string __top = (string)(from r in ((DataTable)cbTOP.Properties.DataSource).Rows.Cast<DataRow>()
                                                    where r["code"].ToString() == dtFill4.Rows[0]["cg_cust_top"].ToString().Trim()
                                                    select r["code"].ToString())
                                    .ToList()
                                    .FirstOrDefault();
                            if (!__top.IsNullOrEmptyOrWhiteSpace())
                                cbTOP.EditValue = __top;
                        }

                        if (frm.ArrField[8].ToString().Trim() == "Y")
                        {
                            DataTable dtFill5 = new DataTable();
                            strSQL = "select top 1 cm_top_id from SO_CUST_MASTER WITH (NOLOCK)  ";
                            strSQL = strSQL + " where cm_cust_code1='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'";
                            dtFill5 = _clsGlobal.ExecDT(strSQL);

                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string cdPF = dtFill5.Rows[0]["cm_top_id"].ToString().Trim();
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (vTOP == cdPF)
                                {
                                    break;
                                }
                            }
                        }

                        string vPF1 = Convert.ToString(cbPembyFaktur.EditValue);
                        //jika tipe pembayaran T, TOP tidak berubah, 21-03-2019
                        if (vPF1 == "T")
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                        }

                    }
                    #endregion

                    #region TOP BY DIVISION
                    else if (dtFillcheckTopByDivision.Rows.Count > 0)
                    {

                        string vPF1 = Convert.ToString(cbPembyFaktur.EditValue);
                        //jika tipe pembayaran T, TOP tidak berubah, 21-03-2019

                        if (vPF1 == "T")
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                        }
                        else if (frm.ArrField[8].ToString().Trim() == "Y")
                        {
                            DataTable dtFill5 = new DataTable();
                            strSQL = "select top 1 cm_top_id from SO_CUST_MASTER WITH (NOLOCK)  ";
                            strSQL = strSQL + " where cm_cust_code1='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'";
                            dtFill5 = _clsGlobal.ExecDT(strSQL);

                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string cdPF = dtFill5.Rows[0]["cm_top_id"].ToString().Trim();
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (vTOP == cdPF)
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            DataTable dtFill3 = new DataTable();
                            if (dgvSalesDetail.Rows.Count > 0)
                            {
                                string product = "";
                                for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                                {
                                    if (string.IsNullOrEmpty(product))
                                    {
                                        if (!string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim()))
                                        {
                                            product = "'" + dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim() + "'";
                                        }

                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim()))
                                        {
                                            product += ",'" + dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim() + "'";
                                        }
                                    }


                                }

                                //strSQL = "select top 1 SMT.smt_top_id,PTC.pptc_no_of_days ";
                                //strSQL += "   from SO_CUST_MASTER CM WITH (NOLOCK)";
                                //strSQL += "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                                //strSQL += "       ON  ";
                                //strSQL += "           TSC.csc_entity          = CM.cm_entity ";
                                //strSQL += "           AND TSC.csc_branch      = CM.cm_branch ";
                                //strSQL += "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 ";
                                //strSQL += "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 ";
                                //strSQL += "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK)";
                                //strSQL += "       ON	";
                                //strSQL += "           SPG.sgm_spgm_id  = TSC.csc_salesman_id ";
                                //strSQL += "           AND SPG.sgm_entity_id = TSC.csc_entity ";
                                //strSQL += "           AND SPG.sgm_branch_id = TSC.csc_branch ";
                                //strSQL += "   left join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK)";
                                //strSQL += "       ON ";
                                //strSQL += "           SMT.smt_entity_id      = CM.cm_entity ";
                                //if (!isMultibranch)
                                //{
                                //    strSQL += "           and SMT.smt_branch_id  = CM.cm_branch ";
                                //    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                //    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                //}
                                //else
                                //{
                                //    strSQL += "           and SMT.smt_branch_id  = CM.cm_cust_bill_branch ";
                                //    strSQL += "           and SMT.smt_cust_code1 = CM.cm_bill_to_code1 ";
                                //    strSQL += "           and SMT.smt_cust_code2 = CM.cm_bill_to_code2 ";
                                //}
                                //strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                                //strSQL += "       ON ";
                                //strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                                //strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                //strSQL += "       ON ";
                                //strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                //strSQL += "   where ";
                                //strSQL += "   prm_prd_master_code in (" + product + ") ";
                                //strSQL += "   and SPG.sgm_spgm_id = '" + txtSalesID.Text + "' ";
                                //strSQL += "   and SPG.sgm_entity_id = '" + xEntityID + "' ";
                                //strSQL += "   and SPG.sgm_branch_id = '" + xBranchID + "' ";
                                //strSQL += "   and CM.cm_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                                //strSQL += "   and CM.cm_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                                //strSQL += "   order by pptc_no_of_days asc ";

                                strSQL = "select top 1 SMT.smt_top_id,PTC.pptc_no_of_days ";
                                strSQL += "   from SO_CUST_MASTER CM WITH (NOLOCK)";
                                strSQL += "   left join TBL_SD_CUSTCOVER TSC WITH (NOLOCK) ";
                                strSQL += "       ON  ";
                                strSQL += "           TSC.csc_entity          = CM.cm_entity ";
                                strSQL += "           AND TSC.csc_branch      = CM.cm_branch ";
                                strSQL += "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 ";
                                strSQL += "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 ";
                                strSQL += "   left join SO_SPG_GIRL_MAN SPG WITH (NOLOCK)";
                                strSQL += "       ON	";
                                strSQL += "           SPG.sgm_spgm_id  = TSC.csc_salesman_id ";
                                strSQL += "           AND SPG.sgm_entity_id = TSC.csc_entity ";
                                strSQL += "           AND SPG.sgm_branch_id = TSC.csc_branch ";
                                strSQL += "   left join SO_MAPPING_TOPBYPRDLINE SMT WITH (NOLOCK)";
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
                                strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                                strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                strSQL += "   where ";
                                strSQL += "   prm_prd_master_code in (" + product + ") ";
                                strSQL += "   and SPG.sgm_spgm_id = '" + txtSalesID.Text + "' ";
                                strSQL += "   and SPG.sgm_entity_id = '" + xEntityID + "' ";
                                strSQL += "   and SPG.sgm_branch_id = '" + xBranchID + "' ";
                                strSQL += "   and CM.cm_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                                strSQL += "   and CM.cm_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                                strSQL += "   order by pptc_no_of_days asc ";
                            }
                            else
                            {
                                strSQL = "Select cm_bill_to_code1,cm_bill_to_code2";
                                if (isMultibranch)
                                {
                                    strSQL += ",cm_cust_bill_branch ";
                                }
                                strSQL += " From SO_CUST_MASTER WITH(NOLOCK) where cm_cust_code1 = '" + frm.ArrField[0].Trim() + "' and cm_cust_code2 = '" + frm.ArrField[1].Trim() + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' ";
                                DataTable dtCustMaster = _clsGlobal.ExecDT(strSQL);
                                string billtocode1 = "";
                                string billtocode2 = "";
                                string billtobranch = "";

                                if (dtCustMaster.Rows.Count > 0)
                                {
                                    billtocode1 = dtCustMaster.Rows[0]["cm_bill_to_code1"].ToString();
                                    billtocode2 = dtCustMaster.Rows[0]["cm_bill_to_code2"].ToString();
                                    if (isMultibranch)
                                    {
                                        billtobranch = dtCustMaster.Rows[0]["cm_cust_bill_branch"].ToString();
                                    }
                                }

                                //strSQL = "select top 1 smt_top_id from SO_MAPPING_TOPBYPRDLINE WITH(NOLOCK) ";
                                //strSQL = strSQL + " where smt_cust_code1 ='" + frm.ArrField[0].Trim() + "' and smt_cust_code2 = '" + frm.ArrField[1].Trim() + "' and smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + xBranchID + "' ";

                                strSQL = "select top 1 smt_top_id from SO_MAPPING_TOPBYPRDLINE WITH(NOLOCK) ";
                                if (!isMultibranch)
                                {
                                    strSQL = strSQL + " where smt_cust_code1 ='" + billtocode1 + "' and smt_cust_code2 = '" + billtocode2 + "' and smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + xBranchID + "' ";
                                }
                                else
                                {
                                    strSQL = strSQL + " where smt_cust_code1 ='" + billtocode1 + "' and smt_cust_code2 = '" + billtocode2 + "' and smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + billtobranch + "' ";
                                }
                            }

                            dtFill3 = _clsGlobal.ExecDT(strSQL);

                            if (dtFill3.Rows.Count > 0)
                            {
                                for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                                {
                                    cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                    string cdPF = dtFill3.Rows[0]["smt_top_id"].ToString().Trim();
                                    string vTOP = Convert.ToString(cbTOP.EditValue);
                                    if (Convert.ToString(cbTOP.EditValue) == cdPF)
                                    {
                                        cbTOP.Enabled = false;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("TOP Belum Disetting", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                txtKdOutlet1.Text = "";
                                lblOutletDesc.Text = "";
                                lblKdOutlet2.Text = "";
                                txtStatusOutlet.Text = "";
                                cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                            }
                        }

                    }
                    #endregion

                    #region TOP BY CUSTOMER
                    else
                    {
                        if (frm.ArrField[6].ToString().Trim() == "")
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                        }
                        else
                        {
                            g_bEdit = false;
                            //jika top by cust
                            if (frm.ArrField[8].ToString().Trim() == "Y")
                            {
                                for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                                {
                                    cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                    string __top = frm.ArrField[6].ToString().Trim();
                                    if (Convert.ToString(cbTOP.EditValue) == __top)
                                    {
                                        g_bEdit = true;

                                        if (statusSAP == true)
                                        {
                                            string vPF1 = Convert.ToString(cbPembyFaktur.EditValue);
                                            if (vPF1 == "T")
                                            {
                                                cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                                            }
                                        }
                                        break;
                                    }
                                }

                                if (g_bEdit == false)
                                {
                                    MessageBox.Show("Default TOP tidak ada untuk customer " + frm.ArrField[0].Trim(), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    txtKdOutlet1.Text = "";
                                    lblOutletDesc.Text = "";
                                    lblKdOutlet2.Text = "";
                                    txtStatusOutlet.Text = "";
                                    cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                                }
                            }

                        }
                    }
                    #endregion
                }
                #endregion

                #region SELAIN DC
                else
                {

                    ////jika top by prdline harusnya tidak set default TOP
                    DataTable dtFill6 = new DataTable();
                    strSQL = "select sgl_group_top from SO_SPG_GIRL_MAN WITH (NOLOCK)  left join SO_SPG_GROUP on sgm_spgm_group = sgl_group_code ";
                    strSQL = strSQL + " where sgm_spgm_id='" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                    dtFill6 = _clsGlobal.ExecDT(strSQL);

                    if (dtFill6.Rows.Count > 0)
                    {
                        for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                            string vTOP = Convert.ToString(cbTOP.EditValue);
                            if (vTOP == frm.ArrField[6].ToString().Trim())
                            {
                                cbTOP.Enabled = false;
                                break;
                            }
                        }
                    }

                    DataTable dtFillcheckTopByDivision = new DataTable();
                    strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH(NOLOCK) where gh_function_name='TOPBYDIVISION' and gh_sys='H' and gh_function_code='Y' ";
                    dtFillcheckTopByDivision = _clsGlobal.ExecDT(strSQL);

                    //KALAU SUBDIST CUMA LIHAT TUNAI ATAU KREDIT BY CUSTOMER ,,,,GA ADA NGELIHAT KE PRDLINE
                    if (Convert.ToString(cbPembyFaktur.EditValue) == "T")
                    {
                        cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                    }
                    else if (dtFillcheckTopByDivision.Rows.Count > 0)
                    {
                        string vPF1 = Convert.ToString(cbPembyFaktur.EditValue);
                        //jika tipe pembayaran T, TOP tidak berubah, 21-03-2019

                        if (vPF1 == "T")
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                        }
                        else if (frm.ArrField[8].ToString().Trim() == "Y")
                        {
                            DataTable dtFill5 = new DataTable();
                            strSQL = "select top 1 cm_top_id from SO_CUST_MASTER WITH (NOLOCK)  ";
                            strSQL = strSQL + " where cm_cust_code1='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'";
                            dtFill5 = _clsGlobal.ExecDT(strSQL);

                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string cdPF = dtFill5.Rows[0]["cm_top_id"].ToString().Trim();
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (vTOP == cdPF)
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            DataTable dtFill3 = new DataTable();
                            strSQL = "select top 1 smt_top_id from SO_MAPPING_TOPBYPRDLINE WITH(NOLOCK) ";
                            strSQL = strSQL + " where smt_cust_code1 ='" + frm.ArrField[0].Trim() + "' and smt_cust_code2 = '" + frm.ArrField[1].Trim() + "' and smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + xBranchID + "' ";
                            dtFill3 = _clsGlobal.ExecDT(strSQL);

                            if (dtFill3.Rows.Count > 0)
                            {
                                for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                                {
                                    cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                    string cdPF = dtFill3.Rows[0]["smt_top_id"].ToString().Trim();
                                    string vTOP = Convert.ToString(cbTOP.EditValue);
                                    if (Convert.ToString(cbTOP.EditValue) == cdPF)
                                    {
                                        cbTOP.Enabled = false;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("TOP Belum Disetting " + frm.ArrField[0].Trim(), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                txtKdOutlet1.Text = "";
                                lblOutletDesc.Text = "";
                                lblKdOutlet2.Text = "";
                                txtStatusOutlet.Text = "";
                                cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                            }
                        }
                    }

                    //DataTable dtFill7 = new DataTable();
                    //strSQL = "select cg_cust_top from SO_CUST_MASTER WITH (NOLOCK)  inner join SO_CUST_GROUP WITH (NOLOCK)   on cm_cust_group = cg_cust_group ";
                    //strSQL = strSQL + " where cg_cust_top is not null and cm_cust_code1='" + txtKdOutlet1.Text + "'";
                    //dtFill7 = _clsGlobal.ExecDT(strSQL);

                    //if (dtFill7.Rows.Count > 0)
                    //{
                    //    for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                    //    {
                    //        cbTOP.ItemIndex = lCounter;
                    //        string vTOP = Convert.ToString(cbTOP.EditValue);
                    //        if (vTOP == dtFill7.Rows[0]["cg_cust_top"].ToString().Trim())
                    //        {
                    //            break;
                    //        }
                    //    }
                    //}
                    ////ibnu 7 maret 2017 TOP by CUst
                    //string vPF1 = Convert.ToString(cbPembyFaktur.EditValue);
                    //if (vPF1 == "T")
                    //{
                    //    cbTOP.ItemIndex = 0;
                    //}
                }
                #endregion

                g_bChangeOutlet = false;

            }
            //added 26 12 2018
            btnOk.Enabled = true;

            initDateDelvExp(Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd"), txtKdOutlet1.Text, lblKdOutlet2.Text, 1);
        }

        private void txtKdOutlet1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4)
            {
                btnPopUpOutlet_Click(sender, e);
            }
        }


        private sealed class CalcDiscBackupState
        {
            public DataTable SalesDetail;
            public string KdOutlet1;
            public string OutletDesc;
            public string KdOutlet2;
            public string NoPajak;
            public string StatusOutlet;
            public string NoPO;
            public string RefNoRetur;
            public string Note;
            public string SalesID;
            public string SalesDesc;
            public string TipeSales;
            public string WHLoc1;
            public string WHLoc2;
            public string TglTrans;
            public string Employee;
            public object Pembayaran;
            public object TOP;
            public object TipeOrder;
        }

        private CalcDiscBackupState CreateCalcDiscBackupState()
        {
            CalcDiscBackupState state = new CalcDiscBackupState();

            try
            {
                DataTable source = dgvSalesDetail.DataSource as DataTable;
                if (source == null)
                    source = dtGridSODetail;

                if (source != null)
                    state.SalesDetail = source.Copy();

                state.KdOutlet1 = txtKdOutlet1.Text;
                state.OutletDesc = lblOutletDesc.Text;
                state.KdOutlet2 = lblKdOutlet2.Text;
                state.NoPajak = txtNoPajak.Text;
                state.StatusOutlet = txtStatusOutlet.Text;
                state.NoPO = txtNoPO.Text;
                state.RefNoRetur = txtRefNoRetur.Text;
                state.Note = txtNote.Text;
                state.SalesID = txtSalesID.Text;
                state.SalesDesc = txtSalesDesc.Text;
                state.TipeSales = lblTipeSales.Text;
                state.WHLoc1 = lblWHLoc1.Text;
                state.WHLoc2 = lblWHLoc2.Text;
                state.TglTrans = txtTglTrans.Text;
                state.Employee = lblEmployee.Text;
                state.Pembayaran = cbPembyFaktur.EditValue;
                state.TOP = cbTOP.EditValue;
                state.TipeOrder = cbTipeOrder.EditValue;
            }
            catch
            {
                // Backup hanya safety net; jangan gagalkan proses utama.
            }

            return state;
        }

        private DataTable CloneSalesDetailBeforeCalcDisc()
        {
            CalcDiscBackupState state = CreateCalcDiscBackupState();
            return state == null ? null : state.SalesDetail;
        }

        private bool IsOutletBlankAfterCalcDisc()
        {
            return string.IsNullOrWhiteSpace(txtKdOutlet1.Text)
                && string.IsNullOrWhiteSpace(lblKdOutlet2.Text)
                && string.IsNullOrWhiteSpace(lblOutletDesc.Text);
        }

        private void RestoreOutletAfterCalcDiscIfBlank(CalcDiscBackupState state)
        {
            try
            {
                if (state == null || !IsOutletBlankAfterCalcDisc())
                    return;

                txtKdOutlet1.Text = state.KdOutlet1;
                lblOutletDesc.Text = state.OutletDesc;
                lblKdOutlet2.Text = state.KdOutlet2;
                txtNoPajak.Text = state.NoPajak;
                txtStatusOutlet.Text = state.StatusOutlet;
                txtNoPO.Text = state.NoPO;
                txtRefNoRetur.Text = state.RefNoRetur;
                txtNote.Text = state.Note;
                txtSalesID.Text = state.SalesID;
                txtSalesDesc.Text = state.SalesDesc;
                lblTipeSales.Text = state.TipeSales;
                lblWHLoc1.Text = state.WHLoc1;
                lblWHLoc2.Text = state.WHLoc2;
                txtTglTrans.Text = state.TglTrans;
                lblEmployee.Text = state.Employee;

                if (state.Pembayaran != null) cbPembyFaktur.EditValue = state.Pembayaran;
                if (state.TOP != null) cbTOP.EditValue = state.TOP;
                if (state.TipeOrder != null) cbTipeOrder.EditValue = state.TipeOrder;
            }
            catch
            {
                // Safety only.
            }
        }

        private void RestoreCalcDiscStateIfRefreshFailed(CalcDiscBackupState state)
        {
            if (state == null)
                return;

            RestoreSalesDetailProductIfBlankAfterCalcDisc(state.SalesDetail);
            RestoreOutletAfterCalcDiscIfBlank(state);
        }

        private void RestoreSalesDetailProductIfBlankAfterCalcDisc(DataTable beforeCalc)
        {
            try
            {
                if (beforeCalc == null || beforeCalc.Rows.Count == 0 || !beforeCalc.Columns.Contains("wsd_prd_master_code"))
                    return;

                bool beforeHasProduct = false;
                for (int i = 0; i < beforeCalc.Rows.Count; i++)
                {
                    if (!string.IsNullOrWhiteSpace(Convert.ToString(beforeCalc.Rows[i]["wsd_prd_master_code"])))
                    {
                        beforeHasProduct = true;
                        break;
                    }
                }

                if (!beforeHasProduct)
                    return;

                DataTable current = dgvSalesDetail.DataSource as DataTable;
                bool needFullRestore = false;

                if (current == null || current.Rows.Count == 0 || !current.Columns.Contains("wsd_prd_master_code"))
                {
                    needFullRestore = true;
                }
                else
                {
                    bool currentHasProduct = false;
                    for (int i = 0; i < current.Rows.Count; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(Convert.ToString(current.Rows[i]["wsd_prd_master_code"])))
                        {
                            currentHasProduct = true;
                            break;
                        }
                    }
                    needFullRestore = !currentHasProduct;
                }

                if (needFullRestore)
                {
                    dtGridSODetail.Rows.Clear();
                    foreach (DataRow src in beforeCalc.Rows)
                    {
                        DataRow dst = dtGridSODetail.NewRow();
                        foreach (DataColumn col in dtGridSODetail.Columns)
                        {
                            if (beforeCalc.Columns.Contains(col.ColumnName))
                                dst[col.ColumnName] = src[col.ColumnName];
                        }
                        dtGridSODetail.Rows.Add(dst);
                    }

                    dgvSalesDetail.DataSource = dtGridSODetail;
                    SetGridSODetailCss(dgvSalesDetail);
                    return;
                }

                int max = Math.Min(current.Rows.Count, beforeCalc.Rows.Count);
                string[] identityColumns = new string[]
                {
                    "wsd_prd_master_code",
                    "wsd_grade",
                    "wsd_prd_size",
                    "prm_prd_desc",
                    "wsd_uom_convert_big",
                    "wsd_uom_convert_mid",
                    "tax_code",
                    "pct_tax",
                    "wsd_prd_line_code",
                    "wsd_sled",
                    "wsd_req_order_xqty",
                    "wsd_real_order_xqty",
                    "wsd_so_xqty",
                    "tgl_price",
                    "reason"
                };

                for (int i = 0; i < max; i++)
                {
                    string currentPCode = current.Columns.Contains("wsd_prd_master_code")
                        ? Convert.ToString(current.Rows[i]["wsd_prd_master_code"])
                        : string.Empty;

                    string beforePCode = beforeCalc.Columns.Contains("wsd_prd_master_code")
                        ? Convert.ToString(beforeCalc.Rows[i]["wsd_prd_master_code"])
                        : string.Empty;

                    if (!string.IsNullOrWhiteSpace(currentPCode) || string.IsNullOrWhiteSpace(beforePCode))
                        continue;

                    foreach (string colName in identityColumns)
                    {
                        if (current.Columns.Contains(colName) && beforeCalc.Columns.Contains(colName))
                            current.Rows[i][colName] = beforeCalc.Rows[i][colName];
                    }
                }
            }
            catch
            {
                // Safety only: jangan gagalkan proses Calc Disc hanya karena restore tampilan gagal.
            }
        }

        private void btnCalcDisc_Click(object sender, EventArgs e)
        {
            try
            {


                string _pcode;
                string _grade;
                string _size;
                string _sQty;
                string _conv1;
                string _conv2;

                string sErrDesc = "";
                //double lAFSQty = 0;
                btnCalcDisc.Enabled = false;

                if (txtCashDiscP.Text == "")
                {
                    txtCashDiscP.Text = "0";
                }

                if (Convert.ToDecimal(txtCashDiscP.Text) > 100)
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Cash Disc yang di input lebih dari 100 %, Apakah mau di lanjutkan ?", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.No)
                    {
                        btnCalcDisc.Enabled = true;
                        return;
                    }
                }

                bool chkM = fCekMandatory();
                if (chkM == false)
                {
                    MarkMandatoryValidationErrors();
                    btnCalcDisc.Enabled = true;
                    return;
                }


                //BEGIN OF COUNTING AFS

                if (dgvSalesDetail.Rows.Count > 0)
                {
                    int rowindex = dgvSalesDetail.CurrentCell.RowIndex;
                    string g_PCodee = dgvSalesDetail.Rows[rowindex].Cells["wsd_prd_master_code"].Value.ToString().Trim();

                    if (g_PCodee != "")
                    {
                        _clsGlobal.BeginTrans();
                        for (int lCounter1 = 0; lCounter1 < dgvSalesDetail.Rows.Count; lCounter1++)
                        {
                            _pcode = dgvSalesDetail.Rows[lCounter1].Cells["wsd_prd_master_code"].Value.ToString().Trim();
                            _grade = dgvSalesDetail.Rows[lCounter1].Cells["wsd_grade"].Value.ToString().Trim();
                            _sQty = dgvSalesDetail.Rows[lCounter1].Cells["wsd_real_order_xqty"].Value.ToString().Trim();
                            _size = dgvSalesDetail.Rows[lCounter1].Cells["wsd_prd_size"].Value.ToString().Trim();
                            _conv1 = dgvSalesDetail.Rows[lCounter1].Cells["wsd_uom_convert_big"].Value.ToString().Trim();
                            _conv2 = dgvSalesDetail.Rows[lCounter1].Cells["wsd_uom_convert_mid"].Value.ToString().Trim();

                            bool _afs = fGetAFS(_pcode, _grade, _size, lblWHLoc1.Text, lblWHLoc2.Text);
                            if (_afs == true)
                            {
                                fConvertQty(_sQty, Convert.ToInt32(_conv1), Convert.ToInt32(_conv2));
                                double convqty = Convert.ToDouble(_fConvertQty);
                                if (!isNefoKAM)
                                {
                                    if (convqty > lAFSQty)
                                    {
                                        DialogResult dr;
                                        dr = MessageBox.Show("Stock PCode : " + _pcode + " tidak mencukupi, " + "Apakah mau di lanjutkan ?", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                                        if (dr == DialogResult.Cancel)
                                        {
                                            btnCalcDisc.Enabled = true;
                                            pRollbackProcess();
                                            return;
                                        }
                                        else
                                        {
                                            //1 juni 2018
                                            if (lAFSQty > 0)
                                            {
                                                dgvSalesDetail.Rows[lCounter1].Cells["wsd_so_xqty"].Value = fPC2Qty(lAFSQty, Convert.ToInt16(_conv1), Convert.ToInt16(_conv2));
                                            }
                                            else
                                            {
                                                dgvSalesDetail.Rows[lCounter1].Cells["wsd_so_xqty"].Value = fPC2Qty(0, Convert.ToInt16(_conv1), Convert.ToInt16(_conv2));
                                            }
                                        }
                                    }
                                }


                                //if (convqty > lAFSQty)
                                //{
                                //    DialogResult dr;
                                //    dr = MessageBox.Show("Stock PCode : " + _pcode + " tidak mencukupi, " + "Apakah mau di lanjutkan ?", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                                //    if (dr == DialogResult.Cancel)
                                //    {
                                //        btnCalcDisc.Enabled = true;
                                //        pRollbackProcess();
                                //        return;
                                //    }
                                //    else
                                //    {
                                //        //1 juni 2018
                                //        if (lAFSQty > 0)
                                //        {
                                //            dgvSalesDetail.Rows[lCounter1].Cells["wsd_so_xqty"].Value = fPC2Qty(lAFSQty, Convert.ToInt16(_conv1), Convert.ToInt16(_conv2));
                                //        }
                                //        else
                                //        {
                                //            dgvSalesDetail.Rows[lCounter1].Cells["wsd_so_xqty"].Value = fPC2Qty(0, Convert.ToInt16(_conv1), Convert.ToInt16(_conv2));
                                //        }
                                //    }
                                //}
                            }
                            else
                            {
                                pRollbackProcess();
                                MessageBox.Show("Stock PCode : " + _pcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                break;
                            }

                        }

                        _clsGlobal.CommitTrans();
                    }
                }

                // BEGIN OF CALCDISC PROCESS
                // Backup seluruh state input sebelum Calc Disc.
                // fcalcdisc() memanggil pRefresh(false, true) yang melakukan Clear pada grid dan outlet,
                // sehingga state dikembalikan jika refresh/pFillData gagal membaca ulang data.
                CalcDiscBackupState calcDiscBackup = CreateCalcDiscBackupState();
                bool calcDiscOk = fcalcdisc();
                RestoreCalcDiscStateIfRefreshFailed(calcDiscBackup);

                if (calcDiscOk == false)
                {
                    pClearDiscount();
                    pClearTPR();
                    pClearTax();
                    RestoreCalcDiscStateIfRefreshFailed(calcDiscBackup);
                }

            }
            catch (Exception ex)
            {
                if (_clsGlobal.tr != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnCalcDisc.Enabled = true;
            }


        }

        private void btnProses_Click(object sender, EventArgs e)
        {
            if (xFlagDC == "Y")
            {
                if (string.IsNullOrEmpty(strFlagLimit))
                {

                }
            }
        }

        private void cmdSales_Click(object sender, EventArgs e)
        {
            frmSOManualSalesman frm = new frmSOManualSalesman();
            //Form frm 
            frm.xSalesID = txtSalesID.Text;
            frm.xEdit = "1";
            frm.xFlagDCS = _xFlagDC;
            frm.ShowDialog();
            //this.Close();
        }

        private void cbTipeOrder_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (g_bSOType == true)
            {
                if (cbTipeOrder.ItemIndex > -1)
                {
                    bool cst = fChangeSOType();
                    if (cst == false)
                    {
                        return;
                    }
                }
                else
                {
                    lblWHLoc1.Text = "";
                    lblWHLoc2.Text = "";
                }
            }
        }

        private void cbTipeOrder_Click(object sender, EventArgs e)
        {
            //event cuma di selected index
            //if (g_bSOType == true)
            //{
            //    if (cbTipeOrder.ItemIndex > -1)
            //    {
            //        bool cst = fChangeSOType();
            //        if (cst == false)
            //        {
            //            return;
            //        }
            //    }
            //    else
            //    {
            //        lblWHLoc1.Text = "";
            //        lblWHLoc2.Text = "";
            //    }

            //    pRefresh(true);
            //}
        }

        private void cbPembyFaktur_Click(object sender, EventArgs e)
        {
            bool bExists;

            if (_xFlagDC == "Y")
            {
                if (cbPembyFaktur.Text == "")
                {
                    DataTable dtFill1 = new DataTable();
                    strSQL = "SELECT gh_group_menu_id FROM GS_GEN_HARDCODED WITH (NOLOCK)  " +
                        "WHERE gh_function_name ='CUSTPAYTYPE' AND gh_function_code = '" + cbPembyFaktur.EditValue + "'";
                    dtFill1 = _clsGlobal.ExecDT(strSQL);

                    if (dtFill1.Rows.Count > 0)
                    {
                        DataTable dtFill2 = new DataTable();
                        strSQL = "SELECT gh_group_menu_id FROM GS_GEN_HARDCODED WITH (NOLOCK)  " +
                            "WHERE gh_function_name ='CUSTPAYTYPE' AND gh_function_code = '" + cbPembyFaktur.EditValue + "'";
                        dtFill2 = _clsGlobal.ExecDT(strSQL);

                        if (dtFill2.Rows.Count > 0)
                        {

                        }
                        else
                        {
                            if (dtFill1.Rows[0]["gh_group_menu_id"].ToString().Trim() != "C")
                            {
                                bExists = false;
                                for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                                {
                                    cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);

                                    string tops = cbTOP.Text;
                                    string stops = ExtractTopDaysSafe(tops);
                                    if (stops == "0")
                                    {
                                        bExists = true;
                                        cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                                        break;
                                    }
                                }

                                if (bExists == false)
                                {
                                    MessageBox.Show("TOP 0 has not been set", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    btnOk.Enabled = false;
                                    btnProses.Enabled = false;
                                }

                                cbTOP.Enabled = false;
                                cbTOP.EditValue = LookupValueAt(cbTOP, 0);

                            }
                            else
                            {
                                cbTOP.Enabled = false;
                            }
                        }
                    }
                }
            }
            else
            {
                cbTOP.Enabled = false;
            }
        }

        //cbPembyFaktur_LostFocus
        private void cbPembyFaktur_Leave(object sender, EventArgs e)
        {
            if (cbPembyFaktur.EditValue == "T")
            {
                cbTOP.EditValue = "0";
            }
            else if (cbPembyFaktur.EditValue == "K")
            {
                if (txtSalesID.Text == "")
                {
                    MessageBox.Show("Pilih Salesman terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                if (txtKdOutlet1.Text == "")
                {
                    MessageBox.Show("Pilih Outlet Terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                DataTable dtFill1 = new DataTable();
                strSQL = "select cm_payment_type, coalesce(cg_cust_top,'') as topnya from SO_CUST_MASTER WITH (NOLOCK)  left join SO_CUST_GROUP WITH (NOLOCK)   on cm_cust_group = cg_cust_group " +
                    "where cm_cust_code1 ='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'";
                dtFill1 = _clsGlobal.ExecDT(strSQL);

                if (dtFill1.Rows.Count > 0)
                {
                    if (dtFill1.Rows[0]["topnya"].ToString().Trim() != "")
                    {
                        for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                            if (cbTOP.EditValue == dtFill1.Rows[0]["topnya"].ToString().Trim())
                            {
                                cbTOP.Enabled = false;
                            }
                        }
                    }
                    else
                    {
                        DataTable dtFill2 = new DataTable();
                        strSQL = "select coalesce(sgl_group_top,'') as topnya from SO_SPG_GIRL_MAN WITH (NOLOCK)  left join SO_SPG_GROU WITH (NOLOCK) P on sgm_spgm_group = sgl_group_code " +
                            "where sgm_spgm_id='" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + _xBranchID + "'";
                        dtFill2 = _clsGlobal.ExecDT(strSQL);

                        if (dtFill2.Rows.Count > 0)
                        {
                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                if (cbTOP.EditValue == dtFill2.Rows[0]["topnya"].ToString().Trim())
                                {
                                    cbTOP.Enabled = false;
                                }
                            }
                        }

                    }
                }

            }
        }

        private void cbPembyFaktur_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void cbTOP_Click(object sender, EventArgs e)
        {
            if (g_bChangeOutlet == false)
            {
                if (g_bTOPFlag == true)
                {
                    pRefreshDetail();
                }

            }
        }

        private void cbTOP_SelectedIndexChanged(object sender, EventArgs e)
        {
            int iIndex = 0;
            //long lYes;
            if (cbTOP.ItemIndex != 0)
            {
                if (g_bChangeOutlet == false)
                {
                    if (g_bTOPFlag == true)
                    {
                        iIndex = cbTOP.ItemIndex;
                        DialogResult dr;
                        dr = MessageBox.Show("Perubahan TOP akan mereset detail product pemebelian", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                        if (dr == DialogResult.OK)
                        {
                            pRefreshDetail();
                        }

                    }
                }
            }
        }

        //cbTOP Focus
        private void cbTOP_Enter(object sender, EventArgs e)
        {
            int iIndex = 0;
            //long lYes;

            if (g_bChangeOutlet == false)
            {
                if (g_bTOPFlag == true)
                {
                    iIndex = cbTOP.ItemIndex;
                    DialogResult dr;
                    dr = MessageBox.Show("Perubahan TOP akan mereset detail product pembelian, apakah akan di lanjutkan", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.OK)
                    {
                        pRefreshDetail();
                    }
                    else
                    {
                        txtKdOutlet1.Focus();
                    }

                }
            }
        }

        private void rbCashDiscP_CheckedChanged(object sender, EventArgs e)
        {
            rbCashDiscU.Checked = !rbCashDiscP.Checked;
            if (rbCashDiscP.Checked == true)
            {
                txtCashDiscP.Enabled = true;
                txtCashDiscU.Enabled = false;
                txtCashDiscU.Text = "0";
            }
            else
            {
                txtCashDiscP.Enabled = false;
                txtCashDiscU.Enabled = true;
                if (isNefoKAM)
                {
                    txtCashDiscP.Text = "0";

                }

            }
        }

        private void rbCashDiscU_CheckedChanged(object sender, EventArgs e)
        {
            rbCashDiscP.Checked = !rbCashDiscU.Checked;

            if (rbCashDiscP.Checked == true)
            {
                txtCashDiscP.Enabled = true;
                txtCashDiscU.Enabled = false;
                txtCashDiscU.Text = "0";
            }
            else
            {
                txtCashDiscP.Enabled = true;
                txtCashDiscU.Enabled = true;
                if (isNefoKAM)
                {
                    txtCashDiscP.Enabled = false;
                    txtCashDiscP.Text = "0";
                }

            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                string sCredLimit = "";
                btnOk.Enabled = false;

                if (isMultisource)
                {
                    if (cbFlagSloc.ItemIndex < 0)
                    {
                        DialogResult dRes;
                        dRes = MessageBox.Show("Flag SLOC harus di pilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        if (dRes == DialogResult.OK)
                        {
                            btnOk.Enabled = true;
                        }
                        return;
                    }
                }

                ClearFormValidationErrors();

                if (txtNoPO.Text == "")
                {
                    SetControlError(txtNoPO, "No PO harus di isi");
                    DialogResult dRes;
                    dRes = MessageBox.Show("No PO harus di isi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    if (dRes == DialogResult.OK)
                    {
                        btnOk.Enabled = true;
                    }
                    return;
                }

                if (txtCashDiscP.Text == "")
                {
                    txtCashDiscP.Text = "0";
                }
                Decimal _cdp = Convert.ToDecimal(txtCashDiscP.Text);

                if (_cdp > 100)
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Cash Disc yang di input lebih dari 100 % " + "Apakah mau di lanjutkan ?", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.OK)
                    {
                        btnOk.Enabled = true;
                        //break;
                    }
                }

                //Validasi untuk DearOTC 
                //cek outlet jika cm_dear_notification= Y nomor hp harus ada dan nomor salesman harus ada
                DataTable dtDearOTC = new DataTable();

                strSQL = "";
                strSQL = $@"select isnull(cm_dear_notification,'N') cm_dear_notification,ISNULL(cm_delv_phone,'') outlet_phone from SO_CUST_MASTER WITH (NOLOCK) 
                           WHERE isnull(cm_dear_notification,'N')='Y'
                           AND cm_entity='{xEntityID}' and cm_branch='{xBranchID}' and cm_cust_code1='{txtKdOutlet1.Text.ToString().Trim()}'";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                {
                    dtDearOTC = _clsGlobal.ExecDT(strSQL);
                }
                else
                {
                    dtDearOTC = _clsGlobal.ExecDTTrans(strSQL);
                }
                if (dtDearOTC.Rows.Count > 0)
                {
                    if (string.IsNullOrWhiteSpace(dtDearOTC.Rows[0]["outlet_phone"].ToString()))
                    {
                        DialogResult dRes;
                        dRes = MessageBox.Show("No. HP Outlet " + txtKdOutlet1.Text.ToString().Trim() + " Masih Kosong, Tolong dilengkapi karena akan dikirimkan Notifikasi WA", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        if (dRes == DialogResult.OK)
                        {
                            btnOk.Enabled = true;
                        }
                        return;
                    }
                    else
                    {
                        //Cek Nomor HP Salesman
                        DataTable dtDearOTCSalesman = new DataTable();
                        strSQL = "";
                        strSQL = $@"select ISNULL(sgm_hpno,'') as no_hp_salesman from SO_SPG_GIRL_MAN WITH (NOLOCK) 
                           WHERE sgm_entity_id='{xEntityID}' and sgm_branch_id='{xBranchID}' and sgm_spgm_id='{txtSalesID.Text.ToString().Trim()}'";
                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        {
                            dtDearOTCSalesman = _clsGlobal.ExecDT(strSQL);
                        }
                        else
                        {
                            dtDearOTCSalesman = _clsGlobal.ExecDTTrans(strSQL);
                        }
                        if (dtDearOTCSalesman.Rows.Count > 0)
                        {
                            if (string.IsNullOrWhiteSpace(dtDearOTCSalesman.Rows[0]["no_hp_salesman"].ToString()))
                            {
                                DialogResult dRes;
                                dRes = MessageBox.Show("No. HP Salesman " + txtSalesID.Text.ToString().Trim() + " Masih Kosong, Tolong dilengkapi karena akan dikirimkan Notifikasi WA", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                if (dRes == DialogResult.OK)
                                {
                                    btnOk.Enabled = true;
                                }
                                return;
                            }
                        }

                    }
                }

                //if (chkSLED.Checked)
                //{

                //    if (txtSLED.Text == "")
                //    {
                //        DialogResult dRes;
                //        dRes = MessageBox.Show("SLED Harus di isi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        if (dRes == DialogResult.OK)
                //        {
                //            btnOk.Enabled = true;
                //        }
                //        return;
                //    }
                //    else if (txtSLED.Text == "0")
                //    {
                //        DialogResult dRes;
                //        dRes = MessageBox.Show("SLED Harus lebih dari 0", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        if (dRes == DialogResult.OK)
                //        {
                //            btnOk.Enabled = true;
                //        }
                //        return;
                //    }
                //}

                _clsGlobal.BeginTrans();
                if (clsGlobal.MODE_TRX == 1)
                {
                    if (fCekMandatory() == false)
                    {
                        MarkMandatoryValidationErrors();
                        btnOk.Enabled = true;
                        return;
                    }
                    else
                    {

                        if (fSaveSO(sCredLimit, false) == false)
                        {
                            btnOk.Enabled = true;
                            //break;
                        }
                    }
                }
                else if (clsGlobal.MODE_TRX == 2)
                {
                    if (fCekMandatory() == false)
                    {
                        MarkMandatoryValidationErrors();
                        btnOk.Enabled = true;
                        return;
                    }
                    else
                    {

                        if (fSaveSO(sCredLimit, false) == false)
                        {
                            btnOk.Enabled = true;
                            //break;
                        }
                    }
                }

                if (sCredLimit != "*" && g_bReqProses == false) //g_bReqProses
                {
                    if (fProsesOrderType(sCredLimit, "1") == false)
                    {
                        btnOk.Enabled = true;
                        if (g_bProses == true)
                        {
                            btnProses.Enabled = true;
                            btnCalcDisc.Enabled = true;
                            g_bProses = false;
                        }
                    }
                }

                _clsGlobal.CommitTrans();

                if (sCredLimit == "*") //g_bReqProses
                {
                    MessageBox.Show("Data berhasil disimpan " + txtNoOrder.Text + " dengan status Over Kredit Limit" + " Untuk mem-proses order ini dapat di lakukan pada menu Approval Credit Limit ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Data berhasil disimpan " + txtNoOrder.Text + "", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (xFlagDC == "Y" && clsGlobal.MODE_TRX == 2)
                    {
                        this.Close();
                        return;
                    }
                }
                btnOk.Enabled = true;
                pRefresh(true, false);
                txtKdOutlet1.Enabled = true;
                //comented 26 12 2018
                //DataTable dtFill2 = new DataTable();
                ////strSQL = "select '0' as code, '' as gh_function_code union all Select gh_function_code as code,gh_function_code + ' ~ ' + gh_function_desc gh_function_code from GS_GEN_HARDCODED " +
                ////         "  where gh_function_name = 'CUSTPAYTYPE'"; //g_sPayType
                //strSQL = "select gh_function_code as code,gh_function_code + ' ~ ' + gh_function_name as name from GS_GEN_HARDCODED where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' "; //g_sPayType
                ////gh_function_code as code,gh_function_code + ' ~ ' + gh_function_name gh_function_code
                //dtFill2 = _clsGlobal.ExecDT(strSQL);
                //if (dtFill2.Rows.Count > 0)
                //{
                //    cbPembyFaktur.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                //    cbPembyFaktur.Properties.ValueMember = "code";
                //    cbPembyFaktur.Properties.DisplayMember = "name";
                //}

                txtNoPO.Enabled = true;
                cbTipeOrder.Enabled = true;
                cmdSales.Enabled = true;
                btnPopUpOutlet.Enabled = true;
                btnCalcDisc.Enabled = true;
                btnOk.Enabled = true;
                if (clsGlobal.MODE_TRX == 2)
                {
                    this.Close();
                }


            }
            catch (Exception ex)
            {
                if (_clsGlobal.tr != null && _clsGlobal.tr.Connection != null)
                {
                    _clsGlobal.RollbackTrans();
                }


                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void dgvSalesDetail_KeyDown(object sender, KeyEventArgs e)
        {
            if (APIBlockingSKU)
            {
                return;
            }
            //if (e.KeyCode == Keys.Insert)
            //{
            //    //int rowIndex = dgvSalesDetail.Rows.Add(dtGridSODetail);
            //    //var row = this.dgvSalesDetail.Rows[rowIndex];
            //    //row.Selected = true;
            //    //row.Cells[0].Selected = true;

            //    DataTable dt = dgvSalesDetail.DataSource as DataTable;
            //    dt.Rows.Add();
            //    //dgvSalesDetail.DataSource = dt;

            //}
            if (e.KeyCode == Keys.Delete)
            {
                foreach (DataGridViewRow row in dgvSalesDetail.SelectedRows)
                {
                    if (!row.IsNewRow && row.Index > -1)
                    {
                        dgvSalesDetail.Rows.Remove(row);
                        dtTglPrice.Visible = false;


                        if (dgvSalesDetail.Rows.Count > 1)
                        {
                            dgvSalesDetail.CurrentRow.Selected = true;
                        }
                        if (isTOPDivision == "Y")
                        {
                            TopTOPDivisionRenew();
                        }
                    }

                }
            }
            if (e.KeyCode == Keys.Enter)
            {
                string sTypeOrder = "";
                string sTypeSales = "";
                ////int rowIndex = dgvSalesDetail.Rows.Add(dtGridSODetail);
                //var row = this.dgvSalesDetail.Rows[dgvSalesDetail.CurrentCell.RowIndex];
                ////row.Selected = true;
                //row.Cells[1].Selected = true;
                int row = dgvSalesDetail.CurrentCell.RowIndex;
                if (dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Selected == true)
                {

                    //DataRow selectedDataRow = ((DataRowView)cbTipeOrder.GetSelectedDataRow()).Row;
                    int idO = Convert.ToInt32(Convert.ToString(cbTipeOrder.EditValue));
                    string NameO = cbTipeOrder.Text;

                    DataTable dtFill1 = new DataTable();
                    strSQL = " SELECT ot_def_oprtype FROM SO_ORDER_TYPE WITH (NOLOCK)  ";
                    strSQL = strSQL + " WHERE ot_order_type = '" + idO + "'";
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill1.Rows.Count > 0)
                    {
                        sTypeOrder = dtFill1.Rows[0]["ot_def_oprtype"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("Default operation untuk order type (" + idO + ") tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //break;
                    }

                    DataTable dtFill2 = new DataTable();
                    strSQL = " select sgm_type_operasi from SO_SPG_GIRL_MAN WITH (NOLOCK)  ";
                    strSQL = strSQL + " where sgm_spgm_id = '" + txtSalesID.Text + "' and  sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                    dtFill2 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill2.Rows.Count > 0)
                    {
                        sTypeSales = dtFill2.Rows[0]["sgm_type_operasi"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("Type operation untuk sales (" + txtSalesID.Text + ") tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //break;
                    }



                    DataTable dtFillC = new DataTable();
                    strSQL = "select prm_prd_master_code as PCode, prm_prd_desc as [Prd Name], ISNULL(pg_prd_group_desc, '') AS [PrdGroup] , ";
                    strSQL = strSQL + "prm_grade AS [Grade], prm_prd_size AS [Size], prm_conversion_purc [Conv1], prm_conversion_sales [Conv2], ";
                    if (isNefoKAM)
                    {
                        strSQL = strSQL + " case when cust_tax_code='PPN0' THEN cust_tax_code ELSE  stc_tax_code END [Tax Code], CASE WHEN cust_tax_code='PPN0' THEN isnull(cust_tax_amount,0) ELSE  isnull(stc_tax_amount, 0) END [Tax] ";

                    }
                    else
                    {
                        strSQL = strSQL + " prm_tax_code [Tax Code],case when isnull(excludeTax,'N') = 'Y' then 0 else isnull(stc_tax_amount, 0) end [Tax] ";
                    }
                    strSQL = strSQL + " ,prm_prd_line_code";

                    strSQL = strSQL + " from IM_PRD_MASTER WITH (NOLOCK)  ";
                    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
                    //if (isNefoKAM)
                    //{
                    //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON  CAST(CONVERT(DATE, '" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "', 112) AS DATETIME) >=cast(stc_tax_valid_from as datetime) ";
                    //    strSQL = strSQL + "  AND CAST(CONVERT(DATE,'" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "', 112) AS DATETIME) <=cast(stc_tax_valid_to as datetime)  ";
                    //}
                    //else
                    //{
                    //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
                    //}                    
                    strSQL = strSQL + " inner join IM_STOCK_BALANCE on ";
                    strSQL = strSQL + "  wlb_loc_Id1 = '" + lblWHLoc1.Text + "'";
                    strSQL = strSQL + "  and wlb_loc_Id2 = '" + lblWHLoc2.Text + "'";
                    strSQL = strSQL + "  and wlb_entity_id = '" + xEntityID + "'";
                    strSQL = strSQL + "  and wlb_branch_id = '" + xBranchID + "'";
                    strSQL = strSQL + "  and wlb_prd_master_code = prm_prd_master_code";
                    strSQL = strSQL + "  and wlb_grade = prm_grade";
                    strSQL = strSQL + "  and wlb_prd_size = prm_prd_size";
                    strSQL = strSQL + " and isnull(wlb_stock_sts, 'N') = 'N' ";
                    strSQL = strSQL + " inner join IM_PRD_GROUP ON ";
                    strSQL = strSQL + " prm_prd_line_code = pg_prd_line_code ";
                    strSQL = strSQL + " AND prm_prd_group_code = pg_prd_group_code ";
                    strSQL = strSQL + " inner join IM_PRD_LINE ON ";
                    strSQL = strSQL + " prm_prd_line_code = pl_prd_line_code ";
                    if (g_bTOPFlag == true)
                    {
                        DataRow selectedDataRow1 = ((DataRowView)cbTOP.GetSelectedDataRow()).Row;
                        string idO1 = Convert.ToString(selectedDataRow1["code"]);
                        //string NameO = selectedDataRow["order_type"].ToString();

                        strSQL = strSQL + " INNER JOIN ( ";
                        strSQL = strSQL + " SELECT  DISTINCT map_prdline_code ";
                        strSQL = strSQL + " FROM    SO_CUST_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " INNER JOIN TBL_SD_MAP_TOP_PRDLINE WITH (NOLOCK) ON ";
                        strSQL = strSQL + " cm_cust_group = map_cust_group ";
                        strSQL = strSQL + " INNER JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON ";
                        strSQL = strSQL + " pptc_term_code = map_term_code ";
                        strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + txtKdOutlet1.Text + "'";
                        strSQL = strSQL + "  AND cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                        strSQL = strSQL + "  AND pptc_term_code >= '" + idO1 + "'";
                        strSQL = strSQL + "  AND cm_entity = '" + xEntityID + "'";
                        strSQL = strSQL + "  AND cm_branch = '" + xBranchID + "'";
                        strSQL = strSQL + " ) TOPPRD ON prm_prd_line_code = map_prdline_code ";
                    }
                    if (sTypeOrder == "C" && sTypeSales == "C")
                    {
                        strSQL = strSQL + " inner join TBL_CVS_PROD_REG ";
                        strSQL = strSQL + " on cpr_sld_id = '" + txtSalesID.Text + "'";
                        strSQL = strSQL + " and cpr_product_id = prm_prd_master_code";
                        strSQL = strSQL + " and cpr_product_grade = prm_grade";
                        strSQL = strSQL + " and cpr_product_size = prm_prd_size";
                        strSQL = strSQL + " and cpr_wh_loc_id1 = '" + lblWHLoc1.Text + "'";
                        strSQL = strSQL + " and cpr_wh_loc_id2 = '" + lblWHLoc2.Text + "'";
                        strSQL = strSQL + " and cpr_entity_id = '" + xEntityID + "'";
                        strSQL = strSQL + " and cpr_branch_id = '" + xBranchID + "'";
                    }
                    strSQL = strSQL + " LEFT JOIN TBL_IM_MAP_PRD_CUST ";
                    strSQL = strSQL + " ON  impc_cust_code1 = '" + txtKdOutlet1.Text + "'";
                    strSQL = strSQL + "     AND impc_cust_code2 = '" + lblKdOutlet2.Text + "'";
                    strSQL = strSQL + "     AND impc_prd_code = prm_prd_master_code ";
                    strSQL = strSQL + "     AND impc_prd_grade = prm_prd_size ";
                    strSQL = strSQL + "     AND impc_prd_size = prm_prd_size ";
                    strSQL = strSQL + "     AND impc_del_flag = 'N' ";
                    strSQL = strSQL + "     AND impc_entity_id = '" + xEntityID + "' ";
                    strSQL = strSQL + "     AND impc_branch_id = '" + xBranchID + "' ";

                    strSQL = strSQL + " INNER JOIN SO_SPG_GIRL_MAN ";
                    strSQL = strSQL + " ON  sgm_spgm_id = '" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "' ";
                    strSQL = strSQL + " INNER JOIN TBL_SD_SPGGROUP_PRDLINE ";
                    strSQL = strSQL + " ON  sgm_spgm_group = ssp_spggroup_code ";
                    strSQL = strSQL + "     AND prm_prd_line_code = ssp_prdline_code ";
                    strSQL = strSQL + "     AND ssp_del_flag = 'N' ";

                    strSQL = strSQL + " LEFT JOIN ( ";
                    strSQL = strSQL + " select gh_function_code as cocline from GS_GEN_HARDCODED WITH (NOLOCK) ";
                    strSQL = strSQL + " where gh_function_name like 'LINE_COLD_CHAIN' and gh_sys = 'H' ";
                    strSQL = strSQL + " ) COC";
                    strSQL = strSQL + " ON  prm_prd_line_code = cocline ";


                    if (isNefoKAM)
                    {
                        /** customer PPN **/
                        strSQL = strSQL + " LEFT JOIN  (SELECT cm_entity,cm_branch, cm_ship_tax_id AS cust_tax_code , stc_tax_amount as cust_tax_amount FROM SO_CUST_MASTER WITH (NOLOCK)  ";
                        strSQL = strSQL + "  INNER JOIN dbo.SO_TAX_CODE  WITH (NOLOCK) ON cm_ship_tax_id = stc_tax_code  ";
                        strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + txtKdOutlet1.Text + "'";
                        strSQL = strSQL + "  AND cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                        strSQL = strSQL + "  AND cm_entity = '" + xEntityID + "'";
                        strSQL = strSQL + "  AND cm_branch = '" + xBranchID + "'";
                        strSQL = strSQL + " )CUST ON cm_entity = '" + _xEntityID + "' and cm_branch = '" + _xBranchID + "' ";
                        /** End customer PPN **/
                    }
                    else
                    {
                        /** exclude tax **/
                        strSQL = strSQL + " LEFT JOIN ( SELECT gh_group_menu_id as __entity,gh_parent_menu __branch,gh_function_code excludeTax FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE GH_FUNCTION_NAME = 'Exclude-TAX-KAM' AND GH_SYS = 'S' ) FLAGTAX ON __entity = '" + _xEntityID + "' and __branch = '" + _xBranchID + "' ";
                        /** end exclude tax **/
                    }

                    //strSQL = strSQL + " WHERE prm_status = 'A'";
                    strSQL = strSQL + " WHERE prm_status IN ('A','D')";
                    strSQL = strSQL + " AND ( prm_prd_cust_flag = 'N' OR ( prm_prd_cust_flag = 'Y' AND impc_cust_code1 IS NOT NULL ) )";
                    //if (isMultisource)
                    //{
                    //    if (cbFlagSloc.Text.ToUpper().Equals("COC"))
                    //    {
                    //        strSQL = strSQL + " AND cocline is not null";
                    //    }
                    //    else
                    //    {
                    //        strSQL = strSQL + " AND cocline is null";
                    //    }
                    //}


                    strSQL = strSQL + " order by prm_prd_master_code";


                    dtFillC = _clsGlobal.ExecDT(strSQL);
                    if (dtFillC.Rows.Count > 0)
                    {
                        if (pCheckValue(dtFillC.Rows[0].ToString(), dgvSalesDetail.Rows.Count, dgvSalesDetail) == true)
                        {
                            if (isNefoKAM)
                            {
                                if (string.IsNullOrEmpty(dtFillC.Rows[7].ToString()))
                                {
                                    MessageBox.Show("Parameter untuk so_tax_code valid date '" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "' tidak ditemukan!.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    return;
                                }

                            }


                            dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = dtFillC.Rows[0].ToString();
                            dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = dtFillC.Rows[3].ToString();
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = dtFillC.Rows[4].ToString();
                            dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = dtFillC.Rows[1].ToString();
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = dtFillC.Rows[5].ToString();
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = dtFillC.Rows[6].ToString();
                            dgvSalesDetail.Rows[row].Cells["tax_code"].Value = dtFillC.Rows[7].ToString();
                            dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = dtFillC.Rows[8].ToString();
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = dtFillC.Rows[9].ToString();
                            pChangePCode(row);
                        }
                        else
                        {
                            MessageBox.Show("Data sudah ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";

                            pChangePCode(row);
                        }

                        dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];
                        dgvSalesDetail.Rows[row].Cells["wsd_req_order_xqty"].Value = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];
                    }
                    else
                    {
                        MessageBox.Show("PCode tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }

            }
        }

        private void TopTOPDivisionRenew()
        {
            string top_by_customer = "N";
            strSQL = "Select cm_top_by_cust,cm_top_id from SO_CUST_MASTER WITH (NOLOCK) where cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' and cm_cust_code1 = '" + txtKdOutlet1.Text + "' and cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
            DataTable dtCheck = _clsGlobal.ExecDT(strSQL);
            if (dtCheck.Rows.Count > 0)
            {
                top_by_customer = dtCheck.Rows[0]["cm_top_by_cust"].ToString();
            }
            bool recal = true;
            if (Convert.ToString(cbPembyFaktur.EditValue) == "T")
                recal = false;
            if (top_by_customer == "Y")
            {
                recal = false;
            }

            if (recal)
            {
                if (dgvSalesDetail.Rows.Count > 0)
                {
                    string product = "";
                    for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                    {
                        if (string.IsNullOrEmpty(product))
                        {
                            if (!string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim()))
                            {
                                product = "'" + dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim() + "'";
                            }

                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim()))
                            {
                                product += ",'" + dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim() + "'";
                            }
                        }


                    }
                    strSQL = "select top 1 SMT.smt_top_id,PTC.pptc_no_of_days ";
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
                    strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                    strSQL += "       ON ";
                    strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                    strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                    strSQL += "       ON ";
                    strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                    strSQL += "   where ";
                    strSQL += "   prm_prd_master_code in (" + product + ") ";
                    strSQL += "   and SPG.sgm_spgm_id = '" + txtSalesID.Text + "' ";
                    strSQL += "   and SPG.sgm_entity_id = '" + xEntityID + "' ";
                    strSQL += "   and SPG.sgm_branch_id = '" + xBranchID + "' ";
                    strSQL += "   and CM.cm_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                    strSQL += "   and CM.cm_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                    strSQL += "   order by pptc_no_of_days asc ";
                    DataTable dtCheck2 = _clsGlobal.ExecDT(strSQL);
                    if (dtCheck2.Rows.Count > 0)
                    {
                        if (Convert.ToString(cbTOP.EditValue).Trim() != dtCheck2.Rows[0].ToString().Trim())
                        {
                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string cdPF = dtCheck2.Rows[0][0].ToString().Trim();
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (Convert.ToString(cbTOP.EditValue) == cdPF)
                                {
                                    cbTOP.Enabled = false;
                                    break;
                                }
                            }
                        }
                    }


                }


            }
        }

        private void IsBlockingAPI()
        {
            DataTable dtCheckIsAPI = new DataTable();
            strSQL = " select wsh_qasir_flag from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
            strSQL = strSQL + " WHERE wsh_seq_no = '" + txtNoOrder.Text.ToString().Trim() + "' ";
            dtCheckIsAPI = _clsGlobal.ExecDT(strSQL);
            if (dtCheckIsAPI.Rows.Count > 0)
            {
                if (dtCheckIsAPI.Rows[0]["wsh_qasir_flag"].ToString().Trim() == "S" || dtCheckIsAPI.Rows[0]["wsh_qasir_flag"].ToString().Trim() == "W")
                {
                    APIBlockingSKU = true;
                }
                else
                {
                    APIBlockingSKU = false;
                }
            }
        }


        private string GetGridColumnName(DataGridView grid, DataGridViewCellEventArgs e)
        {
            if (grid == null || e == null || e.ColumnIndex < 0)
                return string.Empty;

            try
            {
                DataTable dt = grid.DataSource as DataTable;
                if (dt != null && e.ColumnIndex < dt.Columns.Count)
                    return dt.Columns[e.ColumnIndex].ColumnName;
            }
            catch
            {
            }

            try
            {
                if (e.ColumnIndex < grid.Columns.Count)
                    return grid.Columns[e.ColumnIndex].Name;
            }
            catch
            {
            }

            return string.Empty;
        }

        private bool IsGridColumn(DataGridView grid, DataGridViewCellEventArgs e, string columnName)
        {
            return string.Equals(GetGridColumnName(grid, e), columnName, StringComparison.OrdinalIgnoreCase);
        }

        private void dgvSalesDetail_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (APIBlockingSKU)
            {
                return;
            }

            if (IsGridColumn(dgvSalesDetail, e, COL_TGL_PRICE))
            {
                dtTglPrice = new DevExpress.XtraEditors.DateEdit();
                //dtTglPrice.ShowCheckBox = true;

                if (isNefoKAM)
                {
                    //string pcode_code= dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_master_code"].Value.ToString();
                    //string pcode_line = dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_line_code"].Value.ToString();
                    if (e.RowIndex > -1 && txtKdOutlet1.Text != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_master_code"].Value.ToString() != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_line_code"].Value.ToString() != "")
                    {
                        DataTable dtPriceDate = new DataTable();
                        strSQL = "";
                        strSQL = strSQL + " SELECT smt_createat_sap from SO_MAPPING_TOPBYPRDLINE WITH (NOLOCK) ";
                        strSQL = strSQL + " where smt_entity_id = '" + xEntityID + "'";
                        strSQL = strSQL + "  and smt_branch_id = '" + xBranchID + "'";
                        strSQL = strSQL + "  and smt_cust_code1 = '" + txtKdOutlet1.Text.Trim() + "'";
                        strSQL = strSQL + "  and smt_cust_code2 = '" + lblKdOutlet2.Text.Trim() + "'";
                        strSQL = strSQL + "  and smt_prdline_id = '" + dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_line_code"].Value.ToString() + "'";

                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                            dtPriceDate = _clsGlobal.ExecDT(strSQL);
                        else
                            dtPriceDate = _clsGlobal.ExecDTTrans(strSQL);

                        if (dtPriceDate.Rows.Count > 0)
                        {
                            if (string.IsNullOrEmpty(dtPriceDate.Rows[0]["smt_createat_sap"].ToString()))
                            {
                                g_CreateAtSAP = g_dTglGudang;
                            }
                            else
                            {
                                g_CreateAtSAP = Convert.ToDateTime(dtPriceDate.Rows[0]["smt_createat_sap"]);
                            }

                        }
                        else
                        {
                            g_CreateAtSAP = g_dTglGudang;

                        }

                    }
                    else
                    {
                        g_CreateAtSAP = g_dTglGudang;
                    }

                    dtTglPrice.Properties.MinValue = g_CreateAtSAP;

                }

                if (clsGlobal.MODE_TRX == 2)
                {
                    dtTglPrice.EditValue = Convert.ToDateTime(dgvSalesDetail.Rows[e.RowIndex].Cells["tgl_price"].Value);
                }


                //Adding DateTimePicker control into DataGridView   
                dgvSalesDetail.Controls.Add(dtTglPrice);
                // Setting the format (i.e. 2014-10-10)  

                // It returns the retangular area that represents the Display area for a cell  
                Rectangle oRectangle = dgvSalesDetail.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                //Setting area for DateTimePicker Control  
                dtTglPrice.Size = new Size(oRectangle.Width, oRectangle.Height);

                // Setting Location  
                dtTglPrice.Location = new Point(oRectangle.X, oRectangle.Y);

                // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed  

                // An event attached to dateTimePicker Control which is fired when any date is selected  
                dtTglPrice.TextChanged += new EventHandler(dateTimePicker_OnTextChange);

                // Now make it visible  
                dtTglPrice.Visible = true;
                dgvSalesDetail.Columns["tgl_price"].ReadOnly = true;
            }

            if (IsGridColumn(dgvSalesDetail, e, COL_REASON))
            {
                EnsureReasonSearchLookUpEditor();
                RefreshReasonEditState(e.RowIndex);

                if (IsReasonAllowed(e.RowIndex))
                {
                    GridView view = GetGridView(dgvSalesDetail);
                    if (view != null)
                    {
                        GridColumn col = view.Columns.ColumnByFieldName("reason") ?? view.Columns["reason"];
                        view.FocusedRowHandle = e.RowIndex;
                        if (col != null) view.FocusedColumn = col;
                        view.ShowEditor();
                        SearchLookUpEdit editor = view.ActiveEditor as SearchLookUpEdit;
                        if (editor != null)
                            editor.ShowPopup();
                    }
                }
            }
        }
        private void cmbReasonGrid_CloseUp(object sender, EventArgs e)
        {
            cmbReasonGrid.Visible = false;
        }

        private void cmbReasonGrid_OnTextChange(object sender, EventArgs e)
        {
            if (dgvSalesDetail.CurrentRow == null) return;

            int i = dgvSalesDetail.CurrentRow.Index;
            string selectedReason = Convert.ToString(cmbReasonGrid.EditValue);

            if (!string.IsNullOrEmpty(selectedReason))
            {
                dgvSalesDetail["reason", i].Value = selectedReason;
            }
            else
            {
                dgvSalesDetail["reason", i].Value = "";
            }
        }


        private void dtTglPrice_CloseUp(object sender, EventArgs e)
        {
            dtTglPrice.Visible = false;
        }

        private void dateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            int i = dgvSalesDetail.CurrentRow.Index;
            if (dtTglPrice.EditValue != null)
            {
                dgvSalesDetail.Rows[i].Cells["tgl_price"].Value = Convert.ToDateTime(dtTglPrice.EditValue).ToString("dd MMM yyyy");
            }
            else
            {
                dgvSalesDetail.Rows[i].Cells["tgl_price"].Value = string.Empty;
            }


        }

        private void dtTglPrice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                //btnPopUpSalesCode_Click(sender, e);
                dtTglPrice.Visible = false;

            }
            else if (e.KeyCode == Keys.Return)
            {
                dtTglPrice.Visible = false;
            }
        }

        private void dtTglPrice_Leave(object sender, EventArgs e)
        {
            dtTglPrice.Visible = false;
        }

        private void dgvSalesDetail_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //if (dgvSalesDetail.CurrentCell.IsInEditMode)
            //{
            //    if (dgvSalesDetail.IsCurrentCellDirty)
            //    {
            //        dgvSalesDetail.EndEdit();
            //    }
            //}

            if (IsGridColumn(dgvSalesDetail, e, COL_PCODE) || IsGridColumn(dgvSalesDetail, e, "wsd_prd_line_code"))
            {
                if (e.RowIndex > -1)
                    RefreshSubstitutionFlag(e.RowIndex);
            }
            else if (IsGridColumn(dgvSalesDetail, e, COL_TGL_PRICE))
            {
                if (e.RowIndex > -1 && txtKdOutlet1.Text != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_master_code"].Value.ToString() != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString() != "")
                {
                    fQtyFormat(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_mid"].Value.ToString()));
                    string _cpcode = dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_master_code"].Value.ToString();
                    string _cgrade = dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_grade"].Value.ToString();
                    string _csize = dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_size"].Value.ToString();
                    string _ctglprice = dgvSalesDetail.Rows[e.RowIndex].Cells["tgl_price"].Value.ToString();
                    Int32 __qtyy = 0;
                    if (_ctglprice != "")
                    {

                        if (fPricePCode(_cpcode, _cgrade, _csize, txtKdOutlet1.Text, lblKdOutlet2.Text, _ctglprice) == true)
                        {
                            //decimal _iqty = Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString());
                            if (isNefoKAM)
                            {
                                dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value = curPrice.ToString("#,##0.0000");
                            }
                            else
                            {
                                dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value = curPrice.ToString("#,##0.00");
                            }

                            fConvertQty(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_mid"].Value.ToString()));
                            __qtyy = Convert.ToInt32(_fConvertQty);
                            //decimal _jmlhrg = __qtyy * Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value);
                            //if (isNefoKAM)
                            //{
                            // rounding mekanism

                            decimal _jmlhrg = __qtyy * Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value);
                            if (isRoundingMekanism)
                            {
                                _jmlhrg = Math.Round(__qtyy * Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value), 0, MidpointRounding.AwayFromZero);
                            }
                            //}
                            //dgvSalesDetail.Rows[row].Cells["wsd_total_so_amount"].Value = _jmlhrg.ToString("#,##0.00");//g_iJmlHarga
                            dgvSalesDetail.Rows[e.RowIndex].Cells["ijumlahharga"].Value = _jmlhrg.ToString("#,##0.00");//g_iJmlHarga

                            if (dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_tot_disc_pc"].Value.ToString() != "")
                            {
                                decimal _jmlrpnett = _jmlhrg - Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_tot_disc_pc"].Value);
                                dgvSalesDetail.Rows[e.RowIndex].Cells["jml_rp_netto"].Value = _jmlrpnett.ToString("#,##0.00");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Periode validate harga sudah habis atau harga barang belum di setting", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value = "";
                            dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_total_so_amount"].Value = "0";
                            dgvSalesDetail.Rows[e.RowIndex].Cells["jml_rp_netto"].Value = "0";
                        }
                    }
                }
            }
            else if (IsGridColumn(dgvSalesDetail, e, "wsd_tot_disc_pc"))
            {
                if (e.RowIndex > -1 && txtKdOutlet1.Text != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_master_code"].Value.ToString() != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString() != "")
                {
                    //decimal _jmlrpnett = Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_total_so_amount"].Value) - Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_tot_disc_pc"].Value);
                    decimal _jmlrpnett = Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["ijumlahharga"].Value) - Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_tot_disc_pc"].Value);
                    dgvSalesDetail.Rows[e.RowIndex].Cells["jml_rp_netto"].Value = Convert.ToDecimal(_jmlrpnett).ToString("#,##0.00");
                }
            }
            else if (IsGridColumn(dgvSalesDetail, e, "wsd_so_xqty"))
            {
                if (e.RowIndex > -1 && txtKdOutlet1.Text != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_master_code"].Value.ToString() != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString() != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value.ToString() != "")
                {
                    //decimal _iqty = Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString());
                    fConvertQty(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_mid"].Value.ToString()));
                    //decimal _jmlhrg = Convert.ToInt64(_fConvertQty) * Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value);
                    //if (isNefoKAM)
                    //{
                    //rounding mekanism
                    decimal _jmlhrg = Convert.ToInt64(_fConvertQty) * Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value);
                    if (isRoundingMekanism)
                    {
                        _jmlhrg = Math.Round(Convert.ToInt64(_fConvertQty) * Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value), 0, MidpointRounding.AwayFromZero);
                    }
                    // }
                    //dgvSalesDetail.Rows[row].Cells["wsd_total_so_amount"].Value = _jmlhrg.ToString("#,##0.00");//g_iJmlHarga
                    dgvSalesDetail.Rows[e.RowIndex].Cells["ijumlahharga"].Value = _jmlhrg.ToString("#,##0.00");//g_iJmlHarga

                    decimal _jmlrpnett = Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["ijumlahharga"].Value) - Convert.ToDecimal(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_tot_disc_pc"].Value);
                    dgvSalesDetail.Rows[e.RowIndex].Cells["jml_rp_netto"].Value = Convert.ToDecimal(_jmlrpnett).ToString("#,##0.00");
                }
            }
            else if (IsGridColumn(dgvSalesDetail, e, "wsd_req_order_xqty"))
            {
                double QtyReal = 0;
                double QtyReq = 0;
                if (e.RowIndex > -1 && txtKdOutlet1.Text != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_prd_master_code"].Value.ToString() != "" && dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value.ToString() != "")
                {

                    fQtyFormat(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_mid"].Value.ToString()));
                    QtyReal = Convert.ToDouble(_fConvertQty);

                    fQtyFormat(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_req_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_mid"].Value.ToString()));
                    QtyReq = Convert.ToDouble(_fConvertQty);

                    if (QtyReq > QtyReal)
                    {
                        dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_req_order_xqty"].Value = dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_real_order_xqty"].Value;

                        MessageBox.Show("Qty Req harus lebih kecil atau sama dengan Qty RealOrder", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //bExists = false;
                        return;
                    }
                }
            }

        }

        private void dgvSalesDetail_KeyUp(object sender, KeyEventArgs e)
        {
            string sTypeOrder = "";
            string sTypeSales = "";
            bool iscoldchain = false;
            bool sValidDivProd = false;
            IsBlockingAPI();
            if (APIBlockingSKU)
            {
                return;
            }

            if (e.KeyCode == Keys.Insert || e.KeyCode == Keys.Down)
            {
                int NewRow = 0;
                if (dgvSalesDetail.Rows.Count > 0)
                {
                    int row = dgvSalesDetail.Rows.Count - 1;
                    NewRow = dgvSalesDetail.Rows.Count;
                    if (dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString() != "" &&
                        dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"].Value.ToString() != "" &&
                        dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value.ToString() != "")
                    {
                        if (clsGlobal.MODE_TRX == 1)
                        {
                            DataTable dt = dgvSalesDetail.DataSource as DataTable;
                            dt.Rows.Add();
                        }
                        else
                        {
                            DataTable dt = dgvSalesDetail.DataSource as DataTable;
                            dt.Rows.Add();
                            dgvSalesDetail.DataSource = dt;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Masih ada data yang kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else
                {
                    DataTable dt = dgvSalesDetail.DataSource as DataTable;
                    dt.Rows.Add();
                }
                dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[NewRow].Cells["wsd_prd_master_code"];
            }
            else if (e.KeyCode == Keys.Delete)
            {
                if (dgvSalesDetail.Rows.Count > 2)
                {
                    dgvSalesDetail.Rows.Remove(dgvSalesDetail.CurrentRow);
                    dtTglPrice.Visible = false;
                    if (isTOPDivision == "Y")
                    {
                        TopTOPDivisionRenew();
                    }
                }
                if (dtTglPrice != null)
                {
                    dgvSalesDetail.Controls.Remove(dtTglPrice);
                }
                //dtTglPrice.Visible = false;
            }
            else if (e.KeyCode == Keys.F4)
            {
                if (dgvSalesDetail.CurrentCell == null)
                {
                    return;
                }
                int row = dgvSalesDetail.CurrentCell.RowIndex;
                if (dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Selected == true)
                {
                    if (dgvSalesDetail.Rows.Count > 1)
                    {
                        string pcoderow1 = dgvSalesDetail.Rows[0].Cells["wsd_prd_master_code"].Value.ToString();
                        strSQL = "SELECT * FROM IM_PRD_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " INNER JOIN ( ";
                        strSQL = strSQL + " select gh_function_code as cocline from GS_GEN_HARDCODED WITH (NOLOCK) ";
                        strSQL = strSQL + " where gh_function_name like 'LINE_COLD_CHAIN' and gh_sys = 'H' ";
                        strSQL = strSQL + " ) COC";
                        strSQL = strSQL + " ON  prm_prd_line_code = cocline ";
                        strSQL = strSQL + " WHERE  prm_prd_master_code = '" + pcoderow1 + "' ";

                        DataTable dtCheck = _clsGlobal.ExecDT(strSQL);
                        if (dtCheck.Rows.Count > 0)
                        {
                            iscoldchain = true;
                        }
                    }


                    //DataRow selectedDataRow = ((DataRowView)cbTipeOrder.GetSelectedDataRow()).Row;
                    int idO = Convert.ToInt32(Convert.ToString(cbTipeOrder.EditValue));
                    string NameO = cbTipeOrder.Text;

                    DataTable dtFill1 = new DataTable();
                    strSQL = " SELECT ot_def_oprtype FROM SO_ORDER_TYPE WITH (NOLOCK)  ";
                    strSQL = strSQL + " WHERE ot_order_type = '" + idO + "'";
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill1.Rows.Count > 0)
                    {
                        sTypeOrder = dtFill1.Rows[0]["ot_def_oprtype"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("Default operation untuk order type (" + idO + ") tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //break;
                    }

                    DataTable dtFill2 = new DataTable();
                    strSQL = " select sgm_type_operasi from SO_SPG_GIRL_MAN WITH (NOLOCK)  ";
                    strSQL = strSQL + " where sgm_spgm_id = '" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                    dtFill2 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill2.Rows.Count > 0)
                    {
                        sTypeSales = dtFill2.Rows[0]["sgm_type_operasi"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("Type operation untuk sales (" + txtSalesID.Text + ") tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //break;
                    }




                    strSQL = "select distinct prm_prd_master_code as PCode, prm_prd_desc as [Prd Name], ISNULL(pg_prd_group_desc, '') AS [PrdGroup] , ";
                    strSQL = strSQL + "prm_grade AS [Grade], prm_prd_size AS [Size], prm_conversion_purc [Conv1], prm_conversion_sales [Conv2], ";

                    if (isNefoKAM)
                    {
                        strSQL = strSQL + " case when cust_tax_code='PPN0' THEN cust_tax_code ELSE  stc_tax_code END [Tax Code], CASE WHEN cust_tax_code='PPN0' THEN isnull(cust_tax_amount,0) ELSE  isnull(stc_tax_amount, 0) END [Tax] ";
                    }
                    else
                    {
                        strSQL = strSQL + " prm_tax_code [Tax Code],case when isnull(excludeTax,'N') = 'Y' then 0 else isnull(stc_tax_amount, 0) end [Tax] ";
                    }
                    strSQL = strSQL + " ,prm_prd_line_code";



                    strSQL = strSQL + " from IM_PRD_MASTER WITH (NOLOCK)  ";
                    if (isNefoKAM)
                    {
                        strSQL = strSQL + " INNER JOIN SO_MAPPING_TOPBYPRDLINE WITH (NOLOCK) on ";
                        strSQL = strSQL + " smt_entity_id = '" + xEntityID + "' ";
                        strSQL = strSQL + " and smt_branch_id = '" + xBranchID + "' ";
                        strSQL = strSQL + " and smt_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                        strSQL = strSQL + " and smt_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                        strSQL = strSQL + " and smt_prdline_id = prm_prd_line_code ";
                    }
                    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
                    //if (isNefoKAM)
                    //{
                    //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON  CAST(CONVERT(DATE, '" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "', 112) AS DATETIME) >=cast(stc_tax_valid_from as datetime) ";
                    //    strSQL = strSQL + "  AND CAST(CONVERT(DATE,'" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "', 112) AS DATETIME) <=cast(stc_tax_valid_to as datetime)  ";
                    //}
                    //else
                    //{
                    //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
                    //}

                    strSQL = strSQL + " inner join IM_STOCK_BALANCE  on ";
                    strSQL = strSQL + "  wlb_loc_Id1 = '" + lblWHLoc1.Text + "'";
                    strSQL = strSQL + "  and wlb_loc_Id2 = '" + lblWHLoc2.Text + "'";
                    strSQL = strSQL + "  and wlb_entity_id = '" + xEntityID + "'";
                    strSQL = strSQL + "  and wlb_branch_id = '" + xBranchID + "'";
                    strSQL = strSQL + "  and wlb_prd_master_code = prm_prd_master_code";
                    strSQL = strSQL + "  and wlb_grade = prm_grade";
                    strSQL = strSQL + "  and wlb_prd_size = prm_prd_size";
                    strSQL = strSQL + " and isnull(wlb_stock_sts, 'N') = 'N' ";
                    strSQL = strSQL + " inner join IM_PRD_GROUP WITH (NOLOCK) ON ";
                    strSQL = strSQL + " prm_prd_line_code = pg_prd_line_code ";
                    strSQL = strSQL + " AND prm_prd_group_code = pg_prd_group_code ";
                    strSQL = strSQL + " inner join IM_PRD_LINE WITH (NOLOCK) ON ";
                    strSQL = strSQL + " prm_prd_line_code = pl_prd_line_code ";
                    if (g_bTOPFlag == true)
                    {
                        //DataRow selectedDataRow1 = ((DataRowView)cbTOP.GetSelectedDataRow()).Row;
                        //int idO1 = Convert.ToInt32(selectedDataRow1["code"]);
                        //string NameO = selectedDataRow["order_type"].ToString();

                        strSQL = strSQL + " INNER JOIN ( ";
                        strSQL = strSQL + " SELECT  DISTINCT map_prdline_code ";
                        strSQL = strSQL + " FROM    SO_CUST_MASTER WITH (NOLOCK)  ";
                        strSQL = strSQL + " INNER JOIN TBL_SD_MAP_TOP_PRDLINE WITH (NOLOCK)  ON ";
                        strSQL = strSQL + " cm_cust_group = map_cust_group ";
                        strSQL = strSQL + " INNER JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK)  ON ";
                        strSQL = strSQL + " pptc_term_code = map_term_code ";
                        strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + txtKdOutlet1.Text + "'";
                        strSQL = strSQL + "  AND cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                        strSQL = strSQL + "  AND cm_entity = '" + xEntityID + "'";
                        strSQL = strSQL + "  AND cm_branch = '" + xBranchID + "'";
                        strSQL = strSQL + "  AND pptc_term_code >= '" + Convert.ToString(cbTOP.EditValue).Trim() + "'";
                        strSQL = strSQL + " ) TOPPRD ON prm_prd_line_code = map_prdline_code ";
                    }
                    if (sTypeOrder == "C" && sTypeSales == "C")
                    {
                        strSQL = strSQL + " inner join TBL_CVS_PROD_REG WITH (NOLOCK)  ";
                        strSQL = strSQL + " on cpr_sld_id = '" + txtSalesID.Text + "'";
                        strSQL = strSQL + " and cpr_product_id = prm_prd_master_code";
                        strSQL = strSQL + " and cpr_product_grade = prm_grade";
                        strSQL = strSQL + " and cpr_product_size = prm_prd_size";
                        strSQL = strSQL + " and cpr_wh_loc_id1 = '" + lblWHLoc1.Text + "'";
                        strSQL = strSQL + " and cpr_wh_loc_id2 = '" + lblWHLoc2.Text + "'";
                        strSQL = strSQL + " and cpr_entity_id = '" + xEntityID + "'";
                        strSQL = strSQL + " and cpr_branch_id = '" + xBranchID + "'";

                    }
                    strSQL = strSQL + " LEFT JOIN TBL_IM_MAP_PRD_CUST WITH (NOLOCK)  ";
                    strSQL = strSQL + " ON  impc_cust_code1 = '" + txtKdOutlet1.Text + "'";
                    strSQL = strSQL + "     AND impc_cust_code2 = '" + lblKdOutlet2.Text + "'";
                    strSQL = strSQL + "     AND impc_entity_id = '" + xEntityID + "' ";
                    strSQL = strSQL + "     AND impc_branch_id = '" + xBranchID + "' ";
                    strSQL = strSQL + "     AND impc_prd_code = prm_prd_master_code ";
                    strSQL = strSQL + "     AND impc_prd_grade = prm_prd_size ";
                    strSQL = strSQL + "     AND impc_prd_size = prm_prd_size ";
                    strSQL = strSQL + "     AND impc_del_flag = 'N' ";

                    strSQL = strSQL + " INNER JOIN SO_SPG_GIRL_MAN WITH (NOLOCK)  ";
                    strSQL = strSQL + " ON  sgm_spgm_id = '" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "' ";
                    strSQL = strSQL + " INNER JOIN TBL_SD_SPGGROUP_PRDLINE WITH (NOLOCK)  ";
                    strSQL = strSQL + " ON  sgm_spgm_group = ssp_spggroup_code ";
                    strSQL = strSQL + "     AND prm_prd_line_code = ssp_prdline_code ";
                    strSQL = strSQL + "     AND ssp_del_flag = 'N' ";

                    strSQL = strSQL + " LEFT JOIN ( ";
                    strSQL = strSQL + " select gh_function_code as cocline from GS_GEN_HARDCODED WITH (NOLOCK) ";
                    strSQL = strSQL + " where gh_function_name like 'LINE_COLD_CHAIN' and gh_sys = 'H' ";
                    strSQL = strSQL + " ) COC";
                    strSQL = strSQL + " ON  prm_prd_line_code = cocline ";

                    if (isNefoKAM)
                    {
                        /** customer PPN **/
                        strSQL = strSQL + " LEFT JOIN  (SELECT cm_entity,cm_branch, cm_ship_tax_id AS cust_tax_code , stc_tax_amount as cust_tax_amount FROM SO_CUST_MASTER WITH (NOLOCK)  ";
                        strSQL = strSQL + "  INNER JOIN dbo.SO_TAX_CODE  WITH (NOLOCK) ON cm_ship_tax_id = stc_tax_code  ";
                        strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + txtKdOutlet1.Text + "'";
                        strSQL = strSQL + "  AND cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                        strSQL = strSQL + "  AND cm_entity = '" + xEntityID + "'";
                        strSQL = strSQL + "  AND cm_branch = '" + xBranchID + "'";
                        strSQL = strSQL + " )CUST ON cm_entity = '" + _xEntityID + "' and cm_branch = '" + _xBranchID + "' ";
                        /** End customer PPN **/
                    }
                    else
                    {
                        /** exclude tax **/
                        strSQL = strSQL + " LEFT JOIN ( SELECT gh_group_menu_id as __entity,gh_parent_menu __branch,gh_function_code excludeTax FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE GH_FUNCTION_NAME = 'Exclude-TAX-KAM' AND GH_SYS = 'S' ) FLAGTAX ON __entity = '" + xEntityID + "' and __branch = '" + xBranchID + "' ";
                        /** end exclude tax **/

                    }



                    //strSQL = strSQL + " WHERE prm_status = 'A'";
                    strSQL = strSQL + " WHERE prm_status IN ('A','D')";
                    strSQL = strSQL + " AND ( prm_prd_cust_flag = 'N' OR ( prm_prd_cust_flag = 'Y' AND impc_cust_code1 IS NOT NULL ) )";
                    //if (isMultisource)
                    //{
                    //    if (Convert.ToString(cbFlagSloc.EditValue).ToUpper().Equals("COC"))
                    //    {
                    //        strSQL = strSQL + " AND cocline is not null";
                    //    }
                    //    else
                    //    {
                    //        strSQL = strSQL + " AND cocline is null";
                    //    }
                    //}
                    //else
                    //{

                    if (dgvSalesDetail.Rows.Count > 1)
                    {
                        if (iscoldchain)
                        {
                            strSQL = strSQL + " AND cocline is not null";
                        }
                        else
                        {
                            strSQL = strSQL + " AND cocline is null";
                        }
                    }

                    //}

                    strSQL = strSQL + " order by prm_prd_master_code";



                    frmPopUp frm = new frmPopUp();
                    frm.FrmText = "Search Product";
                    frm.Query = strSQL;
                    frm.ShowDialog();
                    if (frm.ArrField != null)
                    {
                        if (isNefoKAM)
                        {
                            // untuk yang top division
                            if (isTOPDivision == "Y")
                            {
                                string top_by_customer = "N";
                                strSQL = "Select cm_top_by_cust,cm_top_id from SO_CUST_MASTER WITH (NOLOCK) where cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' and cm_cust_code1 = '" + txtKdOutlet1.Text + "' and cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                                DataTable dtCheck = _clsGlobal.ExecDT(strSQL);
                                if (dtCheck.Rows.Count > 0)
                                {
                                    top_by_customer = dtCheck.Rows[0]["cm_top_by_cust"].ToString();
                                }
                                if (Convert.ToString(cbPembyFaktur.EditValue) == "T" && dtCheck.Rows.Count > 0)
                                {
                                    sValidDivProd = true;
                                    if (dgvSalesDetail.Rows.Count > 1)
                                    {
                                        //string existprd = dgvSalesDetail.Rows[0].Cells["wsd_prd_master_code"].Value.ToString();
                                        //string nowprd = frm.ArrField[0].Trim();
                                        //strSQL = "SELECT Distinct prm_prd_line_code FROM IM_PRD_MASTER WHERE PRM_PRD_MASTER_CODE IN ('" + existprd + "','" + nowprd + "')";
                                        //if (_clsGlobal.ExecDT(strSQL).Rows.Count != 1)
                                        //{
                                        //    sValidDivProd = false;
                                        //    MsgErrorDivision = "Product berbeda Division";
                                        //}


                                    }
                                }
                                else
                                {
                                    if (dgvSalesDetail.Rows.Count > 1)
                                    {
                                        sValidDivProd = checkKAMTOPKAM(dgvSalesDetail.Rows.Count, frm.ArrField[0].Trim(), frm.ArrField[3].Trim(), frm.ArrField[4].Trim(), txtKdOutlet1.Text, lblKdOutlet2.Text, txtSalesID.Text);
                                    }
                                    else
                                    {
                                        sValidDivProd = checkKAMTOPKAM(0, frm.ArrField[0].Trim(), frm.ArrField[3].Trim(), frm.ArrField[4].Trim(), txtKdOutlet1.Text, lblKdOutlet2.Text, txtSalesID.Text);
                                    }

                                }

                                if (sValidDivProd == false)
                                {
                                    MessageBox.Show(MsgErrorDivision, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";
                                    return;
                                }
                                else
                                {
                                    if (top_by_customer == "N")
                                    {
                                        string vPF = Convert.ToString(cbPembyFaktur.EditValue);
                                        //jika tipe pembayaran T, TOP tidak berubah, 21-03-2019
                                        if (vPF == "T")
                                        {
                                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                                        }
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(TOPDivision))
                                            {
                                                for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                                                {
                                                    cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                                    string cdPF = TOPDivision;
                                                    string vTOP = Convert.ToString(cbTOP.EditValue);
                                                    if (Convert.ToString(cbTOP.EditValue) == cdPF)
                                                    {
                                                        cbTOP.Enabled = false;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            // untuk yang non top division
                            else
                            {
                                if (dgvSalesDetail.Rows.Count > 1)
                                {
                                    string prd1 = dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString();
                                    string grd1 = dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value.ToString();
                                    string sz1 = dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value.ToString();

                                    strSQL = " select CASE WHEN P1.prm_prd_line_code <> P2.prm_prd_line_code THEN 'N' ELSE 'Y' END as result" +
                                             " from " +
                                             " ( select prm_prd_line_code from IM_PRD_MASTER WITH (NOLOCK) where prm_prd_master_code = '" + prd1 + "' and prm_grade = '" + grd1 + "' and prm_prd_size = '" + sz1 + "') P1 " +
                                             " left join " +
                                             " ( select prm_prd_line_code from IM_PRD_MASTER WITH (NOLOCK) where prm_prd_master_code = '" + frm.ArrField[0].Trim() + "' and prm_grade = '" + frm.ArrField[3].Trim() + "' and prm_prd_size = '" + frm.ArrField[4].Trim() + "') P2 " +
                                             " ON 1 = 1";
                                    DataTable dtPL = _clsGlobal.ExecDT(strSQL);
                                    if (dtPL.Rows.Count > 0)
                                    {
                                        if (dtPL.Rows[0]["result"].ToString().Trim() == "N")
                                        {
                                            sValidDivProd = false;
                                        }
                                        else
                                        {
                                            sValidDivProd = true;
                                        }
                                    }
                                    //sValidDivProd = checkKAMTOPKAM(dgvSalesDetail.Rows.Count, frm.ArrField[0].Trim(), frm.ArrField[3].Trim(), frm.ArrField[4].Trim(), txtKdOutlet1.Text, lblKdOutlet2.Text, txtSalesID.Text);
                                }
                                if (sValidDivProd == false && dgvSalesDetail.Rows.Count > 1)
                                {
                                    MessageBox.Show("Product berbeda Division", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";
                                    return;
                                }

                            }
                        }


                        if (pCheckValue(frm.ArrField[0].Trim(), dgvSalesDetail.Rows.Count - 1, dgvSalesDetail) == true)
                        {
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = frm.ArrField[0].Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = frm.ArrField[3].Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = frm.ArrField[4].Trim();
                            dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = frm.ArrField[1].Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = frm.ArrField[5].Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = frm.ArrField[6].Trim();
                            dgvSalesDetail.Rows[row].Cells["tax_code"].Value = frm.ArrField[7].Trim();
                            dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = frm.ArrField[8].Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = frm.ArrField[9].Trim();

                            pChangePCode(row);
                        }
                        else
                        {
                            MessageBox.Show("Data sudah ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";
                            pChangePCode(row);
                        }

                        // dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];
                        // diminus 1 cell dikarenakan di keyup enter belum di eksekusi.
                        this.BeginInvoke(new MethodInvoker(() =>
                        {
                            //dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"];
                            //permintaan dari mas arief utk langsung ke sebelumnya
                            dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];
                        }));
                    }

                }
            }
        }

        // Mode 0 -- untuk yang pertama masuk  ; mode 1 -- untuk yang sudah ada paling atas produknya
        private bool checkKAMTOPKAM(int mode, string sPCode, string sGrade, string sSize, string sCustCode1, string sCustCode2, string sSpgmCode)
        {
            bool result = true;
            #region Mode 0
            if (mode == 0)
            {
                strSQL = "select PM.prm_prd_master_code pcode,PM.prm_grade,PM.prm_prd_size,PM.prm_prd_line_code pcodelinenew,SMT.smt_prdline_id plinemap,SMT.smt_top_id ";
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
                strSQL += "   where ";
                strSQL += "   prm_prd_master_code = '" + sPCode + "' ";
                strSQL += "   and prm_grade = '" + sGrade + "' ";
                strSQL += "   and prm_prd_size = '" + sSize + "' ";
                strSQL += "   and SPG.sgm_spgm_id = '" + sSpgmCode + "' ";
                strSQL += "   and SPG.sgm_entity_id = '" + xEntityID + "' ";
                strSQL += "   and SPG.sgm_branch_id = '" + xBranchID + "' ";
                strSQL += "   and CM.cm_cust_code1 = '" + sCustCode1 + "' ";
                strSQL += "   and CM.cm_cust_code2 = '" + sCustCode2 + "'";
                DataTable dtNewProduct = _clsGlobal.ExecDT(strSQL);
                if (dtNewProduct.Rows.Count > 0)
                {
                    string plmapping = dtNewProduct.Rows[0]["plinemap"].ToString().Trim();

                    if (string.IsNullOrEmpty(plmapping))
                    {
                        MsgErrorDivision = "TOP Belum di Setting";
                        return false;
                    }
                    TOPDivision = dtNewProduct.Rows[0]["smt_top_id"].ToString().Trim();
                }
                else
                {
                    MsgErrorDivision = "TOP Belum di Setting";
                    return false;
                }
            }
            #endregion

            #region mode 1
            else
            {
                bool isChangeTOP = false;
                strSQL = "select PM.prm_prd_master_code pcode,PM.prm_grade,PM.prm_prd_size,PM.prm_prd_line_code pcodelinenew,SMT.smt_prdline_id plinemap,SMT.smt_top_id, PTC.pptc_no_of_days ";
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
                strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                strSQL += "       ON ";
                strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                strSQL += "       ON ";
                strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                strSQL += "   where ";
                strSQL += "   prm_prd_master_code = '" + sPCode + "' ";
                strSQL += "   and prm_grade = '" + sGrade + "' ";
                strSQL += "   and prm_prd_size = '" + sSize + "' ";
                strSQL += "   and SPG.sgm_spgm_id = '" + sSpgmCode + "' ";
                strSQL += "   and SPG.sgm_entity_id = '" + xEntityID + "' ";
                strSQL += "   and SPG.sgm_branch_id = '" + xBranchID + "' ";
                strSQL += "   and CM.cm_cust_code1 = '" + sCustCode1 + "' ";
                strSQL += "   and CM.cm_cust_code2 = '" + sCustCode2 + "'";
                DataTable dtNewProduct = _clsGlobal.ExecDT(strSQL);

                string existprd = dgvSalesDetail.Rows[dgvSalesDetail.Rows.Count - 2].Cells["wsd_prd_master_code"].Value.ToString();
                string existgrd = dgvSalesDetail.Rows[dgvSalesDetail.Rows.Count - 2].Cells["wsd_grade"].Value.ToString();
                string existsize = dgvSalesDetail.Rows[dgvSalesDetail.Rows.Count - 2].Cells["wsd_prd_size"].Value.ToString();

                strSQL = "select PM.prm_prd_master_code pcode,PM.prm_grade,PM.prm_prd_size,PM.prm_prd_line_code pcodelinenew,SMT.smt_prdline_id plinemap,SMT.smt_top_id, PTC.pptc_no_of_days ";
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
                strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                strSQL += "       ON ";
                strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                strSQL += "       ON ";
                strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                strSQL += "   where ";
                strSQL += "   prm_prd_master_code = '" + existprd + "' ";
                strSQL += "   and prm_grade = '" + existgrd + "' ";
                strSQL += "   and prm_prd_size = '" + existsize + "' ";
                strSQL += "   and SPG.sgm_spgm_id = '" + sSpgmCode + "' ";
                strSQL += "   and SPG.sgm_entity_id = '" + xEntityID + "' ";
                strSQL += "   and SPG.sgm_branch_id = '" + xBranchID + "' ";
                strSQL += "   and CM.cm_cust_code1 = '" + sCustCode1 + "' ";
                strSQL += "   and CM.cm_cust_code2 = '" + sCustCode2 + "'";
                DataTable dtExistProduct = _clsGlobal.ExecDT(strSQL);

                if (dtNewProduct.Rows.Count > 0 && dtExistProduct.Rows.Count > 0)
                {
                    string plmapping = dtNewProduct.Rows[0]["plinemap"].ToString().Trim();
                    string topmapping = dtNewProduct.Rows[0]["smt_top_id"].ToString().Trim();
                    int noadays = Convert.ToInt32(dtNewProduct.Rows[0]["pptc_no_of_days"].ToString().Trim());

                    string plexistlinemapping = dtExistProduct.Rows[0]["plinemap"].ToString().Trim();
                    string topexistingmapping = dtExistProduct.Rows[0]["smt_top_id"].ToString().Trim();
                    int existsnoadays = Convert.ToInt32(Convert.ToString(cbTOP.EditValue).Replace("P", "").ToString().Trim());
                    if (!string.IsNullOrEmpty(plmapping) && !string.IsNullOrEmpty(plexistlinemapping))
                    {
                        if (noadays < existsnoadays)
                        {
                            isChangeTOP = true;
                        }
                    }
                    else
                    {
                        MsgErrorDivision = "TOP Belum di Setting";
                        return false;
                    }
                    if (string.IsNullOrEmpty(plexistlinemapping))
                    {
                        MsgErrorDivision = "TOP Belum di Setting";
                        return false;
                    }
                }
                else
                {
                    MsgErrorDivision = "TOP Belum di Setting";
                    return false;
                }
                if (isChangeTOP)
                {
                    TOPDivision = dtNewProduct.Rows[0]["smt_top_id"].ToString().Trim();
                }
                else
                {
                    TOPDivision = Convert.ToString(cbTOP.EditValue).Trim();
                }

            }
            #endregion
            return result;
        }


        //private void dgvSalesDetail_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        //{
        //    //vsList_StartEdit

        //}

        private void dgvSalesDetail_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string sTypeOrder = "";
            string sTypeSales = "";
            bool iscoldchain = false;
            bool sValidDivProd = false;
            int row = dgvSalesDetail.CurrentCell.RowIndex;

            if (APIBlockingSKU)
            {
                return;
            }

            if (IsGridColumn(dgvSalesDetail, e, COL_PCODE))
            {


                if (dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString() != "")
                {
                    /** SKU COLDCHAIN **/
                    if (dgvSalesDetail.Rows.Count > 1)
                    {
                        string pcoderow1 = dgvSalesDetail.Rows[0].Cells["wsd_prd_master_code"].Value.ToString();
                        strSQL = "SELECT * FROM IM_PRD_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " INNER JOIN ( ";
                        strSQL = strSQL + " select gh_function_code as cocline from GS_GEN_HARDCODED WITH (NOLOCK) ";
                        strSQL = strSQL + " where gh_function_name like 'LINE_COLD_CHAIN' and gh_sys = 'H' ";
                        strSQL = strSQL + " ) COC";
                        strSQL = strSQL + " ON  prm_prd_line_code = cocline ";
                        strSQL = strSQL + " WHERE  prm_prd_master_code = '" + pcoderow1 + "' ";

                        DataTable dtCheck = _clsGlobal.ExecDT(strSQL);
                        if (dtCheck.Rows.Count > 0)
                        {
                            iscoldchain = true;
                        }
                    }
                    /** END OF SKU COLDCHAIN **/

                    //DataRow selectedDataRow = ((DataRowView)cbTipeOrder.GetSelectedDataRow()).Row;
                    string idO = Convert.ToString(Convert.ToString(cbTipeOrder.EditValue));

                    DataTable dtFill1 = new DataTable();
                    strSQL = " SELECT ot_def_oprtype FROM SO_ORDER_TYPE WITH (NOLOCK)  ";
                    strSQL = strSQL + " WHERE ot_order_type = '" + idO + "'";
                    dtFill1 = _clsGlobal.ExecDT(strSQL);

                    if (dtFill1.Rows.Count > 0)
                    {
                        sTypeOrder = dtFill1.Rows[0]["ot_def_oprtype"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("Default operation untuk order type (" + idO + ") tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //break;
                    }

                    DataTable dtFill2 = new DataTable();
                    strSQL = " select sgm_type_operasi from SO_SPG_GIRL_MAN WITH (NOLOCK)  ";
                    strSQL = strSQL + " where sgm_spgm_id = '" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                    dtFill2 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill2.Rows.Count > 0)
                    {
                        sTypeSales = dtFill2.Rows[0]["sgm_type_operasi"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("Type operation untuk sales (" + txtSalesID.Text + ") tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //break;
                    }

                    DataTable dtFill3 = new DataTable();
                    strSQL = "select prm_prd_master_code as PCode, prm_prd_desc as PrdNameS, ISNULL(pg_prd_group_desc, '') AS PrdGroup , ";
                    strSQL = strSQL + "prm_grade AS GradeS, prm_prd_size AS SizeS, prm_conversion_purc as Conv1, prm_conversion_sales as Conv2, ";

                    if (isNefoKAM)
                    {
                        strSQL = strSQL + " case when cust_tax_code='PPN0' THEN cust_tax_code ELSE  stc_tax_code END TaxCodeS, CASE WHEN cust_tax_code='PPN0' THEN isnull(cust_tax_amount,0) ELSE  isnull(stc_tax_amount, 0) END TaxS ";
                        //strSQL = strSQL + " ,smt_createat_sap ";
                    }
                    else
                    {
                        strSQL = strSQL + " prm_tax_code as TaxCodeS, case when isnull(excludeTax,'N') = 'Y' then 0 else isnull(stc_tax_amount, 0) end as TaxS ";
                    }
                    strSQL = strSQL + " ,prm_prd_line_code, mcps_prd_sled as [sled]";


                    strSQL = strSQL + " from IM_PRD_MASTER WITH (NOLOCK)  ";
                    if (isNefoKAM)
                    {
                        strSQL = strSQL + " INNER JOIN SO_MAPPING_TOPBYPRDLINE WITH (NOLOCK) on ";
                        strSQL = strSQL + " smt_entity_id = '" + xEntityID + "' ";
                        strSQL = strSQL + " and smt_branch_id = '" + xBranchID + "' ";
                        strSQL = strSQL + " and smt_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                        strSQL = strSQL + " and smt_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                        strSQL = strSQL + " and smt_prdline_id = prm_prd_line_code ";
                    }
                    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
                    strSQL = strSQL + " LEFT JOIN TBL_MAP_CUST_PRD_SLED WITH (NOLOCK) ON mcps_entity_id='" + xEntityID + "'";
                    strSQL = strSQL + " and mcps_branch_id = '" + xBranchID + "' ";
                    strSQL = strSQL + " and mcps_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                    strSQL = strSQL + " and mcps_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                    strSQL = strSQL + " and mcps_prd_code = prm_prd_master_code ";
                    strSQL = strSQL + " and mcps_prd_grade = prm_grade ";
                    strSQL = strSQL + " and mcps_prd_size = prm_prd_size ";

                    //if (isNefoKAM)
                    //{
                    //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON  CAST(CONVERT(DATE, '" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "', 112) AS DATETIME) >=cast(stc_tax_valid_from as datetime) ";
                    //    strSQL = strSQL + "  AND CAST(CONVERT(DATE,'" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "', 112) AS DATETIME) <=cast(stc_tax_valid_to as datetime)  ";
                    //}
                    //else
                    //{
                    //    strSQL = strSQL + " LEFT JOIN SO_TAX_CODE WITH(NOLOCK) ON prm_tax_code = stc_tax_code ";
                    //}      
                    strSQL = strSQL + " inner join IM_STOCK_BALANCE  on ";
                    strSQL = strSQL + "  wlb_loc_Id1 = '" + lblWHLoc1.Text + "'";
                    strSQL = strSQL + "  and wlb_loc_Id2 = '" + lblWHLoc2.Text + "'";
                    strSQL = strSQL + "  and wlb_prd_master_code = prm_prd_master_code";
                    strSQL = strSQL + "  and wlb_grade = prm_grade";
                    strSQL = strSQL + "  and wlb_prd_size = prm_prd_size";
                    strSQL = strSQL + "  and wlb_entity_id = '" + xEntityID + "'";
                    strSQL = strSQL + "  and wlb_branch_id = '" + xBranchID + "'";
                    strSQL = strSQL + " and isnull(wlb_stock_sts, 'N') = 'N' ";
                    strSQL = strSQL + " inner join IM_PRD_GROUP WITH (NOLOCK) ON ";
                    strSQL = strSQL + " prm_prd_line_code = pg_prd_line_code ";
                    strSQL = strSQL + " AND prm_prd_group_code = pg_prd_group_code ";
                    strSQL = strSQL + " inner join IM_PRD_LINE WITH (NOLOCK) ON ";
                    strSQL = strSQL + " prm_prd_line_code = pl_prd_line_code ";
                    if (g_bTOPFlag == true)
                    {
                        //DataRow selectedDataRow1 = ((DataRowView)cbTOP.GetSelectedDataRow()).Row;
                        //string idO1 = Convert.ToString(selectedDataRow["code"]);
                        //string NameO = selectedDataRow["order_type"].ToString();

                        strSQL = strSQL + " INNER JOIN ( ";
                        strSQL = strSQL + " SELECT  DISTINCT map_prdline_code ";
                        strSQL = strSQL + " FROM    SO_CUST_MASTER WITH (NOLOCK)  ";
                        strSQL = strSQL + " INNER JOIN TBL_SD_MAP_TOP_PRDLINE WITH (NOLOCK)  ON ";
                        strSQL = strSQL + " cm_cust_group = map_cust_group ";
                        strSQL = strSQL + " INNER JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK)  ON ";
                        strSQL = strSQL + " pptc_term_code = map_term_code ";
                        strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + txtKdOutlet1.Text + "'";
                        strSQL = strSQL + "  AND cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                        strSQL = strSQL + "  AND cm_entity = '" + xEntityID + "'";
                        strSQL = strSQL + "  AND cm_branch = '" + xBranchID + "'";
                        strSQL = strSQL + "  AND pptc_term_code >= '" + Convert.ToString(cbTOP.EditValue).Trim() + "'";
                        strSQL = strSQL + " ) TOPPRD ON prm_prd_line_code = map_prdline_code ";
                    }
                    if (sTypeOrder == "C" && sTypeSales == "C")
                    {
                        strSQL = strSQL + " inner join TBL_CVS_PROD_REG WITH (NOLOCK)  ";
                        strSQL = strSQL + " on cpr_sld_id = '" + txtSalesID.Text + "'";
                        strSQL = strSQL + " and cpr_product_id = prm_prd_master_code";
                        strSQL = strSQL + " and cpr_product_grade = prm_grade";
                        strSQL = strSQL + " and cpr_product_size = prm_prd_size";
                        strSQL = strSQL + " and cpr_wh_loc_id1 = '" + lblWHLoc1.Text + "'";
                        strSQL = strSQL + " and cpr_wh_loc_id2 = '" + lblWHLoc2.Text + "'";
                        strSQL = strSQL + " and cpr_entity_id = '" + xEntityID + "'";
                        strSQL = strSQL + " and cpr_branch_id = '" + xBranchID + "'";
                    }
                    strSQL = strSQL + " LEFT JOIN TBL_IM_MAP_PRD_CUST WITH (NOLOCK)  ";
                    strSQL = strSQL + " ON  impc_cust_code1 = '" + txtKdOutlet1.Text + "'";
                    strSQL = strSQL + "     AND impc_cust_code2 = '" + lblKdOutlet2.Text + "'";
                    strSQL = strSQL + "     AND impc_prd_code = prm_prd_master_code ";
                    strSQL = strSQL + "     AND impc_prd_grade = prm_grade ";
                    strSQL = strSQL + "     AND impc_prd_size = prm_prd_size ";
                    strSQL = strSQL + "     AND impc_del_flag = 'N' ";
                    strSQL = strSQL + "     AND impc_entity_id = '" + xEntityID + "' ";
                    strSQL = strSQL + "     AND impc_branch_id = '" + xBranchID + "' ";
                    strSQL = strSQL + " INNER JOIN SO_SPG_GIRL_MAN WITH (NOLOCK) ";
                    strSQL = strSQL + " ON  sgm_spgm_id = '" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "' ";
                    strSQL = strSQL + " INNER JOIN TBL_SD_SPGGROUP_PRDLINE WITH (NOLOCK) ";
                    strSQL = strSQL + " ON  sgm_spgm_group = ssp_spggroup_code ";
                    strSQL = strSQL + "     AND prm_prd_line_code = ssp_prdline_code ";
                    strSQL = strSQL + "     AND ssp_del_flag = 'N' ";

                    strSQL = strSQL + " LEFT JOIN ( ";
                    strSQL = strSQL + " select gh_function_code as cocline from GS_GEN_HARDCODED WITH (NOLOCK) ";
                    strSQL = strSQL + " where gh_function_name like 'LINE_COLD_CHAIN' and gh_sys = 'H' ";
                    strSQL = strSQL + " ) COC";
                    strSQL = strSQL + " ON  prm_prd_line_code = cocline ";
                    if (isNefoKAM)
                    {
                        /** customer PPN **/
                        strSQL = strSQL + " LEFT JOIN  (SELECT cm_entity,cm_branch, cm_ship_tax_id AS cust_tax_code , stc_tax_amount as cust_tax_amount FROM SO_CUST_MASTER WITH (NOLOCK)  ";
                        strSQL = strSQL + "  INNER JOIN dbo.SO_TAX_CODE  WITH (NOLOCK) ON cm_ship_tax_id = stc_tax_code  ";
                        strSQL = strSQL + "  WHERE  cm_cust_code1 = '" + txtKdOutlet1.Text + "'";
                        strSQL = strSQL + "  AND cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                        strSQL = strSQL + "  AND cm_entity = '" + xEntityID + "'";
                        strSQL = strSQL + "  AND cm_branch = '" + xBranchID + "'";
                        strSQL = strSQL + " )CUST ON cm_entity = '" + _xEntityID + "' and cm_branch = '" + _xBranchID + "' ";
                        /** End customer PPN **/
                    }
                    else
                    {
                        /** exclude tax **/
                        strSQL = strSQL + " LEFT JOIN ( SELECT gh_group_menu_id as __entity,gh_parent_menu __branch,gh_function_code excludeTax FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE GH_FUNCTION_NAME = 'Exclude-TAX-KAM' AND GH_SYS = 'S' ) FLAGTAX ON __entity = '" + xEntityID + "' and __branch = '" + xBranchID + "' ";
                        /** end exclude tax **/
                    }

                    //strSQL = strSQL + " WHERE prm_status = 'A'";
                    strSQL = strSQL + " WHERE prm_status IN ('A','D')";
                    strSQL = strSQL + " AND ( prm_prd_cust_flag = 'N' OR ( prm_prd_cust_flag = 'Y' AND impc_cust_code1 IS NOT NULL ) ) and prm_prd_master_code='" + dgvSalesDetail.Rows[e.RowIndex].Cells[COL_PCODE].Value.ToString() + "'";
                    //if (isMultisource)
                    //{
                    //    if (Convert.ToString(cbFlagSloc.EditValue).ToUpper().Equals("COC"))
                    //    {
                    //        strSQL = strSQL + " AND cocline is not null";
                    //    }
                    //    else
                    //    {
                    //        strSQL = strSQL + " AND cocline is null";
                    //    }
                    //}
                    //else
                    //{
                    if (dgvSalesDetail.Rows.Count > 1)
                    {
                        if (iscoldchain)
                        {
                            strSQL = strSQL + " AND cocline is not null";
                        }
                        else
                        {
                            strSQL = strSQL + " AND cocline is null";
                        }
                    }
                    //}


                    dtFill3 = _clsGlobal.ExecDT(strSQL);

                    if (dtFill3.Rows.Count > 0)
                    {
                        //KAM Segment
                        if (isNefoKAM)
                        {

                            if (isTOPDivision == "Y")
                            {
                                string top_by_customer = "N";
                                strSQL = "Select cm_top_by_cust,cm_top_id from SO_CUST_MASTER WITH (NOLOCK) where cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' and cm_cust_code1 = '" + txtKdOutlet1.Text + "' and cm_cust_code2 = '" + lblKdOutlet2.Text + "'";
                                DataTable dtCheck = _clsGlobal.ExecDT(strSQL);
                                if (dtCheck.Rows.Count > 0)
                                {
                                    top_by_customer = dtCheck.Rows[0]["cm_top_by_cust"].ToString();
                                }
                                if (cbPembyFaktur.EditValue == "T" && dtCheck.Rows.Count > 0)
                                {
                                    sValidDivProd = true;
                                    if (dgvSalesDetail.Rows.Count > 1)
                                    {
                                        string existprd = dgvSalesDetail.Rows[0].Cells["wsd_prd_master_code"].Value.ToString();
                                        string nowprd = dtFill3.Rows[0]["PrdNameS"].ToString().Trim();

                                        //strSQL = "SELECT Distinct prm_prd_line_code FROM IM_PRD_MASTER WHERE PRM_PRD_MASTER_CODE IN ('" + existprd + "','" + nowprd + "')";
                                        //if (_clsGlobal.ExecDT(strSQL).Rows.Count != 1)
                                        //{
                                        //    sValidDivProd = false;
                                        //    MsgErrorDivision = "Product berbeda Division";
                                        //}


                                    }
                                }
                                else
                                {
                                    if (dgvSalesDetail.Rows.Count > 1)
                                    {
                                        sValidDivProd = checkKAMTOPKAM(dgvSalesDetail.Rows.Count, dtFill3.Rows[0]["PCode"].ToString(), dtFill3.Rows[0]["GradeS"].ToString().Trim(), dtFill3.Rows[0]["SizeS"].ToString().Trim(), txtKdOutlet1.Text, lblKdOutlet2.Text, txtSalesID.Text);
                                    }
                                    else
                                    {
                                        sValidDivProd = checkKAMTOPKAM(0, dtFill3.Rows[0]["PCode"].ToString(), dtFill3.Rows[0]["GradeS"].ToString().Trim(), dtFill3.Rows[0]["SizeS"].ToString().Trim(), txtKdOutlet1.Text, lblKdOutlet2.Text, txtSalesID.Text);
                                    }
                                }
                                if (sValidDivProd == false)
                                {
                                    MessageBox.Show(MsgErrorDivision, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_sled"].Value = "";
                                    return;
                                }
                                else
                                {
                                    if (top_by_customer == "N")
                                    {
                                        string vPF = Convert.ToString(cbPembyFaktur.EditValue);
                                        //jika tipe pembayaran T, TOP tidak berubah, 21-03-2019
                                        if (vPF == "T")
                                        {
                                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                                        }
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(TOPDivision))
                                            {
                                                for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                                                {
                                                    cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                                    string cdPF = TOPDivision;
                                                    string vTOP = Convert.ToString(cbTOP.EditValue);
                                                    if (Convert.ToString(cbTOP.EditValue) == cdPF)
                                                    {
                                                        cbTOP.Enabled = false;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // kalo bukan top by division
                            else
                            {
                                if (dgvSalesDetail.Rows.Count > 1)
                                {
                                    string prd1 = dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString();
                                    string grd1 = dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value.ToString();
                                    string sz1 = dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value.ToString();

                                    strSQL = " select CASE WHEN P1.prm_prd_line_code <> P2.prm_prd_line_code THEN 'N' ELSE 'Y' END as result" +
                                             " from " +
                                             " ( select prm_prd_line_code from IM_PRD_MASTER WITH (NOLOCK) where prm_prd_master_code = '" + prd1 + "' and prm_grade = '" + grd1 + "' and prm_prd_size = '" + sz1 + "') P1 " +
                                             " left join " +
                                             " ( select prm_prd_line_code from IM_PRD_MASTER WITH (NOLOCK) where prm_prd_master_code = '" + dtFill3.Rows[0]["PCode"].ToString().Trim() + "' and prm_grade = '" + dtFill3.Rows[0]["GradeS"].ToString().Trim() + "' and prm_prd_size = '" + dtFill3.Rows[0]["SizeS"].ToString().Trim() + "') P2 " +
                                             " ON 1 = 1";
                                    DataTable dtPL = _clsGlobal.ExecDT(strSQL);
                                    if (dtPL.Rows.Count > 0)
                                    {
                                        if (dtPL.Rows[0]["result"].ToString().Trim() == "N")
                                        {
                                            sValidDivProd = false;
                                        }
                                        else
                                        {
                                            sValidDivProd = true;
                                        }
                                    }
                                    //sValidDivProd = checkKAMTOPKAM(dgvSalesDetail.Rows.Count, frm.ArrField[0].Trim(), frm.ArrField[3].Trim(), frm.ArrField[4].Trim(), txtKdOutlet1.Text, lblKdOutlet2.Text, txtSalesID.Text);
                                }
                                if (sValidDivProd == false && dgvSalesDetail.Rows.Count > 1)
                                {
                                    MessageBox.Show("Product berbeda Division", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";
                                    dgvSalesDetail.Rows[row].Cells["wsd_sled"].Value = "";
                                    return;
                                }
                            }
                        }
                        if (pCheckValue(dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString(), row, dgvSalesDetail) == true)
                        {
                            if (isNefoKAM)
                            {
                                if (string.IsNullOrEmpty(dtFill3.Rows[0]["TaxCodeS"].ToString()))
                                {
                                    MessageBox.Show("Parameter untuk so_tax_code valid date '" + Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd") + "' tidak ditemukan!.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    return;
                                }

                            }
                            //dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = dtFill3.Rows[0]["ot_def_oprtype"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = dtFill3.Rows[0]["GradeS"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = dtFill3.Rows[0]["SizeS"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = dtFill3.Rows[0]["PrdNameS"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = dtFill3.Rows[0]["Conv1"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = dtFill3.Rows[0]["Conv2"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["tax_code"].Value = dtFill3.Rows[0]["TaxCodeS"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = dtFill3.Rows[0]["TaxS"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = dtFill3.Rows[0]["prm_prd_line_code"].ToString().Trim();
                            dgvSalesDetail.Rows[row].Cells["wsd_sled"].Value = dtFill3.Rows[0]["sled"].ToString().Trim();
                            pChangePCode(row);
                            dgvSalesDetail.Columns.Equals(9);
                            isPcodeGridExists = true;
                            editedgvRowid = row;
                            dgvSalesDetail.ClearSelection();
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];
                            }));
                            //dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];
                            //dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Selected=true;

                        }
                        else
                        {
                            MessageBox.Show("Data sudah ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_sled"].Value = "";
                            pChangePCode(row);
                        }
                    }
                    else
                    {
                        dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["prm_prd_desc"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["tax_code"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["pct_tax"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["wsd_prd_line_code"].Value = "";
                        dgvSalesDetail.Rows[row].Cells["wsd_sled"].Value = "";
                        pChangePCode(row);
                    }
                    dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value = "";
                    //this.Invoke((MethodInvoker)delegate { dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"]; });
                    //dgvSalesDetail[9, e.RowIndex].Selected = true;
                    //dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value = "";
                    //dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];

                    //
                    //dgvSalesDetail.CurrentCell.Selected = true;



                }
            }
            if (IsGridColumn(dgvSalesDetail, e, "wsd_req_order_xqty"))
            {
                if (dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString() != "")
                {

                    if (dgvSalesDetail.Rows[row].Cells["wsd_req_order_xqty"].Value.ToString() != "")
                    {
                        fQtyFormat(dgvSalesDetail.Rows[row].Cells["wsd_req_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        dgvSalesDetail.Rows[row].Cells["wsd_req_order_xqty"].Value = _fQtyFormat;

                        double QtyReal = 0;
                        double QtyReq = 0;
                        fQtyFormat(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        QtyReal = Convert.ToDouble(_fConvertQty);

                        fQtyFormat(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_req_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        QtyReq = Convert.ToDouble(_fConvertQty);

                        RefreshReasonEditState(e.RowIndex);


                    }
                }
            }
            if (IsGridColumn(dgvSalesDetail, e, "wsd_real_order_xqty"))
            {
                if (dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString() != "")
                {
                    if (dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value.ToString() != "")
                    {
                        //dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value =fqty
                        fQtyFormat(dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        if (_fQtyFormat.Equals("00000.000.0000"))
                        {
                            dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value = "";
                            MessageBox.Show("Product Order Harus Lebih Besar Dari 0.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                        dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value = _fQtyFormat;
                        dgvSalesDetail.Rows[row].Cells["wsd_req_order_xqty"].Value = _fQtyFormat;
                        //dgvSalesDetail.Rows[row].Cells["wsd_so_xqty"].Value = String.Format("{0:n2}",dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"].Value.ToString());
                        dgvSalesDetail.Rows[row].Cells["wsd_so_xqty"].Value = _fQtyFormat;
                        string _cpcode = dgvSalesDetail.Rows[row].Cells["wsd_prd_master_code"].Value.ToString();
                        string _cgrade = dgvSalesDetail.Rows[row].Cells["wsd_grade"].Value.ToString();
                        string _csize = dgvSalesDetail.Rows[row].Cells["wsd_prd_size"].Value.ToString();
                        string _ctglprice = dgvSalesDetail.Rows[row].Cells["tgl_price"].Value.ToString();

                        if (fPricePCode(_cpcode, _cgrade, _csize, txtKdOutlet1.Text, lblKdOutlet2.Text, _ctglprice) == true)
                        {
                            //decimal _iqty = Convert.ToDecimal(dgvSalesDetail.Rows[row].Cells["wsd_so_xqty"].Value.ToString());
                            fConvertQty(_fQtyFormat, Convert.ToInt16(dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[row].Cells["wsd_uom_convert_mid"].Value.ToString()));
                            long _xx = Convert.ToInt64(_fConvertQty);
                            if (isNefoKAM)
                            {
                                dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"].Value = curPrice.ToString("#,##0.0000");
                            }
                            else
                            {
                                dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"].Value = curPrice.ToString("#,##0.00");
                            }

                            //decimal _jmlhrg = _xx * Convert.ToDecimal(dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"].Value);
                            //if (isNefoKAM)
                            //{
                            // rounding mekanism
                            decimal _jmlhrg = (_xx * Convert.ToDecimal(dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"].Value));
                            if (isRoundingMekanism)
                            {
                                _jmlhrg = Math.Round(_xx * Convert.ToDecimal(dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"].Value), 0, MidpointRounding.AwayFromZero);
                            }
                            //}
                            //dgvSalesDetail.Rows[row].Cells["wsd_total_so_amount"].Value = _jmlhrg.ToString("#,##0.00");//g_iJmlHarga
                            dgvSalesDetail.Rows[row].Cells["ijumlahharga"].Value = _jmlhrg.ToString("#,##0.00");//g_iJmlHarga

                            decimal _jmlrpnett = _jmlhrg - Convert.ToDecimal(dgvSalesDetail.Rows[row].Cells["wsd_tot_disc_pc"].Value);
                            dgvSalesDetail.Rows[row].Cells["jml_rp_netto"].Value = _jmlrpnett.ToString("#,##0.00");
                        }
                        else
                        {
                            MessageBox.Show("Periode validate harga sudah habis atau harga barang belum di setting", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            dgvSalesDetail.Rows[row].Cells["wsd_het_unit_price"].Value = "";
                            dgvSalesDetail.Rows[row].Cells["wsd_total_so_amount"].Value = "0";
                            dgvSalesDetail.Rows[row].Cells["jml_rp_netto"].Value = "0";
                            dgvSalesDetail.Rows[row].Cells["ijumlahharga"].Value = "0";
                        }
                    }

                }
                else
                {
                    dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_so_xqty"].Value = "";
                    dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_real_order_xqty"].Value = "";
                    dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_req_order_xqty"].Value = "";
                    dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_het_unit_price"].Value = "";
                    dgvSalesDetail.Rows[e.RowIndex].Cells["wsd_total_so_amount"].Value = "0";
                    dgvSalesDetail.Rows[e.RowIndex].Cells["jml_rp_netto"].Value = "0";
                }
            }
        }

        private void txtKdOutlet1_Leave(object sender, EventArgs e)
        {
            string sKDOutlet1;
            string sTOPByCust;
            //string BLMode ="N";
            //strSQL = "select gh_function_code, gh_function_desc from GS_GEN_HARDCODED WHERE gh_function_name = 'TIRASND_BL'";
            //DataTable dtstatusbl = _clsGlobal.ExecDT(strSQL);
            //if (dtstatusbl.Rows.Count > 0)
            //{
            //    BLMode = dtstatusbl.Rows[0]["gh_function_code"].ToString();
            //}

            if (txtKdOutlet1.Text != g_sOutlet)
            {
                APIBlockingSKU = false;
                dgvSalesDetail.Columns["wsd_prd_master_code"].ReadOnly = false;
                dgvSalesDetail.Columns["wsd_real_order_xqty"].ReadOnly = false;

                sKDOutlet1 = txtKdOutlet1.Text;
                DataTable dtFill1 = new DataTable();
                strSQL = "select cm_cust_code1, cm_cust_code2,  cm_cust_name, gh_function_desc, isnull(cm_payment_type, '') cm_payment_type, cm_top_id, cm_top_by_cust, isnull(cm_fullfilment,'N') cm_fullfilment from SO_CUST_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + "  inner join TBL_SD_CUSTCOVER  WITH (NOLOCK) on  cm_cust_code1 = csc_cust_code1 and cm_cust_code2 = csc_cust_code2 and cm_entity = csc_entity and cm_branch = csc_branch left join ( select gh_function_code, gh_function_desc ";
                strSQL = strSQL + "  from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_sys = 'H' ";//OUTLETSTATUS
                strSQL = strSQL + "  and gh_function_name ='OUTLETSTATUS')  GS on cm_active_flag = gh_function_code  ";
                strSQL = strSQL + "  where csc_salesman_id = '" + txtSalesID.Text + "' ";
                strSQL = strSQL + "  and cm_cust_code1 = '" + sKDOutlet1 + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' ";
                strSQL = strSQL + "  and cm_active_flag <> 'D' ";
                dtFill1 = _clsGlobal.ExecDT(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    txtKdOutlet1.Text = dtFill1.Rows[0]["cm_cust_code1"].ToString();
                    lblOutletDesc.Text = dtFill1.Rows[0]["cm_cust_name"].ToString();
                    lblKdOutlet2.Text = dtFill1.Rows[0]["cm_cust_code2"].ToString();
                    txtStatusOutlet.Text = dtFill1.Rows[0]["gh_function_desc"].ToString();
                    sTOPByCust = dtFill1.Rows[0]["cm_top_by_cust"].ToString();
                    // cbfullfilment.Checked = false; 
                    // if (dtFill1.Rows[0]["cm_fullfilment"].ToString().Trim().ToUpper().Equals("Y"))
                    // {
                    //     cbfullfilment.Checked = true;
                    // }

                    if (dtFill1.Rows[0]["cm_payment_type"].ToString() == "")
                    {
                        cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, 0);
                    }
                    else
                    {
                        g_bEdit = false;

                        if (!_clsGlobal.IsNefoForBL)
                        {


                            for (lCounter = 0; lCounter < LookupRowCount(cbPembyFaktur); lCounter++)
                            {
                                cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, lCounter);
                                string cdPF = dtFill1.Rows[0]["cm_payment_type"].ToString().Trim();
                                string vTOP = Convert.ToString(cbPembyFaktur.EditValue);
                                if (vTOP == cdPF)
                                {
                                    g_bEdit = true;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            g_bEdit = true;
                        }

                        if (g_bEdit == false)
                        {
                            MessageBox.Show("Default tipe Pembayaran Faktur tidak ada untuk customer " + dtFill1.Rows[0]["cm_cust_code1"].ToString(), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                    }

                    g_bChangeOutlet = true;
                    //if (dtFill1.Rows[0]["cm_top_id"].ToString() == "")
                    //{
                    //    cbTOP.ItemIndex = 0;
                    //}
                    //else
                    //{
                    g_bEdit = false;
                    DataTable dtFill2 = new DataTable();
                    strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
                    //strSQL = strSQL + "  ge_entity_id = '" + txtEntityCode.Text + "'"; //AND ge_branch_flag = '" + txtBranchCode.Text + "'";
                    dtFill2 = _clsGlobal.ExecDT(strSQL);

                    DataTable dtCheckTopDivision = new DataTable();
                    strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name='TOPBYDIVISION' and gh_sys='H' and gh_function_code='Y'";
                    dtCheckTopDivision = _clsGlobal.ExecDT(strSQL);

                    #region TOP BY SALESMAN GROUP
                    if (dtFill2.Rows.Count > 0)
                    {
                        // d_Branch = dtFill2.Rows[0]["ge_branch_flag"].ToString();
                        DataTable dtFill3 = new DataTable();
                        strSQL = "select sgl_group_top from SO_SPG_GIRL_MAN WITH (NOLOCK)  left join SO_SPG_GROUP WITH (NOLOCK)  on sgm_spgm_group = sgl_group_code ";
                        strSQL = strSQL + " where sgm_spgm_id='" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                        dtFill3 = _clsGlobal.ExecDT(strSQL);
                        if (dtFill3.Rows.Count > 0)
                        {
                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string cdPF = dtFill3.Rows[0]["sgl_group_top"].ToString().Trim();
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (Convert.ToString(cbTOP.EditValue) == cdPF)
                                {
                                    cbTOP.Enabled = false;
                                    break;
                                }
                            }

                            //top harusnya tidak refresh
                            if (Convert.ToString(cbPembyFaktur.EditValue) == "T")
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                            }

                            DataTable dtFill4 = new DataTable();
                            strSQL = "select cg_cust_top from SO_CUST_MASTER WITH (NOLOCK)  inner join SO_CUST_GROUP WITH (NOLOCK)   on cm_cust_group = cg_cust_group ";
                            strSQL = strSQL + " where cg_cust_top is not null and cm_cust_code1='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'"; //AND ge_branch_flag = '" + txtBranchCode.Text + "'";
                            dtFill4 = _clsGlobal.ExecDT(strSQL);
                            //if (dtFill4.Rows.Count > 0)
                            //{
                            //    for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            //    {
                            //        cbTOP.ItemIndex = lCounter;
                            //        string cdPF = dtFill4.Rows[0]["cg_cust_top"].ToString().Trim();
                            //        string vTOP = Convert.ToString(cbTOP.EditValue);
                            //        if (Convert.ToString(cbTOP.EditValue) == cdPF)
                            //        {
                            //            //cbTOP.Enabled = false;
                            //            break;
                            //        }
                            //    }
                            //}
                            if (dtFill4.Rows.Count > 0)
                            {
                                string __top = (string)(from r in ((DataTable)cbTOP.Properties.DataSource).Rows.Cast<DataRow>()
                                                        where r["code"].ToString() == dtFill4.Rows[0]["cg_cust_top"].ToString().Trim()
                                                        select r["code"].ToString())
                                        .ToList()
                                        .FirstOrDefault();
                                if (!__top.IsNullOrEmptyOrWhiteSpace())
                                    cbTOP.EditValue = __top;
                            }


                            if (sTOPByCust == "Y")
                            {
                                DataTable dtFill5 = new DataTable();
                                strSQL = "select top 1 cm_top_id from SO_CUST_MASTER WITH (NOLOCK)  ";
                                strSQL = strSQL + " where cm_cust_code1='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'";
                                dtFill5 = _clsGlobal.ExecDT(strSQL);

                                for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                                {
                                    cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                    string cdPF = dtFill5.Rows[0]["cm_top_id"].ToString().Trim();
                                    string vTOP = Convert.ToString(cbTOP.EditValue);
                                    if (vTOP == cdPF)
                                    {
                                        break;
                                    }
                                }
                            }

                            //top harusnya tidak refresh
                            if (Convert.ToString(cbPembyFaktur.EditValue) == "T")
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                            }
                        }
                    }
                    #endregion
                    #region TOP BY DIVISION (KAM)
                    else if (dtCheckTopDivision.Rows.Count > 0)
                    {
                        if (Convert.ToString(cbPembyFaktur.EditValue) == "T")
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                        }
                        //Top Cust di setting Y 
                        else if (sTOPByCust == "Y")
                        {
                            DataTable dtFill5 = new DataTable();
                            strSQL = "select top 1 cm_top_id from SO_CUST_MASTER WITH (NOLOCK)  ";
                            strSQL = strSQL + " where cm_cust_code1='" + txtKdOutlet1.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "'";
                            dtFill5 = _clsGlobal.ExecDT(strSQL);

                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string cdPF = dtFill5.Rows[0]["cm_top_id"].ToString().Trim();
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (vTOP == cdPF)
                                {
                                    break;
                                }
                            }
                        }
                        // Top Cust Division check to mapping
                        else
                        {
                            DataTable dtFillCheckMapping = new DataTable();
                            if (dgvSalesDetail.Rows.Count > 0)
                            {
                                //string __prd = dgvSalesDetail.Rows[0].Cells["wsd_prd_master_code"].Value.ToString();
                                //string __grade = dgvSalesDetail.Rows[0].Cells["wsd_grade"].Value.ToString();
                                //string __size = dgvSalesDetail.Rows[0].Cells["wsd_prd_size"].Value.ToString();

                                //strSQL = "select prm_prd_line_code from IM_PRD_master where prm_prd_master_code = '" + __prd + "' and prm_grade = '" + __grade + "' and prm_prd_size = '"+ __size +"' ";
                                //DataTable _dtPRDLINE = _clsGlobal.ExecDT(strSQL);

                                //strSQL = "select top 1 smt_top_id from SO_MAPPING_TOPBYPRDLINE where smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + xBranchID + "' and smt_cust_code1 = '" + txtKdOutlet1.Text + "' and smt_cust_code2 = '" + lblKdOutlet2.Text + "' and smt_prdline_id = '" + _dtPRDLINE.Rows[0]["prm_prd_line_code"].ToString() + "'";
                                // ERROR BUG YUDA

                                string product = "";
                                for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                                {
                                    if (string.IsNullOrEmpty(product))
                                    {
                                        if (!string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim()))
                                        {
                                            product = "'" + dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim() + "'";
                                        }

                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim()))
                                        {
                                            product += ",'" + dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString().Trim() + "'";
                                        }
                                    }


                                }
                                //strSQL = "select top 1 SMT.smt_top_id,PTC.pptc_no_of_days " +
                                //         "   from SO_CUST_MASTER CM " +
                                //         "   left join TBL_SD_CUSTCOVER TSC " +
                                //         "       ON  " +
                                //         "           TSC.csc_entity          = CM.cm_entity " +
                                //         "           AND TSC.csc_branch      = CM.cm_branch " +
                                //         "           AND TSC.csc_cust_code1  = CM.cm_cust_code1 " +
                                //         "           AND TSC.csc_cust_code2  = CM.cm_cust_code2 " +
                                //         "   left join SO_SPG_GIRL_MAN SPG " +
                                //         "       ON	" +
                                //         "           SPG.sgm_spgm_id  = TSC.csc_salesman_id " +
                                //         "           AND SPG.sgm_entity_id = TSC.csc_entity " +
                                //         "           AND SPG.sgm_branch_id = TSC.csc_branch " +
                                //         "   left join SO_MAPPING_TOPBYPRDLINE SMT " +
                                //         "       ON " +
                                //         "           SMT.smt_entity_id      = CM.cm_entity " +
                                //         "           and SMT.smt_branch_id  = CM.cm_branch " +
                                //         "           and SMT.smt_cust_code1 = CM.cm_cust_code1 " +
                                //         "           and SMT.smt_cust_code2 = CM.cm_cust_code2 " +
                                //         "   left Join PO_PAYMENT_TERM_CODES PTC " +
                                //         "       ON " +
                                //         "           PTC.pptc_term_code = SMT.smt_top_id " +
                                //         "   left join IM_PRD_MASTER PM " +
                                //         "       ON " +
                                //         "           SMT.smt_prdline_id = PM.prm_prd_line_code " +
                                //         "   where " +
                                //         "   prm_prd_master_code in (" + product + ") " +
                                //         "   and SPG.sgm_spgm_id = '" + txtSalesID.Text + "' " +
                                //         "   and SPG.sgm_entity_id = '" + xEntityID + "' " +
                                //         "   and SPG.sgm_branch_id = '" + xBranchID + "' " +
                                //         "   and CM.cm_cust_code1 = '" + txtKdOutlet1.Text + "' " +
                                //         "   and CM.cm_cust_code2 = '" + lblKdOutlet2.Text + "' " +
                                //         "   order by pptc_no_of_days asc ";
                                strSQL = "select top 1 SMT.smt_top_id,PTC.pptc_no_of_days ";
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
                                strSQL += "   left Join PO_PAYMENT_TERM_CODES PTC WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           PTC.pptc_term_code = SMT.smt_top_id ";
                                strSQL += "   left join IM_PRD_MASTER PM WITH (NOLOCK) ";
                                strSQL += "       ON ";
                                strSQL += "           SMT.smt_prdline_id = PM.prm_prd_line_code ";
                                strSQL += "   where ";
                                strSQL += "   prm_prd_master_code in (" + product + ") ";
                                strSQL += "   and SPG.sgm_spgm_id = '" + txtSalesID.Text + "' ";
                                strSQL += "   and SPG.sgm_entity_id = '" + xEntityID + "' ";
                                strSQL += "   and SPG.sgm_branch_id = '" + xBranchID + "' ";
                                strSQL += "   and CM.cm_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                                strSQL += "   and CM.cm_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                                strSQL += "   order by pptc_no_of_days asc ";
                            }
                            else
                            {
                                //strSQL = " select top 1 smt_top_id from SO_MAPPING_TOPBYPRDLINE where smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + xBranchID + "' and smt_cust_code1 = '" + txtKdOutlet1.Text + "' and smt_cust_code2 = '" + lblKdOutlet2.Text + "'";
                                strSQL = "Select cm_bill_to_code1,cm_bill_to_code2";
                                if (isMultibranch)
                                {
                                    strSQL += ",cm_cust_bill_branch";
                                }
                                strSQL += " From SO_CUST_MASTER WITH(NOLOCK) where cm_cust_code1 = '" + txtKdOutlet1.Text + "' and cm_cust_code2 = '" + lblKdOutlet2.Text + "' and cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' ";
                                DataTable dtCustMaster = _clsGlobal.ExecDT(strSQL);
                                string billtocode1 = "";
                                string billtocode2 = "";
                                string billtobranch = "";

                                if (dtCustMaster.Rows.Count > 0)
                                {
                                    billtocode1 = dtCustMaster.Rows[0]["cm_bill_to_code1"].ToString();
                                    billtocode2 = dtCustMaster.Rows[0]["cm_bill_to_code2"].ToString();
                                    if (isMultibranch)
                                    {
                                        billtobranch = dtCustMaster.Rows[0]["cm_cust_bill_branch"].ToString();
                                    }
                                }
                                strSQL = "select top 1 smt_top_id  from SO_MAPPING_TOPBYPRDLINE WITH(NOLOCK) ";
                                if (!isMultibranch)
                                {
                                    strSQL = strSQL + " where smt_cust_code1 ='" + billtocode1 + "' and smt_cust_code2 = '" + billtocode2 + "' and smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + xBranchID + "' ";
                                }
                                else
                                {
                                    strSQL = strSQL + " where smt_cust_code1 ='" + billtocode1 + "' and smt_cust_code2 = '" + billtocode2 + "' and smt_entity_id = '" + xEntityID + "' and smt_branch_id = '" + billtobranch + "' ";
                                }


                            }

                            dtFillCheckMapping = _clsGlobal.ExecDT(strSQL);

                            if (dtFillCheckMapping.Rows.Count > 0)
                            {
                                string __top = (string)(from r in ((DataTable)cbTOP.Properties.DataSource).Rows.Cast<DataRow>()
                                                        where r["code"].ToString() == dtFillCheckMapping.Rows[0]["smt_top_id"].ToString().Trim()
                                                        select r["code"].ToString())
                                                .ToList()
                                                .FirstOrDefault();
                                if (!__top.IsNullOrEmptyOrWhiteSpace())
                                    cbTOP.EditValue = __top;


                            }
                            else
                            {

                                txtKdOutlet1.Text = "";
                                lblOutletDesc.Text = "";
                                lblKdOutlet2.Text = "";
                                txtStatusOutlet.Text = "";
                                cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                                MessageBox.Show("TOP Belum di Setting", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;

                            }
                        }

                    }
                    #endregion
                    #region TOP BY CUSTOMER
                    else
                    {
                        for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                            string cdPF = dtFill1.Rows[0]["cm_top_id"].ToString().Trim();
                            string vTOP = Convert.ToString(cbTOP.EditValue);
                            if (Convert.ToString(cbTOP.EditValue) == cdPF)
                            {
                                g_bEdit = true;
                                break;
                            }
                        }

                        if (g_bEdit == false)
                        {
                            MessageBox.Show("Default TOP tidak ada untuk customer  " + dtFill1.Rows[0]["cm_cust_code1"].ToString(), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            txtKdOutlet1.Text = "";
                            lblOutletDesc.Text = "";
                            lblKdOutlet2.Text = "";
                            txtStatusOutlet.Text = "";
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                            return;
                        }
                    }
                    #endregion

                    initDateDelvExp(Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd"), txtKdOutlet1.Text, lblKdOutlet2.Text, 1);
                    g_bChangeOutlet = false;
                    //}

                }
                else
                {
                    txtKdOutlet1.Text = "";
                    lblOutletDesc.Text = "";
                    lblKdOutlet2.Text = "";
                    txtStatusOutlet.Text = "";
                    cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                }
            }

        }

        #endregion

        #region Function

        private GridView GetGridView(DataGridView dgv)
        {
            if (dgv == null) return null;
            return dgv.MainView as GridView;
        }

        private void EnsureReasonSearchLookUpEditor()
        {
            try
            {
                GridView view = GetGridView(dgvSalesDetail);
                if (view == null) return;

                GridColumn reasonColumn = view.Columns.ColumnByFieldName("reason") ?? view.Columns["reason"];
                if (reasonColumn == null) return;

                if (_repoReasonSearchLookUp == null)
                {
                    _repoReasonSearchLookUp = new RepositoryItemSearchLookUpEdit();
                    _repoReasonSearchLookUp.Name = "repoReasonSearchLookUp";
                    _repoReasonSearchLookUp.NullText = "";
                    _repoReasonSearchLookUp.DisplayMember = "reason";
                    _repoReasonSearchLookUp.ValueMember = "reason";
                    _repoReasonSearchLookUp.TextEditStyle = TextEditStyles.Standard;
                    _repoReasonSearchLookUp.PopupFilterMode = PopupFilterMode.Contains;
                    _repoReasonSearchLookUp.ImmediatePopup = true;
                    _repoReasonSearchLookUp.PopupFormSize = new Size(420, 280);

                    GridView reasonView = new GridView();
                    reasonView.OptionsView.ShowGroupPanel = false;
                    reasonView.OptionsView.ShowIndicator = false;
                    reasonView.OptionsView.ColumnAutoWidth = false;
                    reasonView.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
                    reasonView.OptionsSelection.EnableAppearanceFocusedCell = false;
                    _repoReasonSearchLookUp.View = reasonView;
                }

                if (_dtReasonSearchLookUp == null)
                {
                    string sql = @"
SELECT '' AS reason
UNION ALL
SELECT gh_function_code + ' ~ ' + gh_function_desc AS reason
FROM GS_GEN_HARDCODED WITH (NOLOCK)
WHERE gh_function_name = 'REASONKAM'
ORDER BY reason";
                    _dtReasonSearchLookUp = _clsGlobal.ExecDT(sql);
                }

                _repoReasonSearchLookUp.DataSource = _dtReasonSearchLookUp;

                GridView popupView = _repoReasonSearchLookUp.View as GridView;
                if (popupView != null)
                {
                    popupView.Columns.Clear();
                    popupView.Columns.AddVisible("reason", "Reason");
                    popupView.Columns["reason"].Width = 380;
                    popupView.BestFitColumns();
                }

                if (!dgvSalesDetail.RepositoryItems.Contains(_repoReasonSearchLookUp))
                    dgvSalesDetail.RepositoryItems.Add(_repoReasonSearchLookUp);

                reasonColumn.ColumnEdit = _repoReasonSearchLookUp;
                reasonColumn.OptionsColumn.AllowEdit = true;
                reasonColumn.OptionsColumn.ReadOnly = false;

                if (!_isReasonSearchLookUpHooked)
                {
                    view.ShowingEditor += SalesDetailView_ShowingEditorReason;
                    _isReasonSearchLookUpHooked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void SalesDetailView_ShowingEditorReason(object sender, CancelEventArgs e)
        {
            try
            {
                GridView view = sender as GridView;
                if (view == null || view.FocusedColumn == null) return;

                if (!string.Equals(view.FocusedColumn.FieldName, "reason", StringComparison.OrdinalIgnoreCase))
                    return;

                if (!IsReasonAllowed(view.FocusedRowHandle))
                {
                    e.Cancel = true;
                    SetGridCellValueSafe(view.FocusedRowHandle, "reason", "");
                    return;
                }

                EnsureReasonSearchLookUpEditor();
            }
            catch
            {
                e.Cancel = true;
            }
        }

        private bool IsReasonAllowed(int rowHandle)
        {
            try
            {
                GridView view = GetGridView(dgvSalesDetail);
                if (view == null || rowHandle < 0) return false;

                object pcode = view.GetRowCellValue(rowHandle, "wsd_prd_master_code");
                if (string.IsNullOrWhiteSpace(Convert.ToString(pcode))) return false;

                string realQtyText = Convert.ToString(view.GetRowCellValue(rowHandle, "wsd_real_order_xqty"));
                string reqQtyText = Convert.ToString(view.GetRowCellValue(rowHandle, "wsd_req_order_xqty"));
                string bigText = Convert.ToString(view.GetRowCellValue(rowHandle, "wsd_uom_convert_big"));
                string midText = Convert.ToString(view.GetRowCellValue(rowHandle, "wsd_uom_convert_mid"));

                if (string.IsNullOrWhiteSpace(realQtyText) || string.IsNullOrWhiteSpace(reqQtyText)) return false;

                int big = 0;
                int mid = 0;
                int.TryParse(bigText, out big);
                int.TryParse(midText, out mid);

                fQtyFormat(realQtyText, big, mid);
                double qtyReal = Convert.ToDouble(_fConvertQty);

                fQtyFormat(reqQtyText, big, mid);
                double qtyReq = Convert.ToDouble(_fConvertQty);

                return qtyReq != qtyReal;
            }
            catch
            {
                return false;
            }
        }

        private void RefreshReasonEditState(int rowHandle)
        {
            try
            {
                bool allowReason = IsReasonAllowed(rowHandle);
                GridView view = GetGridView(dgvSalesDetail);
                if (view == null) return;

                GridColumn reasonColumn = view.Columns.ColumnByFieldName("reason") ?? view.Columns["reason"];
                if (reasonColumn != null)
                {
                    reasonColumn.OptionsColumn.AllowEdit = true;
                    reasonColumn.OptionsColumn.ReadOnly = false;
                    if (reasonColumn.ColumnEdit == null)
                        EnsureReasonSearchLookUpEditor();
                }

                if (!allowReason)
                    SetGridCellValueSafe(rowHandle, "reason", "");
            }
            catch
            {
            }
        }

        private void EnsureSubstitutionSearchLookUpEditor()
        {
            try
            {
                GridView view = GetGridView(dgvSalesDetail);
                if (view == null) return;

                GridColumn subsColumn = view.Columns.ColumnByFieldName(COL_SUBS_LOOKUP) ?? view.Columns[COL_SUBS_LOOKUP];
                if (subsColumn == null) return;
                GridColumn flagColumn = view.Columns.ColumnByFieldName(COL_IS_SUBS) ?? view.Columns[COL_IS_SUBS];
                if (flagColumn != null)
                {
                    flagColumn.OptionsColumn.AllowEdit = false;
                    flagColumn.OptionsColumn.ReadOnly = true;
                }

                if (_repoSubsSearchLookUp == null)
                {
                    _repoSubsSearchLookUp = new RepositoryItemSearchLookUpEdit();
                    _repoSubsSearchLookUp.Name = "repoSubsSearchLookUp";
                    _repoSubsSearchLookUp.NullText = "";
                    _repoSubsSearchLookUp.Buttons.Clear();
                    _repoSubsSearchLookUp.Buttons.Add(new EditorButton(ButtonPredefines.Search));
                    _repoSubsSearchLookUp.DisplayMember = "prds_child";
                    _repoSubsSearchLookUp.ValueMember = "prds_child";
                    _repoSubsSearchLookUp.TextEditStyle = TextEditStyles.Standard;
                    _repoSubsSearchLookUp.PopupFilterMode = PopupFilterMode.Contains;
                    _repoSubsSearchLookUp.ImmediatePopup = true;
                    _repoSubsSearchLookUp.PopupFormSize = new Size(760, 320);

                    GridView subsView = new GridView();
                    subsView.OptionsView.ShowGroupPanel = false;
                    subsView.OptionsView.ShowIndicator = false;
                    subsView.OptionsView.ColumnAutoWidth = false;
                    subsView.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
                    subsView.OptionsSelection.EnableAppearanceFocusedCell = false;
                    _repoSubsSearchLookUp.View = subsView;
                }

                if (!dgvSalesDetail.RepositoryItems.Contains(_repoSubsSearchLookUp))
                    dgvSalesDetail.RepositoryItems.Add(_repoSubsSearchLookUp);

                subsColumn.ColumnEdit = _repoSubsSearchLookUp;
                subsColumn.OptionsColumn.AllowEdit = true;
                subsColumn.OptionsColumn.ReadOnly = false;

                ConfigureSubstitutionPopupColumns();

                if (!_isSubsSearchLookUpHooked)
                {
                    view.ShowingEditor += SalesDetailView_ShowingEditorSubstitution;
                    _repoSubsSearchLookUp.QueryPopUp += RepoSubsSearchLookUp_QueryPopUp;
                    _isSubsSearchLookUpHooked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void ConfigureSubstitutionPopupColumns()
        {
            GridView popupView = _repoSubsSearchLookUp == null ? null : _repoSubsSearchLookUp.View as GridView;
            if (popupView == null) return;

            popupView.Columns.Clear();
            popupView.Columns.AddVisible("prds_entity_id", "Entity");
            popupView.Columns.AddVisible("prds_branch_id", "Branch");
            popupView.Columns.AddVisible("prds_line", "Line");
            popupView.Columns.AddVisible("prds_parent", "Parent");
            popupView.Columns.AddVisible("prds_child", "Child");
            popupView.Columns.AddVisible("prds_index", "Index");
            popupView.Columns.AddVisible("prds_valid_from", "Valid From");
            popupView.Columns.AddVisible("prds_valid_to", "Valid To");
            popupView.Columns.AddVisible("prds_mrp_ind", "MRP");
            popupView.BestFitColumns();
        }

        private void SalesDetailView_ShowingEditorSubstitution(object sender, CancelEventArgs e)
        {
            try
            {
                GridView view = sender as GridView;
                if (view == null || view.FocusedColumn == null) return;

                if (!string.Equals(view.FocusedColumn.FieldName, COL_SUBS_LOOKUP, StringComparison.OrdinalIgnoreCase))
                    return;

                RefreshSubstitutionFlag(view.FocusedRowHandle);
                if (!IsSubstitutionAllowed(view.FocusedRowHandle))
                    e.Cancel = true;
            }
            catch
            {
                e.Cancel = true;
            }
        }

        private void RepoSubsSearchLookUp_QueryPopUp(object sender, CancelEventArgs e)
        {
            try
            {
                GridView view = GetGridView(dgvSalesDetail);
                if (view == null) return;

                int rowHandle = view.FocusedRowHandle;
                if (!IsSubstitutionAllowed(rowHandle))
                {
                    e.Cancel = true;
                    return;
                }

                string parent = Convert.ToString(view.GetRowCellValue(rowHandle, COL_PCODE));
                string line = Convert.ToString(view.GetRowCellValue(rowHandle, "wsd_prd_line_code"));
                _repoSubsSearchLookUp.DataSource = GetSubstitutionLookup(parent, line);
                ConfigureSubstitutionPopupColumns();
            }
            catch
            {
                e.Cancel = true;
            }
        }

        private bool IsSubstitutionAllowed(int rowHandle)
        {
            GridView view = GetGridView(dgvSalesDetail);
            if (view == null || rowHandle < 0) return false;
            return string.Equals(Convert.ToString(view.GetRowCellValue(rowHandle, COL_IS_SUBS)), "Y", StringComparison.OrdinalIgnoreCase);
        }

        private void RefreshSubstitutionFlag(int rowHandle)
        {
            try
            {
                GridView view = GetGridView(dgvSalesDetail);
                if (view == null || rowHandle < 0) return;

                string parent = Convert.ToString(view.GetRowCellValue(rowHandle, COL_PCODE)).Trim();
                string line = Convert.ToString(view.GetRowCellValue(rowHandle, "wsd_prd_line_code")).Trim();
                string flag = HasSubstitutionParent(parent, line) ? "Y" : "N";
                view.SetRowCellValue(rowHandle, COL_IS_SUBS, flag);
                if (flag != "Y")
                    view.SetRowCellValue(rowHandle, COL_SUBS_LOOKUP, "");
            }
            catch { }
        }

        private void RefreshAllSubstitutionFlags()
        {
            GridView view = GetGridView(dgvSalesDetail);
            if (view == null) return;

            for (int i = 0; i < view.DataRowCount; i++)
                RefreshSubstitutionFlag(i);
        }

        private bool HasSubstitutionParent(string parent, string line)
        {
            if (string.IsNullOrWhiteSpace(parent) || string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(xEntityID) || string.IsNullOrWhiteSpace(xBranchID))
                return false;

            string sql = "select top 1 1 from TBL_PRD_SUBS WITH(NOLOCK) " +
                "where prds_entity_id = '" + xEntityID.Replace("'", "''") + "' " +
                "and prds_branch_id = '" + xBranchID.Replace("'", "''") + "' " +
                "and prds_line = '" + line.Replace("'", "''") + "' " +
                "and prds_parent = '" + parent.Replace("'", "''") + "'";
            DataTable dt = (_clsGlobal.Connect.State == ConnectionState.Closed) ? _clsGlobal.ExecDT(sql) : _clsGlobal.ExecDTTrans(sql);
            return dt.Rows.Count > 0;
        }

        private DataTable GetSubstitutionLookup(string parent, string line)
        {
            string sql = "select prds_entity_id, prds_branch_id, prds_line, prds_parent, prds_child, prds_index, prds_valid_from, prds_valid_to, prds_mrp_ind " +
                "from TBL_PRD_SUBS WITH(NOLOCK) " +
                "where prds_entity_id = '" + xEntityID.Replace("'", "''") + "' " +
                "and prds_branch_id = '" + xBranchID.Replace("'", "''") + "' " +
                "and prds_line = '" + Convert.ToString(line).Replace("'", "''") + "' " +
                "and prds_parent = '" + Convert.ToString(parent).Replace("'", "''") + "' " +
                "order by prds_index, prds_child";
            return (_clsGlobal.Connect.State == ConnectionState.Closed) ? _clsGlobal.ExecDT(sql) : _clsGlobal.ExecDTTrans(sql);
        }

        private void SetGridCellValueSafe(int rowHandle, string fieldName, object value)
        {
            try
            {
                GridView view = GetGridView(dgvSalesDetail);
                if (view != null && rowHandle >= 0)
                    view.SetRowCellValue(rowHandle, fieldName, value);
            }
            catch
            {
            }
        }

        private void SetGridSODetail(DataGridView dgv)
        {
            dtGridSODetail.Columns.Add("wsd_prd_master_code", typeof(string));
            dtGridSODetail.Columns.Add("wsd_real_order_amt", typeof(string));
            dtGridSODetail.Columns.Add("wsd_total_so_amount", typeof(string));
            dtGridSODetail.Columns.Add("wsd_grade", typeof(string));
            dtGridSODetail.Columns.Add("wsd_prd_size", typeof(string));
            dtGridSODetail.Columns.Add("wsd_uom_convert_big", typeof(string));
            dtGridSODetail.Columns.Add("wsd_uom_convert_mid", typeof(string));
            dtGridSODetail.Columns.Add("prm_prd_desc", typeof(string));
            dtGridSODetail.Columns.Add("wsd_het_unit_price", typeof(string));
            dtGridSODetail.Columns.Add("wsd_real_order_xqty", typeof(string));
            dtGridSODetail.Columns.Add("wsd_so_xqty", typeof(string));
            dtGridSODetail.Columns.Add("ijumlahharga", typeof(string));
            dtGridSODetail.Columns.Add("wsd_tot_disc_pc", typeof(string));
            dtGridSODetail.Columns.Add("tgl_price", typeof(string));
            dtGridSODetail.Columns.Add("jml_rp_netto", typeof(string));
            dtGridSODetail.Columns.Add("tpr_flag", typeof(string));
            dtGridSODetail.Columns.Add("cash_disc", typeof(string));
            dtGridSODetail.Columns.Add("gross_net", typeof(string));
            dtGridSODetail.Columns.Add("tax_code", typeof(string));
            dtGridSODetail.Columns.Add("pct_tax", typeof(string));
            dtGridSODetail.Columns.Add("rp_tax", typeof(string));
            dtGridSODetail.Columns.Add("discpc", typeof(string));
            dtGridSODetail.Columns.Add("pctdiscpc", typeof(string));
            dtGridSODetail.Columns.Add("disc1", typeof(string));
            dtGridSODetail.Columns.Add("disc2", typeof(string));
            dtGridSODetail.Columns.Add("reason", typeof(string));
            dtGridSODetail.Columns.Add("wsd_sap_unit_price", typeof(string));
            dtGridSODetail.Columns.Add("wsd_sled", typeof(string));
            dtGridSODetail.Columns.Add("wsd_req_order_xqty", typeof(string));//29
            dtGridSODetail.Columns.Add("wsd_NetSales", typeof(string));
            dtGridSODetail.Columns.Add("wsd_prd_line_code", typeof(string));
            dtGridSODetail.Columns.Add(COL_IS_SUBS, typeof(string));
            dtGridSODetail.Columns.Add(COL_SUBS_LOOKUP, typeof(string));
            dtGridSODetail.Columns[COL_IS_SUBS].DefaultValue = "N";
            dtGridSODetail.Columns[COL_SUBS_LOOKUP].DefaultValue = "";
            dgv.UseDesignTimeColumns = true;
            dgv.DataSource = dtGridSODetail;


            //hide column


            if (isNefoKAM || IsNefoForDC)
            {
            }
            else
            {
            }




            //dgv.Columns["wsd_prd_master_code"].ReadOnly = true;
            dgv.Columns["prm_prd_desc"].ReadOnly = true;
            dgv.Columns["wsd_het_unit_price"].ReadOnly = true;
            dgv.Columns["wsd_so_xqty"].ReadOnly = true;
            dgv.Columns["ijumlahharga"].ReadOnly = true;
            dgv.Columns["wsd_tot_disc_pc"].ReadOnly = true;
            dgv.Columns["reason"].ReadOnly = true;

            if (chkSLED.Checked)
            {
                dgv.Columns["wsd_sled"].ReadOnly = false;
            }
            else
            {
                dgv.Columns["wsd_sled"].ReadOnly = true;
            }

            dgv.Columns["wsd_NetSales"].ReadOnly = true;
            dgv.Columns["wsd_prd_line_code"].ReadOnly = true;
            dgv.Columns[COL_IS_SUBS].ReadOnly = true;




            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //dgv.Columns["wsd_prd_master_code"].SortMode = DataGridViewColumnSortMode.NotSortable;


            EnsureReasonSearchLookUpEditor();
            EnsureSubstitutionSearchLookUpEditor();

            //width header
            dgv.Columns["wsd_het_unit_price"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["ijumlahharga"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["jml_rp_netto"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgv.Columns["wsd_so_xqty"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["wsd_real_order_xqty"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["wsd_req_order_xqty"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgv.Columns["wsd_tot_disc_pc"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgv.Columns["opl_validity_date_to"].Width = 125;

            //alignment cell


            ////type data cell            


            dgv.Columns["reason"].DisplayIndex = 28;
            EnsureReasonSearchLookUpEditor();
            //dgv.Columns["wsd_NetSales"].DisplayIndex = 28;
        }

        private void SetGridSODetailCss(DataGridView dgv)
        {
            //hide column
            dgv.Columns["wsd_real_order_amt"].Visible = false;
            dgv.Columns["wsd_total_so_amount"].Visible = false;
            dgv.Columns["wsd_grade"].Visible = false;
            dgv.Columns["wsd_prd_size"].Visible = false;
            dgv.Columns["wsd_uom_convert_big"].Visible = false;
            dgv.Columns["wsd_uom_convert_mid"].Visible = false;
            dgv.Columns["wsd_prd_master_code"].Visible = true;
            dgv.Columns["prm_prd_desc"].Visible = true;
            dgv.Columns["wsd_het_unit_price"].Visible = true;
            dgv.Columns["wsd_real_order_xqty"].Visible = true;
            dgv.Columns["wsd_so_xqty"].Visible = true;
            dgv.Columns["ijumlahharga"].Visible = true;
            dgv.Columns["wsd_tot_disc_pc"].Visible = true;

            dgv.Columns["tgl_price"].Visible = true;
            dgv.Columns["jml_rp_netto"].Visible = true;
            dgv.Columns["tpr_flag"].Visible = false;
            dgv.Columns["cash_disc"].Visible = false;
            dgv.Columns["gross_net"].Visible = false;
            dgv.Columns["tax_code"].Visible = false;
            dgv.Columns["pct_tax"].Visible = false;
            dgv.Columns["rp_tax"].Visible = false;
            dgv.Columns["discpc"].Visible = false;
            dgv.Columns["pctdiscpc"].Visible = false;
            dgv.Columns["disc1"].Visible = false;
            dgv.Columns["disc2"].Visible = false;
            dgv.Columns["reason"].Visible = true;
            dgv.Columns["wsd_sled"].Visible = true;
            dgv.Columns["wsd_req_order_xqty"].Visible = true;
            dgv.Columns["wsd_NetSales"].Visible = true;
            dgv.Columns["wsd_prd_line_code"].Visible = false;
            dgv.Columns[COL_IS_SUBS].Visible = true;
            dgv.Columns[COL_SUBS_LOOKUP].Visible = true;

            //dgv.Columns["wsd_prd_master_code"].ReadOnly = true; commented 23/01/19
            dgv.Columns["prm_prd_desc"].ReadOnly = true;
            dgv.Columns["wsd_het_unit_price"].ReadOnly = true;
            //dgv.Columns["wsd_real_order_xqty"].Visible = true;
            dgv.Columns["wsd_so_xqty"].ReadOnly = true;
            dgv.Columns["ijumlahharga"].ReadOnly = true;
            dgv.Columns["wsd_tot_disc_pc"].ReadOnly = true;
            dgv.Columns["jml_rp_netto"].ReadOnly = true;
            dgv.Columns["reason"].ReadOnly = true;

            if (chkSLED.Checked)
            {
                dgv.Columns["wsd_sled"].ReadOnly = false;
            }
            else
            {
                dgv.Columns["wsd_sled"].ReadOnly = true;
            }

            dgv.Columns["wsd_NetSales"].ReadOnly = true;
            dgv.Columns["wsd_prd_line_code"].ReadOnly = true;
            dgv.Columns[COL_IS_SUBS].ReadOnly = true;

            dgv.Columns["wsd_prd_master_code"].HeaderText = "PCODE";
            dgv.Columns["wsd_real_order_amt"].HeaderText = "real ordera";
            dgv.Columns["wsd_total_so_amount"].HeaderText = "ttl soa";
            dgv.Columns["wsd_grade"].HeaderText = "grade";
            dgv.Columns["wsd_prd_size"].HeaderText = "size";
            dgv.Columns["wsd_uom_convert_big"].HeaderText = "conv1";
            dgv.Columns["wsd_uom_convert_mid"].HeaderText = "conv2";
            dgv.Columns["prm_prd_desc"].HeaderText = "Nama Barang";
            dgv.Columns["wsd_het_unit_price"].HeaderText = "Harga";
            dgv.Columns["wsd_real_order_xqty"].HeaderText = "RealOrder Qty";
            dgv.Columns["wsd_so_xqty"].HeaderText = "Order Qty";
            dgv.Columns["ijumlahharga"].HeaderText = "Jumlah Harga";
            dgv.Columns["wsd_tot_disc_pc"].HeaderText = "Discount";

            dgv.Columns["tgl_price"].HeaderText = "Tgl Price";
            dgv.Columns["jml_rp_netto"].HeaderText = "Jml Rp Netto";
            dgv.Columns["tpr_flag"].HeaderText = "TPR Flag";
            dgv.Columns["cash_disc"].HeaderText = "Cash Disc";
            dgv.Columns["gross_net"].HeaderText = "Gross Net";
            dgv.Columns["tax_code"].HeaderText = "TAX Code";
            dgv.Columns["pct_tax"].HeaderText = "Pct TAX";
            dgv.Columns["rp_tax"].HeaderText = "Rp TAX";
            dgv.Columns["discpc"].HeaderText = "Disc pc";
            dgv.Columns["pctdiscpc"].HeaderText = "Pct DiscPC";
            dgv.Columns["disc1"].HeaderText = "Disc1";
            dgv.Columns["disc2"].HeaderText = "Disc2";
            dgv.Columns["reason"].HeaderText = "Reason";
            EnsureReasonSearchLookUpEditor();
            dgv.Columns["wsd_sled"].HeaderText = "Sled";
            dgv.Columns["wsd_req_order_xqty"].HeaderText = "Qty Req";
            dgv.Columns["wsd_NetSales"].HeaderText = "Net Sales";
            dgv.Columns["wsd_prd_line_code"].HeaderText = "Product Line";
            dgv.Columns[COL_IS_SUBS].HeaderText = "Subs";
            dgv.Columns[COL_SUBS_LOOKUP].HeaderText = "";
            EnsureSubstitutionSearchLookUpEditor();
            //width header
            dgv.Columns["wsd_het_unit_price"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["ijumlahharga"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["jml_rp_netto"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgv.Columns["wsd_so_xqty"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["wsd_real_order_xqty"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["wsd_req_order_xqty"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["wsd_tot_disc_pc"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgv.Columns["opl_validity_date_to"].Width = 125;

            //alignment cell
            dgv.Columns["wsd_het_unit_price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["ijumlahharga"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["jml_rp_netto"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dgv.Columns["wsd_so_xqty"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["wsd_real_order_xqty"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["wsd_req_order_xqty"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            dgv.Columns["wsd_tot_disc_pc"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["wsd_NetSales"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["wsd_prd_line_code"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            ////type data cell            
            dgv.Columns["wsd_het_unit_price"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["ijumlahharga"].DefaultCellStyle.Format = "#,##0.00";
            dgv.Columns["jml_rp_netto"].DefaultCellStyle.Format = "#,##0.00";
            //dgv.Columns["wsd_so_xqty"].DefaultCellStyle.Format = "0.00##";
            //dgv.Columns["wsd_real_order_xqty"].DefaultCellStyle.Format = "0.00##";

            //dgv.Columns["wsd_tot_disc_pc"].DefaultCellStyle.Format = "#,##0.0000";
            dgv.Columns["tgl_price"].DefaultCellStyle.Format = "dd MMM yyyy";
            dgv.Columns["wsd_NetSales"].DefaultCellStyle.Format = "#,##0.00";
        }

        private void SetGridTPRB(DataGridView dgv)
        {
            dtGridTPRB.Columns.Add("wsd_prd_master_code", typeof(string));
            dtGridTPRB.Columns.Add("prm_prd_desc", typeof(string));
            dtGridTPRB.Columns.Add("wsd_prd_size", typeof(string));
            dtGridTPRB.Columns.Add("wsd_uom_convert_big", typeof(string));
            dtGridTPRB.Columns.Add("wsd_uom_convert_mid", typeof(string));
            dtGridTPRB.Columns.Add("wsd_het_unit_price", typeof(string));
            dtGridTPRB.Columns.Add("wsd_so_xqty", typeof(string));

            // Kolom didefinisikan di Designer (gvPromosiQty) -> lewati auto-populate.
            dgv.UseDesignTimeColumns = true;
            dgv.DataSource = dtGridTPRB;
        }

        private void SetGridTPRBCss(DataGridView dgv)
        {

            dgv.Columns["wsd_prd_master_code"].HeaderText = "PCODE";
            dgv.Columns["prm_prd_desc"].HeaderText = "Nama Barang";
            dgv.Columns["wsd_het_unit_price"].HeaderText = "Harga";
            dgv.Columns["wsd_so_xqty"].HeaderText = "Qty Promosi";

            //hide column
            dgv.Columns["wsd_prd_size"].Visible = false;
            dgv.Columns["wsd_uom_convert_big"].Visible = false;
            dgv.Columns["wsd_uom_convert_mid"].Visible = false;

            //width header
            dgv.Columns["wsd_het_unit_price"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgv.Columns["opl_validity_date_to"].Width = 125;

            //alignment cell
            dgv.Columns["wsd_het_unit_price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            ////type data cell            
            dgv.Columns["wsd_het_unit_price"].DefaultCellStyle.Format = "#,##0.00";
        }

        private void SetGridTPRU(DataGridView dgv)
        {
            //if (isNefoKAM)
            //{
            dtGridTPRU.Columns.Add("pt_promo_id", typeof(string));
            dtGridTPRU.Columns.Add("pt_product_id", typeof(string));
            dtGridTPRU.Columns.Add("prm_prd_desc", typeof(string));
            dtGridTPRU.Columns.Add("pt_amount", typeof(string));
            dtGridTPRU.Columns.Add("pt_pct_promo", typeof(string));

            dgv.UseDesignTimeColumns = true;
            dgv.DataSource = dtGridTPRU;




            //}
            //else
            //{

            //    dtGridTPRU.Columns.Add("pt_product_id", typeof(string));
            //    dtGridTPRU.Columns.Add("prm_prd_desc", typeof(string));
            //    dtGridTPRU.Columns.Add("pt_amount", typeof(string));

            //    dgv.DataSource = dtGridTPRU;


            //    hide column

            //    width header

            //    alignment cell
            //    //type data cell            
            //}

        }

        private void SetGridTPRUCss(DataGridView dgv)
        {

            dgv.Columns["pt_product_id"].HeaderText = "PCODE";
            dgv.Columns["prm_prd_desc"].HeaderText = "Nama Barang";
            dgv.Columns["pt_amount"].HeaderText = "Promosi Rp";

            //hide column
            dgv.Columns["pt_product_id"].Visible = true;
            dgv.Columns["prm_prd_desc"].Visible = true;
            dgv.Columns["pt_amount"].Visible = true;

            //width header
            dgv.Columns["pt_amount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgv.Columns["opl_validity_date_to"].Width = 125;

            //alignment cell
            dgv.Columns["pt_amount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            ////type data cell            
            dgv.Columns["pt_amount"].DefaultCellStyle.Format = "#,##0.00";
        }

        private void SetGridDisc(DataGridView dgv)
        {
            dtGridDisc.Columns.Add("tdt_product_master_id", typeof(string));
            dtGridDisc.Columns.Add("g_iDiscCode1", typeof(string));
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


            dgv.UseDesignTimeColumns = true;
            dgv.DataSource = dtGridDisc;




        }

        private void SetGridDiscCss(DataGridView dgv)
        {

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
            dgv.Columns["g_iDiscCode1"].Visible = false;

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

        //1 Juni 2018
        private void SetGridTax(DataGridView dgv)
        {
            dtGridTax.Columns.Add("wsd_prd_master_code", typeof(string));
            dtGridTax.Columns.Add("grade", typeof(string));
            dtGridTax.Columns.Add("size", typeof(string));
            dtGridTax.Columns.Add("TPRCode", typeof(string));
            dtGridTax.Columns.Add("conv1", typeof(string));
            dtGridTax.Columns.Add("conv2", typeof(string));
            dtGridTax.Columns.Add("TPRType", typeof(string));
            dtGridTax.Columns.Add("jml_promosi", typeof(string));

            dgv.DataSource = dtGridTax;

            //dgv.Columns["wsd_prd_master_code"].HeaderText = "PCODE";
            //dgv.Columns["prm_prd_desc"].HeaderText = "Nama Barang";
            //dgv.Columns["wsd_het_unit_price"].HeaderText = "Harga";
            //dgv.Columns["wsd_so_xqty"].HeaderText = "Qty Promosi";

            //hide column
            //dgv.Columns["wsd_prd_size"].Visible = false;
            //dgv.Columns["wsd_uom_convert_big"].Visible = false;
            //dgv.Columns["wsd_uom_convert_mid"].Visible = false;

            ////width header
            //dgv.Columns["jml_promosi"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ////dgv.Columns["opl_validity_date_to"].Width = 125;

            ////alignment cell
            //dgv.Columns["jml_promosi"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //////type data cell            
            //dgv.Columns["jml_promosi"].DefaultCellStyle.Format = "#,##0.0000";
        }

        //1 Juni 2018
        private void SetGridPromoTax(DataGridView dgv)
        {
            dtGridPromoTax.Columns.Add("wsd_prd_master_code", typeof(string));
            dtGridPromoTax.Columns.Add("grade", typeof(string));
            dtGridPromoTax.Columns.Add("size", typeof(string));
            dtGridPromoTax.Columns.Add("TPRCode", typeof(string));
            dtGridPromoTax.Columns.Add("conv1", typeof(string));
            dtGridPromoTax.Columns.Add("conv2", typeof(string));
            dtGridPromoTax.Columns.Add("TPRType", typeof(string));
            dtGridPromoTax.Columns.Add("TPRQty", typeof(string));
            dtGridPromoTax.Columns.Add("TPRPrice", typeof(string));
            dtGridPromoTax.Columns.Add("TPRAmount", typeof(string));
            dtGridPromoTax.Columns.Add("TPRPct", typeof(string));

            dgv.DataSource = dtGridPromoTax;

            //dgv.Columns["wsd_prd_master_code"].HeaderText = "PCODE";
            //dgv.Columns["prm_prd_desc"].HeaderText = "Nama Barang";
            //dgv.Columns["wsd_het_unit_price"].HeaderText = "Harga";
            //dgv.Columns["wsd_so_xqty"].HeaderText = "Qty Promosi";

            //hide column
            //dgv.Columns["wsd_prd_size"].Visible = false;
            //dgv.Columns["wsd_uom_convert_big"].Visible = false;
            //dgv.Columns["wsd_uom_convert_mid"].Visible = false;

            //width header
            //dgv.Columns["TPRAmount"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            ////dgv.Columns["opl_validity_date_to"].Width = 125;

            ////alignment cell
            //dgv.Columns["TPRAmount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //////type data cell            
            //dgv.Columns["TPRAmount"].DefaultCellStyle.Format = "#,##0.0000";
        }

        private void ShowMe(string sNoOrder)
        {
            try
            {
                DataTable dtFill1 = new DataTable();
                strSQL = " select wsh_seq_no, isnull(wsh_flag_so_limit, '') as wsh_flag_so_limit, ";
                strSQL = strSQL + " isnull(wsh_approval_rovi_by, '') as wsh_approval_rovi_by, ";
                strSQL = strSQL + " isnull(wsh_approval_rovi_date, '') as wsh_approval_rovi_date, wsh_entity_id, wsh_branch_id ";
                strSQL = strSQL + " from SO_WEB_SALES_HEADER WITH (NOLOCK) ";
                strSQL = strSQL + " where wsh_seq_no = '" + xSONo + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "'";
                dtFill1 = _clsGlobal.ExecDT(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    pRefresh(true, false);
                    _xEntityID = dtFill1.Rows[0]["wsh_entity_id"].ToString();
                    _xBranchID = dtFill1.Rows[0]["wsh_branch_id"].ToString();
                    _clsGlobal.BeginTrans();
                    pFillData(xSONo, _xEntityID, _xBranchID, false);
                    _clsGlobal.CommitTrans();
                    strFlagLimit = dtFill1.Rows[0]["wsh_flag_so_limit"].ToString();
                    strRoviUpdUser = dtFill1.Rows[0]["wsh_approval_rovi_by"].ToString();
                    strRoviUpdDate = dtFill1.Rows[0]["wsh_approval_rovi_date"].ToString();
                    txtNoOrder.Enabled = false;
                    txtStatusOutlet.Enabled = false;
                    dtTglOrder.Enabled = false;

                }
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindTOP()
        {
            try
            {

                if (g_bTOPFlag == true)
                {
                    DataTable dtFill1 = new DataTable();
                    strSQL = "SELECT DISTINCT pptc_no_of_days, pptc_term_code as code, pptc_term_code + ' ~ ' + CAST(pptc_no_of_days AS VARCHAR) + ' - ' + pptc_term_desc AS pptc_term_des";
                    strSQL = strSQL + " FROM SO_SPG_GIRL_MAN SPG WITH (NOLOCK) ";
                    strSQL = strSQL + " JOIN TBL_SD_SPGGROUP_PRDLINE WITH (NOLOCK) ON sgm_spgm_group=ssp_spggroup_code";
                    strSQL = strSQL + " JOIN TBL_SD_MAP_TOP_PRDLINE WITH (NOLOCK) ON ssp_prdline_code = map_prdline_code";
                    strSQL = strSQL + " JOIN PO_PAYMENT_TERM_CODES WITH (NOLOCK) ON map_term_code=pptc_term_code";
                    strSQL = strSQL + " WHERE sgm_spgm_id=" + _clsGlobal.FmtStr(xSalesID) + " and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                    strSQL = strSQL + " order by pptc_no_of_days ASC";

                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill1.Rows.Count > 0)
                    {
                        //cbTOP.Items.Clear();
                        cbTOP.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                        cbTOP.Properties.ValueMember = "code";
                        cbTOP.Properties.DisplayMember = "pptc_term_des";
                        cbTOP.Properties.ForceInitialize();
                    }

                    cbTOP.Enabled = true;

                }
                else
                {
                    DataTable dtFill2 = new DataTable();
                    strSQL = " Select pptc_no_of_days, pptc_term_code as code, pptc_term_code + ' ~ ' + CAST(pptc_no_of_days AS VARCHAR) + ' - ' + pptc_term_desc AS pptc_term_desc " +
                        " FROM PO_PAYMENT_TERM_CODES WITH (NOLOCK)   order by pptc_no_of_days ";
                    dtFill2 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill2.Rows.Count > 0)
                    {
                        //cbTOP.Properties.DataSource = null;
                        //cbTOP.Items.Clear();
                        cbTOP.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                        cbTOP.Properties.ValueMember = "code";
                        cbTOP.Properties.DisplayMember = "pptc_term_desc";
                        cbTOP.Properties.ForceInitialize();
                    }

                    cbTOP.Enabled = false;
                }


            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindTOPFrmSalesman()
        {
            try
            {
                DataTable dtFill = new DataTable();
                DataTable dtFill2 = new DataTable();
                DataTable dtFill3 = new DataTable();
                strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
                //strSQL = "select gh_function_name from GS_GEN_HARDCODED where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";

                dtFill = _clsGlobal.ExecDT(strSQL);

                if (dtFill.Rows.Count > 0)
                {

                    strSQL = " Select pptc_no_of_days, pptc_term_code as code, pptc_term_code + ' ~ ' + CAST(pptc_no_of_days AS VARCHAR) + ' - ' + pptc_term_desc AS pptc_term_desc " +
                        " FROM PO_PAYMENT_TERM_CODES WITH (NOLOCK)   order by pptc_no_of_days ";
                    dtFill2 = _clsGlobal.ExecDT(strSQL);
                    if (dtFill2.Rows.Count > 0)
                    {
                        cbTOP.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                        cbTOP.Properties.ValueMember = "code";
                        cbTOP.Properties.DisplayMember = "pptc_term_desc";
                        cbTOP.Properties.ForceInitialize();

                        strSQL = " select sgl_group_top, sgm_spgm_id from SO_SPG_GIRL_MAN WITH (NOLOCK)  Inner join SO_SPG_GROUP WITH (NOLOCK)  on sgm_spgm_group = sgl_group_code " +
                        " where sgm_spgm_id='" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                        dtFill3 = _clsGlobal.ExecDT(strSQL);

                        if (dtFill3.Rows.Count > 0)
                        {

                            for (int i = 0; i < LookupRowCount(cbTOP); i++)
                            {
                                cbTOP.EditValue = LookupValueAt(cbTOP, i);
                                //DataRow selectedDataRow = ((DataRowView)cbTOP.GetSelectedDataRow()).Row;
                                string idO = Convert.ToString(cbTOP.EditValue);
                                //string NameO = selectedDataRow["pptc_term_desc"].ToString();
                                if (idO == dtFill3.Rows[0]["sgl_group_top"].ToString().Trim())
                                {
                                    cbTOP.Enabled = false;
                                    break;
                                }
                            }



                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void BindBebanDisc()
        {
            string __arr = "";
            string strSQL = "";
            strSQL = " select '0' as code,'' as Descr UNION ALL select gh_function_code as code, gh_function_desc as Descr from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'T_BEBAN_DISC'";
            DataTable dt = _clsGlobal.ExecDT(strSQL);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    __arr += dt.Rows[i]["Descr"].ToString() + "|";
                }

                __arr = Microsoft.VisualBasic.Strings.Left(__arr, __arr.Length - 1);

                arr_BebanDisc = __arr.ToString().Split('|');
            }
            cmbBebanDisc.Properties.ValueMember = "code";
            cmbBebanDisc.Properties.DisplayMember = "Descr";
            cmbBebanDisc.Properties.DataSource = dt;
        }

        private void BindPembyFaktur()
        {
            DataTable dtFill2 = new DataTable();
            //strSQL = "select '0' as code, '' as gh_function_code union all Select gh_function_code as code,gh_function_code + ' ~ ' + gh_function_desc gh_function_code from GS_GEN_HARDCODED " +
            //         "  where gh_function_name = 'CUSTPAYTYPE'"; //g_sPayType
            strSQL = "Select gh_function_code as code,gh_function_code + ' ~ ' + gh_function_desc gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) " +
                     "  where gh_function_name = 'CUSTPAYTYPE' order by gh_sequence_no"; //g_sPayType
            dtFill2 = _clsGlobal.ExecDT(strSQL);
            if (dtFill2.Rows.Count > 0)
            {
                cbPembyFaktur.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                cbPembyFaktur.Properties.ValueMember = "code";
                cbPembyFaktur.Properties.DisplayMember = "gh_function_code";
                cbPembyFaktur.Properties.ForceInitialize();
            }
        }

        private void pFillData(string sNoOrder, string sEntitys, string sBranchs, bool bCalcDisc)
        {
            try
            {

                decimal curJmlHarga = 0;
                bool bHiddenCol = false;

                DataTable dtFill1 = new DataTable();
                strSQL = " select wsh_seq_no, wsh_so_date, wsh_so_type, ";
                strSQL = strSQL + " wsh_cust_code1, wsh_cust_code2, wsh_loc_id1, wsh_loc_id2, ";
                strSQL = strSQL + " wsh_spgm_id, wsh_status_so, wsh_net_amount, wsh_tax_amount, ";
                strSQL = strSQL + " wsh_tot_tpr_amt , wsh_pct_cash_disc, wsh_cash_disc, isnull(wsh_pay_type, '') wsh_pay_type, ";
                strSQL = strSQL + " isnull(cm_cust_name, '') cm_cust_name, isnull(gh_function_desc, '') outlet_status, ";
                strSQL = strSQL + " isnull(sgm_spgm_name, '') sgm_spgm_name, isnull(sgm_type_operasi, '') sgm_type_operasi, ";
                strSQL = strSQL + " isnull(convert(varchar(12), wh_loc_last_work_date, 106), '') wh_loc_last_work_date, isnull(wsh_validate, getdate()) wsh_validate, ";
                strSQL = strSQL + " isnull(wsh_po_no, '') wsh_po_no, isnull(wsh_so_type, '') wsh_so_type, ";
                strSQL = strSQL + " isnull(wsh_employee, '') wsh_employee, isnull(wsh_top_id, '') as wsh_top_id, ";
                strSQL = strSQL + " isnull(cm_flag_blacklist, 'N') as cm_flag_blacklist,wsh_sch_date,wsh_expired_date,wsh_beban_cashdisc,wsh_po_date,wsh_ret_seq_no,wsh_note,isnull(wsh_extract_sap,'N') wsh_extract_sap,wsh_qasir_flag,wsh_flag_sloc,isnull(wsh_is_fulfillment,'N') wsh_is_fulfillment, isnull(wsh_sled_flag,'N') wsh_sled_flag, wsh_sled ";
                //if (isMultisource)
                //{
                //    strSQL = strSQL + "wsh_qasir"
                //}
                strSQL = strSQL + " from SO_WEB_SALES_HEADER WITH (NOLOCK) ";
                strSQL = strSQL + " left join SO_CUST_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + " on wsh_cust_code1 = cm_cust_code1";
                strSQL = strSQL + " and wsh_cust_code2 = cm_cust_code2";
                strSQL = strSQL + " and wsh_entity_id = cm_entity";
                strSQL = strSQL + " and wsh_branch_id = cm_branch";
                strSQL = strSQL + " left join ( select gh_function_code, gh_function_desc";
                strSQL = strSQL + " from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_sys = 'H'";
                strSQL = strSQL + " and gh_function_name = 'OUTLETSTATUS') GS on ";
                strSQL = strSQL + " cm_active_flag = gh_function_code";
                strSQL = strSQL + " left join SO_SPG_GIRL_MAN WITH (NOLOCK)  on wsh_spgm_id = sgm_spgm_id and wsh_entity_id = sgm_entity_id and wsh_branch_id = sgm_branch_id";
                strSQL = strSQL + " Left Join IM_WH_LOC WITH (NOLOCK) ";
                strSQL = strSQL + " on";
                strSQL = strSQL + " sgm_wh_loc1 = wh_loc_id1";
                strSQL = strSQL + " and sgm_wh_loc2 = wh_loc_id2";
                strSQL = strSQL + " and sgm_entity_id = wh_loc_entity";
                strSQL = strSQL + " and sgm_branch_id = wh_branch_id";
                strSQL = strSQL + " where wsh_seq_no = '" + sNoOrder + "'";
                if (sEntitys != "")
                {
                    strSQL = strSQL + " and wsh_entity_id = '" + sEntitys + "'";
                }

                if (sBranchs != "")
                {
                    strSQL = strSQL + " and wsh_branch_id = '" + sBranchs + "'";
                }
                if (_clsGlobal.tr.Connection == null)
                {
                    _clsGlobal.BeginTrans();
                }
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    //jika bukan open tidak bisa edit

                    //if (!bCalcDisc)
                    //{
                    _CurrTOP = dtFill1.Rows[0]["wsh_top_id"].ToString();
                    _CurrPembayaran = dtFill1.Rows[0]["wsh_pay_type"].ToString();
                    //}
                    if (dtFill1.Rows[0]["wsh_status_so"].ToString() == "O" || bCalcDisc == true)
                    {
                        if (dtFill1.Rows[0]["cm_flag_blacklist"].ToString() == "Y")
                        {
                            string nl = System.Environment.NewLine;
                            MessageBox.Show("Outlet : " + dtFill1.Rows[0]["wsh_cust_code1"].ToString() + " ber Status BlackList " + nl + "Proses tidak bisa dilanjutkan !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                        dgvSalesDetail.ReadOnly = false;
                        btnProses.Enabled = true;
                        btnCalcDisc.Enabled = true;
                        cmdSales.Enabled = true;
                        cbTipeOrder.Enabled = true;
                        txtKdOutlet1.Enabled = true;
                        btnPopUpOutlet.Enabled = true;
                        dtTglValidasi.Enabled = true;
                        cbPembyFaktur.Enabled = true;
                        cbTOP.Enabled = true;
                        txtNoPO.Enabled = true;
                        if (rbCashDiscP.Checked == true)
                        {
                            txtCashDiscP.Enabled = true;
                        }
                        else
                        {
                            txtCashDiscU.Enabled = true;
                        }
                        btnOk.Enabled = true;


                    }
                    else
                    {
                        cbSource.Enabled = false;
                        cbFlagSloc.Enabled = false;
                        dgvSalesDetail.ReadOnly = true;
                        btnProses.Enabled = false;
                        btnCalcDisc.Enabled = false;
                        cmdSales.Enabled = false;
                        cbTipeOrder.Enabled = false;
                        txtKdOutlet1.Enabled = false;
                        btnPopUpOutlet.Enabled = false;
                        dtTglValidasi.Enabled = false;
                        cbPembyFaktur.Enabled = false;
                        cbTOP.Enabled = false;
                        txtNoPO.Enabled = false;
                        if (rbCashDiscP.Checked == true)
                        {
                            txtCashDiscP.Enabled = false;
                        }
                        else
                        {
                            txtCashDiscU.Enabled = false;
                        }
                        btnOk.Enabled = false;
                        dgvSalesDetail.KeyUp -= this.dgvSalesDetail_KeyUp;
                    }

                    txtSalesID.Text = dtFill1.Rows[0]["wsh_spgm_id"].ToString();
                    txtSalesDesc.Text = dtFill1.Rows[0]["sgm_spgm_name"].ToString();
                    lblTipeSales.Text = dtFill1.Rows[0]["sgm_type_operasi"].ToString();
                    txtTglTrans.Text = dtFill1.Rows[0]["wh_loc_last_work_date"].ToString();
                    lblWHLoc1.Text = dtFill1.Rows[0]["wsh_loc_id1"].ToString();
                    lblWHLoc2.Text = dtFill1.Rows[0]["wsh_loc_id2"].ToString();

                    if (bCalcDisc == false)
                    {
                        txtNoOrder.Text = dtFill1.Rows[0]["wsh_seq_no"].ToString();
                    }

                    dtTglOrder.Text = dtFill1.Rows[0]["wsh_so_date"].ToString();
                    dtTglPO.Text = dtFill1.Rows[0]["wsh_po_date"].ToString();
                    dtTglValidasi.Text = dtFill1.Rows[0]["wsh_validate"].ToString();
                    txtKdOutlet1.Text = dtFill1.Rows[0]["wsh_cust_code1"].ToString();
                    lblOutletDesc.Text = dtFill1.Rows[0]["cm_cust_name"].ToString();
                    lblKdOutlet2.Text = dtFill1.Rows[0]["wsh_cust_code2"].ToString();
                    lblEmployee.Text = dtFill1.Rows[0]["wsh_employee"].ToString();
                    //***** Update Ticket  23708 *******//
                    txtRefNoRetur.Text = dtFill1.Rows[0]["wsh_ret_seq_no"].ToString();
                    //*****    END *******//
                    //***** Update Ticket 23816 ********//
                    txtNote.Text = dtFill1.Rows[0]["wsh_note"].ToString();
                    //**********************************//

                    //****** Multisource ***************//
                    if (dtFill1.Rows[0]["wsh_extract_sap"].ToString().ToUpper().Equals("N"))
                    {
                        cbResetReq.Enabled = false;
                        cbResetReq.Checked = false;
                    }
                    else
                    {
                        cbResetReq.Checked = true;
                        cbResetReq.Enabled = true;
                    }

                    if (dtFill1.Rows[0]["wsh_sled_flag"].ToString().ToUpper().Equals("N"))
                    {
                        chkSLED.Checked = false;
                    }
                    else
                    {
                        chkSLED.Checked = true;
                    }
                    //txtSLED.Text = dtFill1.Rows[0]["wsh_sled"].ToString();


                    if (isMultisource)
                    {
                        cbSource.EditValue = dtFill1.Rows[0]["wsh_qasir_flag"].ToString();
                        cbFlagSloc.EditValue = dtFill1.Rows[0]["wsh_flag_sloc"].ToString();
                    }
                    cbfullfilment.Checked = false;
                    if (dtFill1.Rows[0]["wsh_is_fulfillment"].ToString().Trim().ToUpper().Equals("Y"))
                    {
                        cbfullfilment.Checked = true;
                    }

                    //********** End *******************//

                    if (bCalcDisc == false)
                    {
                        if (dtFill1.Rows[0]["wsh_pay_type"].ToString() == "")
                        {
                            cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, 0);
                        }
                        else
                        {

                            g_bEdit = false;
                            for (lCounter = 0; lCounter < LookupRowCount(cbPembyFaktur); lCounter++)
                            {
                                cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, lCounter);
                                string vTOP = Convert.ToString(cbPembyFaktur.EditValue);
                                if (vTOP == _CurrPembayaran)
                                {
                                    g_bEdit = true;
                                    break;
                                }

                            }

                            if (g_bEdit == false)
                            {
                                MessageBox.Show("Tipe Pembayaran Faktur tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                        }

                        if (dtFill1.Rows[0]["wsh_top_id"].ToString().Trim() == "")
                        {
                            cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                        }
                        else
                        {
                            g_bEdit = false;
                            for (lCounter = 0; lCounter < LookupRowCount(cbTOP); lCounter++)
                            {
                                g_bTOPFlag = false;
                                cbTOP.EditValue = LookupValueAt(cbTOP, lCounter);
                                string vTOP = Convert.ToString(cbTOP.EditValue);
                                if (vTOP == _CurrTOP) //diganti, supaya tidak reset ketika calcdisk
                                {
                                    g_bEdit = true;
                                    break;
                                }
                            }

                            if (g_bEdit == false)
                            {
                                MessageBox.Show("TOP tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                txtKdOutlet1.Text = "";
                                lblKdOutlet2.Text = "";
                                txtStatusOutlet.Text = "";
                                lblOutletDesc.Text = "";
                                return;
                            }

                        }
                    }

                    DateTime sodate = Convert.ToDateTime(dtFill1.Rows[0]["wsh_so_date"].ToString().Trim());

                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["wsh_sch_date"].ToString().Trim()))
                    {

                        initDateDelvExp(sodate.ToString("yyyyMMdd"), dtFill1.Rows[0]["wsh_cust_code1"].ToString(), dtFill1.Rows[0]["wsh_cust_code2"].ToString(), 2);
                    }
                    else
                    {
                        DateTime TglDelv = Convert.ToDateTime(dtFill1.Rows[0]["wsh_sch_date"]);
                        DateTime TglExp = Convert.ToDateTime(dtFill1.Rows[0]["wsh_Expired_date"]);
                        if (TglDelv < g_dTglGudang)
                        {
                            //dtTglDelv.Text = g_dTglGudang.ToString();                           
                            if (clsGlobal.MODE_TRX == 1)
                            {
                                initDateDelvExp(sodate.ToString("yyyyMMdd"), dtFill1.Rows[0]["wsh_cust_code1"].ToString(), dtFill1.Rows[0]["wsh_cust_code2"].ToString(), 2);
                            }
                            else
                            {
                                dtTglDelv.Text = dtFill1.Rows[0]["wsh_sch_date"].ToString();
                                TglDeliveryDate = Convert.ToDateTime(dtFill1.Rows[0]["wsh_sch_date"]);
                            }

                        }
                        else
                        {
                            dtTglDelv.Text = dtFill1.Rows[0]["wsh_sch_date"].ToString();
                            TglDeliveryDate = Convert.ToDateTime(dtFill1.Rows[0]["wsh_sch_date"]);
                        }


                        if (TglExp < g_dTglGudang)
                        {
                            if (clsGlobal.MODE_TRX == 1)
                            {
                                initDateDelvExp(sodate.ToString("yyyyMMdd"), dtFill1.Rows[0]["wsh_cust_code1"].ToString(), dtFill1.Rows[0]["wsh_cust_code2"].ToString(), 2);
                            }
                            else
                            {
                                dtTglExp.Text = dtFill1.Rows[0]["wsh_Expired_date"].ToString();
                                TglExpiredDate = Convert.ToDateTime(dtFill1.Rows[0]["wsh_Expired_date"]);
                            }

                        }
                        else
                        {

                            dtTglExp.Text = dtFill1.Rows[0]["wsh_Expired_date"].ToString();
                            TglExpiredDate = Convert.ToDateTime(dtFill1.Rows[0]["wsh_Expired_date"]);
                        }


                    }



                    cmbBebanDisc.EditValue = dtFill1.Rows[0]["wsh_beban_cashdisc"].ToString();

                    txtStatusOutlet.Text = dtFill1.Rows[0]["outlet_status"].ToString();
                    txtNoPO.Text = dtFill1.Rows[0]["wsh_po_no"].ToString();
                    if (dtFill1.Rows[0]["wsh_qasir_flag"].ToString().Trim() == "S" || dtFill1.Rows[0]["wsh_qasir_flag"].ToString().Trim() == "W")
                    {
                        txtNoPO.Enabled = false;
                    }
                    else
                    {
                        txtNoPO.Enabled = true;
                    }
                    lblOrderType.Text = dtFill1.Rows[0]["wsh_so_type"].ToString();
                    //form_actived
                    frmSOManualEntryEditor_Activated(null, null);

                    g_bActiveEdit = true;
                    if (dtFill1.Rows[0]["wsh_so_type"].ToString() == "")
                    {
                        g_bSOType = false;
                        cbTipeOrder.EditValue = LookupValueAt(cbTipeOrder, 0);
                        g_bSOType = true;
                    }
                    else
                    {
                        g_bEdit = false;
                        for (lCounter = 0; lCounter < LookupRowCount(cbTipeOrder); lCounter++)
                        {
                            g_bSOType = false;
                            cbTipeOrder.EditValue = LookupValueAt(cbTipeOrder, lCounter);
                            g_bSOType = true;
                            string vTOP = Convert.ToString(cbTipeOrder.EditValue);
                            if (vTOP == dtFill1.Rows[0]["wsh_so_type"].ToString().Trim())
                            {
                                g_bEdit = true;
                                lblOrderType.Text = dtFill1.Rows[0]["wsh_so_type"].ToString();
                                break;
                            }

                        }

                        if (g_bEdit == false)
                        {
                            MessageBox.Show("Tipe Order : " + dtFill1.Rows[0]["wsh_so_type"].ToString() + " tidak ada " + " atau belum di setting untuk salesman : " + dtFill1.Rows[0]["wsh_spgm_id"].ToString(), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }

                    }
                }

                //SO DETAIL
                //if (dgvSalesDetail.Rows.Count > 0)
                //{
                //    dgvSalesDetail.Rows.Clear();

                //}

                //comented 08.01.2019
                //if (dgvSalesDetail.Rows.Count > 0)
                //{
                //    dtGridSODetail.Rows.Clear();
                //}

                dgvSalesDetail.DataSource = dtGridSODetail;
                DataTable dt = dgvSalesDetail.DataSource as DataTable;
                int rowIndex = 0;
                DataTable dtFill2 = new DataTable();
                strSQL = " select wsd_prd_master_code, wsd_grade, wsd_prd_size, isnull(prm_prd_desc, '') as prm_prd_desc, ";
                strSQL = strSQL + " wsd_uom_convert_big, wsd_uom_convert_mid, wsd_het_unit_price, ";
                strSQL = strSQL + " wsd_real_order_xqty, wsd_real_order_amt, isnull(wsd_so_xqty, '') as wsd_so_xqty, isnull(wsd_total_so_amount, 0) as wsd_total_so_amount, ";
                strSQL = strSQL + " wsd_tot_disc_pc, wsd_price_date, wsd_tax_code, wsd_tax_pct,wsd_sap_unit_price,wsd_sled,ISNULL(wsd_req_order_xqty,wsd_real_order_xqty) AS wsd_req_order_xqty, CASE WHEN ISNULL(wsd_reason, '') = '' THEN '' ELSE wsd_reason + ' ~ ' + gh_function_desc END AS wsd_reason,";
                strSQL = strSQL + " (ISNULL(wsd_gross_net,0)-(isnull(wsd_tot_disc_1,0) + isnull(wsd_tot_disc_2,0)+isnull(wsd_tot_disc_3,0)+ isnull(wsd_tot_disc_4,0)+ isnull(wsd_tot_disc_5,0)+ isnull(wsd_tot_disc_6,0)+isnull(wsd_tot_disc_7,0)+isnull(wsd_tot_disc_8,0)+isnull(wsd_tot_disc_9,0)+isnull(wsd_tot_disc_10,0)) + isnull(wsd_tax_amount,0)) as Net_Sales ,wsd_prd_line_code, isnull(wsd_is_subs, 'N') as wsd_is_subs ";
                strSQL = strSQL + " from SO_WEB_SALES_DETAIL WITH (NOLOCK) ";
                strSQL = strSQL + " left join IM_PRD_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + " on wsd_prd_master_code = prm_prd_master_code";
                strSQL = strSQL + " and wsd_grade = prm_grade";
                strSQL = strSQL + " and wsd_prd_size = prm_prd_size";
                strSQL = strSQL + " LEFT JOIN ( SELECT gh_function_code, gh_function_desc FROM GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + "                 where gh_function_name ='REASONKAM' ";
                strSQL = strSQL + "          ) RES ";
                strSQL = strSQL + "        ON wsd_reason = gh_function_code ";
                strSQL = strSQL + " where wsd_so_seq_no = '" + sNoOrder + "'";
                if (sEntitys != "")
                {
                    strSQL = strSQL + " and wsd_entity_id = '" + sEntitys + "' ";
                }

                if (sBranchs != "")
                {
                    strSQL = strSQL + " and wsd_so_branch_id = '" + sBranchs + "' ";
                }
                strSQL = strSQL + " and wsd_line_type = 'N' ";
                strSQL = strSQL + " order by wsd_so_detail_line_no";
                dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill2.Rows.Count > 0)
                {
                    foreach (DataRow rw in dtFill2.Rows)
                    {
                        dt.Rows.Add();

                        var row = this.dgvSalesDetail.Rows[rowIndex];
                        row.Cells["wsd_prd_master_code"].Value = rw["wsd_prd_master_code"].ToString();
                        row.Cells["wsd_grade"].Value = rw["wsd_grade"].ToString();
                        row.Cells["wsd_prd_size"].Value = rw["wsd_prd_size"].ToString();
                        row.Cells["prm_prd_desc"].Value = rw["prm_prd_desc"].ToString();
                        row.Cells["wsd_uom_convert_big"].Value = rw["wsd_uom_convert_big"].ToString();
                        row.Cells["wsd_uom_convert_mid"].Value = rw["wsd_uom_convert_mid"].ToString();
                        row.Cells["wsd_het_unit_price"].Value = rw["wsd_het_unit_price"].ToString();
                        row.Cells["wsd_real_order_xqty"].Value = rw["wsd_real_order_xqty"].ToString();
                        row.Cells["wsd_tot_disc_pc"].Value = "0";
                        row.Cells["wsd_sled"].Value = rw["wsd_sled"].ToString();
                        row.Cells["reason"].Value = rw["wsd_reason"].ToString();
                        row.Cells["wsd_NetSales"].Value = Convert.ToDecimal(rw["Net_Sales"]).ToString("#,##0.00");
                        row.Cells["wsd_prd_line_code"].Value = rw["wsd_prd_line_code"].ToString();
                        row.Cells[COL_IS_SUBS].Value = rw["wsd_is_subs"].ToString();
                        row.Cells[COL_SUBS_LOOKUP].Value = "";

                        //wsd_req_order_xqty = wsd_real_order_xqty TPM-6612
                        row.Cells["wsd_req_order_xqty"].Value = rw["wsd_req_order_xqty"].ToString();
                        if (btnProses.Enabled == false || bCalcDisc == true)
                        {
                            row.Cells["ijumlahharga"].Value = rw["wsd_total_so_amount"].ToString();
                            if (rw["wsd_so_xqty"].ToString() != "")
                            {
                                row.Cells["wsd_so_xqty"].Value = rw["wsd_so_xqty"].ToString();
                            }
                            else
                            {
                                row.Cells["wsd_so_xqty"].Value = "0";
                            }

                        }
                        else
                        {
                            row.Cells["ijumlahharga"].Value = rw["wsd_real_order_amt"].ToString();
                            row.Cells["wsd_so_xqty"].Value = rw["wsd_real_order_xqty"].ToString();
                        }
                        if (rw["wsd_tot_disc_pc"].ToString() == "0.00")
                        {
                            //row.Cells["wsd_tot_disc_pc"].Value = "0";
                        }
                        else
                        {
                            row.Cells["wsd_tot_disc_pc"].Value = rw["wsd_tot_disc_pc"].ToString();
                        }
                        if (bCalcDisc == false)
                        {
                            row.Cells["tgl_price"].Value = Convert.ToDateTime(rw["wsd_price_date"]).ToString("dd MMM yyyy");

                        }
                        else
                        {
                            //dtTglPrice.Text = rw["wsd_price_date"].ToString();
                            //DateTime __dtgl3 = DateTime.Parse(rw["wsd_price_date"].ToString(), CultureInfo.GetCultureInfo("id-ID"));
                            row.Cells["tgl_price"].Value = Convert.ToDateTime(rw["wsd_price_date"]).ToString("dd MMM yyyy");
                        }

                        decimal _ijml = Convert.ToDecimal(row.Cells["ijumlahharga"].Value);
                        decimal _idisc = Convert.ToDecimal(rw["wsd_tot_disc_pc"].ToString());
                        curJmlHarga = _ijml - _idisc;
                        row.Cells["jml_rp_netto"].Value = curJmlHarga;
                        row.Cells["tax_code"].Value = rw["wsd_tax_code"].ToString();
                        row.Cells["pct_tax"].Value = rw["wsd_tax_pct"].ToString();
                        row.Cells["wsd_sap_unit_price"].Value = rw["wsd_sap_unit_price"].ToString();
                        rowIndex++;

                    }

                    if (bCalcDisc == true)
                    {
                        //dgvSalesDetail.Rows[rowIndex].Cells["tgl_price"].Value = dtTglPrice.Text;
                    }
                }

                SetGridSODetailCss(dgvSalesDetail);

                //TPRB
                if (dgvPromosiQty.Rows.Count > 0)
                {
                    dtGridTPRB.Rows.Clear();
                    //dgvPromosiQty.Rows.Clear();
                }

                if (dtGridTPRB.Rows.Count > 0)
                {
                    dtGridTPRB.Rows.Clear();
                }
                dgvPromosiQty.DataSource = dtGridTPRB;
                DataTable dt1 = dgvPromosiQty.DataSource as DataTable;
                int rowIndex1 = 0;
                DataTable dtFill3 = new DataTable();
                strSQL = "select  wsd_prd_master_code, isnull(prm_prd_desc, '') prm_prd_desc, ";
                strSQL = strSQL + " wsd_het_unit_price, wsd_so_xqty";
                strSQL = strSQL + " From SO_WEB_SALES_DETAIL WITH (NOLOCK) ";
                strSQL = strSQL + " left join IM_PRD_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + " on";
                strSQL = strSQL + " prm_prd_master_code = wsd_prd_master_code";
                strSQL = strSQL + " and prm_grade = wsd_grade";
                strSQL = strSQL + " and prm_prd_size = wsd_prd_size";
                strSQL = strSQL + " where wsd_line_type = 'B' ";
                strSQL = strSQL + " and wsd_so_seq_no = '" + sNoOrder + "' ";
                strSQL = strSQL + " order by wsd_so_detail_line_no ";
                dtFill3 = _clsGlobal.ExecDTTrans(strSQL);
                //dgvPromosiQty.Rows.Clear();
                if (dtFill3.Rows.Count > 0)
                {
                    //dgvPromosiQty.Rows.Clear();
                    foreach (DataRow rw in dtFill3.Rows)
                    {
                        if (rw["wsd_so_xqty"].ToString().Trim() != "00000.000.0000")
                        {
                            dt1.Rows.Add();
                            var row = this.dgvPromosiQty.Rows[rowIndex1];
                            row.Cells["wsd_prd_master_code"].Value = rw["wsd_prd_master_code"].ToString();
                            row.Cells["prm_prd_desc"].Value = rw["prm_prd_desc"].ToString();
                            row.Cells["wsd_het_unit_price"].Value = rw["wsd_het_unit_price"].ToString();
                            row.Cells["wsd_so_xqty"].Value = rw["wsd_so_xqty"].ToString();
                            rowIndex1++;
                        }

                    }
                }

                SetGridTPRBCss(dgvPromosiQty);

                //'TPR Uang
                //if (dgvPromosiRp.Rows.Count > 0)
                //{
                //    dgvPromosiRp.Rows.Clear();
                //}

                if (dgvPromosiRp.Rows.Count > 0)
                {
                    dtGridTPRU.Rows.Clear();
                }
                dgvPromosiRp.DataSource = dtGridTPRU;
                DataTable dt2 = dgvPromosiRp.DataSource as DataTable;
                int rowIndex2 = 0;

                if (!isNefoKAM)
                {
                    DataTable dtFill4 = new DataTable();
                    strSQL = "select  pt_product_id, isnull(prm_prd_desc, '') prm_prd_desc,";
                    strSQL = strSQL + " pt_amount,pt_promo_id,pt_pct_promo ";
                    strSQL = strSQL + " From TBL_SD_PROMO_TXN WITH (NOLOCK) ";
                    strSQL = strSQL + " left join IM_PRD_MASTER WITH (NOLOCK) ";
                    strSQL = strSQL + " on";
                    strSQL = strSQL + " prm_prd_master_code = pt_product_id";
                    strSQL = strSQL + " and prm_grade = pt_grade";
                    strSQL = strSQL + " and prm_prd_size = pt_size";
                    strSQL = strSQL + " where pt_flag_proc = '1' ";
                    strSQL = strSQL + " and pt_order_no = '" + sNoOrder + "'";
                    dtFill4 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill4.Rows.Count > 0)
                    {
                        foreach (DataRow rw in dtFill4.Rows)
                        {
                            lRow = 0;
                            if (pCheckValue(rw["pt_product_id"].ToString(), dgvPromosiRp.Rows.Count - 1, dgvPromosiRp) == true)
                            {
                                dt2.Rows.Add();

                                var row = this.dgvPromosiRp.Rows[rowIndex2];

                                row.Cells["pt_product_id"].Value = rw["pt_product_id"].ToString();
                                row.Cells["prm_prd_desc"].Value = rw["prm_prd_desc"].ToString();
                                row.Cells["pt_amount"].Value = Convert.ToDecimal(rw["pt_amount"]).ToString("#,##0.00");
                                row.Cells["pt_promo_id"].Value = rw["pt_promo_id"].ToString();
                                row.Cells["pt_pct_promo"].Value = Convert.ToDecimal(rw["pt_pct_promo"]).ToString("#,##0.00");

                                rowIndex2++;
                            }
                            else
                            {
                                decimal _xyz = Convert.ToDecimal(this.dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value) + Convert.ToDecimal(rw["pt_amount"].ToString());
                                this.dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value = _xyz.ToString("#,##0.00");
                            }
                        }

                    }
                }
                else
                {
                    DataTable dtFill4 = new DataTable();
                    strSQL = "select  dt_product_id pt_product_id, isnull(prm_prd_desc, '') prm_prd_desc,";
                    strSQL = strSQL + " SUM(dt_amount) pt_amount,dt_promo_id as pt_promo_id,dt_pct_promo pt_pct_promo ";
                    strSQL = strSQL + " From SD_SALES_DEAL_TRX WITH (NOLOCK) ";
                    strSQL = strSQL + " left join IM_PRD_MASTER WITH (NOLOCK) ";
                    strSQL = strSQL + " on";
                    strSQL = strSQL + " prm_prd_master_code = dt_product_id";
                    strSQL = strSQL + " and prm_grade = dt_grade";
                    strSQL = strSQL + " and prm_prd_size = dt_size";
                    strSQL = strSQL + " where dt_flag_proc = '1' ";
                    strSQL = strSQL + " and dt_order_no = '" + sNoOrder + "'";
                    strSQL = strSQL + " group by dt_product_id,Dt_promo_id,prm_prd_desc,dt_pct_promo";
                    dtFill4 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill4.Rows.Count > 0)
                    {
                        foreach (DataRow rw in dtFill4.Rows)
                        {
                            lRow = 0;
                            if (pCheckValueSDUangKAM(rw["pt_product_id"].ToString(), rw["pt_promo_id"].ToString(), dgvPromosiRp.Rows.Count - 1, dgvPromosiRp) == true)
                            {
                                dt2.Rows.Add();
                                var row = this.dgvPromosiRp.Rows[rowIndex2];
                                row.Cells["pt_promo_id"].Value = rw["pt_promo_id"].ToString();
                                row.Cells["pt_product_id"].Value = rw["pt_product_id"].ToString();
                                row.Cells["prm_prd_desc"].Value = rw["prm_prd_desc"].ToString();
                                row.Cells["pt_amount"].Value = Convert.ToDecimal(rw["pt_amount"]).ToString("#,##0.00");
                                row.Cells["pt_pct_promo"].Value = Convert.ToDecimal(rw["pt_pct_promo"]).ToString("#,##0.00");
                                rowIndex2++;
                            }
                            else
                            {
                                decimal _xyz = Convert.ToDecimal(this.dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value) + Convert.ToDecimal(rw["pt_amount"].ToString());
                                this.dgvPromosiRp.Rows[lRow].Cells["pt_amount"].Value = _xyz.ToString("#,##0.00");
                            }
                        }

                    }

                }


                SetGridTPRUCss(dgvPromosiRp);

                if (!isNefoKAM)
                {



                    //'Discount
                    if (dgvDisc.Rows.Count > 0)
                    {
                        dtGridDisc.Rows.Clear();
                        //dgvDisc.Rows.Clear();
                    }
                    if (dtGridDisc.Rows.Count > 0)
                    {
                        dtGridDisc.Rows.Clear();
                    }
                    dgvDisc.DataSource = dtGridDisc;
                    DataTable dt3 = dgvDisc.DataSource as DataTable;
                    int rowIndex3 = 0;
                    DataTable dtFill5 = new DataTable();
                    strSQL = "select tdt_product_master_id,";
                    strSQL = strSQL + " tdt_discount_code,";
                    strSQL = strSQL + " tdt_disc_level,";
                    strSQL = strSQL + " isnull(tdt_disc_pct, 0) as tdt_disc_pct, ";
                    strSQL = strSQL + " isnull(tdt_disc_value, 0) as tdt_disc_value ";
                    strSQL = strSQL + " From TBL_DISC_TXN WITH (NOLOCK) ";
                    strSQL = strSQL + " where tdt_order_no = '" + sNoOrder + "' ";
                    strSQL = strSQL + " and tdt_disc_level <> '0'";
                    dtFill5 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill5.Rows.Count > 0)
                    {
                        foreach (DataRow rw in dtFill5.Rows)
                        {
                            lRow = 0;
                            string diskPrsnVal = Convert.ToDecimal(rw["tdt_disc_pct"]).ToString("#,##0.00");
                            string diskRpVal = Convert.ToDecimal(rw["tdt_disc_value"]).ToString("#,##0.00");
                            if (pCheckValueDisc(rw["tdt_product_master_id"].ToString(), dgvDisc.Rows.Count - 1, dgvDisc) == true)
                            {
                                dt3.Rows.Add();

                                var row = this.dgvDisc.Rows[rowIndex3];
                                row.Cells["tdt_product_master_id"].Value = rw["tdt_product_master_id"].ToString();

                                if (rw["tdt_disc_level"].ToString() == "1")
                                {
                                    row.Cells["g_iPctD1"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD1"].Value = diskRpVal;
                                    row.Cells["g_iDiscCode1"].Value = rw["tdt_discount_code"].ToString();
                                }
                                else if (rw["tdt_disc_level"].ToString() == "2")
                                {
                                    row.Cells["g_iPctK1"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK1"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "3")
                                {
                                    row.Cells["g_iPctD2"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD2"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "4")
                                {
                                    row.Cells["g_iPctK2"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK2"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "5")
                                {
                                    row.Cells["g_iPctD3"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD3"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "6")
                                {
                                    row.Cells["g_iPctK3"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK3"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "7")
                                {
                                    row.Cells["g_iPctD4"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD4"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "8")
                                {
                                    row.Cells["g_iPctK4"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK4"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "9")
                                {
                                    row.Cells["g_iPctD5"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD5"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "10")
                                {
                                    row.Cells["g_iPctK5"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK5"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "11")
                                {
                                    row.Cells["g_iPctD6"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD6"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "12")
                                {
                                    row.Cells["g_iPctK6"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK6"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "13")
                                {
                                    row.Cells["g_iPctD7"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD7"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "14")
                                {
                                    row.Cells["g_iPctK7"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK7"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "15")
                                {
                                    row.Cells["g_iPctD8"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD8"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "16")
                                {
                                    row.Cells["g_iPctK8"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK8"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "17")
                                {
                                    row.Cells["g_iPctD9"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD9"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "18")
                                {
                                    row.Cells["g_iPctK9"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK9"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "19")
                                {
                                    row.Cells["g_iPctD10"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD10"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "20")
                                {
                                    row.Cells["g_iPctK10"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK10"].Value = diskRpVal;
                                }
                                rowIndex3++;
                            }
                            else
                            {
                                var row = this.dgvDisc.Rows[lRow];
                                if (rw["tdt_disc_level"].ToString() == "1")
                                {
                                    row.Cells["g_iPctD1"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD1"].Value = diskRpVal;
                                    row.Cells["g_iDiscCode1"].Value = rw["tdt_discount_code"].ToString();
                                }
                                else if (rw["tdt_disc_level"].ToString() == "2")
                                {
                                    row.Cells["g_iPctK1"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK1"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "3")
                                {
                                    row.Cells["g_iPctD2"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD2"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "4")
                                {
                                    row.Cells["g_iPctK2"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK2"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "5")
                                {
                                    row.Cells["g_iPctD3"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD3"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "6")
                                {
                                    row.Cells["g_iPctK3"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK3"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "7")
                                {
                                    row.Cells["g_iPctD4"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD4"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "8")
                                {
                                    row.Cells["g_iPctK4"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK4"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "9")
                                {
                                    row.Cells["g_iPctD5"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD5"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "10")
                                {
                                    row.Cells["g_iPctK5"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK5"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "11")
                                {
                                    row.Cells["g_iPctD6"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD6"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "12")
                                {
                                    row.Cells["g_iPctK6"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK6"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "13")
                                {
                                    row.Cells["g_iPctD7"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD7"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "14")
                                {
                                    row.Cells["g_iPctK7"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK7"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "15")
                                {
                                    row.Cells["g_iPctD8"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD8"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "16")
                                {
                                    row.Cells["g_iPctK8"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK8"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "17")
                                {
                                    row.Cells["g_iPctD9"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD9"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "18")
                                {
                                    row.Cells["g_iPctK9"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK9"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "19")
                                {
                                    row.Cells["g_iPctD10"].Value = diskPrsnVal;
                                    row.Cells["g_iRpD10"].Value = diskRpVal;
                                }
                                else if (rw["tdt_disc_level"].ToString() == "20")
                                {
                                    row.Cells["g_iPctK10"].Value = diskPrsnVal;
                                    row.Cells["g_iRpK10"].Value = diskRpVal;
                                }
                            }
                        }

                        //For lCounter = g_iPctD1 To g_iRpK10
                        //    bHiddenCol = True
                        //    For lBrs = vsList(3).FixedRows To vsList(3).Rows - 1
                        //        If vsList(3).TextMatrix(lBrs, lCounter) <> "" Then
                        //            bHiddenCol = False
                        //            Exit For
                        //        End If
                        //    Next lBrs
                        //    vsList(3).ColHidden(lCounter) = bHiddenCol
                        //Next lCounter
                        DataTable dtGrid = new DataTable();
                        dtGrid = dgvDisc.DataSource as DataTable;
                        //dtGrid = dtGrid.AsEnumerable()
                        //       .GroupBy(r => new { Col1 = r["Col1"], Col2 = r["Col2"] })
                        //       .Select(g => g.OrderBy(r => r["PK"]).First())
                        //       .CopyToDataTable();

                        for (lCounter = 2; lCounter < 42; lCounter++)
                        {
                            bHiddenCol = true;
                            for (lRow = 0; lRow < dgvDisc.Rows.Count; lRow++)
                            {
                                if (dgvDisc.Rows[lRow].Cells[lCounter].Value.ToString() == "")
                                {
                                    bHiddenCol = false;
                                    break;
                                }
                                else
                                {
                                    bHiddenCol = true;
                                    break;
                                }
                            }
                            dgvDisc.Columns[lCounter].Visible = bHiddenCol;
                        }

                    }

                    SetGridDiscCss(dgvDisc);

                }

                curJmlHarga = 0;
                for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                {
                    if (dgvSalesDetail.Rows[i].Cells["wsd_tot_disc_pc"].Value.ToString() == "")
                    {
                        dgvSalesDetail.Rows[i].Cells["wsd_tot_disc_pc"].Value = "0";
                    }

                    curJmlHarga = curJmlHarga + Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["ijumlahharga"].Value) - Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_tot_disc_pc"].Value);
                }

                txtSubTotal.Text = curJmlHarga.ToString("#,##0.00");




                DataTable dtFill6 = new DataTable();
                strSQL = "select wsh_tot_tpr_amt, wsh_tot_disc_1, wsh_tot_disc_2, ";
                strSQL = strSQL + " wsh_tot_disc_3, wsh_tot_disc_4, ";
                strSQL = strSQL + " wsh_tot_disc_5, wsh_tot_disc_6, ";
                strSQL = strSQL + " wsh_tot_disc_7, wsh_tot_disc_8, ";
                strSQL = strSQL + " wsh_tot_disc_9, wsh_tot_disc_10, ";
                strSQL = strSQL + " wsh_pct_cash_disc, ";
                strSQL = strSQL + " wsh_cash_disc, wsh_tax_amount, wsh_net_amount, wsh_dpp_amount";
                strSQL = strSQL + " From dbo.SO_WEB_SALES_HEADER WITH (NOLOCK) ";
                strSQL = strSQL + " where wsh_seq_no = '" + sNoOrder + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "'";
                dtFill6 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill6.Rows.Count > 0)
                {
                    for (int i = 1; i < 11; i++)
                    {
                        if (string.IsNullOrEmpty(dtFill6.Rows[0]["wsh_tot_disc_" + i].ToString()))
                        {
                            dtFill6.Rows[0]["wsh_tot_disc_" + i] = "0";
                        }
                    }
                    txtPromosiU.Text = Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_tpr_amt"]).ToString("#,##0.00");
                    Decimal ttlDisc1dan2 = Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_1"]) + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_2"]) + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_3"]);
                    ttlDisc1dan2 = ttlDisc1dan2 + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_4"]) + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_5"]) + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_6"]);
                    ttlDisc1dan2 = ttlDisc1dan2 + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_7"]) + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_8"]) + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_9"]) + Convert.ToDecimal(dtFill6.Rows[0]["wsh_tot_disc_10"]);
                    txtDisc1.Text = ttlDisc1dan2.ToString("#,##0.00");
                    // txtCashDiscP.Text = Convert.ToDecimal(dtFill6.Rows[0]["wsh_pct_cash_disc"]).ToString("#,##0.00");
                    txtCashDiscP.Text = Convert.ToDecimal(dtFill6.Rows[0]["wsh_pct_cash_disc"]).ToString("#,##0.####");
                    txtCashDiscU.Text = Convert.ToDecimal(dtFill6.Rows[0]["wsh_cash_disc"]).ToString("#,##0.00");
                    txtPPN.Text = Convert.ToDecimal(dtFill6.Rows[0]["wsh_tax_amount"].ToString() == "" ? "0" : dtFill6.Rows[0]["wsh_tax_amount"].ToString()).ToString("#,##0.00");
                    txtDPP.Text = Convert.ToDecimal(dtFill6.Rows[0]["wsh_dpp_amount"].ToString() == "" ? "0" : dtFill6.Rows[0]["wsh_dpp_amount"].ToString()).ToString("#,##0.00");
                    txtTtlInvoice.Text = Convert.ToDecimal(dtFill6.Rows[0]["wsh_net_amount"].ToString() == "" ? "0" : dtFill6.Rows[0]["wsh_net_amount"].ToString()).ToString("#,##0.00");
                }
                else
                {
                    txtPromosiU.Text = "";
                    txtDisc1.Text = "";
                    txtCashDiscP.Text = "";
                    txtCashDiscU.Text = "";
                    txtPPN.Text = "";
                    txtDPP.Text = "";
                    txtTtlInvoice.Text = "";
                }

                g_bEdit = true;
                DataTable dtFill7 = new DataTable();
                strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
                dtFill7 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill7.Rows.Count > 0)
                {
                    cbTOP.Enabled = false;
                    cbPembyFaktur.Enabled = false;
                }






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

        private void pChangePCode(int iIndex)
        {
            dgvSalesDetail.Rows[iIndex].Cells["wsd_het_unit_price"].Value = "";
            //dgvSalesDetail.Rows[iIndex].Cells["wsd_real_order_xqty"].Value = "";
            dgvSalesDetail.Rows[iIndex].Cells["wsd_so_xqty"].Value = "";
            dgvSalesDetail.Rows[iIndex].Cells["wsd_real_order_amt"].Value = "0";
            dgvSalesDetail.Rows[iIndex].Cells["wsd_tot_disc_pc"].Value = "0";
            string _dttglprice = Convert.ToDateTime(dtTglOrder.EditValue).ToString("dd MMM yyyy");
            dgvSalesDetail.Rows[iIndex].Cells["tgl_price"].Value = _dttglprice;
            dgvSalesDetail.Rows[iIndex].Cells["ijumlahharga"].Value = "0";
            dgvSalesDetail.Rows[iIndex].Cells["rp_tax"].Value = "0";
            //dgvSalesDetail.Rows[iIndex].Cells["wsd_sled"].Value = "0";
            dgvSalesDetail.Rows[iIndex].Cells["wsd_NetSales"].Value = "0";
            RefreshSubstitutionFlag(iIndex);

            var checkSled = ((DataTable)dgvSalesDetail.DataSource).Rows.Cast<DataRow>()
                    .Where(x => !x["wsd_prd_master_code"].ToString().IsNullOrEmptyOrWhiteSpace() && decimal.TryParse(x["wsd_sled"]?.ToString(), out var sled) && sled > 0m).ToList();

            if (checkSled.Count() > 0)
            {
                chkSLED.Checked = true;
            }
            else
            {
                chkSLED.Checked = false;
            }
        }

        //proses
        private bool fProsesOrderType(string sCredLimit, string sType)
        {
            bool bYesNo;
            bool keyFound = true;

            if (sType == "1")
            {
                if (g_bReqProses == false)
                {
                    bYesNo = true;
                    //if (btnProses.Enabled == true)
                    //{
                    if (fProsesSO(sCredLimit, bYesNo) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                    //}
                    if (bYesNo == false)
                    {

                    }
                }

                if (sCredLimit == "*")
                {
                    keyFound = true;
                    return keyFound;
                }

                //'Proses Summary Order
                if (g_bReqProses == false && g_bReqSummary == false)
                {
                    if (fProsesSummaryOrder(txtNoOrder.Text) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }

                //'Proses Shipment
                if (g_bReqProses == false && g_bReqSummary == false && g_bReqShipment == false)
                {
                    if (fProsesShipSO(txtNoOrder.Text) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }

                //'Proses POD
                if (g_bReqProses == false && g_bReqSummary == false && g_bReqShipment == false && g_bReqPOD == false)
                {
                    if (fProsesPOD(sSONo) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }

                //'Proses Invoice
                if (g_bReqProses == false && g_bReqSummary == false && g_bReqShipment == false && g_bReqPOD == false && g_bReqInvoice == false)
                {
                    if (fProsesInv(sSONo) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
            }
            else if (sType == "2")
            {
                //'Proses Summary Order
                if (g_bReqProses == false && g_bReqSummary == false)
                {
                    if (fProsesSummaryOrder(txtNoOrder.Text) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }

                //'Proses Shipment
                if (g_bReqProses == false && g_bReqSummary == false && g_bReqShipment == false)
                {
                    if (fProsesShipSO(txtNoOrder.Text) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }

                //'Proses POD
                if (g_bReqProses == false && g_bReqSummary == false && g_bReqShipment == false && g_bReqPOD == false)
                {
                    if (fProsesPOD(sSONo) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }

                //'Proses Invoice
                if (g_bReqProses == false && g_bReqSummary == false && g_bReqShipment == false && g_bReqPOD == false && g_bReqInvoice == false)
                {
                    if (fProsesInv(sSONo) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                }
            }

            //keyFound = true;
            return keyFound;
        }

        //proses
        private bool fProsesSummaryOrder(string sSONo)
        {

            bool keyFound = true;
            try
            {
                DateTime dtShipment;
                string sWHLoc1 = "";
                string sWHLoc2 = "";

                //'Get WH Location
                if (fGetWHLoc(sSONo, sWHLoc1, sWHLoc2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                //'Get Shipment Date //string whloc yang di lempar kosong, jd pst error
                //if(fGetWHDate

                //proses SUMMARY ORDER
                //string __tglShipment = Convert.ToDateTime(dtShipment).ToString("yyyy-MM-dd");
                //DataTable dtFill1 = new DataTable();
                //strSQL = " EXEC SP_PROSES_SUMMARY_ORDER " + sSONo + ", " + __tglShipment + ", " + clsLogin.USERID;
                //dtFill1 = _clsGlobal.ExecDT(strSQL);

                //DataTable dtFill1 = new DataTable();
                //strSQL = "" EXEC SP_INSERT_HISTORY " & FmtStr("2") & " , " & FmtStr(sSummaryNo) & " , " & FmtStr(eUserID);
                //dtFill1 = _clsGlobal.ExecDT(strSQL);    

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

        private bool fProsesShipSO(string sSONo)
        {
            bool keyFound = false;
            double dQtyShip = 0;
            string sWeek;
            string sWHLoc1 = "";
            string sWHLoc2 = "";
            DateTime dtShipment = DateTime.Now;
            string sTypeSales;

            try
            {
                if (fGetWHLoc(sSONo, sWHLoc1, sWHLoc2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (fGetWHDate(sWHLoc1, sWHLoc2, dtShipment) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                string xTransportNopol;
                string xTransportID;
                string xTruckType;
                DataTable dtFill1 = new DataTable();
                strSQL = " select sgm_type_operasi from SO_SPG_GIRL_MAN WITH (NOLOCK)  ";
                strSQL = strSQL + " where sgm_spgm_id = '" + txtSalesID.Text + "' and sgm_entity_id = '" + xEntityID + "' and sgm_branch_id = '" + xBranchID + "'";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    sTypeSales = dtFill1.Rows[0]["sgm_type_operasi"].ToString();
                }
                else
                {
                    MessageBox.Show("Type operation untuk sales (" + txtSalesID.Text + ")" + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                DataTable dtFill2 = new DataTable();
                if (sTypeSales == "C")
                {
                    strSQL = "select  top 1 coalesce(stm_id,'') as stm_id,coalesce(stm_type_id,'') as stm_type_id ,coalesce(stm_police_no,'') as stm_police_no from TBL_SD_TRANSPORT_MASTER WITH (NOLOCK)   where stm_canvas_slsman=" + txtSalesID.Text + " and stm_canvas_flag='Y'";
                }
                else
                {
                    strSQL = "select  top 1 coalesce(stm_id,'') as stm_id,coalesce(stm_type_id,'') as stm_type_id ,coalesce(stm_police_no,'') as stm_police_no from TBL_SD_TRANSPORT_MASTER WITH (NOLOCK)   where stm_id in  (select top 1 gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name='DUMMYTRUCK') ";
                }
                dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill2.Rows.Count > 0)
                {
                    xTransportID = dtFill2.Rows[0]["stm_id"].ToString();
                    xTruckType = dtFill2.Rows[0]["stm_type_id"].ToString();
                    xTransportNopol = dtFill2.Rows[0]["stm_police_no"].ToString();

                }
                else
                {
                    if (sTypeSales == "C")
                    {
                        MessageBox.Show("Kendaraan untuk salesman " + txtSalesID.Text + " tidak ditemukan !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                    {
                        MessageBox.Show("Parameter DUMMYTRUCK belum di setting di GST3", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                    keyFound = false;
                    return keyFound;

                }

                //DataTable dtFill3 = new DataTable();
                strSQL = " EXEC SP_PROSES_SHIPMENT_MULTIBRANCH '";
                strSQL = strSQL + "1','";
                strSQL = strSQL + "2','";
                strSQL = strSQL + sSONo + "','";
                strSQL = strSQL + dtShipment.ToString("yyyy-MM-dd") + "','";
                strSQL = strSQL + clsLogin.USERID + "','";
                strSQL = strSQL + "','";
                strSQL = strSQL + xTransportNopol + "','";
                strSQL = strSQL + "','";
                strSQL = strSQL + xTransportID + "','";
                strSQL = strSQL + "I" + "','";
                strSQL = strSQL + "" + "','";
                strSQL = strSQL + xTruckType + "',";
                strSQL = strSQL + "0" + ",";
                strSQL = strSQL + "0" + ",";
                strSQL = strSQL + "0" + ",";
                strSQL = strSQL + "0" + ",";
                strSQL = strSQL + "0" + ",";
                strSQL = strSQL + "0" + ",'";
                strSQL = strSQL + xEntityID + "','";
                strSQL = strSQL + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //DataTable dtFill4 = new DataTable();
                strSQL = " EXEC SP_INSERT_HISTORY_MULTIBRANCH " + "'2','" + sSummaryNo + "','" + clsLogin.USERID + "', '" + xEntityID + "', '" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                keyFound = true;

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

        private bool fProsesPOD(string sSONo)
        {
            bool keyFound = false;
            double dQty = 0;
            string sWeek;
            string sWHLoc1 = "";
            string sWHLoc2 = "";
            DateTime dtShipment = DateTime.Now;
            string sTypeSales;

            try
            {
                if (fGetWHLoc(sSONo, sWHLoc1, sWHLoc2) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                if (fGetWHDate(sWHLoc1, sWHLoc2, dtShipment) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                //DataTable dtFill3 = new DataTable();
                strSQL = " EXEC SP_PROSES_POD_KAM " + "'1','" + sSONo + "','" + dtShipment.ToString("yyyy-MM-dd") + ",'" + xEntityID + "','" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //DataTable dtFill4 = new DataTable();
                strSQL = " EXEC SP_INSERT_HISTORY_MULTIBRANCH " + "'1','" + sSONo + "','" + clsLogin.USERID + "','" + xEntityID + "', '" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                keyFound = true;

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

        private bool fProsesInv(string sSONo)
        {
            bool keyFound = false;
            string sInvNo;
            string sWHLoc1 = "";
            string sWHLoc2 = "";
            DateTime dtInv;

            try
            {

                //DataTable dtFill3 = new DataTable();
                strSQL = " EXEC SP_INVOICE_PROSES_KAM " + "'" + sSONo + "','" + clsLogin.USERID + "','" + xEntityID + "','" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //DataTable dtFill4 = new DataTable();
                strSQL = " EXEC SP_INSERT_HISTORY_MULTIBRANCH " + "'1','" + sSONo + "','" + clsLogin.USERID + "','" + xEntityID + "','" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                keyFound = true;

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

        private bool fProsesSO(string sCredLimit, bool bYesNo)
        {
            bool keyFound = false;

            int iTpr = 0;
            string sGroup;
            string sType;
            //string sTPR();
            decimal curJumlah;
            long iYesNo = 0;
            double lAFSQty = 0;
            long tmpPCodeCnt = 0;
            int tmpPCodeCntOOS = 0;
            string tmpPCode = "";
            bool ContinueProcess = false;

            string cPcode = "";
            string cGrade = "";
            string cSize = "";
            string cQty = "";
            int cConv1 = 0;
            int cConv2 = 0;

            try
            {
                xSaveOOS = false;
                xCountOOS = 0;

                if (dgvSalesDetail.Rows.Count > 0)
                {
                    for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                    {
                        if (dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() != "")
                        {
                            cPcode = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString();
                            cGrade = dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString();
                            cSize = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString();
                            cQty = dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString();
                            cConv1 = Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString());
                            cConv2 = Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString());
                            lAFSQty = 0;

                            if (fGetAFS(cPcode, cGrade, cSize, lblWHLoc1.Text, lblWHLoc2.Text) == true)
                            {
                                fConvertQty(cQty, cConv1, cConv2);
                                if (Convert.ToInt64(_fConvertQty) > lAFSQty)
                                {
                                    tmpPCodeCnt = tmpPCodeCnt + 1;
                                    tmpPCode = tmpPCode.ToString() + cPcode + ",";

                                    DataTable dtFill2 = new DataTable();
                                    strSQL = " select top 1  rm_reason_code from IM_STOCK_BALANCE inner join TBL_GS_REASON_MASTER WITH (NOLOCK)  on rm_alias = wlb_status_oos where isnull(wlb_status_oos ,'') <> '' and wlb_prd_master_code='" + cPcode + "' and wlb_entity_id = '" + xEntityID + "' and wlb_branch_id = '" + xBranchID + "'";
                                    dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                                    if (dtFill2.Rows.Count > 0)
                                    {
                                        dgvSalesDetail.Rows[lCounter].Cells["reason"].Value = dtFill2.Rows[0]["rm_reason_code"].ToString();
                                    }
                                    else
                                    {
                                        dgvSalesDetail.Rows[lCounter].Cells["reason"].Value = "SC06";
                                    }

                                    if (lAFSQty > 0)
                                    {
                                        dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value = fPC2Qty(lAFSQty, cConv1, cConv2);
                                    }
                                    else
                                    {
                                        dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value = fPC2Qty(0, cConv1, cConv2);
                                        tmpPCodeCntOOS = tmpPCodeCntOOS + 1;
                                    }
                                }
                                else
                                {
                                    dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value = dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString();
                                }

                            }
                            else
                            {
                                pRollbackProcess();
                                _clsGlobal.RollbackTrans();
                                MessageBox.Show("Stock PCode : " + cPcode + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                        }
                    }
                }

                bolOOS = false;
                if (tmpPCodeCnt == 0)
                {
                    keyFound = true;
                    ContinueProcess = true;
                }
                else if ((lCounter - 1) == tmpPCodeCnt && ((lCounter - 1) != tmpPCodeCntOOS))
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Nomor SO " + txtNoOrder.Text + " tidak terpenuhi semua, Apakah mau di lanjutkan ?", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.OK)
                    {
                        keyFound = true;
                        ContinueProcess = true;
                    }
                    else
                    {
                        ContinueProcess = false;
                        keyFound = false;
                        pRollbackProcess();
                    }
                }
                else if ((lCounter - 1) > tmpPCodeCnt && ((lCounter - 1) != tmpPCodeCntOOS))
                {
                    DialogResult dr;
                    dr = MessageBox.Show("Stock PCode : " + tmpPCode.Substring(1, tmpPCode.Length - 1) + " tidak mencukupi untuk No Order : " + txtNoOrder.Text + ", Apakah mau di lanjutkan ?", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.OK)
                    {
                        keyFound = true;
                        ContinueProcess = true;
                    }
                    else
                    {
                        bolCloseSO = false;
                        //ContinueProcess = false;
                        keyFound = false;
                        pRollbackProcess();
                    }
                }
                else if ((lCounter - 1) == tmpPCodeCntOOS)
                {
                    bolCloseSO = true;
                    bolOOS = true;
                    keyFound = true;
                    ContinueProcess = true;

                    DialogResult dr;
                    dr = MessageBox.Show("Nomor SO " + txtNoOrder.Text + " tidak terpenuhi semua, Apakah mau di lanjutkan ?", clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.OK)
                    {
                        keyFound = true;
                        ContinueProcess = true;
                    }
                    else
                    {
                        ContinueProcess = false;
                        keyFound = false;
                        pRollbackProcess();
                    }

                }

                if (ContinueProcess == true)
                {
                    for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                    {
                        if (dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() != "")
                        {
                            cPcode = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString();
                            cGrade = dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString();
                            cSize = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString();
                            cQty = dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString();
                            cConv1 = Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString());
                            cConv2 = Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString());

                            DataTable dtFill3 = new DataTable();
                            strSQL = "  EXEC SP_INSERT_STOCK_OUT '" + lblWHLoc1.Text + "','";
                            strSQL = strSQL + lblWHLoc2.Text + "','";
                            strSQL = strSQL + cPcode + "','";
                            strSQL = strSQL + cGrade + "','";
                            strSQL = strSQL + cSize + "','";
                            strSQL = strSQL + clsLogin.USERID + "','";
                            strSQL = strSQL + xEntityID + "','";
                            strSQL = strSQL + xBranchID + "'";
                            dtFill3 = _clsGlobal.ExecDTTrans(strSQL);
                        }

                    }
                }

                if (fCekMandatory() == false)
                {
                    keyFound = false;

                }

                pClearTPR();
                pClearDiscount();
                pClearTax();

                if (sCredLimit != "CALCDISC")
                {
                    bYesNo = true;

                    if (fSaveProsesSO(sCredLimit, bYesNo) == false)
                    {
                        keyFound = false;
                    }
                }
                else
                {
                    if (fcalcdisc() == false)
                    {
                        keyFound = false;
                    }
                }


                if (sCredLimit != "CALCDISC")
                {
                    //btnProses.Enabled = false;
                    btnCalcDisc.Enabled = false;
                }

                g_bProses = true;
                //keyFound = true;

            }
            catch (Exception ex)
            {
                if (_clsGlobal.tr != null && _clsGlobal.tr.Connection != null)
                {
                    _clsGlobal.RollbackTrans();
                }

                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                keyFound = false;
                return keyFound;

            }



            return keyFound;
        }

        private bool fPricePCode(string sPcode, string sGrade, string sSize, string sCust1, string sCust2, string sTgl)
        {
            bool keyFound = false;

            try
            {

                DateTime __dt = DateTime.Parse(sTgl, CultureInfo.GetCultureInfo("en-US"));
                //DateTime.TryParseExact(sTgl, "dd/MM/yyyy",
                //    System.Globalization.CultureInfo.InvariantCulture,
                //    System.Globalization.DateTimeStyles.NoCurrentDateDefault, out __dt);

                DataTable dtFill1 = new DataTable();
                strSQL = "IP_PRICEPCODE_MBRANCH '" + sPcode + "','";
                strSQL = strSQL + sGrade + "','";
                strSQL = strSQL + sSize + "','";
                strSQL = strSQL + sCust1 + "','";
                strSQL = strSQL + sCust2 + "','";
                strSQL = strSQL + __dt.ToString("yyyyMMdd") + "','";
                strSQL = strSQL + xEntityID + "','";
                strSQL = strSQL + xBranchID + "'";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                else
                    dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

                if (dtFill1.Rows.Count > 0)
                {
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["opl_het"].ToString()))
                    {
                        keyFound = false;
                    }
                    else
                    {
                        curPrice = Convert.ToDecimal(dtFill1.Rows[0]["opl_het"].ToString());
                        keyFound = true;
                    }
                }
            }
            catch (Exception ex)
            {
                //_clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                keyFound = false;
                return keyFound;

            }

            return keyFound;
        }

        private bool fGetWHLoc(string sSONo, string sWHLoc1, string sWHLoc2)
        {

            bool keyFound = false;
            try
            {


                DataTable dtFill1 = new DataTable();

                strSQL = " select wsh_loc_id1, wsh_loc_id2 ";
                strSQL = strSQL + " from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                strSQL = strSQL + " where wsh_seq_no = '" + sSONo + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "'";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    sWHLoc1 = dtFill1.Rows[0]["wsh_loc_id1"].ToString();
                    sWHLoc2 = dtFill1.Rows[0]["wsh_loc_id2"].ToString();
                    keyFound = true;
                }
                else
                {
                    MessageBox.Show("SO Header : " + sSONo + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                //keyFound = false;
                //return keyFound;

            }

            return keyFound;
        }

        private bool fGetWHDate(string sWHLoc1, string sWHLoc2, DateTime dtWHDate)
        {

            bool keyFound = false;
            try
            {


                DataTable dtFill1 = new DataTable();

                strSQL = " select cast(convert(varchar(10), wh_loc_last_work_date, 120) as datetime) as wh_loc_last_work_date ";
                strSQL = strSQL + " from IM_WH_LOC WITH (NOLOCK)  ";
                strSQL = strSQL + " where wh_loc_id1 = '" + sWHLoc1 + "'";
                strSQL = strSQL + " and wh_loc_id2 = '" + sWHLoc2 + "' ";
                strSQL = strSQL + " and wh_loc_entity = '" + xEntityID + "' and wh_branch_id = '" + xBranchID + "'";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    dtWHDate = Convert.ToDateTime(dtFill1.Rows[0]["wh_loc_last_work_date"].ToString());
                    keyFound = true;
                }
                else
                {
                    MessageBox.Show("Tanggal untuk Gudang : " + sWHLoc1 + "," + sWHLoc2 + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //keyFound = true;
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                //keyFound = false;
                //return keyFound;

            }

            return keyFound;
        }

        private bool fChangeSOType()
        {
            bool keyFound = false;

            try
            {
                //DataRow selectedDataRow = ((DataRowView)cbTipeOrder.GetSelectedDataRow()).Row;
                int idO = Convert.ToInt32(cbTipeOrder.EditValue);
                string NameO = cbTipeOrder.Text;

                lblOrderType.Text = idO.ToString();

                DataTable dtFill = new DataTable();
                //int key = Convert.ToInt32(cbTipeOrder.EditValue);
                strSQL = " select msow_wh_loc1, msow_wh_loc2 from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK)   where msow_so_type = '" + idO + "'  and msow_sld_id = '" + txtSalesID.Text + "' and msow_entity_id = '" + xEntityID + "' and msow_branch_id = '" + xBranchID + "'";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtFill = _clsGlobal.ExecDT(strSQL);
                else
                    dtFill = _clsGlobal.ExecDTTrans(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    lblWHLoc1.Text = dtFill.Rows[0]["msow_wh_loc1"].ToString();
                    lblWHLoc2.Text = dtFill.Rows[0]["msow_wh_loc2"].ToString();

                    keyFound = true;

                    DataTable dtFill2 = new DataTable();
                    strSQL = " select cast(convert(varchar(10), wh_loc_last_work_date, 120) as datetime) as wh_loc_last_work_date  from IM_WH_LOC WITH (NOLOCK)  ";
                    strSQL = strSQL + " where wh_loc_id1 = '" + lblWHLoc1.Text + "'";
                    strSQL = strSQL + " AND wh_loc_id2 = '" + lblWHLoc2.Text + "'";
                    strSQL = strSQL + " AND wh_loc_entity = '" + xEntityID + "'";
                    strSQL = strSQL + " AND wh_branch_id = '" + xBranchID + "'";
                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        dtFill2 = _clsGlobal.ExecDT(strSQL);
                    else
                        dtFill2 = _clsGlobal.ExecDTTrans(strSQL);

                    if (dtFill2.Rows.Count > 0)
                    {
                        //dtTglOrder.Text = Convert.ToDateTime(dtFill2.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                        //dtTglPO.Text = Convert.ToDateTime(dtFill2.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                        if (clsGlobal.MODE_TRX == 1)
                        {
                            dtTglOrder.Text = Convert.ToDateTime(dtFill2.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                            dtTglPO.Text = Convert.ToDateTime(dtFill2.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                            initDateDelvExp(Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd"), "", "", 1);
                        }
                        keyFound = true;
                    }
                    else
                    {
                        lblWHLoc1.Text = "";
                        lblWHLoc2.Text = "";
                        lblOrderType.Text = "";
                        MessageBox.Show("Tanggal Gudang : " + lblWHLoc1.Text + " - " + lblWHLoc2.Text + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }

                    DataTable dtFill3 = new DataTable();
                    strSQL = "select isnull(ot_process_req, '') as  ot_process_req ,isnull(ot_summary_req, '') as ot_summary_req,isnull(ot_shipment_req, '') as ot_shipment_req ,isnull(ot_pod_req, '') as ot_pod_req ,isnull(ot_inv_req, '') as ot_inv_req ";
                    strSQL = strSQL + " From SO_ORDER_TYPE WITH (NOLOCK)  ";
                    strSQL = strSQL + " where ot_order_type = '" + lblOrderType.Text + "'";
                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        dtFill3 = _clsGlobal.ExecDT(strSQL);
                    else
                        dtFill3 = _clsGlobal.ExecDTTrans(strSQL);

                    if (dtFill3.Rows.Count > 0)
                    {
                        //dtTglOrder.Text = Convert.ToDateTime(dtFill3.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                        //dtTglPO.Text = Convert.ToDateTime(dtFill3.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                        if (dtFill3.Rows[0]["ot_process_req"].ToString() == "Y")
                        {
                            g_bReqProses = true;
                            btnProses.Enabled = true;
                        }
                        else
                        {
                            g_bReqProses = false;
                            btnProses.Enabled = false;
                        }

                        if (dtFill3.Rows[0]["ot_summary_req"].ToString() == "Y")
                        {
                            g_bReqSummary = true;
                        }
                        else
                        {
                            g_bReqSummary = false;
                        }

                        if (dtFill3.Rows[0]["ot_shipment_req"].ToString() == "Y")
                        {
                            g_bReqShipment = true;
                        }
                        else
                        {
                            g_bReqShipment = false;
                        }

                        if (dtFill3.Rows[0]["ot_pod_req"].ToString() == "Y")
                        {
                            g_bReqPOD = true;
                        }
                        else
                        {
                            g_bReqPOD = false;
                        }

                        if (dtFill3.Rows[0]["ot_inv_req"].ToString() == "Y")
                        {
                            g_bReqInvoice = true;
                        }
                        else
                        {
                            g_bReqInvoice = false;
                        }

                    }
                    else
                    {
                        lblWHLoc1.Text = "";
                        lblWHLoc2.Text = "";
                        lblOrderType.Text = "";
                        MessageBox.Show("Data order type : " + cbTipeOrder.Text + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }

                }
                else
                {
                    lblWHLoc1.Text = "";
                    lblWHLoc2.Text = "";
                    lblOrderType.Text = "";
                    if (isMsgGudangShown == false)
                    {
                        DialogResult dr = MessageBox.Show("Gudang pada Mapping Order type : " + cbTipeOrder.Text + " tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        if (dr == DialogResult.OK)
                        {
                            isMsgGudangShown = true;
                        }

                    }

                    keyFound = false;
                }

            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fGetTOPFlag()
        {
            bool keyFound = false;

            try
            {
                DataTable dtFill = new DataTable();

                strSQL = "SELECT gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK)  WHERE gh_sys = 'H' AND gh_function_name = 'TOPBYPRDLINE'";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtFill = _clsGlobal.ExecDT(strSQL);
                else
                    dtFill = _clsGlobal.ExecDTTrans(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    if (dtFill.Rows[0]["gh_function_code"].ToString() == "Y")
                    {
                        g_bTOPFlag = true;
                        keyFound = true;
                    }
                    else
                    {
                        g_bTOPFlag = false;
                        keyFound = true;
                    }
                }
                else
                {
                    MessageBox.Show("TOP By Product Line has not been setting on Parameter & Report Setting", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fCekMandatory()
        {
            bool keyFound = true;
            string __product = "";

            try
            {
                ClearFormValidationErrors();

                if (txtKdOutlet1.Text == "")
                {
                    MessageBox.Show("Outlet masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //txtKdOutlet1.Focus();                
                    keyFound = false;
                    //return keyFound;
                }

                if (rbCashDiscP.Checked)
                {
                    if (!string.IsNullOrEmpty(txtCashDiscP.Text.Trim()))
                    {
                        if (Convert.ToDecimal(txtCashDiscP.Text.Trim()) > 0)
                        {
                            if (Convert.ToString(cmbBebanDisc.EditValue) == "0")
                            {
                                MessageBox.Show("Apabila Cash Disc % diisi Beban Cash Disc harus diisi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                //txtKdOutlet1.Focus();                
                                keyFound = false;
                            }
                        }
                    }
                }


                if (rbCashDiscU.Checked)
                {
                    if (!string.IsNullOrEmpty(txtCashDiscU.Text.Trim()))
                    {
                        if (Convert.ToDecimal(txtCashDiscU.Text.Trim()) > 0)
                        {
                            if (Convert.ToString(cmbBebanDisc.EditValue) == "0")
                            {
                                MessageBox.Show("Apabila Rp Cash Disc diisi Beban Cash Disc harus diisi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                //txtKdOutlet1.Focus();                
                                keyFound = false;
                            }
                        }
                    }
                }


                if (cbPembyFaktur.Text == "")
                {
                    MessageBox.Show("Pembayaran Faktur masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //cbPembyFaktur.Focus();                
                    keyFound = false;
                }

                if (txtNoPO.Text == "")
                {
                    MessageBox.Show("No PO masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //txtNoPO.Focus();                
                    keyFound = false;
                }
                else
                {
                    if (isNefoKAM)
                    {
                        if (txtNoPO.Text.Length > 35)
                        {
                            MessageBox.Show("No PO Tidak Boleh Dari 35 Karakter!!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            //txtNoPO.Focus();                
                            keyFound = false;
                        }
                    }
                }

                if (cbTOP.Text == "")
                {
                    MessageBox.Show("TOP masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //txtKdOutlet1.Focus();                
                    keyFound = false;
                }

                if (dgvSalesDetail.Rows.Count <= 0)
                {
                    MessageBox.Show("Product masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //txtKdOutlet1.Focus();                
                    keyFound = false;
                }
                if (cbTipeOrder.Text == "")
                {
                    MessageBox.Show("Order Type masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //txtKdOutlet1.Focus();                
                    keyFound = false;
                }

                for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++) //di buat minus 1 karena bertambah ROW
                {
                    if (dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() == "")
                    {
                        MessageBox.Show("Product masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //txtKdOutlet1.Focus();                
                        keyFound = false;
                        break;
                    }

                    double QtyRealOrder = 0;
                    double QtyReqOrder = 0;
                    fQtyFormat(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                    QtyRealOrder = Convert.ToDouble(_fConvertQty);

                    fQtyFormat(dgvSalesDetail.Rows[lCounter].Cells["wsd_req_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                    QtyReqOrder = Convert.ToDouble(_fConvertQty);

                    if (QtyRealOrder != QtyReqOrder)
                    {
                        if (dgvSalesDetail.Rows[lCounter].Cells["reason"].Value.ToString() == "" || string.IsNullOrEmpty(dgvSalesDetail.Rows[lCounter].Cells["reason"].Value.ToString()))
                        {
                            MessageBox.Show("Reason masih kosong [PCODE: " + dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() + "]", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            break;
                        }
                    }


                    /** untuk validasi agar kan tidak bisa diproses **/
                    if (isNefoKAM)
                    {
                        string __prod = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString();
                        strSQL = " Select smt_prdline_id ";
                        strSQL += " FROM IM_PRD_MASTER WITH (NOLOCK) ";
                        strSQL += " INNER JOIN SO_MAPPING_TOPBYPRDLINE WITH (NOLOCK) ";
                        strSQL += " ON smt_prdline_id = prm_prd_line_code ";
                        strSQL += " WHERE  ";
                        strSQL += " smt_entity_id = '" + xEntityID + "' ";
                        strSQL += " and smt_branch_id = '" + xBranchID + "' ";
                        strSQL += " and smt_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                        strSQL += " and smt_cust_code2 = '" + lblKdOutlet2.Text + "' ";
                        strSQL += " and prm_prd_master_code = '" + __prod + "' ";

                        DataTable dtPRDLINE = new DataTable();
                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                            dtPRDLINE = _clsGlobal.ExecDT(strSQL);
                        else
                            dtPRDLINE = _clsGlobal.ExecDTTrans(strSQL);

                        if (dtPRDLINE.Rows.Count == 0)
                        {
                            MessageBox.Show("Data Produk Line untuk Produk " + __prod + "  Belum di Mapping ke Customer");
                            keyFound = false;
                            break;
                        }

                        /** UNTUK COLD CHAIN **/
                        if (string.IsNullOrEmpty(__product))
                        {
                            __product = "'" + dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() + "'";
                        }
                        else
                        {
                            __product = __product + ",'" + dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() + "'";
                        }


                    }


                    /************** Check Coldchain ****************/
                    else
                    {
                        if (string.IsNullOrEmpty(__product))
                        {
                            __product = "'" + dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() + "'";
                        }
                        else
                        {
                            __product = __product + ",'" + dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() + "'";
                        }
                    }
                    /************** end coldchain *****************/

                    if (dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString() == "")
                    {
                        //Harga masih kosong
                        MessageBox.Show("Harga masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //txtKdOutlet1.Focus();                
                        keyFound = false;
                        break;
                    }

                    if (dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString() == "")
                    {
                        MessageBox.Show("Real Qty masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //txtKdOutlet1.Focus();                
                        keyFound = false;
                        break;
                    }

                    if (dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value.ToString() == "")
                    {
                        MessageBox.Show("Order Qty masih kosong", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //txtKdOutlet1.Focus();                
                        keyFound = false;
                        break;
                    }


                }


                //Cek ShippingPlant
                if (isCrossSite)
                {
                    DataTable dtSP = new DataTable();
                    strSQL = "SELECT 1 from SO_CUST_MASTER WITH(NOLOCK) ";
                    strSQL = strSQL + " WHERE cm_ship_plant = '" + txtShippingPlantId.Text + "'";
                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        dtSP = _clsGlobal.ExecDT(strSQL);
                    else
                        dtSP = _clsGlobal.ExecDTTrans(strSQL);

                    if (dtSP.Rows.Count == 0)
                    {
                        MessageBox.Show("Tolong dicek master customer shipping plant tidak sesuai ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return false;
                    }
                }

                DataTable dtFill1 = new DataTable();
                strSQL = "select fwd_entity From IM_WH_LOC WITH (NOLOCK)  inner join TBL_GS_WORKING_DAY WITH (NOLOCK)  on convert(varchar(10), fwd_work_day, 120) = convert(varchar(10), wh_loc_last_work_date, 120)" +
                    "  and fwd_week_no = wh_loc_week_closing and fwd_periode = wh_loc_month_closing and fwd_year = wh_loc_year_closing and fwd_entity = wh_loc_entity and fwd_branch = wh_branch_id " +
                    " where wh_loc_id1 ='" + lblWHLoc1.Text + "'  and wh_loc_id2 ='" + lblWHLoc2.Text + "' and wh_loc_entity = '" + xEntityID + "' and wh_branch_id = '" + xBranchID + "'";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                else
                    dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

                if (dtFill1.Rows.Count < 1)
                {
                    MessageBox.Show("Tanggal transaksi Gudang tidak di pekan/periode/year berjalan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                }


                DataTable dtNIK = new DataTable();
                strSQL = "select sah_NIK from tbl_salesman_history" +
                         " where" +
                         " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                         " and sah_entity_id = '" + xEntityID + "'" +
                         " and sah_branch_id = '" + xBranchID + "'" +
                         " and sah_salesmen_id = '" + xSalesID + "'";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtNIK = _clsGlobal.ExecDT(strSQL);
                else
                    dtNIK = _clsGlobal.ExecDTTrans(strSQL);

                if (dtNIK.Rows.Count == 1)
                {
                }
                else if (dtNIK.Rows.Count > 1)
                {
                    MessageBox.Show("Salesman History Pada Periode Entry Lebih dari 1 ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                }
                else
                {
                    MessageBox.Show("Salesman History Tidak Ditemukan ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;

                }

                //***** kam jika PO sama di tolak ****//
                if (isNefoKAM)
                {
                    if (clsGlobal.MODE_TRX == 1)
                    {
                        DataTable dtCheckPO = new DataTable();
                        strSQL = "select wsh_po_no from so_web_sales_header WITH (NOLOCK)" +
                                 " where" +
                                 " wsh_entity_id = '" + xEntityID + "'" +
                                 " and wsh_branch_id = '" + xBranchID + "'" +
                                 " and wsh_po_no = '" + txtNoPO.Text + "'";
                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                            dtCheckPO = _clsGlobal.ExecDT(strSQL);
                        else
                            dtCheckPO = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtCheckPO.Rows.Count > 0)
                        {
                            MessageBox.Show("Nomor PO Sudah Ada Mohon Gunakan Nomor PO yang lain!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                        }
                    }
                    else if (clsGlobal.MODE_TRX == 2)
                    {
                        DataTable dtCheckPO = new DataTable();
                        strSQL = "select wsh_po_no from so_web_sales_header WITH (NOLOCK)" +
                                 " where" +
                                 " wsh_entity_id = '" + xEntityID + "'" +
                                 " and wsh_branch_id = '" + xBranchID + "'" +
                                 " and wsh_seq_no = '" + txtNoOrder.Text + "'";
                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                            dtCheckPO = _clsGlobal.ExecDT(strSQL);
                        else
                            dtCheckPO = _clsGlobal.ExecDTTrans(strSQL);

                        if (dtCheckPO.Rows.Count > 0)
                        {
                            string __po_no = dtCheckPO.Rows[0]["wsh_po_no"].ToString();
                            if (__po_no.ToString() != txtNoPO.Text)
                            {
                                DataTable dtCheckPO2 = new DataTable();
                                strSQL = "select wsh_po_no from so_web_sales_header WITH (NOLOCK)" +
                                         " where" +
                                         " wsh_entity_id = '" + xEntityID + "'" +
                                         " and wsh_branch_id = '" + xBranchID + "'" +
                                         " and wsh_po_no = '" + txtNoPO.Text + "'";
                                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                                    dtCheckPO2 = _clsGlobal.ExecDT(strSQL);
                                else
                                    dtCheckPO2 = _clsGlobal.ExecDTTrans(strSQL);
                                if (dtCheckPO2.Rows.Count > 0)
                                {
                                    MessageBox.Show("Nomor PO Sudah Ada Mohon Gunakan Nomor PO yang lain!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    keyFound = false;
                                }
                            }
                        }
                    }

                }

                //***********************************//

                /******** multisource ************/
                if (isMultisource)
                {
                    DataTable dtCheckMultisource = new DataTable();
                    strSQL = " select * from GS_PRM_MULTI_SOURCE WITH (NOLOCK)  ";
                    strSQL += " where pms_entity_id = '" + xEntityID + "' ";
                    strSQL += " and pms_branch_id = '" + xBranchID + "' ";
                    strSQL += " and pms_flag_source = '" + Convert.ToString(cbSource.EditValue) + "' ";
                    strSQL += " and pms_flag_sloc = '" + Convert.ToString(cbFlagSloc.EditValue) + "' ";

                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        dtCheckMultisource = _clsGlobal.ExecDT(strSQL);
                    else
                        dtCheckMultisource = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtCheckMultisource.Rows.Count == 0)
                    {
                        MessageBox.Show("Multisource : Kombinasi Source dan Flag Sloc tidak ada pada Master!!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }


                    /*** check query coldchain ***/
                    if (Convert.ToString(cbFlagSloc.EditValue).ToUpper().Equals("COC"))
                    {
                        if (!string.IsNullOrEmpty(__product))
                        {
                            dtCheckMultisource = new DataTable();
                            strSQL = "select top 1 prm_prd_master_code ";
                            strSQL += "from IM_PRD_MASTER ";
                            strSQL += "LEFT JOIN ( ";
                            strSQL += "	SELECT gh_function_code ";
                            strSQL += "	from GS_GEN_HARDCODED WITH (NOLOCK) ";
                            strSQL += "	where gh_function_name = 'LINE_COLD_CHAIN' ";
                            strSQL += "	and gh_sys = 'H' ";
                            strSQL += ") CC ON CC.gh_function_code = prm_prd_line_code ";
                            strSQL += "where CC.gh_function_code is null and prm_prd_master_code in (" + __product + ") ";

                            if (_clsGlobal.Connect.State == ConnectionState.Closed)
                                dtCheckMultisource = _clsGlobal.ExecDT(strSQL);
                            else
                                dtCheckMultisource = _clsGlobal.ExecDTTrans(strSQL);
                            if (dtCheckMultisource.Rows.Count > 0)
                            {
                                MessageBox.Show("Multisource : Terdapat Product yang bukan ColdChain!!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                            }
                        }
                    }
                    // yang bukan barang coldchain
                    //else
                    //{
                    //    if (!string.IsNullOrEmpty(__product))
                    //    {
                    //        dtCheckMultisource = new DataTable();
                    //        strSQL = "select top 1 prm_prd_master_code ";
                    //        strSQL += "from IM_PRD_MASTER ";
                    //        strSQL += "LEFT JOIN ( ";
                    //        strSQL += "	SELECT gh_function_code ";
                    //        strSQL += "	from GS_GEN_HARDCODED ";
                    //        strSQL += "	where gh_function_name = 'LINE_COLD_CHAIN' ";
                    //        strSQL += "	and gh_sys = 'H' ";
                    //        strSQL += ") CC ON CC.gh_function_code = prm_prd_line_code ";
                    //        strSQL += "where CC.gh_function_code is not null and prm_prd_master_code in (" + __product + ") ";

                    //        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    //            dtCheckMultisource = _clsGlobal.ExecDT(strSQL);
                    //        else
                    //            dtCheckMultisource = _clsGlobal.ExecDTTrans(strSQL);
                    //        if (dtCheckMultisource.Rows.Count > 0)
                    //        {
                    //            MessageBox.Show("Multisource : Terdapat Product yang ColdChain!!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //            keyFound = false;
                    //        }
                    //    }
                    //}
                    /*** end check query coldchain ***/
                }

                /*********************************/
                else // non multisource
                {
                    bool iscoldchain = false;
                    bool iscoldchainot = false;
                    string sCCOrderType = "";
                    int totalsku = dgvSalesDetail.Rows.Count;
                    /** CHECK COLD CHAIN **/
                    DataTable dtCheckColdChain = new DataTable();
                    strSQL = "select distinct prm_prd_master_code ";
                    strSQL += "from IM_PRD_MASTER WITH (NOLOCK) ";
                    strSQL += "LEFT JOIN ( ";
                    strSQL += "	SELECT gh_function_code ";
                    strSQL += "	from GS_GEN_HARDCODED WITH (NOLOCK) ";
                    strSQL += "	where gh_function_name = 'LINE_COLD_CHAIN' ";
                    strSQL += "	and gh_sys = 'H' ";
                    strSQL += ") CC ON CC.gh_function_code = prm_prd_line_code ";
                    strSQL += "where CC.gh_function_code is not null and prm_prd_master_code in (" + __product + ") ";

                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        dtCheckColdChain = _clsGlobal.ExecDT(strSQL);
                    else
                        dtCheckColdChain = _clsGlobal.ExecDTTrans(strSQL);

                    if (dtCheckColdChain.Rows.Count > 0)
                    {
                        iscoldchain = true;
                    }

                    if (iscoldchain)
                    {
                        if (totalsku != dtCheckColdChain.Rows.Count)
                        {
                            MessageBox.Show("Coldchain Validation : terdapat barang yang bukan coldchain ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                        }

                        DataTable dtCheckFlagCCOT = new DataTable();
                        strSQL = "";
                        strSQL += "SELECT * FROM GS_GEN_HARDCODED WITH (NOLOCK) ";
                        strSQL += "WHERE gh_function_name = 'FLAGCOLDCHAIN_OT'";
                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                            dtCheckFlagCCOT = _clsGlobal.ExecDT(strSQL);
                        else
                            dtCheckFlagCCOT = _clsGlobal.ExecDTTrans(strSQL);

                        if (dtCheckFlagCCOT.Rows.Count > 0)
                        {
                            if (dtCheckFlagCCOT.Rows[0]["gh_function_code"].ToString().ToUpper().Equals("Y"))
                            {
                                iscoldchainot = true;
                                sCCOrderType = dtCheckFlagCCOT.Rows[0]["gh_min_seq_no"].ToString();
                            }
                        }

                        if (iscoldchainot)
                        {
                            if (string.IsNullOrEmpty(sCCOrderType.Trim()))
                            {
                                MessageBox.Show("ColdChain Validation : Data tipe order belum diisi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                            }
                            else
                            {
                                DataTable dtCheckTipeOrder = new DataTable();
                                strSQL = "select * from SO_ORDER_TYPE WITH (NOLOCK) where ot_order_type = '" + sCCOrderType + "' ";
                                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                                    dtCheckTipeOrder = _clsGlobal.ExecDT(strSQL);
                                else
                                    dtCheckTipeOrder = _clsGlobal.ExecDTTrans(strSQL);

                                if (dtCheckTipeOrder.Rows.Count == 0)
                                {
                                    MessageBox.Show("ColdChain Tipe Order : Data tipe order " + sCCOrderType + " tidak ada di master order type mohon di create terlebih dahulu.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    keyFound = false;

                                }
                                else
                                {
                                    strSQL = "select * from TBL_MAP_SLD_OPRT_WH WITH (NOLOCK) where msow_sld_id = '" + xSalesID + "' and msow_entity_id = '" + xEntityID + "' and msow_branch_id = '" + xBranchID + "' and  msow_so_type = '" + sCCOrderType + "' AND msow_wh_loc1 +'|'+msow_wh_loc2 = '" + lblWHLoc1.Text + "|" + lblWHLoc2.Text + "' ";
                                    DataTable dtCheckMAPSLD = new DataTable();
                                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                                        dtCheckMAPSLD = _clsGlobal.ExecDT(strSQL);
                                    else
                                        dtCheckMAPSLD = _clsGlobal.ExecDTTrans(strSQL);

                                    if (dtCheckMAPSLD.Rows.Count == 0)
                                    {
                                        MessageBox.Show("ColdChain Tipe Order : Data tipe order tidak ada pada mapping salesman ke Tipe Order", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        keyFound = false;
                                    }
                                }
                            }


                        }
                    }

                    /** END OF COLDCHAIN **/
                }


                //VALIDASI SLED
                if (chkSLED.Checked)
                {
                    //check sled tidak boleh 0 (Product Expired)
                    var checkSled = ((DataTable)dgvSalesDetail.DataSource).Rows.Cast<DataRow>()
                   .Where(x => !x["wsd_prd_master_code"].ToString().IsNullOrEmptyOrWhiteSpace() && (!string.IsNullOrWhiteSpace(x["wsd_sled"]?.ToString()) && decimal.TryParse(x["wsd_sled"]?.ToString(), out var sled) && sled == 0m)).ToList();
                    string ProductCodes = string.Join(",", checkSled.Select(p => p["wsd_prd_master_code"].ToString()));

                    if (checkSled.Count() > 0)
                    {
                        MessageBox.Show($"Sled pada Product : {ProductCodes} tidak boleh 0", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }

                    var checkSled_blank_count = ((DataTable)dgvSalesDetail.DataSource).Rows.Cast<DataRow>();
                    var checkSled_blank = ((DataTable)dgvSalesDetail.DataSource).Rows.Cast<DataRow>()
                    .Where(x => !x["wsd_prd_master_code"].ToString().IsNullOrEmptyOrWhiteSpace() && string.IsNullOrWhiteSpace(x["wsd_sled"]?.ToString())).ToList();


                    if (checkSled_blank_count.Count() == checkSled_blank.Count())
                    {
                        MessageBox.Show($"Sled Masih Blank", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                    }



                }

            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (!keyFound)
                MarkMandatoryValidationErrors();
            else
                ClearFormValidationErrors();

            return keyFound;
        }

        //mdlutility
        private bool fGetAFS(string sPcode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;

            DataTable dtFill1 = new DataTable();
            strSQL = " select dbo.F_GET_AFS('";
            strSQL = strSQL + xEntityID + "','";
            strSQL = strSQL + xBranchID + "','";
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

        //mdlutility
        private void fConvertQty(string sQty, int iConv1, int iConv2)
        {
            //double keyFound = 0;
            _fConvertQty = "0";
            int iSplit = 0;
            string[] sQuantity = sQty.Split('.');
            iSplit = sQuantity.GetUpperBound(0);

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

            if (iSplit == 0)
            {
                cNol = "00000" + sQuantity[0].ToString();
                sFormatQty = cNol.Substring(cNol.Length - 5) + ".000.000";

            }
            else if (iSplit == 1)
            {
                cNol = "00000" + sQuantity[0].ToString();
                cSatu = "000" + sQuantity[1].ToString();
                //cDua = "000" + sQuantity[2].ToString();
                sFormatQty = cNol.Substring(cNol.Length - 5) + "." + cSatu.Substring(cSatu.Length - 3) + ".000";
            }
            else if (iSplit == 2)
            {
                cNol = "00000" + sQuantity[0].ToString();
                cSatu = "000" + sQuantity[1].ToString();
                cDua = "0000000" + sQuantity[2].ToString();
                sFormatQty = cNol.Substring(cNol.Length - 5) + "." + cSatu.Substring(cSatu.Length - 3) + "." + cDua.Substring(cDua.Length - 7);
            }
            else if (iSplit > 2)
            {
                cNol = "00000" + sQuantity[0].ToString();
                cSatu = "000" + sQuantity[1].ToString();
                cDua = "0000000" + sQuantity[2].ToString();
                sFormatQty = cNol.Substring(cNol.Length - 5) + "." + cSatu.Substring(cSatu.Length - 3) + "." + cDua.Substring(cDua.Length - 7);

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
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                else
                    dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    _fQtyFormat = dtFill1.Rows[0]["Qty_Format"].ToString();
                }
            }

        }

        //mdlUtility //1 juni 2018
        private string fPC2Qty(double sQty, int iConv1, int iConv2)
        {
            string qtyformat = "";
            DataTable dtFill1 = new DataTable();
            strSQL = " select dbo.SQTY_FORMAT (";
            strSQL = strSQL + sQty + ",";
            strSQL = strSQL + iConv1 + ",";
            strSQL = strSQL + iConv2 + ") as Qty_Format";
            if (_clsGlobal.Connect.State == ConnectionState.Closed)
                dtFill1 = _clsGlobal.ExecDT(strSQL);
            else
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

            if (dtFill1.Rows.Count > 0)
            {
                qtyformat = dtFill1.Rows[0]["Qty_Format"].ToString();
            }

            return qtyformat;
        }

        //1 juni 2018
        private bool fcalcdisc()
        {
            bool keyFound = false;

            string sOpenStatus = "";
            string sErrDesc = "";
            bool bExists = false;
            bool bYesNo = false;

            if (txtNoOrder.Text != g_sDefaultSONo)
            {
                //check so status
                DataTable dtFill1 = new DataTable();
                strSQL = " select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_sys = 'H' ";
                strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
                strSQL = strSQL + " and gh_sequence_no = 1";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtFill1 = _clsGlobal.ExecDT(strSQL);
                else
                    dtFill1 = _clsGlobal.ExecDTTrans(strSQL);

                if (dtFill1.Rows.Count > 0)
                {
                    sOpenStatus = dtFill1.Rows[0]["gh_function_code"].ToString();
                    keyFound = true;
                    //return keyFound;
                }
                else
                {
                    MessageBox.Show("Status SO sudah bukan pada Open ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }
            }

            try
            {
                _clsGlobal.BeginTrans();
                if (fSaveSO("", true) == false)
                {
                    keyFound = false;
                    return keyFound;
                }

                //DataTable DTUJICOBA1 = _clsGlobal.ExecDTTrans("Select * from so_web_sales_header with(nolock) where wsh_seq_no = '" + g_sDefaultSONo + "'");

                DataTable dtFill2 = new DataTable();
                if (isNefoKAM)
                {
                    strSQL = " EXEC SP_EXEC_GET_TPR_KAM '" + g_sDefaultSONo + "','" + xEntityID + "','" + xBranchID + "'";
                }
                else
                {
                    strSQL = " EXEC SP_EXEC_GET_TPR '" + g_sDefaultSONo + "'";
                }

                dtFill2 = _clsGlobal.ExecDTTrans(strSQL);

                bExists = false;
                bYesNo = true;
                if (dtFill2.Rows.Count > 0)
                {
                    for (int i = 0; i < dtFill2.Rows.Count; i++)
                    {
                        if (dtFill2.Rows[0]["tpr_budget_flag"].ToString() == "1" && dtFill2.Rows[0]["result_type"].ToString() == "2")
                        {
                            bYesNo = false;
                        }
                        bExists = true;
                    }

                }

                //if (bExists == true)
                //{

                //}

                //if (bYesNo == false)
                //{

                //}



                //// buat Subtitusi code //
                //if (isNefoKAM)
                //{
                //    strSQL = "select wsd_prd_master_code from so_web_sales_detail a " +
                //             "inner join TBL_PRD_SUBS b " +
                //             "   ON " +
                //             "       b.prds_entity_id = a.wsd_entity_id " +
                //             "       AND b.prds_branch_id = a.wsd_so_branch_id " +
                //             "       AND b.prds_parent = a.wsd_prd_master_code " +
                //             "       AND b.prds_grade = a.wsd_grade " +
                //             "       AND b.prds_size = a.wsd_prd_size " +
                //             "                       where " +
                //            "wsd_so_seq_no = '" + g_sDefaultSONo + "' and wsd_entity_id='" + xEntityID + "' and wsd_so_branch_id = '" + xBranchID + "'";

                //    DataTable dtCheckSubs = _clsGlobal.ExecDTTrans(strSQL);

                //    if (dtCheckSubs.Rows.Count > 0)
                //    {
                //        strSQL = "IP_PROSES_SUBTITUSI_KAM '" + g_sDefaultSONo + "','" + xEntityID + "','" + xBranchID + "', '" + clsLogin.USERID + "' ";
                //        _clsGlobal.ExecuteTrans(strSQL);

                //    }
                //}

                //'PROSES SO YANG SUDAH DI SAVE
                DataTable dtFill3 = new DataTable();
                if (rbCashDiscP.Checked == true)
                {
                    strSQL = " exec SP_PROSES_SO_MBRANCH " + "'1','" + g_sDefaultSONo + "','" + clsLogin.USERID + "','N','" + xEntityID + "','" + xBranchID + "'";
                }
                else //untuk disc yang di pilih cash disc uang
                {
                    strSQL = " exec SP_PROSES_SO_MBRANCH " + "'1','" + g_sDefaultSONo + "','" + clsLogin.USERID + "','Y','" + xEntityID + "','" + xBranchID + "'";
                }
                dtFill3 = _clsGlobal.ExecDTTrans(strSQL);

                strSQL = " SELECT * FROM VW_SO_DIFF_TAX where wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "' and wsh_seq_no = '" + g_sDefaultSONo + "' ";
                DataTable dtCheckDiffTax = _clsGlobal.ExecDTTrans(strSQL);
                if (dtCheckDiffTax.Rows.Count > 0)
                {
                    strSQL = "exec IP_RECALCULATE_TAX '" + xEntityID + "','" + xBranchID + "' , '" + g_sDefaultSONo + "' ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }

                pRefresh(false, true);
                //SetGridSODetail(dgvSalesDetail);
                //SetGridTPRB(dgvPromosiQty);
                //SetGridTPRU(dgvPromosiRp);
                //SetGridDisc(dgvDisc);
                //SetGridTax(dgvTax);
                //SetGridPromoTax(dgvPromoTax);
                pFillData(g_sDefaultSONo, "", "", true);

                //DataTable dtFill4 = new DataTable();
                /** deadlock SeqCalc **/


                if (isCalcNoLock)
                {
                    strSQL = " EXEC SP_DELETE_SO_TABLE_MBRANCH '" + g_sDefaultSONo + "','" + xEntityID + "','" + xBranchID + "'";
                }
                else
                {
                    strSQL = " EXEC SP_DELETE_SO_TABLE ";
                }

                //strSQL = " EXEC SP_DELETE_SO_TABLE ";
                _clsGlobal.ExecuteTrans(strSQL);

                _clsGlobal.RollbackTrans();
                if (bYesNo == false)
                {
                    frmSOManualEntryEditor2 frm = new frmSOManualEntryEditor2();
                    if (isNefoKAM)
                    {
                        frm.Query = "EXEC SP_EXEC_GET_TPR_KAM '" + g_sDefaultSONo + "','" + xEntityID + "','" + xBranchID + "'";
                    }
                    else
                    {
                        frm.Query = "EXEC SP_EXEC_GET_TPR '" + g_sDefaultSONo + "'";
                    }

                    frm.btnoke = false;
                    frm.btnNo = "&OK";
                    frm.ShowDialog();

                }

            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);

            }


            keyFound = true;

            return keyFound;
        }

        //1 juni 2018
        private bool fSaveSO(string sCredLimit, bool bCalcDisc)
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
            string shipmentType = "reguler";
            string addshipmentCost = "0";
            string extracttranship = "N";
            string runcodetranship = "";
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


            /* update yuda BL (delv so header) 07-09-2021 */
            string add1 = "";
            string add2 = "";
            string add3 = "";
            string add4 = "";
            string prov = "";
            string city = "";
            string kecamatan = "";
            string kelurahan = "";
            string phone = "";
            string participationcode = "";
            /***********************************/


            /** PARAMETER COLDCHAIN **/
            bool iscoldchain = false;
            string newordertype = "";
            string sCCOrderType = "";
            string sFlagColdChainOT = "N";
            /** END OF PARAMETER COLDCHAIN **/


            /** MULTISOURCE CANCEL HOLD^**/
            string countcancel = "0";
            bool ismaingudang = false;
            string ShippingPlant = "0";
            /** END MULTISOURCE CANCEL **/

            //Parameter DearOTC


            try
            {
                keyFound = false;

                pClearTax();
                //pClearTPR();
                pClearDiscount();
                pEmptyTPRFlag();
                DateTime _dtgl1 = DateTime.Parse(Convert.ToDateTime(dtTglOrder.EditValue).ToString(), CultureInfo.GetCultureInfo("en-US"));
                DateTime _dtglDelv = DateTime.Parse(Convert.ToDateTime(dtTglDelv.EditValue).ToString(), CultureInfo.GetCultureInfo("en-US"));
                DateTime _dtglExp = DateTime.Parse(Convert.ToDateTime(dtTglExp.EditValue).ToString(), CultureInfo.GetCultureInfo("en-US"));
                //USER PROCESS SDR34 (ROVI)
                string rovi_od_date = string.Empty;
                string user_rovi_od = string.Empty;

                //SDT25A
                string user_approval_od = string.Empty;
                // wsh_approval_rovi_by 
                // wsh_approval_rovi_date 

                //sdt19
                string usr_proc_order = string.Empty;
                string usr_proc_order_date = string.Empty;

                //SDR34A
                //wsh_user_rovi_cl
                //wsh_rovi_cl_date
                string user_rovi_cl = string.Empty;
                string rovi_cl_date = string.Empty;
                //SDT25B
                // wsh_approval_rovi_cl_by 
                // wsh_approval_rovi_cl_date
                // wsh_user_approval_cl
                string approval_rovi_cl_by = string.Empty;
                string approval_rovi_cl_date = string.Empty;
                string user_approval_cl = string.Empty;

                //SUPOM
                string source_site = string.Empty;
                //DEAROTC
                string cust_phone = string.Empty;
                string spgm_phone = string.Empty;
                string wa_staging = "N";
                string wa_customer = "N";
                string wa_salesman = "N";
                string wa_confirm_order = "Y";

                txtPromosiU.Text = "";
                txtDisc1.Text = "";
                if (rbCashDiscP.Checked == true)
                {
                    txtCashDiscU.Text = "";
                }
                if (rbCashDiscU.Checked == true)
                {
                    if (isNefoKAM)
                    {
                        txtCashDiscP.Text = "";
                    }

                }
                txtPPN.Text = "";
                txtTtlInvoice.Text = "";
                txtDPP.Text = "";


                /** CHECK COLDCHAIN ORDER TYPE **/
                string pcoderow1 = dgvSalesDetail.Rows[0].Cells["wsd_prd_master_code"].Value.ToString();
                strSQL = "SELECT * FROM IM_PRD_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + " INNER JOIN ( ";
                strSQL = strSQL + " select gh_function_code as cocline from GS_GEN_HARDCODED WITH (NOLOCK) ";
                strSQL = strSQL + " where gh_function_name like 'LINE_COLD_CHAIN' and gh_sys = 'H' ";
                strSQL = strSQL + " ) COC";
                strSQL = strSQL + " ON  prm_prd_line_code = cocline ";
                strSQL = strSQL + " WHERE  prm_prd_master_code = '" + pcoderow1 + "' ";

                DataTable dtCheck = _clsGlobal.ExecDTTrans(strSQL);
                if (dtCheck.Rows.Count > 0)
                {
                    iscoldchain = true;
                }

                if (iscoldchain)
                {
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
                        }
                    }

                }


                /** END OF CHECK COLDCHAIN ORDER TYPE **/


                if (bCalcDisc == false)
                {
                    if (clsGlobal.MODE_TRX == 1)//new
                    {
                        if (txtNoOrder.Text == g_sDefaultSONo)
                        {
                            if (fGetSONumber() == false)
                            {
                                keyFound = false;
                                return keyFound;
                            }
                            else
                            {
                                txtNoOrder.Text = sSONo;
                            }
                        }

                        DataTable dtNIK = new DataTable();
                        strSQL = "select sah_NIK from tbl_salesman_history WITH (NOLOCK)" +
                                 " where" +
                                 " convert(varchar(10), sah_from, 120) <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                                 " and convert(varchar(10), sah_to, 120) >= '" + DateTime.Now.ToString("yyyy-MM-dd") + "'" +
                                 " and sah_entity_id = '" + xEntityID + "'" +
                                 " and sah_branch_id = '" + xBranchID + "'" +
                                 " and sah_salesmen_id = '" + xSalesID + "'";
                        if (_clsGlobal.Connect.State == ConnectionState.Closed)
                            dtNIK = _clsGlobal.ExecDT(strSQL);
                        else
                            dtNIK = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtNIK.Rows.Count > 0)
                        {
                            nik = dtNIK.Rows[0]["sah_NIK"].ToString();
                        }

                        DataTable dtDearOTC = new DataTable();
                        strSQL = "";
                        strSQL = $@"select isnull(cm_dear_notification,'N') cm_dear_notification,ISNULL(cm_delv_phone,'') outlet_phone from SO_CUST_MASTER WITH (NOLOCK) 
                           WHERE isnull(cm_dear_notification,'N')='Y'
                           AND cm_entity='{xEntityID}' and cm_branch='{xBranchID}' and cm_cust_code1='{txtKdOutlet1.Text.ToString().Trim()}'";
                        dtDearOTC = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtDearOTC.Rows.Count > 0)
                        {
                            NO_HP_Outlet = dtDearOTC.Rows[0]["outlet_phone"].ToString();
                        }

                        //DEAR OTC Set Nomor Salesman

                        DataTable dtSalesman = new DataTable();
                        strSQL = "";
                        strSQL = $@"select ISNULL(sgm_hpno,'') as no_hp_salesman from SO_SPG_GIRL_MAN WITH (NOLOCK) 
                           WHERE sgm_entity_id='{xEntityID}' and sgm_branch_id='{xBranchID}' and sgm_spgm_id='{txtSalesID.Text.ToString().Trim()}'";

                        dtSalesman = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtSalesman.Rows.Count > 0)
                        {
                            NO_HP_Salesman = dtSalesman.Rows[0]["no_hp_salesman"].ToString().Trim();
                        }

                    }
                    else if (clsGlobal.MODE_TRX == 2) //edit
                    {
                        if (txtNoOrder.Text != g_sDefaultSONo)
                        {
                            //CHECK SO STATUS
                            DataTable dtFill2 = new DataTable();
                            strSQL = " select gh_function_code";
                            strSQL = strSQL + " from GS_GEN_HARDCODED WITH (NOLOCK)  ";
                            strSQL = strSQL + " where gh_sys = 'H'";
                            strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
                            strSQL = strSQL + " and gh_sequence_no = 1";
                            dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                            if (dtFill2.Rows.Count > 0)
                            {
                                sOpenStatus = dtFill2.Rows[0]["gh_function_code"].ToString();
                            }
                            else
                            {
                                MessageBox.Show("Status Open tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                            //DEAR OTC Set Nomor Outlet & Salesman
                            DataTable dtDearOTC = new DataTable();
                            strSQL = "";
                            strSQL = $@"select isnull(cm_dear_notification,'N') cm_dear_notification,ISNULL(cm_delv_phone,'') outlet_phone from SO_CUST_MASTER WITH (NOLOCK) 
                                WHERE isnull(cm_dear_notification,'N')='Y'
                                AND cm_entity='{xEntityID}' and cm_branch='{xBranchID}' and cm_cust_code1='{txtKdOutlet1.Text.ToString().Trim()}'";
                            dtDearOTC = _clsGlobal.ExecDTTrans(strSQL);
                            if (dtDearOTC.Rows.Count > 0)
                            {
                                NO_HP_Outlet = dtDearOTC.Rows[0]["outlet_phone"].ToString();
                            }
                            DataTable dtSalesman = new DataTable();
                            strSQL = "";
                            strSQL = $@"select ISNULL(sgm_hpno,'') as no_hp_salesman from SO_SPG_GIRL_MAN WITH (NOLOCK) 
                           WHERE sgm_entity_id='{xEntityID}' and sgm_branch_id='{xBranchID}' and sgm_spgm_id='{txtSalesID.Text.ToString().Trim()}'";

                            dtSalesman = _clsGlobal.ExecDTTrans(strSQL);
                            if (dtSalesman.Rows.Count > 0)
                            {
                                NO_HP_Salesman = dtSalesman.Rows[0]["no_hp_salesman"].ToString().Trim();
                            }


                            DataTable dtFill3 = new DataTable();
                            strSQL = "select wsh_seq_no,isnull(wsh_pda_no,'') wsh_pda_no,isnull(wsh_pda_flaq,'') wsh_pda_flaq, isnull(wsh_extract_sap,'') wsh_extract_sap, isnull(wsh_runcode_ext,'') wsh_runcode_ext, isnull(wsh_qasir_flag,'') wsh_qasir_flag,wsh_shipment_type,wsh_additional_ship_cost,wsh_extract_tranship,wsh_runcode_tranship,wsh_nik,wsh_delv_address1,wsh_delv_address2,wsh_delv_address3,wsh_delv_address4,wsh_delv_proviency,wsh_delv_city,wsh_delv_kecamatan,wsh_delv_kelurahan,wsh_delv_phone,wsh_participation_code,isnull(wsh_count_cancel,0) as wsh_count_cancel,ot_def_oprtype ";
                            //USER PROCESS (SDR34) & SDT25A
                            strSQL += " ,wsh_rovi_od_date,wsh_user_rovi_od,wsh_user_approval_od";

                            //User Proses Order (SDT19)
                            strSQL += " ,wsh_usr_proc_order,wsh_usr_proc_order_date";

                            // SDR34A 
                            strSQL += " ,wsh_approval_rovi_cl_by,wsh_approval_rovi_cl_date";
                            // SDT25B
                            strSQL += " , wsh_user_rovi_cl,wsh_rovi_cl_date,wsh_user_approval_cl";

                            //SUPOM
                            strSQL += " , ISNULL(wsh_source_site,'') wsh_source_site ";
                            //DEAR OTC
                            strSQL += @" ,isnull(wsh_cust_phone,'') wsh_cust_phone
                                         ,ISNULL(wsh_spgm_phone,'') wsh_spgm_phone 
                                         ,ISNULL(wsh_wa_staging,'N') wsh_wa_staging 
                                         ,ISNULL(wsh_wa_customer,'N') wsh_wa_customer
                                         ,ISNULL(wsh_wa_salesman,'N') wsh_wa_salesman
                                         ,ISNULL(wsh_wa_confirm_order,'Y') wsh_wa_confirm_order";

                            strSQL = strSQL + " from  SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                            strSQL = strSQL + " LEFT JOIN SO_ORDER_TYPE WITH (NOLOCK) ON wsh_so_type = ot_order_type ";
                            strSQL = strSQL + " WHERE wsh_seq_no = '" + txtNoOrder.Text + "'";
                            strSQL = strSQL + "  AND wsh_status_so = '" + sOpenStatus + "' AND wsh_entity_id = '" + xEntityID + "' AND wsh_branch_id = '" + xBranchID + "'";
                            dtFill3 = _clsGlobal.ExecDTTrans(strSQL);
                            if (dtFill3.Rows.Count > 0)
                            {
                                pdaNO = dtFill3.Rows[0]["wsh_pda_no"].ToString();
                                pdaFlag = dtFill3.Rows[0]["wsh_pda_flaq"].ToString();
                                extsap = dtFill3.Rows[0]["wsh_extract_sap"].ToString();
                                runcext = dtFill3.Rows[0]["wsh_runcode_ext"].ToString();
                                QasirFlag = dtFill3.Rows[0]["wsh_qasir_flag"].ToString();
                                shipmentType = dtFill3.Rows[0]["wsh_shipment_type"].ToString();
                                addshipmentCost = dtFill3.Rows[0]["wsh_additional_ship_cost"].ToString();
                                extracttranship = dtFill3.Rows[0]["wsh_extract_tranship"].ToString();
                                runcodetranship = dtFill3.Rows[0]["wsh_runcode_tranship"].ToString();
                                nik = dtFill3.Rows[0]["wsh_nik"].ToString();


                                /* update ticket 22152 */
                                add1 = dtFill3.Rows[0]["wsh_delv_address1"].ToString();
                                add2 = dtFill3.Rows[0]["wsh_delv_address2"].ToString();
                                add3 = dtFill3.Rows[0]["wsh_delv_address3"].ToString();
                                add4 = dtFill3.Rows[0]["wsh_delv_address4"].ToString();
                                prov = dtFill3.Rows[0]["wsh_delv_proviency"].ToString();
                                city = dtFill3.Rows[0]["wsh_delv_city"].ToString();
                                kelurahan = dtFill3.Rows[0]["wsh_delv_kelurahan"].ToString();
                                kecamatan = dtFill3.Rows[0]["wsh_delv_kecamatan"].ToString();
                                phone = dtFill3.Rows[0]["wsh_delv_phone"].ToString();
                                /*******END**********/

                                /********* UPDATE BL VOUCHER NO ***********/
                                participationcode = dtFill3.Rows[0]["wsh_participation_code"].ToString();
                                /********* END ***************************/
                                countcancel = dtFill3.Rows[0]["wsh_count_cancel"].ToString();
                                string __otdef = dtFill3.Rows[0]["ot_def_oprtype"].ToString().Trim();
                                if (!cbResetReq.Checked)
                                {
                                    if (__otdef.ToString().Trim() == "T" && extsap.Trim().ToString() == "Y")
                                    {
                                        /** Gudang Utama **/
                                        strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name like 'MAINWHLOC' and gh_sys = 'H'";
                                        DataTable dtWHmain = _clsGlobal.ExecDTTrans(strSQL);
                                        if (dtWHmain.Rows.Count > 0)
                                        {
                                            string whloc = dtWHmain.Rows[0]["gh_function_code"].ToString().Trim();
                                            lblWHLoc1.Text = whloc.Split('|')[0].ToString().Trim();
                                            lblWHLoc2.Text = whloc.Split('|')[1].ToString().Trim();

                                        }
                                        else
                                        {

                                            lblWHLoc1.Text = "PST";
                                            lblWHLoc2.Text = "000";
                                        }

                                    }

                                    extsap = "N";
                                    runcext = "";
                                    countcancel = "0";
                                    ismaingudang = true;
                                }

                                //USER PROCESS SDR34 & SDT25A
                                rovi_od_date = dtFill3.Rows[0]["wsh_rovi_od_date"] != DBNull.Value ? Convert.ToDateTime(dtFill3.Rows[0]["wsh_rovi_od_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;
                                user_rovi_od = string.IsNullOrEmpty(dtFill3.Rows[0]["wsh_user_rovi_od"].ToString()) ? null : dtFill3.Rows[0]["wsh_user_rovi_od"].ToString();
                                user_approval_od = string.IsNullOrEmpty(dtFill3.Rows[0]["wsh_user_approval_od"].ToString()) ? null : dtFill3.Rows[0]["wsh_user_approval_od"].ToString();

                                //SDT19
                                usr_proc_order = string.IsNullOrEmpty(dtFill3.Rows[0]["wsh_usr_proc_order"].ToString()) ? null : dtFill3.Rows[0]["wsh_usr_proc_order"].ToString();
                                usr_proc_order_date = dtFill3.Rows[0]["wsh_usr_proc_order_date"] != DBNull.Value ? Convert.ToDateTime(dtFill3.Rows[0]["wsh_usr_proc_order_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;



                                //SDR34A                     
                                user_rovi_cl = string.IsNullOrEmpty(dtFill3.Rows[0]["wsh_user_rovi_cl"].ToString()) ? null : dtFill3.Rows[0]["wsh_user_rovi_cl"].ToString();
                                rovi_cl_date = dtFill3.Rows[0]["wsh_rovi_cl_date"] != DBNull.Value ? Convert.ToDateTime(dtFill3.Rows[0]["wsh_rovi_cl_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;

                                //SDT25B
                                approval_rovi_cl_by = string.IsNullOrEmpty(dtFill3.Rows[0]["wsh_approval_rovi_cl_by"].ToString()) ? null : dtFill3.Rows[0]["wsh_approval_rovi_cl_by"].ToString();
                                approval_rovi_cl_date = dtFill3.Rows[0]["wsh_approval_rovi_cl_date"] != DBNull.Value ? Convert.ToDateTime(dtFill3.Rows[0]["wsh_approval_rovi_cl_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;
                                user_approval_cl = string.IsNullOrEmpty(dtFill3.Rows[0]["wsh_user_approval_cl"].ToString()) ? null : dtFill3.Rows[0]["wsh_user_approval_cl"].ToString();

                                //SUPOM 
                                source_site = dtFill3.Rows[0]["wsh_source_site"].ToString();
                                //DearOTC                                
                                wa_staging = dtFill3.Rows[0]["wsh_wa_staging"].ToString();
                                wa_customer = dtFill3.Rows[0]["wsh_wa_customer"].ToString();
                                wa_salesman = dtFill3.Rows[0]["wsh_wa_salesman"].ToString();
                                wa_confirm_order = dtFill3.Rows[0]["wsh_wa_confirm_order"].ToString();

                            }
                            else
                            {
                                if (_xFlagDC != "Y")
                                {
                                    MessageBox.Show("Status SO sudah bukan pada Open ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }
                                else
                                {
                                    MessageBox.Show("Data berhasil disimpan " + " No Order : " + txtNoOrder.Text, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    pRefresh(true, false);
                                    //SetGridSODetail(dgvSalesDetail);
                                    //SetGridTPRB(dgvPromosiQty);
                                    //SetGridTPRU(dgvPromosiRp);
                                    //SetGridDisc(dgvDisc);
                                    //SetGridTax(dgvTax);
                                    //SetGridPromoTax(dgvPromoTax);
                                }
                                keyFound = false;
                                return keyFound;
                            }

                            ///DataTable DTHeader = _clsGlobal.ExecDTTrans("Select * from so_web_sales_header with(nolock) WHERE wsh_seq_no = '" + txtNoOrder.Text + "' AND wsh_entity_id = '" + xEntityID + "' AND wsh_branch_id = '" + xBranchID + "'";

                            //DataTable dtFill4 = new DataTable();
                            strSQL = "delete D FROM SO_WEB_SALES_DETAIL D WITH (NOLOCK) ";
                            strSQL = strSQL + " WHERE wsd_so_seq_no = '" + txtNoOrder.Text + "' AND wsd_entity_id = '" + xEntityID + "' AND wsd_so_branch_id = '" + xBranchID + "'";
                            //strSQL = strSQL + "  AND wsh_branch_id = '" +  + "'"; // di tambahkan branch
                            _clsGlobal.ExecuteTrans(strSQL);

                            //DataTable dtFill5 = new DataTable();
                            strSQL = "delete H FROM SO_WEB_SALES_HEADER H WITH (NOLOCK) ";
                            strSQL = strSQL + " WHERE wsh_seq_no = '" + txtNoOrder.Text + "' AND wsh_entity_id = '" + xEntityID + "' AND wsh_branch_id = '" + xBranchID + "'";
                            //strSQL = strSQL + "  AND wsh_branch_id = '" +  + "'"; // di tambahkan branch
                            _clsGlobal.ExecuteTrans(strSQL);

                        }
                        else
                        {
                            if (fGetSONumber() == false)
                            {
                                keyFound = false;
                                return keyFound;
                            }

                            txtNoOrder.Text = sSONo;
                        }

                        sSONo = txtNoOrder.Text;
                    }
                }
                else
                {
                    sSONo = g_sDefaultSONo;

                    DataTable dtFill6 = new DataTable();
                    strSQL = "select wsh_seq_no,isnull(wsh_pda_no,'') wsh_pda_no,isnull(wsh_pda_flaq,'') wsh_pda_flaq, isnull(wsh_extract_sap,'') wsh_extract_sap, isnull(wsh_runcode_ext,'') wsh_runcode_ext, isnull(wsh_qasir_flag,'') wsh_qasir_flag,wsh_shipment_type,wsh_additional_ship_cost,wsh_extract_tranship,wsh_runcode_tranship from  SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                    strSQL = strSQL + " WHERE wsh_seq_no = '" + txtNoOrder.Text + "' AND wsh_entity_id = '" + xEntityID + "' AND wsh_branch_id = '" + xBranchID + "'";
                    dtFill6 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill6.Rows.Count > 0)
                    {
                        pdaNO = dtFill6.Rows[0]["wsh_pda_no"].ToString();
                        pdaFlag = dtFill6.Rows[0]["wsh_pda_flaq"].ToString();
                        extsap = dtFill6.Rows[0]["wsh_extract_sap"].ToString();
                        runcext = dtFill6.Rows[0]["wsh_runcode_ext"].ToString();
                        QasirFlag = dtFill6.Rows[0]["wsh_qasir_flag"].ToString();
                        shipmentType = dtFill6.Rows[0]["wsh_shipment_type"].ToString();
                        addshipmentCost = dtFill6.Rows[0]["wsh_additional_ship_cost"].ToString();
                        extracttranship = dtFill6.Rows[0]["wsh_extract_tranship"].ToString();
                        runcodetranship = dtFill6.Rows[0]["wsh_runcode_tranship"].ToString();

                    }

                    //DataTable dtFill7 = new DataTable();

                    //strSQL = " EXEC SP_DELETE_SO_TABLE ";
                    /** deadlock SeqCalc **/
                    if (bCalcDisc == false)
                    {
                        if (isCalcNoLock)
                        {
                            strSQL = " EXEC SP_DELETE_SO_TABLE_MBRANCH '" + g_sDefaultSONo + "','" + xEntityID + "','" + xBranchID + "'";
                        }
                        else
                        {
                            strSQL = " EXEC SP_DELETE_SO_TABLE ";
                        }
                        _clsGlobal.ExecuteTrans(strSQL);

                    }


                }

                if (sSONo.Length == 0 || txtNoOrder.Text.Length == 0)
                {
                    MessageBox.Show("No SO tidak Berhasil di Generate. " + System.Environment.NewLine + "Silahkan dicoba kembali untuk Entry Ulang.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }


                //'Get branch, area, wilayah, rayon, cust group dan line code
                DataTable dtFill8 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') cm_line_code, cm_lead_time From SO_CUST_MASTER WITH (NOLOCK) ";
                strSQL = strSQL + " where cm_cust_code1 = '" + txtKdOutlet1.Text + "' AND";
                strSQL = strSQL + " cm_cust_code2 = '" + lblKdOutlet2.Text + "' AND cm_entity = '" + xEntityID + "' and cm_branch = '" + xBranchID + "' ";
                dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill8.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill8.Rows[0]["cm_branch"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Area untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill8.Rows[0]["cm_area"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Wilayah untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill8.Rows[0]["cm_wilayah"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Rayon untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill8.Rows[0]["cm_rayon"].ToString();
                    }

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Cust group untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill8.Rows[0]["cm_cust_group"].ToString();
                    }

                    sLineCode = dtFill8.Rows[0]["cm_line_code"].ToString();

                    if (String.IsNullOrEmpty(dtFill8.Rows[0]["cm_lead_time"].ToString()))
                    {
                        MessageBox.Show("Lead Time untuk Customer (" + txtKdOutlet1.Text + " ) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        lLeadTime = Convert.ToInt64(dtFill8.Rows[0]["cm_lead_time"].ToString());
                    }

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
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK)  ";

                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + _dtgl1.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '" + xEntityID + "' and fwd_branch = '" + xBranchID + "'"; //entity id harus di lempar
                dtFill9 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill9.Rows.Count > 0)
                {
                    sWeekNo = dtFill9.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill9.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill9.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + _dtgl1.ToString() + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get SO Status
                DataTable dtFill10 = new DataTable();
                strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  ";
                strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                if (bCalcDisc == false)
                {
                    strSQL = strSQL + " and gh_sequence_no = '1' ";
                }
                else
                {
                    strSQL = strSQL + " and gh_sequence_no = '3' ";
                }
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
                //strSQL = " exec SP_GET_PRICE_CODE '" + xEntityID + "','" + xBranchID + "','" + txtKdOutlet1.Text + "','" + lblKdOutlet2.Text + "'";
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
                //    MessageBox.Show("Default Price Code untuk Customer " + txtKdOutlet1.Text + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    keyFound = false;
                //    return keyFound;
                //}

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                {
                    if (dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() != "")
                    {
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + Convert.ToInt64(_fConvertQty);
                        curROrdAmt = curROrdAmt + Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["ijumlahharga"].Value.ToString());
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lOrdQty = lOrdQty + Convert.ToInt64(_fConvertQty);
                        curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["ijumlahharga"].Value.ToString());

                        if (fGetCost(dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString(), dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString(), dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        curTotalCost = curTotalCost + _curCost;
                    }
                }

                if (sSONo.Length == 0)
                {
                    MessageBox.Show("nomor SO Blank", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                if (_xFlagDC == "N")
                {
                    // '===> Cek Kredit Limit dihilangkan jika flagDC = Y - Project Improvement TiraSnD
                    //'/*-----------------------------------------------------------------------------------------
                    //'Credit Limit
                    if (fGetCreditLimit(txtKdOutlet1.Text, lblKdOutlet2.Text) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }

                    if (bFlagCr == true)
                    {
                        if (_curAvCredLimit < curROrdAmt)
                        {
                            //sCredLimit = "*";
                        }
                        else
                        {
                            sCredLimit = "";
                        }
                    }
                }

                //DataTable dtFill12 = new DataTable();
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
                if (_xFlagDC == "Y")
                {
                    if (strFlagLimit != "")
                    {
                        strSQL = strSQL + " ,wsh_flag_so_limit ";
                    }
                }
                else
                {
                    strSQL = strSQL + " ,wsh_flag_so_limit ";
                }
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                if (pdaFlag == "" || pdaNO == "")
                {

                }
                else
                {
                    strSQL = strSQL + " , wsh_pda_no ";
                    strSQL = strSQL + " , wsh_pda_flaq  ";
                }

                if (extsap == "N" || runcext == "")
                {

                }
                else
                {
                    strSQL = strSQL + " , wsh_extract_sap ";
                    strSQL = strSQL + " , wsh_runcode_ext  ";
                }
                //Flag Multisource Y
                if (isMultisource)
                {
                    strSQL = strSQL + " , wsh_qasir_flag  ";
                }
                else
                {
                    if (QasirFlag != "")
                    {
                        strSQL = strSQL + " , wsh_qasir_flag  ";
                    }
                }

                strSQL = strSQL + " , wsh_shipment_type ";
                strSQL = strSQL + " , wsh_additional_ship_cost ";
                strSQL = strSQL + " , wsh_extract_tranship ";
                strSQL = strSQL + " , wsh_runcode_tranship ";
                strSQL = strSQL + " , wsh_sch_date ";
                strSQL = strSQL + " , wsh_expired_date ";
                strSQL = strSQL + " , wsh_beban_cashdisc ";
                strSQL = strSQL + " , wsh_nik ";
                /* update ticket 22152 */
                strSQL = strSQL + " , wsh_delv_address1 ";
                strSQL = strSQL + " , wsh_delv_address2 ";
                strSQL = strSQL + " , wsh_delv_address3 ";
                strSQL = strSQL + " , wsh_delv_address4 ";
                strSQL = strSQL + " , wsh_delv_proviency ";
                strSQL = strSQL + " , wsh_delv_city ";
                strSQL = strSQL + " , wsh_delv_kecamatan ";
                strSQL = strSQL + " , wsh_delv_kelurahan ";
                strSQL = strSQL + " , wsh_delv_phone ";
                /*******END**********/

                /* update ticket 23816 */
                strSQL = strSQL + " , wsh_note  ";
                /******* END **********/


                /* Multisource */
                if (isMultisource)
                {
                    strSQL = strSQL + " , wsh_flag_sloc";
                }

                strSQL = strSQL + " , wsh_count_cancel ";

                if (isCrossSite && !String.IsNullOrEmpty(txtShippingPlantId.Text))
                {
                    strSQL = strSQL + " , wsh_ship_plant, wsh_cross_site ";
                }
                strSQL = strSQL + " , wsh_is_fulfillment,wsh_sled_flag,wsh_sled";



                //History Delivey Date dan Expired Date
                if (TglDeliveryDate.ToString("yyyyMMdd") != _dtglDelv.ToString("yyyyMMdd"))
                {
                    strSQL = strSQL + " , wsh_sch_date_update_by,wsh_sch_date_update";
                }
                if (TglExpiredDate.ToString("yyyyMMdd") != _dtglExp.ToString("yyyyMMdd"))
                {
                    strSQL = strSQL + " , wsh_expired_date_by,wsh_expired_date_update";
                }

                //SDR34 & SDT25A
                strSQL = strSQL + " ,wsh_rovi_od_date,wsh_user_rovi_od,wsh_user_approval_od";

                //SDT19
                strSQL = strSQL + " , wsh_usr_proc_order ";
                strSQL = strSQL + " , wsh_usr_proc_order_date ";

                //SDR34A
                strSQL = strSQL + " , wsh_user_rovi_cl ";
                strSQL = strSQL + " , wsh_rovi_cl_date ";

                //SDT25B
                strSQL = strSQL + " , wsh_approval_rovi_cl_by ";
                strSQL = strSQL + " , wsh_approval_rovi_cl_date ";
                strSQL = strSQL + " , wsh_user_approval_cl ";

                //SUPOM 
                strSQL = strSQL + " , wsh_source_site ";
                //DearOTC
                strSQL = strSQL + @",wsh_cust_phone
                                    ,wsh_spgm_phone
                                    ,wsh_wa_staging
                                    ,wsh_wa_customer
                                    ,wsh_wa_salesman
                                    ,wsh_wa_confirm_order  ";


                /***** end multisource *****/
                string __tglorder = _dtgl1.ToString("yyyyMMdd");
                int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + _xEntityID + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + _xBranchID + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                if (sFlagColdChainOT.Trim().Equals("Y"))
                {
                    strSQL = strSQL + sCCOrderType + "' , '";  //'wsh_so_type -- coldchain parameter diisi Y 
                }
                else
                {
                    strSQL = strSQL + lblOrderType.Text + "' , '";  //'wsh_so_type
                }

                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no
                strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                strSQL = strSQL + txtKdOutlet1.Text + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + lblKdOutlet2.Text + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + lblWHLoc1.Text + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + lblWHLoc2.Text + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + txtSalesID.Text + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "0 , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "0 , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "0 , '"; //'wsh_tot_part_line_amount
                strSQL = strSQL + ttlline + "', "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                if (bCalcDisc == false)
                {
                    strSQL = strSQL + "0 , "; //'wsh_total_so_qty
                    strSQL = strSQL + "0 , "; //'wsh_total_so_amt
                }
                else
                {
                    strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                    strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                }
                strSQL = strSQL + "null , "; //'wsh_include_tax
                strSQL = strSQL + "0 , "; //'[wsh_tax_%]
                strSQL = strSQL + "0 , "; //'wsh_total_partisipasi_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_total_margin_disc_amount
                strSQL = strSQL + "0 , "; //'wsh_net_amount
                strSQL = strSQL + "0 , "; //'wsh_dpp_amount
                strSQL = strSQL + "0 , "; //'wsh_tax_amount
                strSQL = strSQL + "'N' , "; //'wsh_invoice_process_flag
                strSQL = strSQL + "null , "; //'wsh_invoice_no
                strSQL = strSQL + "null , "; //'wsh_invoice_date
                strSQL = strSQL + "null , "; //'wsh_invoice_process_by
                strSQL = strSQL + "null , "; //'wsh_disc_margin
                strSQL = strSQL + "0 , '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , "; //'wsh_approval
                strSQL = strSQL + "null , "; //'wsh_approval_by
                strSQL = strSQL + "null , "; //'wsh_approval_date
                strSQL = strSQL + "'" + txtRefNoRetur.Text + "', "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                /******** UPDATE BL VOUCHER NO **********/
                // strSQL = strSQL + "null , "; //'wsh_participation_code
                strSQL = strSQL + "'" + participationcode + "' , "; //'wsh_participation_code
                /******** END UPDATE BL VOUCHER NO ******/
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
                strSQL = strSQL + "0 , "; //'wsh_tax_pct
                if (txtCashDiscP.Text == "")
                {
                    strSQL = strSQL + "0 , "; //'wsh_pct_cash_disc
                }
                else
                {
                    strSQL = strSQL + txtCashDiscP.Text.Replace(",", "") + " , "; //'wsh_pct_cash_disc
                }
                if (txtCashDiscU.Text == "")
                {
                    strSQL = strSQL + "0 , "; //'wsh_cash_disc
                }
                else
                {
                    strSQL = strSQL + txtCashDiscU.Text.Replace(",", "") + " , "; //'wsh_cash_disc
                }
                //strSQL = strSQL + "0 , "; //'wsh_cash_disc
                strSQL = strSQL + "0 , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "0 , '"; //'wsh_tax_pbm_amount
                //DataRow selectedDataRow = ((DataRowView)cbPembyFaktur.GetSelectedDataRow()).Row;
                string idO = Convert.ToString(cbPembyFaktur.EditValue);
                strSQL = strSQL + idO + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , '"; //'wsh_gross_net_amt
                strSQL = strSQL + Convert.ToDateTime(dtTglValidasi.EditValue).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + txtNoPO.Text + "' , '"; //'wsh_po_no
                if (xFlagDC == "Y")
                {
                    if (strFlagLimit != "")
                    {
                        if (clsGlobal.MODE_TRX == 2)
                        {
                            strSQL = strSQL + "' , '"; //'wsh_flag_so_limit
                        }
                        else
                        {
                            strSQL = strSQL + strFlagLimit + "' , '"; //'wsh_flag_so_limit
                        }
                    }
                }
                else
                {
                    strSQL = strSQL + sCredLimit + "' , '"; //'wsh_flag_so_limit
                }
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                string tops = cbTOP.Text;
                string stops = ExtractTopDaysSafe(tops);
                strSQL = strSQL + stops + "' , '"; //'wsh_top
                DataRow selectedDataRow1 = ((DataRowView)cbTOP.GetSelectedDataRow()).Row;
                string idOO = Convert.ToString(selectedDataRow1["Code"]);
                strSQL = strSQL + idOO + "' , '"; //'wsh_top_id
                strSQL = strSQL + Convert.ToDateTime(dtTglPO.EditValue).ToString("yyyy-MM-dd") + "','"; //'wsh_po_date
                strSQL = strSQL + lLeadTime + "'"; //'wsh_cust_leadtime
                if (pdaFlag == "" || pdaNO == "")
                {
                    //strSQL = strSQL + "' , '"; 
                    //strSQL = strSQL + "' , '"; 
                }
                else
                {
                    strSQL = strSQL + ",'" + pdaNO + "' ,'"; //'pdano
                    strSQL = strSQL + pdaFlag + "'"; //'pdaflag
                }
                if (extsap == "" || runcext == "")
                {
                    //strSQL = strSQL + "' , '"; //'extsap
                    //strSQL = strSQL +  "' , '"; //'runcext
                }
                else
                {
                    strSQL = strSQL + ",'" + extsap + "' , '"; //'extsap
                    strSQL = strSQL + runcext + "'"; //'runcext
                }
                //Flag Multisource Y
                if (isMultisource)
                {
                    strSQL = strSQL + ",'" + Convert.ToString(cbSource.EditValue) + "' "; //'qasirflag
                }
                else
                {
                    if (QasirFlag != "")
                    {
                        strSQL = strSQL + ",'" + QasirFlag + "' "; //'qasirflag
                    }
                }

                strSQL = strSQL + ",'" + shipmentType + "' "; //shipmentType
                strSQL = strSQL + ",'" + addshipmentCost + "' "; //Add Shipment Cost
                strSQL = strSQL + ",'" + extracttranship + "' "; //extract tranship
                strSQL = strSQL + ",'" + runcodetranship + "' "; //runcode tranship
                strSQL = strSQL + ",'" + _dtglDelv.ToString("yyyyMMdd") + "' "; //planning date
                strSQL = strSQL + ",'" + _dtglExp.ToString("yyyyMMdd") + "' "; //expired date
                string _beban = "0";
                if (rbCashDiscP.Checked)
                {
                    if (!string.IsNullOrEmpty(txtCashDiscP.Text.Trim().ToString()))
                    {
                        if (Convert.ToDecimal(txtCashDiscP.Text) > 0)
                        {
                            _beban = Convert.ToString(cmbBebanDisc.EditValue);
                        }
                    }

                }

                if (rbCashDiscU.Checked)
                {
                    if (!string.IsNullOrEmpty(txtCashDiscU.Text.Trim().ToString()))
                    {
                        if (Convert.ToDecimal(txtCashDiscU.Text) > 0)
                        {
                            _beban = Convert.ToString(cmbBebanDisc.EditValue);
                        }
                    }

                }

                strSQL = strSQL + ",'" + _beban + "' "; //beban disc
                strSQL = strSQL + ",'" + nik + "' "; // NIK dari salesman
                /* update ticket 22152 */
                strSQL = strSQL + ",'" + add1 + "' "; // wsh_delv_address1
                strSQL = strSQL + ",'" + add2 + "' "; // wsh_delv_address2
                strSQL = strSQL + ",'" + add3 + "' "; // wsh_delv_address3
                strSQL = strSQL + ",'" + add4 + "' "; // wsh_delv_address4
                strSQL = strSQL + ",'" + prov + "' "; // wsh_delv_proviency
                strSQL = strSQL + ",'" + city + "' "; // wsh_delv_city
                strSQL = strSQL + ",'" + kecamatan + "' "; // wsh_delv_kecamatan
                strSQL = strSQL + ",'" + kelurahan + "' "; // wsh_delv_kelurahan
                strSQL = strSQL + ",'" + phone + "' "; // wsh_delv_phone
                /*******END**********/
                ///* update ticket 23816 */
                strSQL = strSQL + ",'" + txtNote.Text + "' "; // wsh_note ";

                /******* END **********/

                /***** Multisource ***/
                if (isMultisource)
                {
                    strSQL = strSQL + ",'" + Convert.ToString(cbFlagSloc.EditValue) + "' "; // wsh_flag_sloc ";
                }

                strSQL = strSQL + ",'" + countcancel + "' "; // wsh_count_cancel ";

                if (isCrossSite && !String.IsNullOrEmpty(txtShippingPlantId.Text))
                {
                    strSQL = strSQL + ",'" + txtShippingPlantId.Text + "', (SELECT pms_cross_site FROM GS_PRM_MULTI_SOURCE WITH(NOLOCK) WHERE pms_branch_id = '" + xBranchID + "' AND pms_ss_site = '" + txtShippingPlantId.Text + "')"; // wsh_ship_plant ";
                }
                string __isfullfilment = "N";
                if (cbfullfilment.Checked)
                {
                    __isfullfilment = "Y";
                }
                strSQL = strSQL + ",'" + __isfullfilment + "' "; // wsh_count_cancel ";

                if (chkSLED.Checked)
                {
                    strSQL = strSQL + ",'Y',NULL "; // wsh_sled_flag, wsh_sled ";
                }
                else
                {
                    strSQL = strSQL + ",'N',NULL "; // wsh_sled_flag, wsh_sled ";
                }

                //History Delivery Date & Expired Date
                if (TglDeliveryDate.ToString("yyyyMMdd") != _dtglDelv.ToString("yyyyMMdd"))
                {
                    strSQL = strSQL + " , '" + clsLogin.USERID + "',GETDATE() ";
                }

                if (TglExpiredDate.ToString("yyyyMMdd") != _dtglExp.ToString("yyyyMMdd"))
                {
                    strSQL = strSQL + " , '" + clsLogin.USERID + "',GETDATE() ";
                }

                //USER PROCESS
                if (!string.IsNullOrEmpty(rovi_od_date))
                {
                    strSQL = strSQL + " , '" + rovi_od_date + "' "; // wsh_rovi_od_date
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_rovi_od_date
                }

                strSQL = strSQL + " , '" + user_rovi_od + "' "; // wsh_user_rovi_od   
                strSQL = strSQL + " , '" + user_approval_od + "' "; // wsh_user_approval_od 
                //SDT19
                strSQL = strSQL + " , '" + usr_proc_order + "' "; // wsh_usr_proc_order
                if (!string.IsNullOrEmpty(usr_proc_order_date))
                {
                    strSQL = strSQL + " ,'" + usr_proc_order_date + "'"; // wsh_usr_proc_order_date
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_rovi_od_date
                }


                //SDR34A            
                strSQL = strSQL + " , '" + user_rovi_cl + "' "; // wsh_user_rovi_cl
                if (!string.IsNullOrEmpty(rovi_cl_date))
                {
                    strSQL = strSQL + " ,'" + rovi_cl_date + "'"; // wsh_rovi_cl_date
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_rovi_cl_date
                }


                //SDT25B
                strSQL = strSQL + " , '" + approval_rovi_cl_by + "' "; // wsh_approval_rovi_cl_by
                if (!string.IsNullOrEmpty(approval_rovi_cl_date))
                {
                    strSQL = strSQL + " ,'" + approval_rovi_cl_date + "'"; // wsh_approval_rovi_cl_date
                }
                else
                {
                    strSQL = strSQL + " ,NULL"; // wsh_approval_rovi_cl_date
                }

                strSQL = strSQL + " , '" + user_approval_cl + "' "; // wsh_user_approval_cl

                //SUPOM 
                strSQL = strSQL + " , '" + source_site + "' "; // wsh_source_site
                //DEAR OTC
                strSQL = strSQL + " , '" + NO_HP_Outlet + "' "; // wsh_cust_phone
                strSQL = strSQL + " , '" + NO_HP_Salesman + "' "; // wsh_spgm_phone
                strSQL = strSQL + " , '" + wa_staging + "' "; // wsh_wa_staging
                strSQL = strSQL + " , '" + wa_customer + "' "; // wsh_wa_customer
                strSQL = strSQL + " , '" + wa_salesman + "' "; // wsh_wa_salesman
                strSQL = strSQL + " , '" + wa_confirm_order + "' "; // wsh_wa_confirm_order

                /***** end Multisource **/

                strSQL = strSQL + " ) ";
                _clsGlobal.ExecuteTrans(strSQL);
                //'END OF Saving SO Header
                //if (dtFill12.Rows.Count > 0)
                //{
                DataTable dtFill13 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK)  ";
                strSQL = strSQL + " where cc_currency_type = 'H' ";
                dtFill13 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill13.Rows.Count > 0)
                {
                    sCurrCode = dtFill13.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill13.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }

                //lCounter = 0;
                for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                {
                    if (dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString() != "")
                    {
                        DataTable dtFill14 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvSalesDetail.Rows[i].Cells["wsd_grade"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvSalesDetail.Rows[i].Cells["wsd_prd_size"].Value.ToString() + "'";
                        dtFill14 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill14.Rows.Count > 0)
                        {
                            cPcode = dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString();
                            cGrade = dgvSalesDetail.Rows[i].Cells["wsd_grade"].Value.ToString();
                            cSize = dgvSalesDetail.Rows[i].Cells["wsd_prd_size"].Value.ToString();
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
                            cPcode = dgvSalesDetail.Rows[i].Cells["wsd_prd_master_code"].Value.ToString();
                            cGrade = dgvSalesDetail.Rows[i].Cells["wsd_grade"].Value.ToString();
                            cSize = dgvSalesDetail.Rows[i].Cells["wsd_prd_size"].Value.ToString();
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
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + _xEntityID + "','" + _xBranchID + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + txtKdOutlet1.Text + "','" + lblKdOutlet2.Text + "','" + __tglorder + "'";
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
                            MessageBox.Show("Default Price Code untuk Customer " + txtKdOutlet1.Text + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }


                        //DataTable dtFill15 = new DataTable();
                        strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                        strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                        strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                        strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                        strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                        strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                        strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                        strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                        strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                        strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_reason,wsd_sap_unit_price,wsd_beban_cashdisc,wsd_sled,wsd_req_order_qty,wsd_req_order_xqty,wsd_req_order_amt,wsd_is_subs)";

                        strSQL = strSQL + " values ( '";
                        strSQL = strSQL + _xEntityID + "' , '"; //'wsd_entity_id
                        strSQL = strSQL + lblOrderType.Text + "' , '"; //'wsd_so_type
                        strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                        strSQL = strSQL + i + "', '"; //'wsd_so_detail_line_no
                        strSQL = strSQL + _xBranchID + "','"; //wsd_so_branch_id NEW
                        strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                        strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                        strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                        strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                        strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                        strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                        strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                        strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                        strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString() + "', '";
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["wsd_real_order_xqty"].Value.ToString() + "', '";
                        fConvertQty(dgvSalesDetail.Rows[i].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[i].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[i].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        strSQL = strSQL + _fConvertQty + "', '";
                        if (dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString() != "" && !string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()))
                        {
                            //Decimal realO = Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty);
                            //rounding mekanism
                            Decimal realO = Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty);
                            if (isRoundingMekanism)
                            {
                                realO = Math.Round(Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                            }
                            strSQL = strSQL + realO + "', '"; //'wsd_real_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + "0" + "', '";
                        }
                        if (bCalcDisc == false)
                        {
                            strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["wsd_so_xqty"].Value.ToString() + "', ";
                            strSQL = strSQL + "NULL, "; //'wsd_qty_sales
                            strSQL = strSQL + "NULL, "; //'wsd_total_so_amount
                        }
                        else
                        {
                            strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["wsd_so_xqty"].Value.ToString() + "', '";
                            strSQL = strSQL + _fConvertQty + "', '";
                            strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["ijumlahharga"].Value.ToString() + "', ";
                        }
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
                        strSQL = strSQL + "null, "; //'wsd_margin
                        strSQL = strSQL + "null, "; //'wsd_unit_cost
                        strSQL = strSQL + "null, "; //'wsd_total_cost
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
                        Decimal orghet = Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(dRateCurr);
                        strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                        strSQL = strSQL + "0, "; //'wsd_org_total_so_amount
                        strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                        strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                        strSQL = strSQL + cConv2 + ", "; //'wsd_uom_convert_big
                        fConvertQty(dgvSalesDetail.Rows[i].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[i].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[i].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                        strSQL = strSQL + uomorg + ", "; //'wsd_uom_org_qty
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["wsd_tot_disc_pc"].Value.ToString() + ", '";

                        lBrs = 0;
                        curDisc1 = 0;
                        curDisc2 = 0;

                        if (pCheckValueDisc(cPcode, dgvDisc.Rows.Count - 1, dgvDisc) == false)
                        {
                            decimal g_iRpD1 = 0;
                            decimal g_iRpK1 = 0;
                            decimal g_iRpD2 = 0;
                            decimal g_iRpK2 = 0;
                            if (dgvDisc.Rows[lRow].Cells["g_iRpD1"].Value.ToString() != "")
                            {
                                g_iRpD1 = Convert.ToDecimal(dgvDisc.Rows[lRow].Cells["g_iRpD1"].Value.ToString());
                            }

                            if (dgvDisc.Rows[lRow].Cells["g_iRpK1"].Value.ToString() != "")
                            {
                                g_iRpK1 = Convert.ToDecimal(dgvDisc.Rows[lRow].Cells["g_iRpK1"].Value.ToString());
                            }

                            if (dgvDisc.Rows[lRow].Cells["g_iRpD2"].Value.ToString() != "")
                            {
                                g_iRpD2 = Convert.ToDecimal(dgvDisc.Rows[lRow].Cells["g_iRpD2"].Value.ToString());
                            }

                            if (dgvDisc.Rows[lRow].Cells["g_iRpK2"].Value.ToString() != "")
                            {
                                g_iRpK2 = Convert.ToDecimal(dgvDisc.Rows[lRow].Cells["g_iRpK2"].Value.ToString());
                            }

                            curDisc1 = g_iRpD1 + g_iRpK1;
                            curDisc2 = g_iRpD2 + g_iRpK2;

                        }
                        strSQL = strSQL + curDisc1 + "', '"; //'wsd_tot_disc_1
                        strSQL = strSQL + curDisc2 + "', "; //'wsd_tot_disc_2
                        strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["tax_code"].Value.ToString() + "', '"; //wsd_tax_code
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["pct_tax"].Value.ToString() + "', '"; //wsd_tax_pct
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["rp_tax"].Value.ToString() + "', '"; //wsd_tax_amount
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["gross_net"].Value.ToString() + "', '"; //wsd_gross_net
                        strSQL = strSQL + dgvSalesDetail.Rows[i].Cells["cash_disc"].Value.ToString() + "', "; //wsd_cash_disc
                        strSQL = strSQL + "'N', '"; //wsd_line_type
                        if (dgvSalesDetail.Rows[i].Cells["tgl_price"].Value.ToString() != "" && !string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["tgl_price"].Value.ToString()))
                        {
                            string _dateP = dgvSalesDetail.Rows[i].Cells["tgl_price"].Value.ToString();

                            DateTime __dtP = DateTime.Parse(_dateP, CultureInfo.GetCultureInfo("en-US"));

                            //DateTime.TryParseExact(_dateP, "dd/MM/yyyy",
                            //    System.Globalization.CultureInfo.InvariantCulture,
                            //    System.Globalization.DateTimeStyles.NoCurrentDateDefault, out __dtP);

                            strSQL = strSQL + __dtP.ToString("yyyyMMdd") + "', ";
                        }
                        else
                        {
                            strSQL = strSQL + DateTime.Now.ToShortDateString() + "', ";
                        }
                        //********* perubahan grup harga*******
                        //sPriceCode = "dbo.F_PRICEGROUPPCODE('" + cPcode + "','" + cGrade + "','" + cSize + "','" + txtKdOutlet1.Text + "','" + lblKdOutlet2.Text + "','" + __tglorder + "')";
                        //strSQL = strSQL + sPriceCode + ", "; //wsd_price_code
                        ////***********grup harga****************
                        strSQL = strSQL + "'" + sPriceCode + "', "; //wsd_price_code

                        string rsn = "";
                        string[] sReason = new string[] { };
                        if (dgvSalesDetail.Rows[i].Cells["reason"].Value != null)
                        {
                            rsn = dgvSalesDetail.Rows[i].Cells["reason"].Value.ToString();
                            sReason = rsn.Split('~');
                        }
                        else
                        {
                            rsn = "";
                            sReason = rsn.Split('~');
                        }
                        strSQL = strSQL + "'" + sReason[0].ToString() + "', "; //wsd_reason

                        strSQL = strSQL + "'" + dgvSalesDetail.Rows[i].Cells["wsd_sap_unit_price"].Value.ToString() + "',"; //wsd_sap_unit_price                       eason
                        strSQL = strSQL + _beban + ","; //wsd_beban_cashdisc

                        if (!dgvSalesDetail.Rows[i].Cells["wsd_sled"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                        {
                            strSQL = strSQL + "'" + dgvSalesDetail.Rows[i].Cells["wsd_sled"].Value.ToString() + "', "; //wsd_sled
                        }
                        else
                        {
                            strSQL = strSQL + "NULL, "; //wsd_sled
                        }


                        fConvertQty(dgvSalesDetail.Rows[i].Cells["wsd_req_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[i].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[i].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        strSQL = strSQL + "'" + _fConvertQty + "',"; // wsd_req_order_qty
                        strSQL = strSQL + "'" + dgvSalesDetail.Rows[i].Cells["wsd_req_order_xqty"].Value.ToString() + "', "; //wsd_req_order_xqty

                        if (dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString() != "" && !string.IsNullOrEmpty(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()))
                        {
                            //Decimal realO = Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty);
                            //rounding mekanism
                            Decimal ReqO = Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty);
                            if (isRoundingMekanism)
                            {
                                ReqO = Math.Round(Convert.ToDecimal(dgvSalesDetail.Rows[i].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                            }
                            strSQL = strSQL + ReqO; //'wsd_req_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + "0"; //'wsd_req_order_amt
                        }
                        strSQL = strSQL + ",'" + Convert.ToString(dgvSalesDetail.Rows[i].Cells[COL_IS_SUBS].Value).Replace("'", "''") + "'";
                        strSQL = strSQL + ")";

                        _clsGlobal.ExecuteTrans(strSQL);
                    }
                }

                if (bCalcDisc == false)
                {
                    DataTable dtFill16 = new DataTable();
                    strSQL = " update SO_SPG_GIRL_MAN set sgm_txn_date = IM_WH_LOC.wh_loc_last_work_date FROM SO_SPG_GIRL_MAN  INNER JOIN ";
                    strSQL = strSQL + " IM_WH_LOC WITH (NOLOCK) ON SO_SPG_GIRL_MAN.sgm_wh_loc1 = IM_WH_LOC.wh_loc_id1 AND SO_SPG_GIRL_MAN.sgm_wh_loc2 = IM_WH_LOC.wh_loc_id2 and SO_SPG_GIRL_MAN.sgm_entity_id = IM_WH_LOC.wh_loc_entity  and SO_SPG_GIRL_MAN.sgm_branch_id = IM_WH_LOC.wh_branch_id ";
                    strSQL = strSQL + " where sgm_spgm_id = '" + txtSalesID.Text + "' and sgm_entity_id = '" + _xEntityID + "' and sgm_branch_id = '" + _xBranchID + "' ";
                    dtFill16 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill16.Rows.Count > 0)
                    {
                        DataTable dtFill17 = new DataTable();
                        strSQL = " select @@ROWCOUNT as JmlRow";
                        dtFill17 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill17.Rows.Count > 0)
                        {
                            if (dtFill17.Rows[0]["JmlRow"].ToString() != "1")
                            {
                                MessageBox.Show("Jumlah baris Sales tidak sama dengan satu ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }

                        }
                        else
                        {
                            MessageBox.Show("Jumlah baris sales tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }
                    }

                }



                keyFound = true;
                //}


            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        private bool fSaveProsesSO(string sCredLimit, bool bYesNo)
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
            string sSONo = "";
            string sSOStatus = "";
            long lROrdQty = 0;
            decimal curROrdAmt = 0;
            long lOrdQty = 0;
            long lBonusQty = 0;
            decimal curOrdAmt = 0;
            decimal curCost = 0;
            decimal curTotalCost = 0;
            //clsPCode
            string sCurrCode = "";
            double dRateCurr = 0;
            long lBrs = 0;
            decimal curDisc1 = 0;
            decimal curDisc2 = 0;
            decimal curTotTPR = 0;
            string sXUserID = "";
            string sXCreateDate = "";
            long lPromo = 0;
            string sErrDesc = "";
            decimal currCredLimit = 0;
            bool bFlagCr = false;
            string sPriceCode = "";
            string sOpenStatus = "";
            bool bExists = false;
            //long lOrdQty = 0;
            long lLeadTime = 0;
            //decimal curOrdAmt = 0;
            string pdaNO = "";
            string pdaFlag = "";
            string extsap = "";
            string runcext = "";
            string sReasonCode = "";
            string sReasonCodeList = "";
            double intRealOrderQty;
            double intQtySales;
            bool flAbtHalfReqDn; //Sebagian Qty terpenuhi dalam 1 prd
            string shipmentType = "reguler";
            string addShipmentCost = "0";
            string extracttranship = "N";
            string runcodetranship = "";

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

            // SLED
            string SLEDValue = "0";
            string SledFlag = "N";

            //USER PROCESS SDR34 (ROVI)
            string rovi_od_date = string.Empty;
            string user_rovi_od = string.Empty;

            //SDT25A
            string user_approval_od = string.Empty;
            // wsh_approval_rovi_by 
            // wsh_approval_rovi_date 

            //sdt19
            string usr_proc_order = string.Empty;
            string usr_proc_order_date = string.Empty;

            //History Delivery Date & Expired Date
            string DeliveryDateHis = string.Empty;
            string DeliveryDateByHis = string.Empty;
            string ExpiredDateHis = string.Empty;
            string ExpiredDateByHis = string.Empty;


            //SDR34A
            //wsh_user_rovi_cl
            //wsh_rovi_cl_date
            string user_rovi_cl = string.Empty;
            string rovi_cl_date = string.Empty;
            //SDT25B
            // wsh_approval_rovi_cl_by 
            // wsh_approval_rovi_cl_date
            // wsh_user_approval_cl
            string approval_rovi_cl_by = string.Empty;
            string approval_rovi_cl_date = string.Empty;
            string user_approval_cl = string.Empty;

            //SUPOM
            string source_site = string.Empty;
            //DEAROTC           
            string wa_staging = "N";
            string wa_customer = "N";
            string wa_salesman = "N";
            string wa_confirm_order = "Y";

            try
            {
                DataTable dtFill1 = new DataTable();
                strSQL = "select  cm_branch, cm_area, cm_wilayah, cm_rayon, cm_cust_group, isnull(cm_line_code, '') as cm_line_code, cm_lead_time  From SO_CUST_MASTER WITH (NOLOCK)  ";
                strSQL = strSQL + " where cm_cust_code1 = '" + txtKdOutlet1.Text + "' ";
                strSQL = strSQL + " and cm_cust_code2 = '" + lblKdOutlet2.Text + "'   ";
                strSQL = strSQL + " and cm_entity = '" + xEntityID + "'   ";
                strSQL = strSQL + " and cm_branch = '" + xBranchID + "'   ";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["cm_branch"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + txtKdOutlet1.Text + ") tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sBranch = dtFill1.Rows[0]["cm_branch"].ToString();
                    }
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["cm_area"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + txtKdOutlet1.Text + ") tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sArea = dtFill1.Rows[0]["cm_area"].ToString();
                    }
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["cm_wilayah"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + txtKdOutlet1.Text + ") tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sWilayah = dtFill1.Rows[0]["cm_wilayah"].ToString();
                    }
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["cm_rayon"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + txtKdOutlet1.Text + ") tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sRayon = dtFill1.Rows[0]["cm_rayon"].ToString();
                    }

                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["cm_cust_group"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + txtKdOutlet1.Text + ") tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sCustGroup = dtFill1.Rows[0]["cm_cust_group"].ToString();
                    }
                    sLineCode = dtFill1.Rows[0]["cm_line_code"].ToString();
                    if (string.IsNullOrEmpty(dtFill1.Rows[0]["cm_line_code"].ToString()))
                    {
                        MessageBox.Show("Branch untuk Customer (" + txtKdOutlet1.Text + ") tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        lLeadTime = Convert.ToInt16(dtFill1.Rows[0]["cm_line_code"].ToString());
                    }
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
                DateTime _dtgl2 = DateTime.Parse(Convert.ToDateTime(dtTglOrder.EditValue).ToString(), CultureInfo.GetCultureInfo("en-US"));
                //'Get Year, Periode, Week untuk Tgl SO
                DataTable dtFill2 = new DataTable();
                strSQL = "select fwd_year, fwd_periode, fwd_week_no  from TBL_GS_WORKING_DAY WITH (NOLOCK)  ";
                strSQL = strSQL + " where convert(varchar(10), fwd_work_day, 112) = '" + _dtgl2.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " AND fwd_entity = '" + xEntityID + "' and fwd_branch = '" + xBranchID + "' "; //entity id harus di lempar
                dtFill2 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill2.Rows.Count > 0)
                {
                    sWeekNo = dtFill2.Rows[0]["fwd_week_no"].ToString();
                    sYear = dtFill2.Rows[0]["fwd_year"].ToString();
                    sPeriode = dtFill2.Rows[0]["fwd_periode"].ToString();
                }
                else
                {
                    MessageBox.Show("Tidak ada tanggal " + Convert.ToDateTime(dtTglOrder.EditValue).ToString("dd MMM yyyy") + " pada tanggal hari kerja ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //get status so
                DataTable dtFill3 = new DataTable();
                if (xSaveOOS == false)
                {

                    strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  ";
                    strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                    strSQL = strSQL + " and gh_sequence_no = '3' ";

                    //dtFill3 = _clsGlobal.ExecDT(strSQL);
                }
                else
                {
                    if (dgvSalesDetail.Rows.Count > 0)
                    {
                        //DataTable dtFill3 = new DataTable();
                        strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  ";
                        strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                        strSQL = strSQL + " and gh_sequence_no = '3' ";
                        //dtFill3 = _clsGlobal.ExecDT(strSQL);
                    }
                    else
                    {
                        //DataTable dtFill3 = new DataTable();
                        strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK)  ";
                        strSQL = strSQL + " where gh_sys = 'H' and gh_function_name = 'SOSTATUS' ";
                        strSQL = strSQL + " and gh_sequence_no = '2' ";
                        //dtFill3 = _clsGlobal.ExecDT(strSQL);
                        sReasonCode = "SC06";
                    }
                }
                dtFill3 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill3.Rows.Count > 0)
                {
                    sSOStatus = dtFill3.Rows[0]["gh_function_code"].ToString();
                }
                else
                {
                    MessageBox.Show("Status SO (Approve) tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //Get Price Code
                DataTable dtFill4 = new DataTable();
                strSQL = " exec SP_GET_PRICE_CODE '" + xEntityID + "','" + xBranchID + "','" + txtKdOutlet1.Text + "','" + lblKdOutlet2.Text + "'";
                dtFill4 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill4.Rows.Count > 0)
                {
                    if (String.IsNullOrEmpty(dtFill4.Rows[0]["cm_def_price_code"].ToString()))
                    {
                        MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                    else
                    {
                        sPriceCode = dtFill4.Rows[0]["cm_def_price_code"].ToString();
                    }
                }
                else
                {
                    MessageBox.Show("Default Price Code untuk Customer " + txtKdOutlet1.Text + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                //'Get total in Grid
                curTotalCost = 0;
                curROrdAmt = 0;
                lROrdQty = 0;
                lOrdQty = 0;
                curOrdAmt = 0;
                for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                {
                    if (dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() != "")
                    {
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lROrdQty = lROrdQty + Convert.ToInt64(_fConvertQty);
                        curROrdAmt = curROrdAmt + (Convert.ToDecimal(_fConvertQty) * Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString()));
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        lOrdQty = lOrdQty + Convert.ToInt64(_fConvertQty);
                        curOrdAmt = curOrdAmt + Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["ijumlahharga"].Value.ToString());

                        if (fGetCost(dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString(), dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString(), dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            //MessageBox.Show("Parameter Default Price Code belum ada pd GS_GEN_HARDCODED ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        //if (isNefoKAM)
                        //{
                        //rounding mekanism
                        _curCost = _curCost * Convert.ToDecimal(_fConvertQty);
                        if (isRoundingMekanism)
                        {
                            _curCost = Math.Round(_curCost * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                        }
                        //}
                        //else
                        //{
                        //    _curCost = _curCost * Convert.ToDecimal(_fConvertQty);
                        //}
                        curTotalCost = curTotalCost + _curCost;
                    }
                }

                if (_xFlagDC == "N")
                {
                    // '===> Cek Kredit Limit dihilangkan jika flagDC = Y - Project Improvement TiraSnD
                    //'/*-----------------------------------------------------------------------------------------
                    //'Credit Limit
                    if (fGetCreditLimit(txtKdOutlet1.Text, lblKdOutlet2.Text) == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }

                    if (bFlagCr == true)
                    {
                        if (_curAvCredLimit < curROrdAmt)
                        {
                            //sCredLimit = "*";
                            if (fSaveSO(sCredLimit, false) == false)
                            {
                                keyFound = false;
                            }
                            keyFound = true;
                            //return keyFound;
                        }
                        else
                        {
                            sCredLimit = "";
                        }
                    }
                }

                if (txtNoOrder.Text == g_sDefaultSONo)
                {
                    if (fGetSONumber() == false)
                    {
                        keyFound = false;
                        return keyFound;
                    }
                    if (txtNoOrder.Text == g_sDefaultSONo)
                    {
                        txtNoOrder.Text = sSONo;
                    }
                    sXCreateDate = DateTime.Now.ToString("yyyy-MM-dd hh:m:ss");
                    sXUserID = clsLogin.USERID;
                }
                else
                {
                    //DEAR OTC Set Nomor Salesman
                    DataTable dtSalesman = new DataTable();
                    strSQL = "";
                    strSQL = $@"select ISNULL(sgm_hpno,'') as no_hp_salesman from SO_SPG_GIRL_MAN WITH (NOLOCK) 
                           WHERE sgm_entity_id='{xEntityID}' and sgm_branch_id='{xBranchID}' and sgm_spgm_id='{txtSalesID.Text.ToString().Trim()}'";
                    if (_clsGlobal.Connect.State == ConnectionState.Closed)
                        dtSalesman = _clsGlobal.ExecDT(strSQL);
                    if (dtSalesman.Rows.Count > 0)
                    {
                        NO_HP_Salesman = dtSalesman.Rows[0]["no_hp_salesman"].ToString().Trim();
                    }

                    sSONo = txtNoOrder.Text;

                    DataTable dtFill5 = new DataTable();
                    strSQL = "select convert(varchar(100), wsh_entry_date, 120) as wsh_entry_date, wsh_user_id,isnull(wsh_pda_no,'') as wsh_pda_no, ";
                    strSQL = strSQL + "isnull(wsh_pda_flaq,'') as wsh_pda_flaq, isnull(wsh_extract_sap,'') as wsh_extract_sap, isnull(wsh_runcode_ext,'') as wsh_runcode_ext,wsh_shipment_type,wsh_additional_ship_cost,wsh_extract_tranship,wsh_runcode_tranship ";
                    //SLED
                    strSQL += ",isnull(wsh_sled,0) wsh_sled ,isnull(wsh_sled_flag,'N') wsh_sled_flag";

                    //USER PROCESS (SDR34) & SDT25A
                    strSQL += " ,wsh_rovi_od_date,wsh_user_rovi_od,wsh_user_approval_od";

                    //User Proses Order (SDT19)
                    strSQL += " ,wsh_usr_proc_order,wsh_usr_proc_order_date";

                    //history delivery date dan expired date
                    strSQL += " , wsh_sch_date_update_by,wsh_sch_date_update,wsh_expired_date_by,wsh_expired_date_update";

                    // SDR34A 
                    strSQL += " ,wsh_approval_rovi_cl_by,wsh_approval_rovi_cl_date";
                    // SDT25B
                    strSQL += " , wsh_user_rovi_cl,wsh_rovi_cl_date,wsh_user_approval_cl";
                    //SUPOM
                    strSQL += " , wsh_source_site";
                    //DearOTC
                    strSQL += @",wsh_wa_staging
                                ,wsh_wa_customer
                                ,wsh_wa_salesman
                                ,wsh_wa_confirm_order";
                    strSQL = strSQL + " from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                    strSQL = strSQL + " where wsh_seq_no = '" + sSONo + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "'";
                    dtFill5 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill5.Rows.Count > 0)
                    {
                        sXCreateDate = dtFill5.Rows[0]["wsh_entry_date"].ToString();
                        sXUserID = dtFill5.Rows[0]["wsh_user_id"].ToString();
                        pdaNO = dtFill5.Rows[0]["wsh_pda_no"].ToString();
                        pdaFlag = dtFill5.Rows[0]["wsh_pda_flaq"].ToString();
                        extsap = dtFill5.Rows[0]["wsh_extract_sap"].ToString();
                        runcext = dtFill5.Rows[0]["wsh_runcode_ext"].ToString();
                        shipmentType = dtFill5.Rows[0]["wsh_shipment_type"].ToString();
                        addShipmentCost = dtFill5.Rows[0]["wsh_additional_ship_cost"].ToString();
                        extracttranship = dtFill5.Rows[0]["wsh_extract_tranship"].ToString();
                        runcodetranship = dtFill5.Rows[0]["wsh_runcode_tranship"].ToString();


                        //SLED
                        SLEDValue = dtFill5.Rows[0]["wsh_sled"].ToString();
                        SledFlag = dtFill5.Rows[0]["wsh_sled_flag"].ToString();

                        //USER PROCESS SDR34 & SDT25A
                        rovi_od_date = dtFill5.Rows[0]["wsh_rovi_od_date"] != DBNull.Value ? Convert.ToDateTime(dtFill5.Rows[0]["wsh_rovi_od_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;
                        user_rovi_od = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_user_rovi_od"].ToString()) ? null : dtFill5.Rows[0]["wsh_user_rovi_od"].ToString();
                        user_approval_od = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_user_approval_od"].ToString()) ? null : dtFill5.Rows[0]["wsh_user_approval_od"].ToString();

                        //SDT19
                        usr_proc_order = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_usr_proc_order"].ToString()) ? null : dtFill5.Rows[0]["wsh_usr_proc_order"].ToString();
                        usr_proc_order_date = dtFill5.Rows[0]["wsh_usr_proc_order_date"] != DBNull.Value ? Convert.ToDateTime(dtFill5.Rows[0]["wsh_usr_proc_order_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;

                        //History Delivery Date & Expired Date

                        DeliveryDateHis = dtFill5.Rows[0]["wsh_sch_date_update"] != DBNull.Value ? Convert.ToDateTime(dtFill5.Rows[0]["wsh_sch_date_update"]).ToString("yyyy-MM-dd HH:mm:ss") : null;
                        DeliveryDateByHis = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_sch_date_update_by"].ToString()) ? null : dtFill5.Rows[0]["wsh_sch_date_update_by"].ToString();
                        ExpiredDateHis = dtFill5.Rows[0]["wsh_expired_date_update"] != DBNull.Value ? Convert.ToDateTime(dtFill5.Rows[0]["wsh_expired_date_update"]).ToString("yyyy-MM-dd HH:mm:ss") : null;
                        ExpiredDateByHis = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_expired_date_by"].ToString()) ? null : dtFill5.Rows[0]["wsh_expired_date_by"].ToString();

                        //SDR34A                     
                        user_rovi_cl = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_user_rovi_cl"].ToString()) ? null : dtFill5.Rows[0]["wsh_user_rovi_cl"].ToString();
                        rovi_cl_date = dtFill5.Rows[0]["wsh_rovi_cl_date"] != DBNull.Value ? Convert.ToDateTime(dtFill5.Rows[0]["wsh_rovi_cl_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;

                        //SDT25B
                        approval_rovi_cl_by = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_approval_rovi_cl_by"].ToString()) ? null : dtFill5.Rows[0]["wsh_approval_rovi_cl_by"].ToString();
                        approval_rovi_cl_date = dtFill5.Rows[0]["wsh_approval_rovi_cl_date"] != DBNull.Value ? Convert.ToDateTime(dtFill5.Rows[0]["wsh_approval_rovi_cl_date"]).ToString("yyyy-MM-dd HH:mm:ss") : null;
                        user_approval_cl = string.IsNullOrEmpty(dtFill5.Rows[0]["wsh_user_approval_cl"].ToString()) ? null : dtFill5.Rows[0]["wsh_user_approval_cl"].ToString();

                        //SUPOM 
                        source_site = dtFill5.Rows[0]["wsh_source_site"].ToString();
                        //DearOTC                      
                        wa_staging = dtFill5.Rows[0]["wsh_wa_staging"].ToString();
                        wa_customer = dtFill5.Rows[0]["wsh_wa_customer"].ToString();
                        wa_salesman = dtFill5.Rows[0]["wsh_wa_salesman"].ToString();
                        wa_confirm_order = dtFill5.Rows[0]["wsh_wa_confirm_order"].ToString();
                    }
                    else
                    {
                        MessageBox.Show("Data Entry Date dan User ID tidak ada untuk No Order : " + sSONo, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                }

                if (txtNoOrder.Text == "")
                {
                    MessageBox.Show("Nomor SO tidak terdefinisi atau blank, cek cara entry datanya ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    keyFound = false;
                    return keyFound;
                }

                if (txtNoOrder.Text != g_sDefaultSONo)
                {
                    //CHECK SO STATUS
                    DataTable dtFill6 = new DataTable();
                    strSQL = " select gh_function_code";
                    strSQL = strSQL + " from GS_GEN_HARDCODED WITH (NOLOCK)  ";
                    strSQL = strSQL + " where gh_sys = 'H'";
                    strSQL = strSQL + " and gh_function_name = 'SOSTATUS'";
                    strSQL = strSQL + " and gh_sequence_no = 1";
                    dtFill6 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill6.Rows.Count > 0)
                    {
                        sOpenStatus = dtFill6.Rows[0]["gh_function_code"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("Status Open tidak ada pada GS_GEN_HARDCODED", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }

                    DataTable dtFill7 = new DataTable();
                    strSQL = "SELECT gh_function_code FROM GS_GEN_HARDCODED WITH (NOLOCK)  WHERE gh_function_name = 'FLAGDC'";

                    dtFill7 = _clsGlobal.ExecDTTrans(strSQL);
                    if (dtFill7.Rows.Count > 0)
                    {
                        if (dtFill7.Rows[0]["gh_function_code"].ToString() == "Y")
                        {
                            DataTable dtFill8 = new DataTable();
                            strSQL = "select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK)  ";
                            strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "' ";
                            strSQL = strSQL + " AND wsh_status_so = '" + sOpenStatus + "' ";
                            strSQL = strSQL + " AND wsh_entity_id = '" + xEntityID + "' ";
                            strSQL = strSQL + " AND wsh_branch_id = '" + xBranchID + "' ";
                            dtFill8 = _clsGlobal.ExecDTTrans(strSQL);
                            if (dtFill8.Rows.Count > 0)
                            {

                            }
                            else
                            {
                                MessageBox.Show("Status SO sudah bukan pada Open ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                keyFound = false;
                                return keyFound;
                            }
                        }
                    }

                }

                //DataTable dtFill9 = new DataTable();
                strSQL = "delete D FROM SO_WEB_SALES_DETAIL D WITH (NOLOCK)";
                strSQL = strSQL + " WHERE wsd_so_seq_no = '" + sSONo + "'";
                strSQL = strSQL + " AND wsd_entity_id = '" + xEntityID + "' AND wsd_so_branch_id = '" + xBranchID + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                //DataTable dtFill10 = new DataTable();
                strSQL = "delete H FROM SO_WEB_SALES_HEADER H WITH (NOLOCK) ";
                strSQL = strSQL + " WHERE wsh_seq_no = '" + sSONo + "'";
                strSQL = strSQL + " AND wsh_entity_id = '" + xEntityID + "'  AND wsh_branch_id = '" + xBranchID + "'"; // di tambahkan branch
                _clsGlobal.ExecuteTrans(strSQL);

                if (!isNefoKAM)
                {
                    //DataTable dtFill11 = new DataTable();
                    strSQL = "delete TBL_DISC_TXN ";
                    strSQL = strSQL + " where tdt_order_no = '" + sSONo + "'";
                    //strSQL = strSQL + "  AND wsh_branch_id = '" +  + "'"; // di tambahkan branch
                    _clsGlobal.ExecuteTrans(strSQL);

                    //DataTable dtFill12 = new DataTable();
                    strSQL = "delete TBL_SD_PROMO_TXN ";
                    strSQL = strSQL + " where pt_order_no = '" + sSONo + "'";
                    //strSQL = strSQL + "  AND wsh_branch_id = '" +  + "'"; // di tambahkan branch
                    _clsGlobal.ExecuteTrans(strSQL);

                    //DataTable dtFill13 = new DataTable();
                    strSQL = "delete TBL_SD_PROMO_CAUSE_TXN ";
                    strSQL = strSQL + " where pct_order_no = '" + sSONo + "'";
                    //strSQL = strSQL + "  AND wsh_branch_id = '" +  + "'"; // di tambahkan branch
                    _clsGlobal.ExecuteTrans(strSQL);

                }

                //save so header
                //DataTable dtFill14 = new DataTable();
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
                //if (_xFlagDC == "Y")
                //{
                //    if (strFlagLimit != "")
                //    {
                strSQL = strSQL + " ,wsh_flag_so_limit ";
                strSQL = strSQL + " ,wsh_approval_rovi_by ";
                strSQL = strSQL + " ,wsh_approval_rovi_date ";
                //    }
                //}
                //else
                //{
                //    strSQL = strSQL + " ,wsh_flag_so_limit ";
                //}
                strSQL = strSQL + " ,wsh_employee, wsh_totl_bonus_qty ";
                strSQL = strSQL + " ,wsh_top, wsh_top_id, wsh_po_date, wsh_cust_leadtime ";
                if (pdaFlag == "" || pdaNO == "")
                {

                }
                else
                {
                    strSQL = strSQL + " , wsh_pda_no ";
                    strSQL = strSQL + " , wsh_pda_flaq  ";
                }
                if (extsap == "N" || runcext == "")
                {

                }
                else
                {
                    strSQL = strSQL + " , wsh_extract_sap ";
                    strSQL = strSQL + " , wsh_runcode_ext  ";
                }
                strSQL = strSQL + " , wsh_shipment_type  ";
                strSQL = strSQL + " , wsh_additional_ship_cost  ";
                strSQL = strSQL + " , wsh_extract_tranship  ";
                strSQL = strSQL + " , wsh_runcode_tranship  ";

                //SLED
                strSQL += ", wsh_sled ,wsh_sled_flag";
                //SDR34 & SDT25A
                strSQL = strSQL + " ,wsh_rovi_od_date,wsh_user_rovi_od,wsh_user_approval_od";

                //SDT19
                strSQL = strSQL + " , wsh_usr_proc_order ";
                strSQL = strSQL + " , wsh_usr_proc_order_date ";

                //History Del Date & Exp Date
                strSQL = strSQL + " , wsh_sch_date_update_by ";
                strSQL = strSQL + " , wsh_sch_date_update ";
                strSQL = strSQL + " , wsh_expired_date_by ";
                strSQL = strSQL + " , wsh_expired_date_update ";
                //SDR34A
                strSQL = strSQL + " , wsh_user_rovi_cl ";
                strSQL = strSQL + " , wsh_rovi_cl_date ";

                //SDT25B
                strSQL = strSQL + " , wsh_approval_rovi_cl_by ";
                strSQL = strSQL + " , wsh_approval_rovi_cl_date ";
                strSQL = strSQL + " , wsh_user_approval_cl ";

                //SUPOM
                strSQL = strSQL + " , wsh_source_site ";
                //DearOTC
                strSQL = strSQL + @" ,wsh_cust_phone
                                     ,wsh_spgm_phone
                                     ,wsh_wa_staging
                                     ,wsh_wa_customer
                                     ,wsh_wa_salesman
                                     , wsh_wa_confirm_order ";

                //if (QasirFlag != "")
                //{
                //    strSQL = strSQL + " , wsh_qasir_flag  ";
                //}
                string __tglorder = _dtgl2.ToString("yyyyMMdd");
                int ttlline = dgvSalesDetail.Rows.Count + dgvPromosiQty.Rows.Count - 2;
                strSQL = strSQL + " ) ";
                strSQL = strSQL + " values ( '";
                strSQL = strSQL + _xEntityID + "' , '";
                strSQL = strSQL + sArea + "' , '";
                strSQL = strSQL + sWilayah + "' , '";
                strSQL = strSQL + sRayon + "' , '";
                strSQL = strSQL + _xBranchID + "' , '";
                strSQL = strSQL + sCustGroup + "' , '";
                strSQL = strSQL + lblOrderType.Text + "' , '";  //'wsh_so_type
                strSQL = strSQL + sSONo + "' , '"; //'wsh_seq_no
                strSQL = strSQL + __tglorder + "' , '"; //'wsh_so_date
                strSQL = strSQL + txtKdOutlet1.Text + "' , '"; //'wsh_cust_code1
                strSQL = strSQL + lblKdOutlet2.Text + "' , '"; //'wsh_cust_code2
                strSQL = strSQL + lblWHLoc1.Text + "' , '"; //'wsh_loc_id1
                strSQL = strSQL + lblWHLoc2.Text + "' , '"; //'wsh_loc_id2
                strSQL = strSQL + sLineCode + "' , "; //'wsh_brand_line
                strSQL = strSQL + "null , "; //'wsh_so_ref_no
                strSQL = strSQL + "null , "; //'wsh_so_txn_date_from
                strSQL = strSQL + "null , '"; //'wsh_so_txn_date_to
                strSQL = strSQL + sWeekNo + "', '"; //'wsh_week_no
                strSQL = strSQL + txtSalesID.Text + "', "; //'wsh_spgm_id
                strSQL = strSQL + "null , "; //'wsh_spgm_group
                strSQL = strSQL + "null , "; //'wsh_spgm_coordinator_id
                strSQL = strSQL + "null , "; //'wsh_spgm_loc1
                strSQL = strSQL + "null , "; //'wsh_spgm_loc2
                strSQL = strSQL + "null , "; //'wsh_shift_no
                strSQL = strSQL + "0 , "; //'[wsh_participation_disc_%]
                strSQL = strSQL + "0 , "; //'wsh_participation_disc_amount
                strSQL = strSQL + "0 , '"; //'wsh_tot_part_line_amount
                strSQL = strSQL + ttlline + "', "; //'wsh_total_line
                strSQL = strSQL + "null , '"; //'wsh_relevan_pod
                strSQL = strSQL + sSOStatus + "', '"; //'wsh_status_so
                strSQL = strSQL + lROrdQty + "', '"; //'wsh_real_order_qty
                strSQL = strSQL + curROrdAmt + "', "; //'wsh_real_order_amt
                //if (bCalcDisc == false)
                //{
                //    strSQL = strSQL + "0 , "; //'wsh_total_so_qty
                //    strSQL = strSQL + "0 , "; //'wsh_total_so_amt
                //}
                //else
                //{
                strSQL = strSQL + lOrdQty + ", "; //'wsh_total_so_qty
                strSQL = strSQL + curOrdAmt + ", "; //'wsh_total_so_amt
                //}
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
                strSQL = strSQL + "null , '"; //'wsh_disc_margin
                strSQL = strSQL + curTotalCost + "' , '"; //'wsh_total_cost
                strSQL = strSQL + sYear + "', '"; //'wsh_year
                strSQL = strSQL + sPeriode + "', "; //'wsh_period
                strSQL = strSQL + "null , "; //'wsh_cancel_date
                strSQL = strSQL + "null , "; //'wsh_cancel_by
                strSQL = strSQL + "null , '"; //'wsh_approval
                strSQL = strSQL + clsLogin.USERID + "' , "; //'wsh_approval_by
                strSQL = strSQL + "getdate() , "; //'wsh_approval_date
                strSQL = strSQL + "null , "; //'wsh_ret_seq_no
                strSQL = strSQL + "null , "; //'wsh_ret_by
                strSQL = strSQL + "null , "; //'wsh_remarks
                strSQL = strSQL + "null , '"; //'wsh_participation_code
                strSQL = strSQL + sXCreateDate + "' , '"; //'wsh_entry_date
                strSQL = strSQL + sXUserID + "', "; //'wsh_user_id
                strSQL = strSQL + "null , "; //'wsh_from_OB_flag
                strSQL = strSQL + "null , "; //'wsh_gross_net
                strSQL = strSQL + "null , "; //'wsh_gen_flag
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_pc
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_1
                strSQL = strSQL + "0 , "; //'wsh_tot_disc_2
                strSQL = strSQL + "0 , "; //'wsh_tot_tpr_amt
                strSQL = strSQL + "null , "; //'wsh_tax_code
                strSQL = strSQL + "null , "; //'wsh_tax_pct
                if (txtCashDiscP.Text == "")
                {
                    strSQL = strSQL + "0 , "; //'wsh_pct_cash_disc
                }
                else
                {
                    strSQL = strSQL + txtCashDiscP.Text + " , "; //'wsh_pct_cash_disc
                }
                //if (txtCashDiscU.Text == "")
                //{
                strSQL = strSQL + "0 , "; //'wsh_cash_disc
                //}
                //else
                //{
                //    strSQL = strSQL + txtCashDiscU.Text + " , "; //'wsh_cash_disc
                //}
                //strSQL = strSQL + "0 , "; //'wsh_cash_disc
                strSQL = strSQL + "null , "; //'wsh_tax_pbm_pct
                strSQL = strSQL + "null , '"; //'wsh_tax_pbm_amount
                //DataRow selectedDataRow = ((DataRowView)cbPembyFaktur.GetSelectedDataRow()).Row;
                string idO = Convert.ToString(cbPembyFaktur.EditValue);
                strSQL = strSQL + idO + "' , "; //'wsh_pay_type
                strSQL = strSQL + "0 , '"; //'wsh_gross_net_amt
                strSQL = strSQL + Convert.ToDateTime(dtTglValidasi.EditValue).ToString("yyyy-MM-dd") + "' , '"; //'wsh_validate
                strSQL = strSQL + txtNoPO.Text + "' , '"; //'wsh_po_no
                if (xFlagDC == "Y")
                {
                    //if (strFlagLimit != "")
                    //{
                    //if (clsGlobal.MODE_TRX == 2)
                    //{
                    //    strSQL = strSQL + "' , '"; //'wsh_flag_so_limit
                    //}
                    //else
                    //{
                    strSQL = strSQL + strFlagLimit + "' , '"; //'wsh_flag_so_limit
                    //}
                    //}
                    if (strRoviUpdUser != "")
                    {
                        strSQL = strSQL + strRoviUpdUser + "' , '";
                    }
                    else
                    {
                        strSQL = strSQL + "' , '";
                    }

                    if (strRoviUpdDate != "")
                    {
                        strSQL = strSQL + Convert.ToDateTime(strRoviUpdDate).ToString("yyyy-MM-dd") + " " + Convert.ToDateTime(strRoviUpdDate).ToString("HH:MM:SS") + "' , '";
                    }
                    else
                    {
                        strSQL = strSQL + "' , '";
                    }
                }
                else
                {
                    strSQL = strSQL + "' , '"; //'wsh_flag_so_limit
                    strSQL = strSQL + "' , '";
                    strSQL = strSQL + "' , '";
                }
                strSQL = strSQL + lblEmployee.Text + "' , "; //'wsh_employee
                strSQL = strSQL + "0 , '"; //'wsh_totl_bonus_qty
                string tops = cbTOP.Text;
                string stops = ExtractTopDaysSafe(tops);
                strSQL = strSQL + stops + "' , '"; //'wsh_top
                DataRow selectedDataRow1 = ((DataRowView)cbTOP.GetSelectedDataRow()).Row;
                string idOO = Convert.ToString(selectedDataRow1["Code"]);
                strSQL = strSQL + idOO + "' , '"; //'wsh_top_id
                strSQL = strSQL + Convert.ToDateTime(dtTglPO.EditValue).ToString("yyyy-MM-dd") + "','"; //'wsh_po_date
                strSQL = strSQL + lLeadTime + "'"; //'wsh_cust_leadtime
                if (pdaFlag == "" || pdaNO == "")
                {
                    //strSQL = strSQL + "' , '"; 
                    //strSQL = strSQL + "' , '"; 
                }
                else
                {
                    strSQL = strSQL + ",'" + pdaNO + "' ,'"; //'pdano
                    strSQL = strSQL + pdaFlag + "'"; //'pdaflag
                }
                if (extsap == "" || runcext == "")
                {
                    //strSQL = strSQL + "' , '"; //'extsap
                    //strSQL = strSQL +  "' , '"; //'runcext
                }
                else
                {
                    strSQL = strSQL + ",'" + extsap + "' , '"; //'extsap
                    strSQL = strSQL + runcext + "'"; //'runcext
                }
                strSQL = strSQL + " , '" + shipmentType + "' ";
                strSQL = strSQL + " , '" + addShipmentCost + "'  ";
                strSQL = strSQL + " , '" + extracttranship + "' ";
                strSQL = strSQL + " , '" + runcodetranship + "'  ";

                //SLED
                strSQL = strSQL + " , '" + SLEDValue + "' "; // wsh_sled
                strSQL = strSQL + " , '" + SledFlag + "' "; // wsh_sled_flag
                //JIRA TPM-5352                            

                //USER PROCESS
                if (!string.IsNullOrEmpty(rovi_od_date))
                {
                    strSQL = strSQL + " , '" + rovi_od_date + "' "; // wsh_rovi_od_date
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_rovi_od_date
                }

                strSQL = strSQL + " , '" + user_rovi_od + "' "; // wsh_user_rovi_od   
                strSQL = strSQL + " , '" + user_approval_od + "' "; // wsh_user_approval_od 
                //SDT19
                strSQL = strSQL + " , '" + usr_proc_order + "' "; // wsh_usr_proc_order
                if (!string.IsNullOrEmpty(usr_proc_order_date))
                {
                    strSQL = strSQL + " ,'" + usr_proc_order_date + "'"; // wsh_usr_proc_order_date
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_rovi_od_date
                }



                //History Del Date & Exp Date
                strSQL = strSQL + " , '" + DeliveryDateByHis + "' "; // wsh_sch_date_update_by
                if (!string.IsNullOrEmpty(DeliveryDateHis))
                {
                    strSQL = strSQL + " ,'" + DeliveryDateHis + "'"; // wsh_sch_date_update
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_sch_date_update
                }

                strSQL = strSQL + " , '" + ExpiredDateByHis + "' "; // wsh_expired_date_by
                if (!string.IsNullOrEmpty(ExpiredDateHis))
                {
                    strSQL = strSQL + " ,'" + ExpiredDateHis + "'"; // wsh_expired_date_update
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_expired_date_update
                }



                //SDR34A            
                strSQL = strSQL + " , '" + user_rovi_cl + "' "; // wsh_user_rovi_cl
                if (!string.IsNullOrEmpty(rovi_cl_date))
                {
                    strSQL = strSQL + " ,'" + rovi_cl_date + "'"; // wsh_rovi_cl_date
                }
                else
                {
                    strSQL = strSQL + " ,null"; // wsh_rovi_cl_date
                }


                //SDT25B
                strSQL = strSQL + " , '" + approval_rovi_cl_by + "' "; // wsh_approval_rovi_cl_by
                if (!string.IsNullOrEmpty(approval_rovi_cl_date))
                {
                    strSQL = strSQL + " ,'" + approval_rovi_cl_date + "'"; // wsh_approval_rovi_cl_date
                }
                else
                {
                    strSQL = strSQL + " ,NULL"; // wsh_approval_rovi_cl_date
                }

                strSQL = strSQL + " , '" + user_approval_cl + "' "; // wsh_user_approval_cl

                //if (QasirFlag != "")
                //{
                //    strSQL = strSQL + ",'" + QasirFlag + "' "; //'qasirflag
                //}

                //SUPOM
                strSQL = strSQL + " , '" + source_site + "' "; // wsh_source_site
                //DEAROTC
                strSQL = strSQL + " , '" + NO_HP_Outlet + "' "; // wsh_cust_phone
                strSQL = strSQL + " , '" + NO_HP_Salesman + "' "; // wsh_spgm_phone
                strSQL = strSQL + " , '" + wa_staging + "' "; // wsh_wa_staging
                strSQL = strSQL + " , '" + wa_customer + "' "; // wsh_wa_customer
                strSQL = strSQL + " , '" + wa_salesman + "' "; // wsh_wa_salesman
                strSQL = strSQL + " , '" + wa_confirm_order + "' "; // wsh_wa_confirm_order
                strSQL = strSQL + " ) ";
                _clsGlobal.ExecuteTrans(strSQL);
                //'END OF Saving SO Header

                DataTable dtFill15 = new DataTable();
                strSQL = "select cc_currency_code,cc_mdl_rate_to_home  from CB_CURR_CODE WITH (NOLOCK)   ";
                strSQL = strSQL + "  where cc_currency_type = 'H'";
                dtFill15 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill15.Rows.Count > 0)
                {
                    sCurrCode = dtFill15.Rows[0]["cc_currency_code"].ToString();
                    dRateCurr = Convert.ToDouble(dtFill15.Rows[0]["cc_mdl_rate_to_home"].ToString());
                }

                lCounter = 0;
                for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                {
                    if (dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() != "")
                    {
                        DataTable dtFill16 = new DataTable();
                        strSQL = "select prm_prd_line_code, prm_prd_group_code, prm_prd_sgroup_code, ";
                        strSQL = strSQL + " prm_prd_model_code , prm_raw_mat_used_code, prm_washing_collor_code, ";
                        strSQL = strSQL + " prm_conversion_purc, prm_conversion_sales  from IM_PRD_MASTER WITH (NOLOCK) ";
                        strSQL = strSQL + " where prm_prd_master_code = '" + dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_grade = '" + dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString() + "'";
                        strSQL = strSQL + "  and prm_prd_size = '" + dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString() + "'";
                        dtFill16 = _clsGlobal.ExecDTTrans(strSQL);
                        if (dtFill16.Rows.Count > 0)
                        {
                            cPcode = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString();
                            cGrade = dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString();
                            cSize = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString();
                            cLineCode = dtFill16.Rows[0]["prm_prd_line_code"].ToString();
                            cGroup = dtFill16.Rows[0]["prm_prd_group_code"].ToString();
                            cSubGroup = dtFill16.Rows[0]["prm_prd_sgroup_code"].ToString();
                            cModel = dtFill16.Rows[0]["prm_prd_model_code"].ToString();
                            cRawMatUsed = dtFill16.Rows[0]["prm_raw_mat_used_code"].ToString();
                            cWashingCollorCode = dtFill16.Rows[0]["prm_washing_collor_code"].ToString();
                            cConv1 = Convert.ToInt16(dtFill16.Rows[0]["prm_conversion_purc"].ToString());
                            cConv2 = Convert.ToInt16(dtFill16.Rows[0]["prm_conversion_sales"].ToString());
                        }
                        else
                        {
                            cPcode = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString();
                            cGrade = dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString();
                            cSize = dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString();
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
                        strSQL = " exec SP_GET_PRICE_CODE_BY_SKU '" + _xEntityID + "','" + _xBranchID + "','" + cPcode + "','" + cGrade + "','" + cSize + "','" + txtKdOutlet1.Text + "','" + lblKdOutlet2.Text + "','" + __tglorder + "'";
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
                            MessageBox.Show("Default Price Code untuk Customer " + txtKdOutlet1.Text + " tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            keyFound = false;
                            return keyFound;
                        }

                        //DataTable dtFill17 = new DataTable();
                        strSQL = "insert into SO_WEB_SALES_DETAIL ( ";
                        strSQL = strSQL + " wsd_entity_id,wsd_so_type,wsd_so_seq_no,wsd_so_detail_line_no,wsd_so_branch_id,wsd_prd_line_code,wsd_prd_group_code";
                        strSQL = strSQL + " ,wsd_prd_sgroup_code,wsd_prd_model_code,wsd_raw_mat_used_code,wsd_washing_collor_code,wsd_prd_master_code";
                        strSQL = strSQL + " ,wsd_grade,wsd_prd_size,wsd_het_unit_price,wsd_real_order_xqty,wsd_real_order_qty,wsd_real_order_amt,wsd_so_xqty";
                        strSQL = strSQL + " ,wsd_qty_sales,wsd_total_so_amount,wsd_shipped_xqty,wsd_shipped_qty,wsd_shipped_amount,wsd_pod_xqty,wsd_pod_qty";
                        strSQL = strSQL + " ,wsd_pod_amount,wsd_invoice_xqty,wsd_invoice_qty,wsd_invoice_amount,[wsd_participation_disc_%],wsd_participation_disc_amt";
                        strSQL = strSQL + " ,wsd_margin,wsd_unit_cost,wsd_total_cost,wsd_remarks,wsd_invoice_process_flag,wsd_invoice_no,wsd_invoice_qty_sales,wsd_invoice_date";
                        strSQL = strSQL + " ,wsd_invoice_process_by,wsd_entry_date,wsd_user_id,wsd_curr_code,wsd_rate,wsd_org_het_unit_price,wsd_org_total_so_amount,wsd_uom_code";
                        strSQL = strSQL + " ,wsd_uom_convert_mid,wsd_uom_convert_big,wsd_uom_org_qty,wsd_tot_disc_pc,wsd_tot_disc_1,wsd_tot_disc_2,wsd_tot_tpr_amt,wsd_tax_code";
                        strSQL = strSQL + " ,wsd_tax_pct,wsd_tax_amount,wsd_gross_net,wsd_cash_disc,wsd_line_type,wsd_price_date,wsd_price_code,wsd_reason,wsd_sap_unit_price,wsd_sled,wsd_req_order_qty,wsd_req_order_xqty,wsd_req_order_amt,wsd_is_subs)";
                        strSQL = strSQL + " values ( '";
                        strSQL = strSQL + _xEntityID + "' , '"; //'wsd_entity_id
                        strSQL = strSQL + lblOrderType.Text + "' , '"; //'wsd_so_type
                        strSQL = strSQL + sSONo + "', '";  //'wsd_so_seq_no
                        strSQL = strSQL + lCounter + "', '"; //'wsd_so_detail_line_no
                        strSQL = strSQL + _xBranchID + "','"; //wsd_so_branch_id NEW
                        strSQL = strSQL + cLineCode + "', '"; //'wsd_prd_line_code
                        strSQL = strSQL + cGroup + "', '"; //'wsd_prd_group_code
                        strSQL = strSQL + cSubGroup + "', '"; //'wsd_prd_sgroup_code
                        strSQL = strSQL + cModel + "', '"; //'wsd_prd_model_code
                        strSQL = strSQL + cRawMatUsed + "', '"; //'wsd_raw_mat_used_code
                        strSQL = strSQL + cWashingCollorCode + "', '"; //'wsd_washing_collor_code
                        strSQL = strSQL + cPcode + "', '"; //'wsd_prd_master_code
                        strSQL = strSQL + cGrade + "', '"; //'wsd_grade
                        strSQL = strSQL + cSize + "', '"; //'wsd_prd_size
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString() + "', '";
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString() + "', '";
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        strSQL = strSQL + _fConvertQty + "', '";
                        Decimal realO = Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToInt64(_fConvertQty);
                        strSQL = strSQL + realO + "', '"; //'wsd_real_order_amt
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value.ToString() + "', '";
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        strSQL = strSQL + _fConvertQty + "', '"; //'wsd_qty_sales
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["ijumlahharga"].Value.ToString() + "', "; //'wsd_total_so_amount
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
                        if (fGetCost(dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString(), dgvSalesDetail.Rows[lCounter].Cells["wsd_grade"].Value.ToString(), dgvSalesDetail.Rows[lCounter].Cells["wsd_prd_size"].Value.ToString(), lblWHLoc1.Text, lblWHLoc2.Text) == false)
                        {
                            keyFound = false;
                            return keyFound;
                        }
                        strSQL = strSQL + _curCost + "', '"; //'wsd_unit_cost
                        decimal ttlcst = _curCost + Convert.ToDecimal(_fConvertQty);
                        strSQL = strSQL + ttlcst + "', "; //'wsd_total_cost
                        strSQL = strSQL + "null, "; //'wsd_remarks
                        strSQL = strSQL + "null, "; //'wsd_invoice_process_flag
                        strSQL = strSQL + "null, "; //'wsd_invoice_no
                        strSQL = strSQL + "null, "; //'wsd_invoice_qty_sales
                        strSQL = strSQL + "null, "; //'wsd_invoice_date
                        strSQL = strSQL + "null, '"; //'wsd_invoice_process_by
                        strSQL = strSQL + sXCreateDate + "' , '"; //'wsd_entry_date
                        strSQL = strSQL + sXUserID + "', '"; //wsd_user_id
                        strSQL = strSQL + sCurrCode + "', "; //'wsd_curr_code
                        strSQL = strSQL + dRateCurr + ", "; //'wsd_rate
                        decimal orghet = Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToInt64(dRateCurr);
                        strSQL = strSQL + orghet + ", "; //'wsd_org_het_unit_price
                        decimal ttlsoamt = Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["ijumlahharga"].Value.ToString()) * Convert.ToInt64(dRateCurr);
                        strSQL = strSQL + ttlsoamt + ", "; //'wsd_org_total_so_amount
                        strSQL = strSQL + "'PCS', "; //'wsd_uom_code
                        strSQL = strSQL + cConv1 + ", "; //'wsd_uom_convert_mid
                        strSQL = strSQL + cConv2 + ", '"; //'wsd_uom_convert_big
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        double uomorg = Convert.ToDouble(_fConvertQty) * dRateCurr;
                        strSQL = strSQL + uomorg + "', "; //'wsd_uom_org_qty
                        strSQL = strSQL + "0, "; //'ormatNumberDB(vsList(0).ValueMatrix(lCounter, g_iDiscount)) & ", " 'wsd_tot_disc_pc
                        strSQL = strSQL + curDisc1 + "', '"; //'wsd_tot_disc_1
                        strSQL = strSQL + curDisc2 + "', "; //'wsd_tot_disc_2
                        strSQL = strSQL + "0, '"; //'wsd_tot_tpr_amt
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["tax_code"].Value.ToString() + "', '"; //wsd_tax_code
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["pct_tax"].Value.ToString() + "', "; //wsd_tax_pct
                        strSQL = strSQL + "0, "; //'wsd_tax_amount
                        strSQL = strSQL + "0, "; //'wsd_gross_net
                        strSQL = strSQL + "0, "; //'wsd_cash_disc
                        strSQL = strSQL + "'N', '"; //wsd_line_type
                        strSQL = strSQL + Convert.ToDateTime(dgvSalesDetail.Rows[lCounter].Cells["tgl_price"].Value.ToString()).ToString("yyyy-MM-dd") + "', '";
                        strSQL = strSQL + sPriceCode + "', '"; //wsd_price_code

                        string rsn = "";
                        string[] sQuantity = new string[] { };
                        if (dgvSalesDetail.Rows[lCounter].Cells["reason"].Value != null)
                        {
                            rsn = dgvSalesDetail.Rows[lCounter].Cells["reason"].Value.ToString();
                            sQuantity = rsn.Split('~');
                        }
                        else
                        {
                            rsn = "";
                            sQuantity = rsn.Split('~');
                        }
                        strSQL = strSQL + sQuantity[0].ToString() + "',' "; //wsd_reason
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["wsd_sap_unit_price"].Value.ToString() + "', "; //wsd_sap_unit_price
                        if (!dgvSalesDetail.Rows[lCounter].Cells["wsd_sled"].Value.ToString().IsNullOrEmptyOrWhiteSpace())
                        {
                            strSQL = strSQL + "'" + dgvSalesDetail.Rows[lCounter].Cells["wsd_sled"].Value.ToString() + "',' "; //wsd_sled
                        }
                        else
                        {
                            strSQL = strSQL + "NULL,' "; //wsd_sled
                        }


                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_req_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        strSQL = strSQL + _fConvertQty + "', '"; // wsd_req_order_qty
                        strSQL = strSQL + dgvSalesDetail.Rows[lCounter].Cells["wsd_req_order_xqty"].Value.ToString() + "',"; //wsd_req_order_xqty

                        if (dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString() != "" && !string.IsNullOrEmpty(dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString()))
                        {

                            Decimal ReqO = Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty);
                            if (isRoundingMekanism)
                            {
                                ReqO = Math.Round(Convert.ToDecimal(dgvSalesDetail.Rows[lCounter].Cells["wsd_het_unit_price"].Value.ToString()) * Convert.ToDecimal(_fConvertQty), 0, MidpointRounding.AwayFromZero);
                            }
                            strSQL = strSQL + ReqO; //'wsd_req_order_amt
                        }
                        else
                        {
                            strSQL = strSQL + "0"; //'wsd_req_order_amt
                        }
                        strSQL = strSQL + ",'" + Convert.ToString(dgvSalesDetail.Rows[lCounter].Cells[COL_IS_SUBS].Value).Replace("'", "''") + "'";
                        strSQL = strSQL + ")";
                        _clsGlobal.ExecuteTrans(strSQL);
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        intRealOrderQty = Convert.ToDouble(_fConvertQty);
                        fConvertQty(dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value.ToString(), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_big"].Value.ToString()), Convert.ToInt16(dgvSalesDetail.Rows[lCounter].Cells["wsd_uom_convert_mid"].Value.ToString()));
                        intQtySales = Convert.ToDouble(_fConvertQty);
                    }
                }

                if (bolOOS == false)
                {
                    //DataTable dtFill18 = new DataTable();
                    strSQL = "update A Set A.wsh_status_so = B.gh_function_code from SO_WEB_SALES_HEADER A ";
                    strSQL = strSQL + "  left join (select gh_function_code, gh_function_desc from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name = 'SOSTATUS' and gh_sequence_no = 3) B on 1=1 ";
                    strSQL = strSQL + " where wsh_seq_no = '" + sSONo + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "' ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }
                else
                {
                    MessageBox.Show("Out Of Stock and SO Document " + sSONo, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //btnProses.Enabled = false;

                    //DataTable dtFill19 = new DataTable();
                    strSQL = "update A Set A.wsh_status_so = B.gh_function_code from SO_WEB_SALES_HEADER A ";
                    strSQL = strSQL + " left join (select gh_function_code, gh_function_desc from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name = 'SOSTATUS' and gh_sequence_no = 7) B on 1=1 ";
                    strSQL = strSQL + " where wsh_seq_no = '" + sSONo + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "' ";
                    _clsGlobal.ExecuteTrans(strSQL);
                }

                //'END OF Saving SO Detail

                DataTable dtFill20 = new DataTable();
                if (isNefoKAM)
                {
                    strSQL = "EXEC SP_EXEC_GET_TPR_KAM '" + sSONo + "','" + xEntityID + "','" + xBranchID + "' ";
                }
                else
                {
                    strSQL = "EXEC SP_EXEC_GET_TPR '" + sSONo + "' ";
                }

                dtFill20 = _clsGlobal.ExecDTTrans(strSQL);
                bYesNo = true;
                if (dtFill20.Rows.Count > 0)
                {
                    if (dtFill20.Rows[0]["tpr_budget_flag"].ToString() == "1" && dtFill20.Rows[0]["result_type"].ToString() == "2")
                    {
                        bYesNo = false;
                    }
                }

                if (bExists == false)
                {
                    if (bYesNo == false)
                    {
                        //frmEditor2.pFillData rs
                    }
                }

                //PROSES SO YANG SUDAH DI SAVE

                //DataTable dtFill21 = new DataTable();
                if (rbCashDiscP.Checked == true)
                {
                    strSQL = " exec SP_PROSES_SO_MBRANCH " + "'1','" + sSONo + "','" + clsLogin.USERID + "','N','" + xEntityID + "','" + xBranchID + "'";
                }
                else //untuk disc yang di pilih cash disc uang
                {
                    strSQL = " exec SP_PROSES_SO_MBRANCH " + "'1','" + sSONo + "','" + clsLogin.USERID + "','Y','" + xEntityID + "','" + xBranchID + "'";
                }
                _clsGlobal.ExecuteTrans(strSQL);

                //'UPDATE STOCK BALANCE
                //DataTable dtFill22 = new DataTable();
                strSQL = " exec SP_SO_UPDATE_BALANCE " + "'2','" + sSONo + "','" + clsLogin.USERID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //'FEFO PROCESS
                //DataTable dtFill23 = new DataTable();
                strSQL = " exec SP_PROCESS_FEFO '" + sSONo + "','" + clsLogin.USERID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //'CHECKING STOCK BALANCE
                //DataTable dtFill24 = new DataTable();
                strSQL = " exec SP_CHECK_BALANCE_SO_MULTIBRANCH '" + "1" + "','" + sSONo + "','" + xEntityID + "','" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //'CHECKING QTY DI DETAIL DENGAN DI BATCH
                //DataTable dtFill25 = new DataTable();
                strSQL = " exec SP_CHECK_QTY_SO_MULTIBRANCH '" + "1" + "','" + sSONo + "','" + xEntityID + "','" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //'TAMBAHAN PENGECEKAN BUDGET TPR
                //DataTable dtFill26 = new DataTable();
                strSQL = " exec SP_TPR_BUDGET_EXTRA '" + sSONo + "','" + clsLogin.USERID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                //Begin Update Credit Limit
                //DataTable dtFill27 = new DataTable();
                strSQL = " exec SP_UPDATE_CUST_CREDIT_LIMIT '" + "'2','" + sSONo + "','" + clsLogin.USERID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                txtNoOrder.Text = sSONo;
                if (bolOOS == false)
                {
                    if (g_bReqSummary == false) //&& btnOk.Enabled = true)
                    {
                        if (fProsesOrderType(sCredLimit, "2") == false)
                        {
                            keyFound = false;
                            return keyFound;
                        }
                    }

                }

                //'Insert ke Table History
                //DataTable dtFill28 = new DataTable();
                strSQL = " EXEC SP_INSERT_HISTORY_MULTIBRANCH '" + "1" + "','" + sSONo + "','" + clsLogin.USERID + "','" + xEntityID + "','" + xBranchID + "'";
                _clsGlobal.ExecuteTrans(strSQL);

                keyFound = true;

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
                strSQL = "select cm_status_cr, isnull(cm_ship_credit_limit, 0) -  isnull(cm_ar_saldo, 0) - isnull(cm_so_comitted, 0) as cred_limit  from SO_CUST_MASTER WITH (NOLOCK)  ";
                strSQL = strSQL + "  where cm_cust_code1 = '" + sCust1 + "' ";
                strSQL = strSQL + " and cm_cust_code2 = '" + sCust2 + "' AND cm_entity = '" + xEntityID + "' AND cm_branch = '" + xBranchID + "' ";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
                if (dtFill1.Rows.Count > 0)
                {
                    _curAvCredLimit = Convert.ToDecimal(dtFill1.Rows[0]["cred_limit"].ToString());

                    if (dtFill1.Rows[0]["cm_status_cr"].ToString() == "N")
                    {
                        bFlagCr = false;
                        keyFound = true;
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

        public string getSoNumSeqUniqueID(string prefix)
        {
            try
            {
                int maxSize = 9;
                char[] chars = new char[62];
                chars = "1234567890".ToCharArray();
                byte[] data = new byte[1];
                using (RNGCryptoServiceProvider crypto = new RNGCryptoServiceProvider())
                {
                    crypto.GetNonZeroBytes(data);
                    data = new byte[maxSize];
                    crypto.GetNonZeroBytes(data);
                }
                StringBuilder result = new StringBuilder(maxSize);
                foreach (byte b in data)
                {
                    result.Append(chars[b % (chars.Length)]);
                }
                return prefix + result.ToString().ToUpper();
            }
            catch (Exception)
            {

                throw;
            }
        }

        //1 Juni 2018
        private bool fGetSONumber()
        {
            bool keyFound = true;
            //bool isRandom = true;

            string sBranchID = "";
            string sErrDesc = "";

            try
            {
                if (!isMultibranch)
                {
                    //DataTable dtFill1 = new DataTable();
                    strSQL = "if not exists( select * FROM GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name = 'SOSEQNUMBER' AND gh_function_name = gh_function_code )";
                    strSQL = strSQL + " Begin";
                    strSQL = strSQL + " insert into GS_GEN_HARDCODED (gh_sys, gh_function_name, gh_sequence_no, gh_function_code, gh_function_desc, ";
                    strSQL = strSQL + " gh_min_seq_no, gh_max_seq_no, gh_last_seq_no ) values ('S', 'SOSEQNUMBER', 1, 'SOSEQNUMBER', 'SOSEQNUMBER', 1, 9999999, 0 ) END";
                    _clsGlobal.ExecuteTrans(strSQL);

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

                    sBranchID = xBranchID;

                    bool checkseq = false;
                    do
                    {
                        //DataTable dtFill4 = new DataTable();
                        strSQL = "Update GS_GEN_HARDCODED ";
                        strSQL = strSQL + " set gh_last_seq_no = gh_last_seq_no + 1";
                        strSQL = strSQL + " where gh_sys = 'S'";
                        strSQL = strSQL + " and gh_function_name = 'SOSEQNUMBER' AND gh_function_name = gh_function_code ";
                        _clsGlobal.ExecuteTrans(strSQL);

                        DataTable dtFillCheck = new DataTable();
                        strSQL = "Select gh_last_seq_no from GS_GEN_HARDCODED WITH (NOLOCK) ";
                        strSQL = strSQL + " where gh_sys = 'S' ";
                        strSQL = strSQL + " and gh_function_name = 'SOSEQNUMBER' AND gh_function_name = gh_function_code";
                        dtFillCheck = _clsGlobal.ExecDTTrans(strSQL);

                        if (dtFillCheck.Rows.Count > 0)
                        {
                            long i = Convert.ToInt64(dtFillCheck.Rows[0]["gh_last_seq_no"].ToString());
                            string cNol = "0000000" + i.ToString();
                            //string cNol = getSoNumSeqUniqueID();
                            sSONo = "S" + sBranchID + cNol.Substring(cNol.Length - 7);

                            DataTable dtSOCheck = new DataTable();
                            dtSOCheck = _clsGlobal.ExecDTTrans("Select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK) where wsh_seq_no = '" + sSONo + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "' ");
                            if (dtSOCheck.Rows.Count == 0)
                            {
                                checkseq = true;
                            }
                        }
                        else
                        {
                            checkseq = true;
                        }

                    } while (checkseq == false);



                    DataTable dtFill5 = new DataTable();
                    strSQL = "Select gh_last_seq_no from GS_GEN_HARDCODED WITH (NOLOCK) ";
                    strSQL = strSQL + " where gh_sys = 'S' ";
                    strSQL = strSQL + " and gh_function_name = 'SOSEQNUMBER' AND gh_function_name = gh_function_code";
                    dtFill5 = _clsGlobal.ExecDTTrans(strSQL);

                    if (dtFill5.Rows.Count > 0)
                    {
                        long i = Convert.ToInt64(dtFill5.Rows[0]["gh_last_seq_no"].ToString());
                        string cNol = "0000000" + i.ToString();
                        //string cNol = getSoNumSeqUniqueID();
                        sSONo = "S" + sBranchID + cNol.Substring(cNol.Length - 7);
                        //sSONo = "S" + sBranchID + cNol;
                        //bool check = false;
                        //do
                        //{
                        //    sSONo = getSoNumSeqUniqueID("S");

                        //    string sqlscript = "Select * from SO_WEB_SALES_HEADER where wsh_seq_no = '" + sSONo + "'";
                        //    DataTable dt = _clsGlobal.ExecDTTrans(sqlscript);
                        //    if (dt.Rows.Count == 0)
                        //    {
                        //        check = true;
                        //    }


                        //} while (check == false);

                    }
                    else
                    {
                        MessageBox.Show("Sequence No Order tidak ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        keyFound = false;
                        return keyFound;
                    }
                }
                //Multibranch
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
                    strSQL += " ,CASE WHEN C.cnt2 > 0  THEN '" + xEntityID + "' ELSE '' END [gh_group_menu_id] ";
                    strSQL += " ,CASE WHEN C.cnt2 > 0  THEN '" + xBranchID + "' ELSE '' END [gh_parent_menu] ";
                    strSQL += "  FROM (SELECT COUNT(gh_function_name) as [cnt2] FROM GS_GEN_HARDCODED WITH (NOLOCK) WHERE gh_group_menu_id = '" + xEntityID + "' AND gh_parent_menu = '" + xBranchID + "' and  gh_function_name = 'SOSEQNUMBER' ) C ";
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
                        dtSOCheck = _clsGlobal.ExecDTTrans("Select wsh_seq_no from SO_WEB_SALES_HEADER WITH (NOLOCK) where wsh_seq_no = '" + sSONo + "' and wsh_entity_id = '" + xEntityID + "' and wsh_branch_id = '" + xBranchID + "' ");
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
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return keyFound;
        }

        //1 Juni 2018
        private bool fGetCost(string sPcode, string sGrade, string sSize, string sWHLoc1, string sWHLoc2)
        {
            bool keyFound = false;
            try
            {
                DataTable dtFill1 = new DataTable();
                strSQL = "select isnull(isnull(wlb_last_avg_ucost, cs_standard_unit_cost), 0) as wlb_last_avg_ucost from IM_STOCK_BALANCE LEFT JOIN IM_COST_STD WITH (NOLOCK)   ON wlb_prd_master_code = cs_prd_master_code ";
                strSQL = strSQL + " AND wlb_grade = cs_grade AND wlb_prd_size = cs_size";
                strSQL = strSQL + " where wlb_loc_Id1 =  '" + sWHLoc1 + "' ";
                strSQL = strSQL + " and wlb_loc_Id2 =  '" + sWHLoc2 + "' ";
                strSQL = strSQL + " and wlb_prd_master_code =  '" + sPcode + "' ";
                strSQL = strSQL + " and wlb_grade =  '" + sGrade + "' ";
                strSQL = strSQL + " and wlb_prd_size =  '" + sSize + "' ";
                strSQL = strSQL + " and wlb_entity_id =  '" + xEntityID + "' ";
                strSQL = strSQL + " and wlb_branch_id =  '" + xBranchID + "' ";
                dtFill1 = _clsGlobal.ExecDTTrans(strSQL);
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

        //1 Juni 2018
        private void pClearTax()
        {
            if (dgvSalesDetail.Rows.Count > 0)
            {
                for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                {
                    dgvSalesDetail.Rows[lCounter].Cells["gross_net"].Value = "0";
                    dgvSalesDetail.Rows[lCounter].Cells["discpc"].Value = "";
                    dgvSalesDetail.Rows[lCounter].Cells["pctdiscpc"].Value = "0";
                    dgvSalesDetail.Rows[lCounter].Cells["cash_disc"].Value = "0";
                }
            }

        }

        //1 Juni 2018
        private void pClearDiscount()
        {
            if (dgvSalesDetail.Rows.Count > 0)
            {
                for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                {
                    dgvSalesDetail.Rows[lCounter].Cells["wsd_tot_disc_pc"].Value = "0";
                    dgvSalesDetail.Rows[lCounter].Cells["rp_tax"].Value = "0";
                }
            }
        }

        //1 Juni 2018
        private void pClearTPR()
        {
            if (dgvPromosiQty.RowCount > 0)
                if (dgvPromosiQty.Rows[0].Cells["wsd_prd_master_code"].Value != null)
                    dtGridTPRB.Rows.Clear();

            if (dgvPromosiRp.RowCount > 0)
                if (dgvPromosiRp.Rows[0].Cells["pt_product_id"].Value != null)
                    dtGridTPRU.Rows.Clear();

            //if (dgvTax.Rows[0].Cells["wsd_prd_master_code"].Value != null)
            //{
            //    dgvTax.Rows.Clear();
            //}
            //if (dgvPromoTax.Rows[0].Cells["wsd_prd_master_code"].Value != null)
            //{
            //    dgvPromoTax.Rows.Clear();
            //}
            // throw new Exception("Test pClearTPR");
        }

        //1 Juni 2018
        private void pEmptyTPRFlag()
        {
            if (dgvSalesDetail.Rows.Count > 0)
            {
                for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
                {
                    dgvSalesDetail.Rows[lCounter].Cells["tpr_flag"].Value = "0";
                }

            }
        }

        private bool pCheckValue(string sValue, int iIndex, DataGridView dgv)
        {
            bool keyFound = true;

            for (lCounter = 0; lCounter < dgv.Rows.Count; lCounter++)
            {
                if (dgv == dgvSalesDetail)
                {
                    if (lCounter != iIndex)
                    {
                        if (dgv.Rows[lCounter].Cells["wsd_prd_master_code"].Value.ToString() == sValue)
                        {
                            keyFound = false;
                            lRow = lCounter;
                            break;
                        }
                    }
                }
                else
                {
                    if (lCounter != iIndex)
                    {
                        if (dgv.Rows[lCounter].Cells["pt_product_id"].Value.ToString() == sValue)
                        {
                            keyFound = false;
                            lRow = lCounter;
                            break;
                        }
                    }
                }
            }

            return keyFound;
        }

        private bool pCheckValueSDUangKAM(string sValue, string sValue2, int iIndex, DataGridView dgv)
        {
            bool keyFound = true;

            for (lCounter = 0; lCounter < dgv.Rows.Count; lCounter++)
            {
                if (lCounter != iIndex)
                {
                    if (dgv.Rows[lCounter].Cells["pt_product_id"].Value.ToString() == sValue && dgv.Rows[lCounter].Cells["pt_promo_id"].Value.ToString() == sValue2)
                    {
                        keyFound = false;
                        lRow = lCounter;
                        break;
                    }
                }
            }

            return keyFound;
        }

        private bool pCheckValueDisc(string sValue, int iIndex, DataGridView dgv)
        {
            bool keyFound = true;

            for (lCounter = 0; lCounter < dgv.Rows.Count; lCounter++)
            {
                if (dgv.Rows[lCounter].Cells["tdt_product_master_id"].Value.ToString() == sValue)
                {
                    keyFound = false;
                    lRow = lCounter;
                    break;
                }

                //if (dgv == dgvDisc)
                //{
                //    if (lCounter != dgv.Rows.Count)
                //    {
                //        if (dgv.Rows[lCounter].Cells["tdt_product_master_id"].Value.ToString() == sValue)
                //        {
                //            keyFound = false;
                //            lRow = lCounter;
                //            break;
                //        }
                //    }
                //}
                //else
                //{
                //    if (dgv.Rows[lCounter].Cells["tdt_product_master_id"].Value.ToString() == sValue)
                //    {
                //        keyFound = false;
                //        lRow = lCounter;
                //        break;
                //    }
                //}
            }

            return keyFound;
        }

        private void pRefresh(bool bChangeSONo, bool isCalcDisc)
        {
            //dgvSalesDetail.DataSource = null;
            //dgvPromosiQty.DataSource = null;
            //dgvPromosiRp.DataSource = null;
            //dgvDisc.DataSource = null;
            //dgvPromoTax.DataSource = null;
            //dgvTax.DataSource = null;
            dtGridSODetail.Rows.Clear();
            dtGridTPRB.Rows.Clear();
            dtGridTPRU.Rows.Clear();
            dtGridDisc.Rows.Clear();
            dtGridPromoTax.Rows.Clear();
            dtGridTax.Rows.Clear();
            //dgvSalesDetail.Rows.Clear();
            //dgvPromosiQty.Rows.Clear();
            //dgvPromosiRp.Rows.Clear();
            //dgvDisc.Rows.Clear();
            //dgvPromoTax.Rows.Clear();
            //dgvTax.Rows.Clear();

            txtKdOutlet1.Text = "";
            lblOutletDesc.Text = "";
            lblKdOutlet2.Text = "";
            txtNoPajak.Text = "";
            txtStatusOutlet.Text = "";
            txtSubTotal.Text = "";
            txtPromosiU.Text = "";
            txtDisc1.Text = "";
            txtCashDiscP.Text = "";
            txtCashDiscU.Text = "";
            txtPPN.Text = "";
            txtNoPO.Text = "";
            txtTtlInvoice.Text = "";
            txtDPP.Text = "";
            //** Update Ticket 23708 **//
            txtRefNoRetur.Text = "";
            txtNote.Text = "";



            //********  END ***********//

            if (bChangeSONo == true)
            {
                txtNoOrder.Text = g_sDefaultSONo;
            }

            DataTable dtFill = new DataTable();

            if (!isCalcDisc)
            {
                strSQL = "select gh_function_name from GS_GEN_HARDCODED WITH (NOLOCK)  where gh_function_name='TOPBYTEAMSALESMAN' and gh_sys='H' and gh_function_code='Y' ";
                if (_clsGlobal.Connect.State == ConnectionState.Closed)
                    dtFill = _clsGlobal.ExecDT(strSQL);
                else
                    dtFill = _clsGlobal.ExecDTTrans(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    cbTOP.EditValue = LookupValueAt(cbTOP, 0);
                    cbPembyFaktur.EditValue = LookupValueAt(cbPembyFaktur, 0);
                }
            }

        }

        private void pRefreshDetail()
        {
            //dgvSalesDetail.Rows.Clear();
            //dgvPromosiQty.Rows.Clear();
            //dgvPromosiRp.Rows.Clear();
            //dgvDisc.Rows.Clear();
            dtGridSODetail.Rows.Clear();
            dtGridTPRB.Rows.Clear();
            dtGridTPRU.Rows.Clear();
            dtGridDisc.Rows.Clear();

            txtSubTotal.Text = "";
            txtPromosiU.Text = "";
            txtDisc1.Text = "";
            txtCashDiscP.Text = "";
            txtCashDiscU.Text = "";
            txtPPN.Text = "";
            txtNoPO.Text = "";
            txtTtlInvoice.Text = "";
            txtDPP.Text = "";
            txtNote.Text = "";
            txtRefNoRetur.Text = "";

        }

        private void pRollbackProcess()
        {
            for (lCounter = 0; lCounter < dgvSalesDetail.Rows.Count; lCounter++)
            {
                dgvSalesDetail.Rows[lCounter].Cells["wsd_so_xqty"].Value = dgvSalesDetail.Rows[lCounter].Cells["wsd_real_order_xqty"].Value;
            }
        }

        #endregion

        private void txtKdOutlet1_Enter(object sender, EventArgs e)
        {
            g_sOutlet = txtKdOutlet1.Text;
        }

        private void txtKdOutlet1_TextChanged(object sender, EventArgs e)
        {
            DataTable dtFill = new DataTable();
            strSQL = " select cm_cust_code1 as custID1, cm_cust_code2 as CustID2, cm_cust_name as CustName from SO_CUST_MASTER WITH(NOLOCK) WHERE cm_active_flag <> 'D' AND cm_cust_code1 ='" + txtKdOutlet1.Text.Trim() + "' AND cm_entity = '" + _xEntityID + "' AND cm_branch = '" + _xBranchID + "' ";
            if (_clsGlobal.Connect.State == ConnectionState.Closed)
            {
                dtFill = _clsGlobal.ExecDT(strSQL);
            }
            else
            {
                dtFill = _clsGlobal.ExecDTTrans(strSQL);
            }
            if (dtFill.Rows.Count > 0)
            {
                //txtOutletCode.Text = dtFill.Rows[0]["BranchDesc"].ToString().Trim();
                lblKdOutlet2.Text = dtFill.Rows[0]["CustID2"].ToString().Trim();
                lblOutletDesc.Text = dtFill.Rows[0]["CustName"].ToString().Trim();
            }
            else
            {
                lblKdOutlet2.Text = "";
                lblOutletDesc.Text = "";
            }
            g_bActiveEdit = true;
            UpdateHeaderPartyFields();
        }

        private void dgvSalesDetail_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvSalesDetail.CurrentCell == null)
            {
                return;
            }
            if (dgvSalesDetail.CurrentCell.IsInEditMode)
            {
                if (dgvSalesDetail.IsCurrentCellDirty)
                {
                    dgvSalesDetail.EndEdit();
                }
            }
        }
        protected void SetFocusCell(int row)
        {
            dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[row].Cells["wsd_real_order_xqty"];
        }

        private void dgvSalesDetail_SelectionChanged(object sender, EventArgs e)
        {
            if (isPcodeGridExists)
            {
                dgvSalesDetail.CurrentCell = dgvSalesDetail.Rows[editedgvRowid].Cells["wsd_real_order_xqty"];
                isPcodeGridExists = false;
                editedgvRowid = 0;
            }
        }

        private void txtCashDiscP_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar < 48 || e.KeyChar > 57 || e.KeyChar.ToString() == "." || e.KeyChar == (char)8) && e.KeyChar != 8 && e.KeyChar != 46))
            {
                e.Handled = true;
                return;
            }

        }

        private void txtCashDiscU_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar < 48 || e.KeyChar > 57 || e.KeyChar.ToString() == "." || e.KeyChar == (char)8) && e.KeyChar != 8 && e.KeyChar != 46))
            {
                e.Handled = true;
                return;
            }
        }

        private void dgvSalesDetail_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (IsGridColumn(dgvSalesDetail, COL_REAL_ORDER_QTY))
            {
                TextBox txtQty = e.Control as TextBox;
                if (txtQty != null)
                {
                    txtQty.KeyPress += new KeyPressEventHandler(txtQty_KeyPress);
                }
            }

            if (IsGridColumn(dgvSalesDetail, COL_QTY_REQ))
            {
                TextBox txtQtyReq = e.Control as TextBox;
                if (txtQtyReq != null)
                {
                    txtQtyReq.KeyPress += new KeyPressEventHandler(txtQtyReq_KeyPress);
                }
            }
        }
        private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar < 48 || e.KeyChar > 57 || e.KeyChar.ToString() == "." || e.KeyChar == (char)8) && e.KeyChar != 8 && e.KeyChar != 46))
            {
                e.Handled = true;
                return;
            }
        }

        private void txtQtyReq_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar < 48 || e.KeyChar > 57 || e.KeyChar.ToString() == "." || e.KeyChar == (char)8) && e.KeyChar != 8 && e.KeyChar != 46))
            {
                e.Handled = true;
                return;
            }
        }

        private void dtTglOrder_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtKdOutlet1.Text))
            {
                initDateDelvExp(Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd"), txtKdOutlet1.Text, lblKdOutlet2.Text, 1);
            }

        }

        private void dtTglOrder_ValueChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtKdOutlet1.Text))
            {
                initDateDelvExp(Convert.ToDateTime(dtTglOrder.EditValue).ToString("yyyyMMdd"), txtKdOutlet1.Text, lblKdOutlet2.Text, 1);
            }


        }

        private void InitialTOPDivision()
        {
            strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name like 'TOPBYDIVISION' and gh_sys = 'H'";
            DataTable dtCheck = _clsGlobal.ExecDT(strSQL);
            if (dtCheck.Rows.Count > 0)
            {
                isTOPDivision = dtCheck.Rows[0]["gh_function_code"].ToString();
            }

        }

        private void dgvSalesDetail_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dgvSalesDetail.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void cbResetReq_CheckedChanged(object sender, EventArgs e)
        {
            if (cbResetReq.Checked)
            {
                cbResetReq.Text = "Ya";
            }
            else
            {
                cbResetReq.Text = "Tidak";
            }
        }
        /** deadlock SeqCalc **/
        private string createSeqCalc(string Branch, string end)
        {
            //int lengthangka = dtFillCheck.Rows[0]["nextid"].ToString().Trim().Length;
            int lengthseq = 10;
            string prefix = Branch;

            int prefixlength = prefix.Length;
            int lengthjumlahnol = lengthseq - (prefixlength + end.Length);
            string cnNol = "";
            string jam = "";
            string menit = "";
            string detik = "";

            jam = DateTime.Now.ToString("HH");
            menit = DateTime.Now.ToString("mm");
            detik = DateTime.Now.ToString("ss");

            cnNol = jam + menit + detik;


            for (int j = cnNol.Length; j < lengthjumlahnol; j++)
            {

                cnNol = cnNol + "X";
            }

            return prefix + cnNol + end;


        }

        /** deadlock SeqCalc **/
        private bool IsNoLockMultibranch()
        {

            DataTable dtGS = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'ISNOLOCKCALCSEQ' and gh_sys = 'H'");
            if (dtGS.Rows.Count > 0)
            {
                if (dtGS.Rows[0][0].ToString().Trim() == "Y")
                {
                    return true;
                }
            }


            return false;
        }

        //private void BindShippingPlant()
        //{
        //    strSQL = "";
        //    //strSQL = "SELECT '0' as id, 'ALL' as descc ";
        //    //strSQL = strSQL + " union ALL ";
        //    // strSQL = strSQL + "Select distinct br_branch_id as id, br_branch_id + ' - ' + br_branch_desc as descc from GS_BRANCH inner join GS_ENTITY on br_gl_entity_initial = ge_entity_id INNER JOIN GS_USERS_SECURITY u on u.gu_entity = ge_entity_id where gu_user_id = '" + clsLogin.USERID + "' and ge_entity_id = '" + cbEntityID.SelectedValue + "' ";
        //    strSQL = strSQL + "SELECT DISTINCT pms_ss_site FROM GS_PRM_MULTI_SOURCE WHERE pms_branch_id = '" + xBranchID + "' AND pms_flag_sloc = 'DED'";

        //    cbShippingPlant.DataSource = _clsGlobal.ExecDT(strSQL);
        //    cbShippingPlant.ValueMember = "pms_ss_site";
        //    cbShippingPlant.DisplayMember = "pms_ss_site";
        //    //cbSOType.SelectedIndex = 0;
        //}

        private void cbFlagSloc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isCrossSite && !String.IsNullOrEmpty(txtKdOutlet1.Text))
            {
                if (cbFlagSloc.EditValue != null && Convert.ToString(cbFlagSloc.EditValue) == "DDE")
                {
                    btnPopUpShippingPlant.Enabled = true;
                    txtShippingPlant.Enabled = true;
                }
                else
                {
                    btnPopUpShippingPlant.Enabled = false;
                    txtShippingPlant.Enabled = false;
                }
            }
            else
            {
                btnPopUpShippingPlant.Enabled = false;
                txtShippingPlant.Enabled = false;
            }
        }

        private void cbShippingPlant_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dtTglValidasi_ValueChanged(object sender, EventArgs e)
        {

        }

        private void txtNoOrder_TextChanged(object sender, EventArgs e)
        {
            UpdateHeaderPartyFields();
        }

        private void lblOutletDesc_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void btnShippingPlant_Click(object sender, EventArgs e)
        {

            strSQL = "SELECT DISTINCT pms_ss_site AS [Shipping Plant Code], [br_branch_desc] AS [Shipping Plant Name] FROM GS_PRM_MULTI_SOURCE WITH (NOLOCK) INNER JOIN [GS_BRANCH] WITH(NOLOCK) ON [br_branch_ud2] = pms_ss_site WHERE pms_branch_id = '" + xBranchID + "' AND pms_flag_sloc = 'DDE' ";

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

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void cbfullfilment_CheckedChanged(object sender, EventArgs e)
        {
            if (cbfullfilment.Checked)
            {
                cbfullfilment.Text = "Car";
            }
            else
            {
                cbfullfilment.Text = "All";
            }
        }

        private void chkSLED_CheckedChanged(object sender, EventArgs e)
        {
            if (_ignoreSLEDCheckEvent) return;
            if (chkSLED.Checked)
            {
                //txtSLED.ReadOnly = false;
                for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                {
                    dgvSalesDetail.Rows[i].Cells["wsd_sled"].ReadOnly = false;

                }
            }
            else
            {

                //jika terdapat mapping sled maka tidak bisa dichange cheked slednya
                var checkSled = ((DataTable)dgvSalesDetail.DataSource).Rows.Cast<DataRow>()
                  .Where(x => !x["wsd_prd_master_code"].ToString().IsNullOrEmptyOrWhiteSpace() && decimal.TryParse(x["wsd_sled"]?.ToString(), out var sled) && sled > 0m).ToList();
                string ProductCodes = string.Join(",", checkSled.Select(p => $"'{p["wsd_prd_master_code"].ToString()}'"));
                string ProductSled = string.Join(",", checkSled.Select(p => p["wsd_prd_master_code"].ToString()));

                if (ProductCodes.IsNullOrEmptyOrWhiteSpace())
                {
                    ProductCodes = "''";
                }
                strSQL = "";
                strSQL += "SELECT mcps_prd_code from TBL_MAP_CUST_PRD_SLED with (nolock) ";
                strSQL += " WHERE mcps_entity_id='" + xEntityID + "' and mcps_branch_id='" + xBranchID + "' ";
                strSQL += " and mcps_cust_code1='" + txtKdOutlet1.Text.Trim() + "' and mcps_cust_code2='" + lblKdOutlet2.Text.Trim() + "' and mcps_prd_code in (" + ProductCodes + ")";

                DataTable dtCheck = _clsGlobal.ExecDT(strSQL);
                if (dtCheck.Rows.Count > 0)
                {
                    for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                    {
                        dgvSalesDetail.Rows[i].Cells["wsd_sled"].ReadOnly = false;

                    }
                    // Kembalikan ceklis
                    _ignoreSLEDCheckEvent = true;
                    chkSLED.Checked = true;
                    _ignoreSLEDCheckEvent = false;
                    MessageBox.Show($"Terdapat Mapping Customer Product SLED : {ProductSled} , Tidak Bisa di Unchecked! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }
                else
                {
                    for (int i = 0; i < dgvSalesDetail.Rows.Count; i++)
                    {
                        dgvSalesDetail.Rows[i].Cells["wsd_sled"].ReadOnly = true;
                        //dgvSalesDetail.Rows[i].Cells["wsd_sled"].Value = "";

                    }
                }


            }
        }



        private void txtSLED_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void getReason()
        {
            //Tipe Operasi
            string strSQL = "";
            strSQL = "SELECT '' AS reason UNION ALL SELECT gh_function_code + ' ~ ' + gh_function_desc AS reason FROM GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name ='REASONKAM' ORDER BY reason";

            cmbReasonGrid.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cmbReasonGrid.Properties.ValueMember = "reason";
            cmbReasonGrid.Properties.DisplayMember = "reason";
        }

        private void WorkDateAccess()
        {
            try
            {
                DataTable dtFill = new DataTable();

                strSQL = "EXEC SP_CURRENT_BOOK_MONTH '" + xEntityID + "','" + xBranchID + "'";
                dtFill = _clsGlobal.ExecDT(strSQL);

                if (dtFill.Rows.Count > 0)
                {
                    if (dtFill.Rows[0]["wh_loc_last_work_date"].ToString() != "")
                    {

                        g_dTglGudang = Convert.ToDateTime(dtFill.Rows[0]["wh_loc_last_work_date"].ToString());
                    }
                    else
                    {

                        g_dTglGudang = DateTime.Now;
                        MessageBox.Show("Tanggal gudang tidak ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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


    }
}

// ============================================================================
// Embedded DevExpress DataGridView compatibility classes.
// Diletakkan setelah class Form supaya Visual Studio Designer tetap membaca Form sebagai class utama.
// Tidak perlu file DevExpressGridCompat.cs terpisah.
// ============================================================================
namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{

    // DevExpress 20.2 grid compatibility members are embedded in this form file,
    // so no additional Compat/Wrapper file is required in the project.
    public enum DataGridViewColumnSortMode { NotSortable, Automatic, Programmatic }
    public enum DataGridViewAutoSizeColumnMode { NotSet, None, AllCells, AllCellsExceptHeader, DisplayedCells, DisplayedCellsExceptHeader, ColumnHeader, Fill }
    public enum DataGridViewAutoSizeColumnsMode { None, AllCells, AllCellsExceptHeader, DisplayedCells, DisplayedCellsExceptHeader, ColumnHeader, Fill }
    public enum DataGridViewColumnHeadersHeightSizeMode { EnableResizing, DisableResizing, AutoSize }
    public enum DataGridViewContentAlignment { NotSet, TopLeft, TopCenter, TopRight, MiddleLeft, MiddleCenter, MiddleRight, BottomLeft, BottomCenter, BottomRight }

    public delegate void DataGridViewCellEventHandler(object sender, DataGridViewCellEventArgs e);
    public delegate void DataGridViewEditingControlShowingEventHandler(object sender, DataGridViewEditingControlShowingEventArgs e);
    public delegate void DataGridViewRowPostPaintEventHandler(object sender, DataGridViewRowPostPaintEventArgs e);

    public class DataGridViewCellEventArgs : EventArgs
    {
        public int ColumnIndex { get; private set; }
        public int RowIndex { get; private set; }
        public DataGridViewCellEventArgs(int columnIndex, int rowIndex) { ColumnIndex = columnIndex; RowIndex = rowIndex; }
    }
    public class DataGridViewCellCancelEventArgs : DataGridViewCellEventArgs
    {
        public bool Cancel { get; set; }
        public DataGridViewCellCancelEventArgs(int columnIndex, int rowIndex) : base(columnIndex, rowIndex) { }
    }
    public class DataGridViewEditingControlShowingEventArgs : EventArgs
    {
        public Control Control { get; set; }
        public DataGridViewEditingControlShowingEventArgs(Control control) { Control = control; }
    }
    public class DataGridViewRowPostPaintEventArgs : EventArgs
    {
        public int RowIndex { get; set; }
        public Rectangle RowBounds { get; set; }
        public DataGridViewCellStyle InheritedRowStyle { get; set; }
        public Graphics Graphics { get; set; }
    }

    public class DataGridViewCellStyle
    {
        public DataGridViewContentAlignment Alignment { get; set; }
        public Color BackColor { get; set; }
        public Color ForeColor { get; set; }
        public Font Font { get; set; }
        public string Format { get; set; }
    }

    public class DataGridViewCell
    {
        private readonly DataGridView _cellOwnerGrid;
        private readonly int _rowIndex;
        private readonly string _fieldName;
        public DataGridViewCell(DataGridView grid, int rowIndex, string fieldName) { _cellOwnerGrid = grid; _rowIndex = rowIndex; _fieldName = fieldName; }
        public object Value { get { return _cellOwnerGrid.GetCellValue(_rowIndex, _fieldName); } set { _cellOwnerGrid.SetCellValue(_rowIndex, _fieldName, value); } }
        public bool Selected { get { return _cellOwnerGrid.CurrentCell != null && _cellOwnerGrid.CurrentCell.RowIndex == _rowIndex && _cellOwnerGrid.CurrentCell.OwningColumnName == _fieldName; } set { if (value) _cellOwnerGrid.SetCurrentCell(_rowIndex, _fieldName); } }
        public bool ReadOnly { get { return _cellOwnerGrid.GetCellReadOnly(_fieldName); } set { _cellOwnerGrid.SetCellReadOnly(_fieldName, value); } }
        internal string OwningColumnName { get { return _fieldName; } }
        public int RowIndex { get { return _rowIndex; } }
        public int ColumnIndex { get { var col = _cellOwnerGrid.FocusedView.Columns.ColumnByFieldName(_fieldName) ?? _cellOwnerGrid.FocusedView.Columns[_fieldName]; return _cellOwnerGrid.GetColumnIndex(col); } }
        public bool IsInEditMode { get { return _cellOwnerGrid.FocusedView.IsEditing; } }
    }

    public class DataGridViewCellCollection
    {
        private readonly DataGridView _grid;
        private readonly int _rowIndex;
        public DataGridViewCellCollection(DataGridView grid, int rowIndex) { _grid = grid; _rowIndex = rowIndex; }
        public DataGridViewCell this[string name] { get { return new DataGridViewCell(_grid, _rowIndex, name); } }
        public DataGridViewCell this[int index] { get { return new DataGridViewCell(_grid, _rowIndex, _grid.Columns[index].DataPropertyNameOrName); } }
    }

    public class DataGridViewRow
    {
        private readonly DataGridView _grid;
        private readonly int _rowIndex;
        public DataGridViewRow(DataGridView grid, int rowIndex) { _grid = grid; _rowIndex = rowIndex; Cells = new DataGridViewCellCollection(grid, rowIndex); }
        public DataGridViewCellCollection Cells { get; private set; }
        public int Index { get { return _rowIndex; } }
        public bool IsNewRow { get { return _grid.FocusedView.IsNewItemRow(_rowIndex); } }
        public bool Selected { get { return _grid.FocusedView.FocusedRowHandle == _rowIndex; } set { if (value) _grid.FocusedView.FocusedRowHandle = _rowIndex; } }
    }

    public class DataGridViewRowCollection : IEnumerable<DataGridViewRow>
    {
        private readonly DataGridView _grid;
        public DataGridViewRowCollection(DataGridView grid) { _grid = grid; }
        public int Count { get { return _grid.FocusedView.DataRowCount; } }
        public DataGridViewRow this[int index] { get { return new DataGridViewRow(_grid, index); } }
        public void Remove(DataGridViewRow row) { _grid.DeleteRow(row.Index); }
        public void Clear() { _grid.ClearRows(); }
        public int Add(object item) { return _grid.AddRow(item); }
        public IEnumerator<DataGridViewRow> GetEnumerator() { for (int i = 0; i < Count; i++) yield return new DataGridViewRow(_grid, i); }
        IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
    }

    public class DataGridViewSelectedRowCollection : IEnumerable<DataGridViewRow>
    {
        private readonly DataGridView _grid;
        public DataGridViewSelectedRowCollection(DataGridView grid) { _grid = grid; }
        public IEnumerator<DataGridViewRow> GetEnumerator()
        {
            int[] handles = _grid.FocusedView.GetSelectedRows();
            if (handles == null || handles.Length == 0) yield return new DataGridViewRow(_grid, _grid.FocusedView.FocusedRowHandle);
            else foreach (int h in handles) if (h >= 0) yield return new DataGridViewRow(_grid, h);
        }
        IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
    }

    public class DataGridViewColumn
    {
        internal GridColumn InnerColumn;
        private string _name;
        private string _dataPropertyName;
        private string _headerText;
        private int _width;
        private bool _visible = true;
        private bool _readOnly;
        private int _displayIndex = -1;
        private DataGridViewColumnSortMode _sortMode;
        private DataGridViewAutoSizeColumnMode _autoSizeMode;
        private bool _capSet, _visSet, _widthSet, _roSet;  // hanya terapkan properti yg eksplisit di-set

        public DataGridViewColumn()
        {
            DefaultCellStyle = new DataGridViewCellStyle();
            _visible = true;
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; ApplyToGridColumn(); }
        }

        public string DataPropertyName
        {
            get { return _dataPropertyName; }
            set { _dataPropertyName = value; ApplyToGridColumn(); }
        }

        public string HeaderText
        {
            get { return _headerText; }
            set { _headerText = value; _capSet = true; ApplyToGridColumn(); }
        }

        public int Width
        {
            get { return _width; }
            set { _width = value; ApplyToGridColumn(); }
        }

        public bool Visible
        {
            get { return _visible; }
            set { _visible = value; ApplyToGridColumn(); }
        }

        public bool ReadOnly
        {
            get { return _readOnly; }
            set { _readOnly = value; ApplyToGridColumn(); }
        }

        public int DisplayIndex
        {
            get { return InnerColumn == null ? _displayIndex : InnerColumn.VisibleIndex; }
            set { _displayIndex = value; if (InnerColumn != null) InnerColumn.VisibleIndex = value; }
        }

        public DataGridViewColumnSortMode SortMode
        {
            get { return _sortMode; }
            set { _sortMode = value; }
        }

        public DataGridViewAutoSizeColumnMode AutoSizeMode
        {
            get { return _autoSizeMode; }
            set
            {
                _autoSizeMode = value;
                if (InnerColumn != null)
                    InnerColumn.OptionsColumn.FixedWidth = value == DataGridViewAutoSizeColumnMode.None;
            }
        }

        public DataGridViewCellStyle DefaultCellStyle { get; set; }

        internal string DataPropertyNameOrName
        {
            get { return string.IsNullOrEmpty(_dataPropertyName) ? _name : _dataPropertyName; }
        }

        internal void ApplyToGridColumn()
        {
            if (InnerColumn == null) return;

            string fieldName = DataPropertyNameOrName;
            if (!string.IsNullOrEmpty(fieldName))
                InnerColumn.FieldName = fieldName;

            if (!string.IsNullOrEmpty(_name))
                InnerColumn.Name = _name;

            if (_capSet)
                InnerColumn.Caption = string.IsNullOrEmpty(_headerText) ? fieldName : _headerText;
            if (_visSet)
                InnerColumn.Visible = _visible;

            if (_widthSet && _width > 0)
                InnerColumn.Width = _width;

            if (_displayIndex >= 0)
                InnerColumn.VisibleIndex = _displayIndex;

            if (_roSet)
                InnerColumn.OptionsColumn.AllowEdit = !_readOnly;
            InnerColumn.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.False;
        }
    }
    public class DataGridViewTextBoxColumn : DataGridViewColumn { }

    public class DataGridViewColumnCollection : IEnumerable<DataGridViewColumn>
    {
        private readonly DataGridView _grid;
        private readonly List<DataGridViewColumn> _manualColumns = new List<DataGridViewColumn>();
        public DataGridViewColumnCollection(DataGridView grid) { _grid = grid; }
        public int Count { get { return _grid.FocusedView.Columns.Count; } }
        public DataGridViewColumn this[string name]
        {
            get
            {
                GridColumn col = _grid.FocusedView.Columns.ColumnByFieldName(name) ?? _grid.FocusedView.Columns[name];
                return _grid.WrapColumn(col, name);
            }
        }
        public DataGridViewColumn this[int index] { get { return _grid.WrapColumn(_grid.FocusedView.Columns[index], null); } }
        public bool Contains(string name) { return _grid.FocusedView.Columns.ColumnByFieldName(name) != null || _grid.FocusedView.Columns[name] != null; }
        public GridColumn ColumnByFieldName(string name) { return _grid.FocusedView.Columns.ColumnByFieldName(name); }
        public void Clear() { _grid.FocusedView.Columns.Clear(); }
        public GridColumn AddVisible(string fieldName, string caption)
        {
            var state = _grid.GetOrCreateColumnState(fieldName);
            state.Name = fieldName;
            state.DataPropertyName = fieldName;
            state.HeaderText = caption;
            GridColumn col = _grid.FocusedView.Columns.ColumnByFieldName(fieldName) ?? _grid.FocusedView.Columns[fieldName];
            if (col == null) col = _grid.FocusedView.Columns.AddVisible(fieldName, caption);
            state.InnerColumn = col;
            state.ApplyToGridColumn();
            return col;
        }
        public int Add(DataGridViewColumn column) { AddRange(column); return Count - 1; }
        public void AddRange(params DataGridViewColumn[] columns)
        {
            foreach (var c in columns)
            {
                _manualColumns.Add(c);
                var key = c.DataPropertyNameOrName;
                _grid.RegisterColumnState(key, c);
                GridColumn gc = _grid.FocusedView.Columns.ColumnByFieldName(key) ?? _grid.FocusedView.Columns[key];
                if (gc == null) gc = _grid.FocusedView.Columns.AddVisible(key, string.IsNullOrEmpty(c.HeaderText) ? key : c.HeaderText);
                c.InnerColumn = gc;
                c.ApplyToGridColumn();
            }
        }
        public IEnumerator<DataGridViewColumn> GetEnumerator() { for (int i = 0; i < Count; i++) yield return this[i]; }
        IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
    }

    public class DataGridViewComboBoxCell { }

    public class DataGridView : GridControl
    {
        private GridView _view;
        private object _dataSource;
        // jika true: kolom didefinisikan di Designer (design-time), lewati auto-PopulateColumns
        public bool UseDesignTimeColumns { get; set; }
        private readonly Dictionary<string, DataGridViewColumn> _columnStates = new Dictionary<string, DataGridViewColumn>(StringComparer.OrdinalIgnoreCase);
        public event DataGridViewCellEventHandler CellClick;
        public event DataGridViewCellEventHandler CellContentClick;
        public event DataGridViewCellEventHandler CellEndEdit;
        public event DataGridViewCellEventHandler CellValueChanged;
        public event DataGridViewEditingControlShowingEventHandler EditingControlShowing;
        public event DataGridViewRowPostPaintEventHandler RowPostPaint;
        public event EventHandler SelectionChanged;

        // Pelacakan sesi edit, agar CellEndEdit hanya menyala saat nilai cell BENAR-BENAR berubah
        // (meniru semantik System.Windows.Forms.DataGridView, bukan HiddenEditor DevExpress yang
        // ikut menyala hanya karena cell diklik lalu ditinggalkan tanpa diubah).
        private GridColumn _editColumn;
        private int _editRowHandle = -1;
        private object _editOriginalValue;

        public DataGridView()
        {
            _view = new GridView(this);
            MainView = _view;
            ViewCollection.Add(_view);
            Columns = new DataGridViewColumnCollection(this);
            Rows = new DataGridViewRowCollection(this);
            SelectedRows = new DataGridViewSelectedRowCollection(this);
            RowHeadersDefaultCellStyle = new DataGridViewCellStyle();
            ConfigureView(_view);
            // Adopsi GridView yang didefinisikan di Designer (jika ada) sebagai view aktif.
            this.ViewRegistered += AdoptDesignerView;
        }

        private void AdoptDesignerView(object sender, DevExpress.XtraGrid.ViewOperationEventArgs e)
        {
            GridView gv = e.View as GridView;
            if (gv == null || gv == _view) return;
            _view = gv;
            MainView = _view;
            ConfigureView(_view);
        }

        private readonly HashSet<GridView> _configuredViews = new HashSet<GridView>();
        private void ConfigureView(GridView v)
        {
            if (v == null || !_configuredViews.Add(v)) return;
            v.OptionsView.ShowGroupPanel = false;
            v.OptionsView.ShowColumnHeaders = true;
            v.OptionsView.ColumnAutoWidth = false;
            v.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            v.VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            v.OptionsBehavior.Editable = true;
            v.OptionsSelection.MultiSelect = true;
            v.FocusedRowChanged += (s, e) => SelectionChanged?.Invoke(this, EventArgs.Empty);
            v.CellValueChanged += (s, e) =>
            {
                AutoBestFitColumn(e.Column);
                CellValueChanged?.Invoke(this, new DataGridViewCellEventArgs(GetColumnIndex(e.Column), e.RowHandle));
            };
            v.RowCellClick += (s, e) => { var args = new DataGridViewCellEventArgs(GetColumnIndex(e.Column), e.RowHandle); CellClick?.Invoke(this, args); CellContentClick?.Invoke(this, args); };
            v.ShownEditor += (s, e) =>
            {
                // Rekam cell yang mulai diedit + nilai awalnya, supaya saat editor ditutup kita
                // bisa tahu apakah benar-benar terjadi perubahan.
                _editColumn = v.FocusedColumn;
                _editRowHandle = v.FocusedRowHandle;
                _editOriginalValue = (_editColumn != null && _editRowHandle >= 0)
                    ? v.GetRowCellValue(_editRowHandle, _editColumn)
                    : null;
                EditingControlShowing?.Invoke(this, new DataGridViewEditingControlShowingEventArgs(v.ActiveEditor));
            };
            // Nomor urut baris pada indikator — pengganti DataGridView.RowPostPaint (yang di DevExpress
            // tidak ada). Memberi efek visual sama: menampilkan (RowIndex + 1) di kolom indikator kiri.
            v.IndicatorWidth = 40;
            v.CustomDrawRowIndicator += (s, e) =>
            {
                if (e.Info != null && e.Info.IsRowIndicator && e.RowHandle >= 0)
                    e.Info.DisplayText = (e.RowHandle + 1).ToString();
            };
            v.HiddenEditor += (s, e) =>
            {
                // Pakai cell yang DIREKAM saat mulai edit (bukan FocusedColumn saat ini, karena fokus
                // bisa sudah pindah ke cell tujuan). Hanya picu CellEndEdit jika nilainya berubah.
                GridColumn col = _editColumn ?? v.FocusedColumn;
                int rowHandle = _editRowHandle >= 0 ? _editRowHandle : v.FocusedRowHandle;

                object newValue = (col != null && rowHandle >= 0)
                    ? v.GetRowCellValue(rowHandle, col)
                    : null;
                bool changed = !CellValuesEqual(_editOriginalValue, newValue);

                _editColumn = null;
                _editRowHandle = -1;
                _editOriginalValue = null;

                if (changed && col != null)
                    CellEndEdit?.Invoke(this, new DataGridViewCellEventArgs(GetColumnIndex(col), rowHandle));
            };
        }
        // Anggap null, DBNull, dan string kosong sebagai nilai yang sama (sehingga klik-tanpa-ubah
        // pada cell kosong maupun terisi tidak dianggap perubahan).
        private static bool CellValuesEqual(object a, object b)
        {
            string sa = (a == null || a == DBNull.Value) ? string.Empty : Convert.ToString(a);
            string sb = (b == null || b == DBNull.Value) ? string.Empty : Convert.ToString(b);
            return string.Equals(sa, sb, StringComparison.Ordinal);
        }

        internal GridView FocusedView
        {
            get
            {
                // MainView adalah view yang ter-bind & ditampilkan (di-set oleh Designer / runtime).
                // Sinkronkan _view ke MainView supaya operasi sel/baris mengenai data yang benar.
                GridView mv = MainView as GridView;
                if (mv != null && mv != _view) { _view = mv; ConfigureView(mv); }
                return _view;
            }
        }

        internal int GetColumnIndex(GridColumn col)
        {
            if (col == null) return -1;

            string fieldName = string.IsNullOrEmpty(col.FieldName) ? col.Name : col.FieldName;
            if (!string.IsNullOrEmpty(fieldName))
            {
                DataTable table = DataSource as DataTable;
                if (table != null && table.Columns.Contains(fieldName))
                    return table.Columns[fieldName].Ordinal;

                CurrencyManager cm = BindingContext[DataSource] as CurrencyManager;
                DataView dv = cm != null ? cm.List as DataView : null;
                if (dv != null && dv.Table != null && dv.Table.Columns.Contains(fieldName))
                    return dv.Table.Columns[fieldName].Ordinal;

                for (int i = 0; i < FocusedView.Columns.Count; i++)
                {
                    GridColumn c = FocusedView.Columns[i];
                    if (string.Equals(c.FieldName, fieldName, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(c.Name, fieldName, StringComparison.OrdinalIgnoreCase))
                        return i;
                }
            }

            return col.VisibleIndex;
        }

        public DataGridViewColumnCollection Columns { get; private set; }
        public DataGridViewRowCollection Rows { get; private set; }
        public DataGridViewSelectedRowCollection SelectedRows { get; private set; }
        public DataGridViewCell this[string columnName, int rowIndex] { get { return new DataGridViewCell(this, rowIndex, columnName); } }
        public DataGridViewCell this[int columnIndex, int rowIndex] { get { return new DataGridViewCell(this, rowIndex, Columns[columnIndex].DataPropertyNameOrName); } }
        public DataGridViewCellStyle RowHeadersDefaultCellStyle { get; set; }
        public bool AllowUserToAddRows { get; set; }
        public bool AllowUserToDeleteRows { get; set; }
        private bool _readOnly;
        public bool ReadOnly
        {
            get { return _readOnly; }
            set
            {
                _readOnly = value;
                FocusedView.OptionsBehavior.Editable = !value;
                foreach (GridColumn col in FocusedView.Columns)
                    col.OptionsColumn.AllowEdit = !value;
            }
        }
        private DataGridViewAutoSizeColumnsMode _autoSizeColumnsMode;
        public DataGridViewAutoSizeColumnsMode AutoSizeColumnsMode
        {
            get { return _autoSizeColumnsMode; }
            set
            {
                _autoSizeColumnsMode = value;
                ApplyScrollSettings();
            }
        }
        public DataGridViewColumnHeadersHeightSizeMode ColumnHeadersHeightSizeMode { get; set; }
        public bool IsCurrentCellDirty { get { return FocusedView.IsEditing; } }
        public int RowCount { get { return Rows.Count; } }
        public new object DataSource
        {
            get { return _dataSource; }
            set
            {
                _dataSource = value;
                base.DataSource = value;
                if (!UseDesignTimeColumns) FocusedView.PopulateColumns();
                ApplyScrollSettings();
                ReapplyColumnStates();
                AutoBestFitColumns();
            }
        }
        public DataGridViewCell CurrentCell
        {
            get { return FocusedView.FocusedColumn == null ? null : new DataGridViewCell(this, FocusedView.FocusedRowHandle, FocusedView.FocusedColumn.FieldName); }
            set { if (value != null) SetCurrentCell(value.RowIndex, value.OwningColumnName); }
        }
        public DataGridViewRow CurrentRow { get { return FocusedView.FocusedRowHandle < 0 ? null : new DataGridViewRow(this, FocusedView.FocusedRowHandle); } }
        internal void SetCurrentCell(int rowIndex, string fieldName)
        {
            FocusedView.FocusedRowHandle = rowIndex;
            var col = FocusedView.Columns.ColumnByFieldName(fieldName) ?? FocusedView.Columns[fieldName];
            if (col != null) FocusedView.FocusedColumn = col;
        }
        internal object GetCellValue(int rowIndex, string fieldName) { return rowIndex < 0 ? null : FocusedView.GetRowCellValue(rowIndex, fieldName); }
        internal void SetCellValue(int rowIndex, string fieldName, object value) { if (rowIndex >= 0) FocusedView.SetRowCellValue(rowIndex, fieldName, value); }
        internal bool GetCellReadOnly(string fieldName)
        {
            GridColumn col = FocusedView.Columns.ColumnByFieldName(fieldName) ?? FocusedView.Columns[fieldName];
            return col == null ? _readOnly : !col.OptionsColumn.AllowEdit;
        }
        internal void SetCellReadOnly(string fieldName, bool value)
        {
            GridColumn col = FocusedView.Columns.ColumnByFieldName(fieldName) ?? FocusedView.Columns[fieldName];
            if (col != null) col.OptionsColumn.AllowEdit = !value;
        }
        internal void RegisterColumnState(string key, DataGridViewColumn state)
        {
            if (string.IsNullOrEmpty(key) || state == null) return;
            _columnStates[key] = state;
        }

        internal DataGridViewColumn GetOrCreateColumnState(string key)
        {
            if (string.IsNullOrEmpty(key)) key = string.Empty;
            DataGridViewColumn state;
            if (!_columnStates.TryGetValue(key, out state))
            {
                state = new DataGridViewColumn();
                state.Name = key;
                state.DataPropertyName = key;
                _columnStates[key] = state;
            }
            return state;
        }

        private void ApplyScrollSettings()
        {
            FocusedView.OptionsView.ShowColumnHeaders = true;
            FocusedView.OptionsView.ColumnAutoWidth = false;
            FocusedView.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            FocusedView.VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            FocusedView.OptionsView.ShowGroupPanel = false;
        }

        private void AutoBestFitColumns()
        {
            try
            {
                FocusedView.OptionsView.ShowColumnHeaders = true;
                FocusedView.OptionsView.ColumnAutoWidth = false;
                FocusedView.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
                FocusedView.BestFitMaxRowCount = 100;

                foreach (GridColumn col in FocusedView.Columns)
                {
                    AutoBestFitColumn(col);
                }
            }
            catch
            {
                // Jangan ganggu business logic jika proses autosize gagal.
            }
        }

        private void AutoBestFitColumn(GridColumn col)
        {
            try
            {
                if (col == null || !col.Visible) return;

                col.OptionsColumn.FixedWidth = false;
                col.MinWidth = 50;
                col.BestFit();

                int captionWidth = TextRenderer.MeasureText(Convert.ToString(col.Caption), SystemFonts.DefaultFont).Width + 24;
                if (col.Width < captionWidth)
                    col.Width = captionWidth;

                if (col.Width < 50)
                    col.Width = 50;
            }
            catch
            {
                // BestFit hanya untuk tampilan, error diabaikan.
            }
        }

        private void ReapplyColumnStates()
        {
            ApplyScrollSettings();
            foreach (GridColumn col in FocusedView.Columns)
            {
                if (col == null) continue;
                var key = string.IsNullOrEmpty(col.FieldName) ? col.Name : col.FieldName;
                DataGridViewColumn state;
                if (_columnStates.TryGetValue(key, out state))
                {
                    state.InnerColumn = col;
                    state.ApplyToGridColumn();
                }
                else
                {
                    col.Caption = key;
                    col.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.False;
                }
            }
        }

        internal DataGridViewColumn WrapColumn(GridColumn gc, string name)
        {
            string key = gc != null ? (string.IsNullOrEmpty(gc.FieldName) ? gc.Name : gc.FieldName) : name;
            var state = GetOrCreateColumnState(key);
            if (gc != null)
            {
                state.InnerColumn = gc;
                if (string.IsNullOrEmpty(state.Name)) state.Name = gc.Name;
                if (string.IsNullOrEmpty(state.DataPropertyName)) state.DataPropertyName = gc.FieldName;
            }
            else
            {
                state.Name = key;
                state.DataPropertyName = key;
            }
            return state;
        }
        public void EndEdit() { FocusedView.CloseEditor(); FocusedView.UpdateCurrentRow(); }
        public void ClearSelection() { FocusedView.ClearSelection(); }
        public void DeleteRow(int rowIndex) { if (rowIndex >= 0) FocusedView.DeleteRow(rowIndex); }
        public void ClearRows()
        {
            DataTable dt = DataSource as DataTable;
            if (dt != null) { dt.Rows.Clear(); return; }
            FocusedView.DeleteSelectedRows();
        }
        public int AddRow(object item)
        {
            DataTable dt = DataSource as DataTable;
            if (dt != null)
            {
                if (item is DataRow) dt.ImportRow((DataRow)item);
                else if (item is DataTable && ((DataTable)item).Rows.Count > 0) dt.ImportRow(((DataTable)item).Rows[0]);
                else dt.Rows.Add(dt.NewRow());
                return dt.Rows.Count - 1;
            }
            FocusedView.AddNewRow();
            return FocusedView.FocusedRowHandle;
        }
        private GridColumn GetGridColumnByCompatIndex(int columnIndex)
        {
            if (columnIndex < 0) return null;

            DataTable table = DataSource as DataTable;
            if (table != null && columnIndex < table.Columns.Count)
            {
                string fieldName = table.Columns[columnIndex].ColumnName;
                return FocusedView.Columns.ColumnByFieldName(fieldName) ?? FocusedView.Columns[fieldName];
            }

            return columnIndex < FocusedView.Columns.Count ? FocusedView.Columns[columnIndex] : null;
        }

        public Rectangle GetCellDisplayRectangle(int columnIndex, int rowIndex, bool cutOverflow)
        {
            try
            {
                GridColumn col = GetGridColumnByCompatIndex(columnIndex);
                if (col == null || rowIndex < 0) return Rectangle.Empty;

                FocusedView.LayoutChanged();
                DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo viewInfo = FocusedView.GetViewInfo() as DevExpress.XtraGrid.Views.Grid.ViewInfo.GridViewInfo;
                if (viewInfo == null) return Rectangle.Empty;

                DevExpress.XtraGrid.Views.Grid.ViewInfo.GridCellInfo cellInfo = viewInfo.GetGridCellInfo(rowIndex, col);
                if (cellInfo == null) return Rectangle.Empty;

                return cellInfo.Bounds;
            }
            catch
            {
                return Rectangle.Empty;
            }
        }
        protected override void OnKeyDown(KeyEventArgs e) { base.OnKeyDown(e); }

        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, Keys keyData)
        {
            // Saat in-place editor aktif, key event mengalir ke editor, bukan ke GridControl, sehingga
            // KeyUp native (tempat F4 "Search Product" ditangani) tidak menyala seperti pada DataGridView lama.
            // Untuk editor non-dropdown (mis. TextEdit kolom pcode): tutup editor lalu teruskan F4 ke KeyUp.
            // Untuk editor dropdown (PopupBaseEdit: SearchLookUpEdit/LookUpEdit/ComboBoxEdit/DateEdit):
            // biarkan F4 berfungsi normal membuka dropdown.
            if (keyData == Keys.F4)
            {
                GridView gv = MainView as GridView;
                if (gv != null && gv.ActiveEditor != null
                    && !(gv.ActiveEditor is DevExpress.XtraEditors.PopupBaseEdit))
                {
                    gv.CloseEditor();
                    OnKeyUp(new KeyEventArgs(Keys.F4));
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }



}
