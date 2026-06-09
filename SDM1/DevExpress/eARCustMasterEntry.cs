using StackedHeader;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.DXErrorProvider;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;
using GridColumn = DevExpress.XtraGrid.Columns.GridColumn;
using DevExpress.XtraTab;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Reflection;
using System.Net;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.IO;


namespace TIRASnDNet.AR.ARCustMaster
{
    public partial class eARCustMasterEntry : DevExpress.XtraEditors.XtraForm
    {
        const bool __isAsk11214 = true;
        const bool __isAsk27918 = true;
        const bool __isAsk28151 = true;
        //const bool __isAsk30724 = true;
        LineHighlight __cursor = null;

        public enum StateEntry
        {
            New = 1,
            Edit = 2,
        };

        clsGlobal _clsGlobal = new clsGlobal();
        private string strSQL, _prmSalesman, _prmSalesmanTo, strSQL1, strSQL2, strSQL3, _prdCustCode1, _prdCustCode2, _prdEntityCode, _prdBrandCode, _Grade, billConde1, billConde2, shipCode1, shipCode2;
        private bool rowDuplicate = false, exitData = false;
        private bool eBlackList = false;
        private readonly int LIMIT_COUNT_HOTT_TAX_CODES = 8;
        readonly Color __COLOR;


        const string QUERY_NIK_BLOCK = @"SELECT gh_sequence_no AS [No],gh_function_code AS [Code] FROM GEN_HARDCODED WITH(NOLOCK) WHERE gh_sys='H' AND gh_function_name='XTUNAINIKFLAG'";
        const string QUERY_SKILLID = @"SELECT TT.stt_id [Kode] ,TT.stt_desc [Nama] FROM [dbo].[TBL_SD_TRANSPORT_TYPE] TT WITH(NOLOCK)";
        const string QUERY_HIRARKI = @"select sk.kl_kelurahan_code [Kode],sk.kl_kelurahan_desc [Nama] from [dbo].[VW_SD_KELURAHAN] sk where sk.kl_kelurahan_code='{0}' and sk.kl_kecamatan_code='{1}' and sk.kl_city_code='{2}' and sk.kl_prov_code='{3}'";
        const string QUERY_OUTLET_STATUS = @"select gs.gh_function_code [kode], gs.gh_function_desc [nama] from [dbo].[GS_GEN_HARDCODED] gs WITH (NOLOCK) where gs.gh_function_name='OUTLETSTATUS' order by gs.gh_sequence_no";
        const string QUERY_OUTLET_STATUS_NEW = @"select gs.gh_function_code [kode], gs.gh_function_desc [nama] from [dbo].[GS_GEN_HARDCODED] gs WITH (NOLOCK) where gs.gh_function_name='OUTLETSTATUS' and gs.gh_function_code in ('O', 'P') order by gs.gh_sequence_no";

        TIRAData<GsGenHardcodedItem> __nik_blocked = null;
        TIRAObject<TIRAItem> __prdlines = null;
        TIRAObject<TIRAItem> __prdlineblocking = null;
        TIRAObject<TIRAItem> __top = null;
        TIRAObject<TIRAItem> __skill = null;
        TIRAData<TIRAItem> __outlet_status = null;

        bool __isAutoTunai = false;
        bool __isFlagKAM = false;
        bool __isFlagDC = false;

        TIRAData<TiraPhoto> photoExist = null;
        PhotoItem nik = null;
        PhotoItem npwp = null;

        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _salesmanSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit _salesmanWeekCheckRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemComboBox _salesmanDayComboRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _groupHargaSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _topPrdLineSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _topPaymentSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _topPlantSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _topSalesEmployeeSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit _groupHargaButtonRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit _taxCodeButtonRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit _gridSearchButtonRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _partnerShipToSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit _partnerDefaultCheckRepository = null;

        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _hirPrdLineSearchLookupRepository = null;
        private DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit _hirCustTypeSearchLookupRepository = null;
        private bool _isUpdatingHirFromSearchLookup = false;

        private bool _isUpdatingSalesmanFromSearchLookup = false;
        private bool _isUpdatingGroupHargaFromSearchLookup = false;
        private bool _isUpdatingTopByPrdLineFromSearchLookup = false;
        private bool _isOpeningGroupHargaPopup = false;
        private bool _isUpdatingPartnerDefault = false;
        private DXValidationProvider dxValidator = null;
        private readonly Dictionary<Control, Color> _validationOriginalBackColorMap = new Dictionary<Control, Color>();
        private readonly Dictionary<Control, Color> _validationOriginalBaseEditBackColorMap = new Dictionary<Control, Color>();
        private readonly Dictionary<Control, XtraTabPage> _validationTabMap = new Dictionary<Control, XtraTabPage>();
        private readonly Dictionary<Control, ValidationRule> _validationRuleMap = new Dictionary<Control, ValidationRule>();
        private readonly List<Control> _registeredValidationControls = new List<Control>();
        private bool _isRefreshingValidationState = false;


        /** SYNC **/
        bool __isFlagSync = false;
        /** END **/

        public string PrmSalesman
        {
            get { return _prmSalesman; }
            set { _prmSalesman = value; }
        }

        public string PrmSalesmanTo
        {
            get { return _prmSalesmanTo; }
            set { _prmSalesmanTo = value; }
        }

        public eARCustMasterEntry()
        {
            if (IsInDesignerMode())
            {
                try
                {
                    InitializeComponent();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Designer InitializeComponent skipped: " + ex.Message);
                }
            }
            else
            {
                InitializeComponent();
                this.Load += new System.EventHandler(this.eARCustMasterEntry_Load);
            }
            __COLOR = Color.FromArgb(245, 245, 255);

            if (!IsInDesignerMode())
            {
                InitializeDxValidation();
            }
        }

        private static bool IsInDesignerMode()
        {
            if (System.ComponentModel.LicenseManager.UsageMode == System.ComponentModel.LicenseUsageMode.Designtime)
                return true;

            string processName = string.Empty;
            try
            {
                processName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            }
            catch
            {
                processName = string.Empty;
            }

            return string.Equals(processName, "devenv", StringComparison.OrdinalIgnoreCase)
                || string.Equals(processName, "XDesProc", StringComparison.OrdinalIgnoreCase);
        }

        public string PrdBrandCode
        {
            get { return _prdBrandCode; }
            set { _prdBrandCode = value; }
        }

        public string PrdCustCode1
        {
            get { return _prdCustCode1; }
            set { _prdCustCode1 = value; }
        }

        public string PrdCustCode2
        {
            get { return _prdCustCode2; }
            set { _prdCustCode2 = value; }
        }

        public string PrdEntityCode
        {
            get { return _prdEntityCode; }
            set { _prdEntityCode = value; }
        }

        public string PrdGrade
        {
            get { return _Grade; }
            set { _Grade = value; }
        }

        StateEntry __state = StateEntry.New;
        public StateEntry State
        {
            get { return __state; }
            set { __state = value; }
        }

        private void btnPopUpTremPayT1_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Kode Pasar ";
            frm.Query = " select pptc_term_code [Term], pptc_term_desc [Desc], pptc_no_of_days [Days] from PO_PAYMENT_TERM_CODES WITH (NOLOCK) ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtTremPaymentT1.Text = frm.ArrField[0].Trim();
                txtTremPaymentToT1.Text = frm.ArrField[1].Trim();
            }

        }

        private void btnPopUpGroupHrgT1_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Group Harga  ";
            frm.Query = "  select ppl_price_list_code [Code], ppl_price_list_description [Description] from OB_PRICE_LIST_HEADER WITH (NOLOCK) ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtGroupHargaT1.Text = frm.ArrField[0].Trim();
                txtGpHargaToT1.Text = frm.ArrField[1].Trim();
            }

        }

        private void btnPopUpIndusT1_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Industri ";
            frm.Query = "select ind_grp_disc [Industri], ind_sort_desc [Description]from TBL_SD_INDUSTRY WITH (NOLOCK) ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtIndustriT1.Text = frm.ArrField[0].Trim();
                txtIndustriToT1.Text = frm.ArrField[1].Trim();
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBoxWarehouseTidak_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxWarehouseTidak.Checked)
            {
                checkBoxWarehouseTidak.Text = "Ya";
            }
            else
            {
                checkBoxWarehouseTidak.Text = "Tidak";
            }
        }

        private void checkBoxInvoiceTidak_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxInvoiceTidak.Checked)
            {
                checkBoxInvoiceTidak.Text = "Ya";
            }
            else
            {
                checkBoxInvoiceTidak.Text = "Tidak";
            }
        }

        private void checkBoxNpwpT1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxNpwpT1.Checked)
            {
                checkBoxNpwpT1.Text = "Ya";
                comboBoxPajakT1.Enabled = true;
                settab2aktif();
            }
            else
            {
                checkBoxNpwpT1.Text = "Tidak";
                comboBoxPajakT1.Enabled = false;
                comboBoxPajakT1.ItemIndex = 1;
                //settab2();
                settab2aktif();
            }
        }

        private void checkBoxKontraBonT1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxKontraBonT1.Checked)
            {
                checkBoxKontraBonT1.Text = "Ya";
            }
            else
            {
                checkBoxKontraBonT1.Text = "Tidak";
            }
        }

        private void checkBoxDiscountBaseT1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxDiscountBaseT1.Checked)
            {
                checkBoxDiscountBaseT1.Text = "Sesudah PPN";
            }
            else
            {
                checkBoxDiscountBaseT1.Text = "Sebelum PPN";
            }
        }

        private void checkBoxTopByCustomerT1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxTopByCustomerT1.Checked)
            {
                checkBoxTopByCustomerT1.Text = "Ya";
                btnPopUpTremPayT1.Enabled = true;
                txtTremPaymentT1.Enabled = true;
            }
            else
            {
                checkBoxTopByCustomerT1.Text = "Tidak";
                btnPopUpTremPayT1.Enabled = false;
                txtTremPaymentT1.Enabled = false;
            }
        }

        private void checkBoxtxtBatasLimitT1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxtxtBatasLimitT1.Checked)
            {
                checkBoxtxtBatasLimitT1.Text = "Dibatasi";
            }
            else
            {
                checkBoxtxtBatasLimitT1.Text = "Tidak dibatasi";
            }
        }

        private void CKBKalender1_CheckedChanged(object sender, EventArgs e)
        {

            if (CKBKalender1.Checked)
            {
                dateTimePicker1.Enabled = true;
            }
            else
            {
                dateTimePicker1.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                dateTimePicker2.Enabled = true;
            }
            else
            {
                dateTimePicker2.Enabled = false;
            }

        }

        private void chkPinSL_CheckedChanged(object sender, EventArgs e)
        {
            if (chkPinSL.Checked)
            {
                chkPinSL.Text = "Ya";
            }
            else
            {
                chkPinSL.Text = "Tidak";
            }
        }

        void setGridViewSalesman()
        {


            foreach (GridColumn col in DVGSalesmanView.Columns)
            {
                if (col.VisibleIndex == 4 || col.VisibleIndex == 9)
                {
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                }
            }
        }

        /** SYNC **/
        private bool IsSync()
        {

            DataTable dtGS = _clsGlobal.ExecDT("select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'FLAGSYNC_SDM1' and gh_sys = 'H'");
            if (dtGS.Rows.Count > 0)
            {
                if (dtGS.Rows[0][0].ToString().Trim() == "Y")
                {
                    return true;
                }
            }


            return false;
        }
        /** END SYNC **/

        private void eARCustMasterEntry_Load(object sender, EventArgs e)
        {
            if (IsInDesignerMode())
                return;
            NormalizeConvertedDevExpressUi();
            NormalizeTaxTabLayout();
            //__cursor = new LineHighlight(this.panel1);
            //foreach (Control __child in this.panel1.Controls)
            //{
            //    if (__child.TabStop) __cursor.Add(__child);
            //}
            //panel1.Focus();
            getSource();

            panel1.AutoScrollMinSize = new Size(0, 500);

            DevExpress.XtraTab.XtraTabPage t = tabControl1.TabPages[3];
            tabControl1.SelectedTabPage = t;

            setGridViewSalesman();
            try
            {

                __isFlagKAM = _clsGlobal.IsNefoForKam;
                __isFlagDC = _clsGlobal.IsNefoForDC;
                if (!__isFlagDC)
                {
                    telf_DA_lbl.Text = telf_DA_lbl.Text.Replace("*", "");
                    skill_DA_lbl.Text = skill_DA_lbl.Text.Replace("*", "");
                }
                pathniklbl.Visible = __isFlagDC;
                pathnpwplbl.Visible = __isFlagDC;
                viewnikbtn.Visible = __isFlagDC;
                viewnpwpbtn.Visible = __isFlagDC;
                //txtSFACustCode.Visible = __isFlagDC;
                //label113.Visible = __isFlagDC;

                __prdlines = new TIRAObject<TIRAItem>("SELECT pl_prd_line_code AS [Kode],pl_prd_line_desc AS [Nama] FROM IM_PRD_LINE WITH(NOLOCK)", 0)
                {
                    Title = "Search Product Line",
                    EndAction = _EndPrdLine,
                };
                __prdlines.Initialize();

                __prdlineblocking = new TIRAObject<TIRAItem>("SELECT pl_prd_line_code AS [Kode],pl_prd_line_desc AS [Nama] FROM IM_PRD_LINE WITH(NOLOCK)", 0)
                {
                    Title = "Search Product Line",
                    EndAction = _EndPrdLineBlocking,
                };
                __prdlineblocking.Initialize();

                __top = new TIRAObject<TIRAItem>("SELECT pptc_term_code AS [Kode],pptc_term_code+' - '+pptc_term_desc AS [Nama] FROM PO_PAYMENT_TERM_CODES WITH(NOLOCK) ORDER BY pptc_no_of_days", 0)
                {
                    Title = "Search TOP",
                    EndAction = _EndTOP,
                };
                __top.Initialize();

                __skill = new TIRAObject<TIRAItem>(QUERY_SKILLID, 0)
                {
                    Title = "Search Skill ID",
                    EndAction = _EndSkillID,
                };
                __skill.AddRangeButton(skillidcodetbn);
                __skill.AddRangeTextBox((DevExpress.XtraEditors.TextEdit)skillidcodetb);
                __skill.Initialize();

                //if (__isAsk30724)
                //{
                __outlet_status = new TIRAData<TIRAItem>(__state == StateEntry.New ?
                    QUERY_OUTLET_STATUS_NEW :
                    QUERY_OUTLET_STATUS, 0);
                __outlet_status.Execute();

                cbstatus.Properties.DisplayMember = "Nama";
                cbstatus.Properties.ValueMember = "Kode";
                cbstatus.Properties.DataSource = __outlet_status.DATA;
                //cbstatus.EditValue = __outlet_status.DATA.FirstOrDefault();
                cbstatus.EditValue = __state == StateEntry.New ? "O" : null;

                cbstatus.Enabled = __state == StateEntry.New;
                cbstatus.Visible = true;
                txtstatus.Visible = false;
                //}
                //else
                //{
                //    cbstatus.Visible = false;
                //    txtstatus.Visible = true;
                //}

                createTopByPrdLine();
                if (topbyprdline_page != null)
                    topbyprdline_page.Text = "Customer by Division";

                if (!__isFlagKAM)
                    tabControl1.TabPages.Remove(topbyprdline_page);

                SetGridViewSchPayD();
                SetGridViewSchPayW();

                SetGridViewPrdLineBlocking();

                /*ZipCode Inisiator agar bisa langsung diisi zipcodenya ketika pilih city */
                popUpProvinsinCity1.Needzipcode(CreateHiddenTextBox(txtPostCodeT4));
                /* end zipcode inisiator */

                /* SYNC */
                __isFlagSync = IsSync();
                /* END OF SYNC */

                if (clsLogin.BTNEDIT == true)
                {
                    btnOk.Enabled = true;
                }
                else
                {
                    btnOk.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            tabControl1.SelectedPageChanged -= tabControl1_SelectedPageChanged;
            tabControl1.SelectedPageChanged += tabControl1_SelectedPageChanged;

            foreach (GridColumn column in DVGSalesmanView.Columns)
            {
                column.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.False;
            }

            foreach (GridColumn column in DGVHIRView.Columns)
            {
                column.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.False;
            }

            _redForeColorMandatory();

            //untuk combo box pilihan
            LoadDropDown();
            comboBoxPajakT1.Enabled = false;
            comboBoxPajakT1.ItemIndex = 1;
            string custcode1 = _prdCustCode1;
            string custcode2 = _Grade;

            if (!__isAsk27918)
            {
                lblTaxAddressChoice.Visible = false;
                cbAddressChoice.Visible = false;
            }

            //fullfilment
            if (!__isAsk28151)
            {
                lblfullfilment.Visible = false;
                cbfullfilment.Visible = false;
            }

            //if (clsGlobal.MODE_TRX == 2)
            if (__state == StateEntry.Edit)
            {

                FillData();

                cmbSource.Enabled = false;

                txtCustCode.Enabled = false;

                ctrlEntityCustMaster2.Enabled = false;

                FillGridHirarki();

                FillGridSalesman();

                FillGridGroupHarga();
                FillGridPartnerFunction();

                FillGridBackList();
                FillGridUNBackList();

                if (__isFlagKAM)
                    FillTopByPrdLine();

                FillCustSchPay();
                FillPrdLineBlocking();

                settab2aktif();
                //added 09042019
                // Native DevExpress: jangan disable GridControl karena tampilannya menjadi abu-abu / terlihat blank.
                // Jika tidak boleh edit saat mode Edit, cukup matikan editing di GridView.
                DGVHIR.Enabled = true;
                DGVHIRView.OptionsBehavior.Editable = false;

                strSQL = "select ISNULL(cm_delivery_days_no,0) cm_delivery_days_no,ISNULL(cm_delivery_day,0) cm_delivery_day,cm_duration_days, ISNULL(cm_sfa_code,'') cm_sfa_code from SO_CUST_MASTER WITH(NOLOCK) where cm_cust_code1 =";
                strSQL += "'" + custcode1 + "' And cm_cust_code2 ='" + custcode2 + "'";

                DataTable dts = new DataTable();

                dts = _clsGlobal.ExecDT(strSQL);
                if (dts.Rows.Count > 0)
                {
                    txtDelivByDays.Text = dts.Rows[0]["cm_delivery_days_no"].ToString().Trim();
                    CBDelivByDay.EditValue = dts.Rows[0]["cm_delivery_day"].ToString().Trim();
                    txtDurationOfDays.Text = dts.Rows[0]["cm_duration_days"].ToString().Trim();

                    if (!__isFlagKAM)
                    {
                        txtSFACustCode.Text = dts.Rows[0]["cm_sfa_code"].ToString().Trim();
                    }
                }
                else
                {
                    CBDelivByDay.ItemIndex = 0;
                    txtDelivByDays.Text = "2";
                }
            }
            else
            {
                //added 09042019

                DGVHIR.Enabled = true;
                checkBoxFakPajakT2.Checked = true;
                checkBoxDiscountBaseT1.Checked = true;
                strSQL = " select a.ge_entity_id , a.ge_entity as Description from GS_ENTITY as a WITH (NOLOCK) WHERE ge_entity_id = '01' order by ge_entity_id asc ";

                DataTable dt = new DataTable();
                dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    //adien
                    //ctrlEntityCustMaster1.txtCM.Text = dt.Rows[0]["ge_entity_id"].ToString().Trim();
                    //ctrlEntityCustMaster1.txtCMTo.Text = dt.Rows[0]["Description"].ToString().Trim();
                }

                strSQL1 = " select a.br_branch_id , a.br_branch_desc as Description from GS_BRANCH as a WITH (NOLOCK) WHERE br_branch_id = '39' order by br_branch_id asc ";

                DataTable dt1 = new DataTable();
                dt1 = _clsGlobal.ExecDT(strSQL1);

                if (dt1.Rows.Count > 0)
                {
                    //adien
                    //ctrlBranchCustMaster1.txtBranchIdCM.Text = dt1.Rows[0]["br_branch_id"].ToString().Trim();
                    //ctrlBranchCustMaster1.txtBranchIdCMTo.Text = dt1.Rows[0]["Description"].ToString().Trim();
                    //adien
                }

                strSQL2 = " select wh_loc_id1 , wh_loc_id2, wh_loc_name [Warehouse Name] from IM_WH_LOC WITH (NOLOCK) where wh_loc_type = 'N'";

                DataTable dt2 = new DataTable();
                dt2 = _clsGlobal.ExecDT(strSQL2);

                if (dt2.Rows.Count > 0)
                {
                    txtWHCode.Text = dt2.Rows[0]["wh_loc_id1"].ToString().Trim();
                    txtWHCodeTo.Text = dt2.Rows[0]["wh_loc_id2"].ToString().Trim();
                }
                strSQL3 = " select wh_loc_last_work_date from IM_WH_LOC WITH (NOLOCK) where wh_loc_id1 = 'PST' and wh_loc_id2 = '000' ";

                DataTable dt3 = new DataTable();
                dt3 = _clsGlobal.ExecDT(strSQL3);

                if (dt3.Rows.Count > 0)
                {
                    // txtTglRegisT1.Text = dt3.Rows[0]["wh_loc_last_work_date"].ToString().Trim();
                    txtTglRegisT1.Text = Convert.ToDateTime(dt3.Rows[0]["wh_loc_last_work_date"].ToString()).ToString("dd MMM yyyy");
                }

                txtCustCode.Text = "0";
                txtSFACustCode.Text = "";

                //settab2();
                settab2aktif();
                txtDelivByDays.Text = "2";

                checkBoxMonitoringInsuranceT1.Checked = false;
                lbltglfromins.Visible = false;
                lbltgltoins.Visible = false;
                dtTglFromIns.Visible = false;
                dtTglToIns.Visible = false;
            }

            FillGridSalesman();
            loadNikBlocked();
            photoCheck();

            //txtCustCodeTo.ReadOnly = true;            
            //txtCustCodeTo.Text = "000";
            //txtLeadtime.Text = "3";
            //txtOutSandingT1.Text = "00000";

            //checkBoxFakPajakT2.Checked = true;
            txtReason.Visible = false;
            label18.Visible = false;
            CKBKalender1.Checked = true;
            checkBox2.Checked = true;
            //txtNpwpT2.Enabled = false;
            //txtNpwpT2Masked.Enabled = false;
            txtRataRppT1.Enabled = false;
            txtTransaksiPertamaT1.Enabled = false;
            txtTglTransaksiAkhirT1.Enabled = false;
            txtTglRegisT1.Enabled = false;
            txtLastUpdateT1.Enabled = false;

            tbNumeric_Leave(this.txtLimitKreditT1, null);
            tbNumeric_Leave(this.txtOutSandingT1, null);

            EnsureAllDevExpressGridColumnsVisible();
            EnsurePartnerFunctionTabRuntime();
            ConfigureHierarchyGridRuntime();
            ConfigureSalesmanCoveredGridRuntime();
            ConfigureGroupHargaGridRuntime();
            ConfigurePartnerFunctionGridRuntime();
            ConfigureTopByPrdLineGridRuntime();
            ConfigureTaxGridRuntime();
            ConfigureSchedulePaymentGridRuntime();
            ConfigurePrdLineBlockingGridRuntime();
            SetEditableGridView();

            this.groupBox1.Focus();
            txtCustName.Focus();
        }


        void loadNikBlocked()
        {
            try
            {
                __nik_blocked = new TIRAData<GsGenHardcodedItem>(QUERY_NIK_BLOCK, 1);
                __nik_blocked.Execute();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void photoCheck()
        {
            try
            {
                string __entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
                string __branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
                string __code1 = txtCustCode.Text.Trim();
                string __code2 = txtCustCodeTo.Text.Trim();

                StringBuilder query = new StringBuilder();
                query.AppendLine(" select scp_custcode1 code");
                query.AppendLine(" ,cast(case when scp_nik_photo is null then 0 else 1 end as bit) isnik");
                query.AppendLine(" ,cast(case when scp_npwp_photo is null then 0 else 1 end as bit) isnpwp");
                query.AppendLine(" from [dbo].[SO_CUST_PHOTO] with(nolock)");
                query.AppendFormat(" where [scp_entity]='{0}'", __entity);
                query.AppendFormat(" and [scp_branch]='{0}'", __branch);
                query.AppendFormat(" and [scp_custcode1]='{0}'", __code1);
                query.AppendFormat(" and [scp_custcode2]='{0}'", __code2);
                photoExist = new TIRAData<TiraPhoto>(query.ToString(), 0);
                //photoExist = new TIRAData<TiraPhoto>("select scp_custcode1 code,cast(case when scp_nik_photo is null then 0 else 1 end as bit) isnik,cast(case when scp_npwp_photo is null then 0 else 1 end as bit) isnpwp from [dbo].[SO_CUST_PHOTO] with(nolock)", 0);
                photoExist.Execute();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void LoadDropDown()
        {
            try
            {
                //Size Group Code

                strSQL = "select gh_sequence_no, CONCAT(gh_sequence_no,' ~ ',gh_function_desc) as gh_function_desc from GS_GEN_HARDCODED WITH (NOLOCK) where gh_sys ='H' AND gh_function_name ='NOTRETPAJAK' order by gh_sequence_no ";
                comboBoxPajakT1.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                comboBoxPajakT1.Properties.ValueMember = "gh_sequence_no";
                comboBoxPajakT1.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            try
            {
                //Size Group Code
                strSQL = "select ghc_sequence_no, CONCAT(ghc_sequence_no,' ~ ',ghc_function_desc) as ghc_function_desc from GS_HARD_CODED WITH(NOLOCK) where ghc_sys ='H' AND ghc_function_name ='SHIPBILL'";
                CBShipnBill.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBShipnBill.Properties.ValueMember = "ghc_sequence_no";
                CBShipnBill.Properties.DisplayMember = "ghc_function_desc";
                CBShipnBill.Properties.NullText = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                strSQL = "select gh_function_code, CONCAT(gh_function_code,' ~ ',gh_function_desc) as gh_function_desc from GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name = 'CUSTPAYTYPE'";
                CBPembyaranFktT1.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPembyaranFktT1.Properties.ValueMember = "gh_function_code";
                CBPembyaranFktT1.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                strSQL = "select gh_function_code, CONCAT(gh_function_code,' ~ ',gh_function_desc) as gh_function_desc from GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name = 'BANGUNAN'";
                CBBagunanT1.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBBagunanT1.Properties.ValueMember = "gh_function_code";
                CBBagunanT1.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //strSQL = "SELECT gh_function_code, gh_function_desc FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name = 'WEEKLYDAYS'";
                strSQL = "SELECT gh_function_code,gh_function_desc FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name='WEEKLYDAYS' ORDER BY gh_function_code";
                //strSQL =  "declare @i int = 0 CREATE TABLE #TABLE (hari varchar(100), angka int) While @i<7 BEGIN ";
                //strSQL += "if(@i = 0) Begin INSERT INTO #TABLE (hari, angka) values ('',99) End ";
                //strSQL += "INSERT INTO #TABLE (hari, angka) values ( DATENAME(WEEKDAY,@i),@i) SET @i = @i+1 END ";
                //strSQL += "select hari, angka from #table drop table #table";
                CBDelivByDay.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBDelivByDay.Properties.ValueMember = "gh_function_code";
                CBDelivByDay.Properties.DisplayMember = "gh_function_desc";
                CBDelivByDay.ItemIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // yuda 06092021 //
            try
            {
                strSQL = "";
                strSQL = "SELECT '' as gh_function_code, '' as gh_function_desc ";
                strSQL = strSQL + " union ALL ";
                strSQL = strSQL + "SELECT gh_function_code,gh_function_desc FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name='PEMILIK_NIK' ORDER BY gh_function_code ASC";
                cbPemilikNIK.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                cbPemilikNIK.Properties.ValueMember = "gh_function_code";
                cbPemilikNIK.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                strSQL = "";
                strSQL = "SELECT '' as gh_function_code, '' as gh_function_desc ";
                strSQL = strSQL + " union ALL ";
                strSQL = strSQL + "SELECT gh_function_code,gh_function_desc FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name='JENIS_IDENTITAS' ORDER BY gh_function_code ASC";
                cbJenisIdentitas.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                cbJenisIdentitas.Properties.ValueMember = "gh_function_code";
                cbJenisIdentitas.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



            // ------end ----//

            /* Ticket 27918 Yuda 09-06-2022*/
            if (__isAsk27918)
            {
                try
                {
                    strSQL = "";
                    //strSQL = "SELECT '' as gh_function_code, '' as gh_function_desc ";
                    //strSQL = strSQL + " union ALL ";
                    strSQL = strSQL + "SELECT gh_function_code,gh_function_code + '~' + gh_function_desc  AS gh_function_desc FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name='TAX_ADDRESS' and gh_sys = 'H' ORDER BY gh_sequence_no ASC";
                    cbAddressChoice.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                    cbAddressChoice.Properties.ValueMember = "gh_function_code";
                    cbAddressChoice.Properties.DisplayMember = "gh_function_desc";

                    if (__isFlagKAM)
                    {
                        cbAddressChoice.ItemIndex = 1;
                    }
                    else
                    {
                        cbAddressChoice.ItemIndex = 0;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            /* end of Ticket */
        }

        void settab2()
        {

            txtTaxNameT2.Enabled = false;
            txtAddress1T2.Enabled = false;
            txtAddress2T2.Enabled = false;
            checkBoxFakPajakT2.Enabled = false;
            txtCountryT2.Enabled = false;
            //txtNpwpT2.Enabled = false;
            txtNpwpT2Masked.Enabled = false;
            txtTaxIdT2.Enabled = false;
            //txtTaxIdT2To.Enabled = false;
            txtBankGroupT2.Enabled = false;
            txtBankCodeT2.Enabled = false;
            txtAccountNoT2.Enabled = false;
            txtPostCodeT2.Enabled = false;
            txtTelephoneT2.Enabled = false;
            txtFaxT2.Enabled = false;

            popUpProvinsinCityTax1.txtPropinsiCMC.Enabled = false;
            popUpProvinsinCityTax1.txtPropinsiCMCTo.Enabled = false;
            popUpProvinsinCityTax1.txtKabupatenCMC.Enabled = false;
            popUpProvinsinCityTax1.txtKabupatenCMCTo.Enabled = false;
            popUpProvinsinCityTax1.btnPopUpPropinsiCMC.Enabled = false;
            popUpProvinsinCityTax1.btnPopUpKabupatenCMC.Enabled = false;

            txtNoReffSupplierT2.Enabled = false;
            txtNIK.Enabled = false;
            BTNBank.Enabled = false;
            BTNbankcode.Enabled = false;
            BTNTaxid.Enabled = false;

        }

        void settab2aktif()
        {

            txtTaxNameT2.Enabled = true;
            txtAddress1T2.Enabled = true;
            txtAddress2T2.Enabled = true;
            checkBoxFakPajakT2.Enabled = true;
            txtCountryT2.Enabled = true;
            //txtNpwpT2.Enabled = true;
            txtNpwpT2Masked.Enabled = true;
            txtTaxIdT2.Enabled = true;
            //txtTaxIdT2To.Enabled = true;
            txtBankGroupT2.Enabled = true;
            txtBankCodeT2.Enabled = true;
            txtAccountNoT2.Enabled = true;
            txtPostCodeT2.Enabled = true;
            txtTelephoneT2.Enabled = true;
            txtFaxT2.Enabled = true;

            popUpProvinsinCityTax1.txtPropinsiCMC.Enabled = true;
            popUpProvinsinCityTax1.txtKabupatenCMC.Enabled = true;
            popUpProvinsinCityTax1.btnPopUpPropinsiCMC.Enabled = true;
            popUpProvinsinCityTax1.btnPopUpKabupatenCMC.Enabled = true;
            txtNoReffSupplierT2.Enabled = true;
            txtNIK.Enabled = true;
            BTNBank.Enabled = true;
            BTNbankcode.Enabled = true;
            BTNTaxid.Enabled = true;

        }

        private void _redForeColorMandatory()
        {
            try
            {
                foreach (Control c in this.Controls.OfType<GroupBox>())
                {
                    foreach (Control ch in c.Controls.OfType<Label>().Where(p => p.Text.Contains("*")))
                        ch.ForeColor = Color.Red;

                    foreach (Control ch in c.Controls.OfType<UserControl>())
                        foreach (Control cx in ch.Controls.OfType<Label>().Where(p => p.Text.Contains("*")))
                            cx.ForeColor = Color.Red;
                }

                foreach (Control c in this.Controls.OfType<TabControl>())
                    foreach (Control cp in c.Controls.OfType<TabPage>())
                    {
                        foreach (Control ch in cp.Controls.OfType<Label>().Where(p => p.Text.Contains("*")))
                            ch.ForeColor = Color.Red;

                        foreach (Control ch in cp.Controls.OfType<UserControl>())
                            foreach (Control cx in ch.Controls.OfType<Label>().Where(p => p.Text.Contains("*")))
                                cx.ForeColor = Color.Red;
                    }
            }
            catch (Exception) { }
        }

        private void CBShipnBill_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBShipnBill.EditValue == null)
                return;

            if (CBShipnBill.ItemIndex < 0)
                return;

            int a = CBShipnBill.ItemIndex;
            IsiTextBill(a, txtCustCode.Text);

        }

        void IsiTextBill_old(int a, string val)
        {
            string strVal = (val == "" ? "0" : val);

            if (a == 0)
            {
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");
                txtShiptoCustCode.Text = (__state == StateEntry.New ? "" : "");

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");//
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "" : "");//

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); // billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); // billConde1

                txtBilltoCustCodeTo.Visible = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2);
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2);

                BTNBillcust.Visible = false;
                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 1)
            {
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCode.Text = (__state == StateEntry.New ? "" : ""); //

                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "" : ""); //
                txtBilltoCustCodeTo.Visible = true;

                BTNBillcust.Visible = true;
                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 2)
            {
                //txtShiptoCustCode.Visible = false;
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCodeTo.Visible = false;
                txtShiptoCustCodeTo.Visible = true;
                //txtBilltoCustCode.Visible = false;
                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCodeTo.Visible = false;
                txtBilltoCustCodeTo.Visible = true;
                //txtBilltoCustCode.Visible = false;
                txtBilltoCustCode.Visible = true;

                //BTNBillcust.Visible = false;
                BTNBillcust.Visible = true;
                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                //label16.Visible = false;
                //label17.Visible = false;
                label16.Visible = true;
                label17.Visible = true;

                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //billConde1
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
            }
            else if (a == 3)
            {
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);                
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //billConde1

                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
                txtBilltoCustCodeTo.Visible = true;

                BTNBillcust.Visible = false;
                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
        }

        void IsiTextBill(int a, string val)
        {
            string strVal = (val == "" ? "0" : val);

            if (a == 0)
            {
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal);

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");//
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);//


                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); // billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); // billConde1

                txtBilltoCustCodeTo.Visible = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2);
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2);

                BTNShipcust.Visible = true;
                BTNShipcust.Enabled = true;
                BTNBillcust.Visible = false;
                BTNBillcust.Enabled = false;
                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 1)
            {
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCode.Text = (__state == StateEntry.New ? "" : ""); //
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //

                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
                txtBilltoCustCodeTo.Visible = true;

                BTNShipcust.Visible = false;
                BTNShipcust.Enabled = false;
                BTNBillcust.Visible = true;
                BTNBillcust.Enabled = true;
                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 2)
            {
                //txtShiptoCustCode.Visible = false;
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCodeTo.Visible = false;
                txtShiptoCustCodeTo.Visible = true;
                //txtBilltoCustCode.Visible = false;
                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCodeTo.Visible = false;
                txtBilltoCustCodeTo.Visible = true;
                txtBilltoCustCodeTo.Enabled = true;
                //txtBilltoCustCode.Visible = false;



                //BTNBillcust.Visible = false;
                BTNShipcust.Visible = false;
                BTNShipcust.Enabled = false;

                BTNBillcust.Visible = false;
                BTNBillcust.Enabled = false;

                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                //label16.Visible = false;
                //label17.Visible = false;
                label16.Visible = true;
                label17.Visible = true;

                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //billConde1
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
            }
            else if (a == 3)
            {
                txtShiptoCustCode.Visible = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);                
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //billConde1

                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
                txtBilltoCustCodeTo.Visible = true;

                BTNBillcust.Visible = false;
                txtWHCode.Enabled = false;
                txtWHCodeTo.Enabled = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
        }

        private void txtCustName_TextChanged(object sender, EventArgs e)
        {
            txtCustNameT3.Text = txtCustName.Text;
            txtCustNameT4.Text = txtCustName.Text;
        }

        private void txtAddress1T3_TextChanged(object sender, EventArgs e)
        {
            txtAddress1T4.Text = txtAddress1T3.Text;
        }

        private void txtAddress2T3_TextChanged(object sender, EventArgs e)
        {
            txtAddress2T4.Text = txtAddress2T3.Text;

        }

        private void txtAddress3T3_TextChanged(object sender, EventArgs e)
        {
            txtAddress3T4.Text = txtAddress3T3.Text;
        }

        private void txtAddress4T3_TextChanged(object sender, EventArgs e)
        {
            txtAddress4T4.Text = txtAddress4T3.Text;
        }

        void FillData()
        {

            try
            {

                if (_prmSalesman == "" && _prmSalesmanTo == "")
                {
                    //strSQL1 = " select * from VW_CUST_MASTER WHERE cm_cust_code1 ='" + _prdBrandCode + "' ";
                    strSQL1 = " select * from VW_CUST_MASTER where cm_entity = '" + _prdEntityCode + "' and cm_branch = '" + _prdBrandCode + "' AND cm_cust_code1  = '" + _prdCustCode1 + "'AND cm_cust_code2  like '" + _prdCustCode2 + "%'";
                }
                else
                {
                    //strSQL1 = " select * from VW_CUST_MASTER_CVR_SLS WHERE cm_cust_code1 ='" + _prdBrandCode + "' ";
                    strSQL1 = " select * from VW_CUST_MASTER_CVR_SLS where cm_entity = '" + _prdEntityCode + "' and cm_branch = '" + _prdBrandCode + "' AND cm_cust_code1  = '" + _prdCustCode1 + "'AND cm_cust_code2  like '" + _prdCustCode2 + "%'";
                }

                DataTable dt1 = new DataTable();
                dt1 = _clsGlobal.ExecDT(strSQL1);

                if (dt1.Rows.Count > 0)
                {
                    if (dt1.Rows[0]["cm_ship_bill_to_flag"] == DBNull.Value) throw new Exception("Field 'cm_ship_bill_to_flag' is empty.");
                    if (dt1.Rows[0]["cm_ship_bill_to_flag"].ToString().IsNullOrEmptyOrWhiteSpace()) throw new Exception("Field 'cm_ship_bill_to_flag' is empty.");

                    // CBShipnBill.EditValue = dt1.Rows[0]["cm_ship_bill_to_flag"].ToString().Trim();
                    string shipBillFlag = Convert.ToString(dt1.Rows[0]["cm_ship_bill_to_flag"]).Trim();
                    SetShipBillValue(shipBillFlag);



                    txtCustCode.Text = dt1.Rows[0]["cm_cust_code1"].ToString().Trim();

                    txtCustCodeTo.Text = dt1.Rows[0]["cm_cust_code2"].ToString().Trim();
                    txtCustName.Text = dt1.Rows[0]["cm_cust_name"].ToString().Trim();
                    txtShortName.Text = dt1.Rows[0]["cm_cust_short_name"].ToString().Trim();
                    txtBarcode.Text = dt1.Rows[0]["cm_barcode_code"].ToString().Trim();

                    //if (__isFlagDC)
                    //{
                    //    txtSFACustCode.Text = dt1.Rows[0]["cm_sfa_code"].ToString().Trim();
                    //}

                    //txtstatus.Text = dt1.Rows[0]["cm_cust_status"].ToString().Trim();
                    //if (__isAsk30724)
                    //{
                    var status = __outlet_status.DATA.FirstOrDefault(p => p.Kode
                        .CompareC((dt1.Rows[0]["cm_active_flag"] != DBNull.Value ? dt1.Rows[0]["cm_active_flag"] : string.Empty)
                        .ToString()
                        .Trim()));
                    cbstatus.EditValue = status != null ? status.Kode : string.Empty;

                    if (dt1.Rows[0]["cm_source"] != DBNull.Value)
                    {
                        cmbSource.EditValue = dt1.Rows[0]["cm_source"].ToString();
                    }


                    DateTime create = Convert.ToDateTime(dt1.Rows[0]["cm_create_date"] != DBNull.Value ? dt1.Rows[0]["cm_create_date"] : DateTime.MinValue);
                    cbstatus.Enabled = create.Date == DateTime.Today;
                    //}
                    //else
                    //{
                    //    txtstatus.Text = dt1.Rows[0]["cm_cust_status"].ToString().Trim();
                    //}

                    txtLat.Text = dt1.Rows[0]["cm_lat"].ToString().Trim();
                    txtLong.Text = dt1.Rows[0]["cm_lng"].ToString().Trim();

                    txtPoCustCode.Text = dt1.Rows[0]["cm_po_cust_code"].ToString().Trim();
                    //adien
                    //ctrlEntityCustMaster1.txtCM.Text = dt1.Rows[0]["cm_entity"].ToString().Trim();
                    //ctrlEntityCustMaster1.txtCMTo.Text = dt1.Rows[0]["ge_entity"].ToString().Trim();

                    ctrlEntityCustMaster2.txtCM.Text = dt1.Rows[0]["cm_entity"].ToString().Trim();
                    ctrlEntityCustMaster2.txtCMTo.Text = dt1.Rows[0]["ge_entity"].ToString().Trim();
                    ctrlEntityCustMaster2.txtBranchIdCM.Text = dt1.Rows[0]["cm_branch"].ToString().Trim();
                    ctrlEntityCustMaster2.txtBranchIdCMTo.Text = dt1.Rows[0]["br_branch_desc"].ToString().Trim();

                    popUpCOA1.EntityId = ctrlEntityCustMaster2.txtCM.Text.Trim();
                    popUpCOA1.GetFormatCOA(popUpCOA1.EntityId, popUpCOA1.BranchId, popUpCOA1.DivisionId, popUpCOA1.DepartmentId, popUpCOA1.Major1, popUpCOA1.Major2, popUpCOA1.Minor, popUpCOA1.Analisys, popUpCOA1.Filler);

                    //ctrlBranchCustMaster1.txtBranchIdCM.Text = dt1.Rows[0]["cm_branch"].ToString().Trim();
                    //ctrlBranchCustMaster1.txtBranchIdCMTo.Text = dt1.Rows[0]["br_branch_desc"].ToString().Trim();
                    //adien
                    ctrlDistrikCustMaster1.txtDistrikCM.Text = dt1.Rows[0]["cm_area"].ToString().Trim();
                    ctrlDistrikCustMaster1.txtDistrikCMTo.Text = dt1.Rows[0]["rg_region_desc"].ToString().Trim();
                    ctrlDistrikCustMaster1.txtBeatCM.Text = dt1.Rows[0]["cm_wilayah"].ToString().Trim();
                    ctrlDistrikCustMaster1.txtBeatCMTo.Text = dt1.Rows[0]["wy_wilayah_desc"].ToString().Trim();
                    ctrlDistrikCustMaster1.txtSubBeatCM.Text = dt1.Rows[0]["cm_rayon"].ToString().Trim();
                    ctrlDistrikCustMaster1.txtSubBeatCMTo.Text = dt1.Rows[0]["ry_rayon_desc"].ToString().Trim();
                    ctrlGroupOutlet1.txtGroupCM.Text = dt1.Rows[0]["cm_cust_group"].ToString().Trim();
                    ctrlGroupOutlet1.txtGroupCMTo.Text = dt1.Rows[0]["cg_cust_group_desc"].ToString().Trim();
                    ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Text = dt1.Rows[0]["cm_cust_type"].ToString().Trim();
                    ctrlpopUpTypeOutletCustMaster1.txtTypeOutletTo.Text = dt1.Rows[0]["ct_cust_type_desc"].ToString().Trim();
                    ctrlKalsifikasiCustMaster1.txtKlasifikasiCM.Text = dt1.Rows[0]["cm_cust_class"].ToString().Trim();
                    ctrlKalsifikasiCustMaster1.txtKlasifikasiCMTo.Text = dt1.Rows[0]["cc_cust_class_desc"].ToString().Trim();
                    ctrlKategoriCustMaster1.txtKategoriCM.Text = dt1.Rows[0]["cm_cust_catg_code"].ToString().Trim();//cm_cust_category
                    ctrlKategoriCustMaster1.txtKategoriCMTo.Text = dt1.Rows[0]["ctg_cust_catg_desc"].ToString().Trim();//cc_cust_class_desc
                    ctrlCounterTipeMC1.txtCounterType.Text = dt1.Rows[0]["cm_cust_counter_type"].ToString().Trim();
                    ctrlCounterTipeMC1.txttxtCounterTypeTo.Text = dt1.Rows[0]["cct_cust_counter_type_desc"].ToString().Trim();
                    txtMerchanid.Text = dt1.Rows[0]["cm_merchant_id"].ToString().Trim();


                    shipCode1 = dt1.Rows[0]["cm_ship_to_code1"].ToString().Trim();
                    shipCode2 = dt1.Rows[0]["cm_ship_to_code2"].ToString().Trim();
                    billConde1 = dt1.Rows[0]["cm_bill_to_code1"].ToString().Trim();
                    billConde2 = dt1.Rows[0]["cm_bill_to_code2"].ToString().Trim();

                    txtShiptoCustCode.Text = shipCode1;
                    txtShiptoCustCodeTo.Text = shipCode2;
                    txtBilltoCustCode.Text = billConde1;
                    txtBilltoCustCodeTo.Text = billConde2;
                    LoadPodRelevantFlag();

                    txtLeadtime.Text = dt1.Rows[0]["cm_lead_time"].ToString().Trim();

                    txtWHCode.Text = dt1.Rows[0]["cm_ship_from_wh_code1"].ToString().Trim();
                    txtWHCodeTo.Text = dt1.Rows[0]["cm_ship_from_wh_code2"].ToString().Trim();
                    txtMoidCode.Text = dt1.Rows[0]["cm_moid_code"].ToString().Trim();
                    ctrlGroupDiscount1.txtGroupDiscountT1.Text = dt1.Rows[0]["cm_group_discount"].ToString().Trim();
                    ctrlGroupDiscount1.txtGpDiscountT1.Text = dt1.Rows[0]["grd_long_desc"].ToString().Trim();
                    ctrlLokasiDesc1.txtLokasiT1.Text = dt1.Rows[0]["cm_location"].ToString().Trim();
                    ctrlLokasiDesc1.txtLokasiToT1.Text = dt1.Rows[0]["loc_long_desc"].ToString().Trim();
                    txtGroupHargaT1.Text = dt1.Rows[0]["cm_def_price_code"].ToString().Trim();
                    //txtGpHargaToT1.Text = dt1.Rows[0]["cm_def_price_code"].ToString().Trim();
                    txtIndustriT1.Text = dt1.Rows[0]["cm_industry"].ToString().Trim();
                    txtIndustriToT1.Text = dt1.Rows[0]["ind_long_desc"].ToString().Trim();
                    ctrlKodepasarDesc.txtKodePasarT1.Text = dt1.Rows[0]["cm_pasar_code"].ToString().Trim();
                    ctrlKodepasarDesc.txtKodePasarToT1.Text = dt1.Rows[0]["psr_long_desc"].ToString().Trim();
                    ctrlGroupPLU1.txtGroupPLUT1.Text = dt1.Rows[0]["cm_group_plu"].ToString().Trim();
                    ctrlGroupPLU1.txtGroupPLUToT1.Text = dt1.Rows[0]["plu_long_desc"].ToString().Trim();
                    txtGroupKonversiT1.Text = dt1.Rows[0]["cm_group_conversi"].ToString().Trim();
                    txtGroupKonversiToT1.Text = dt1.Rows[0]["grc_long_desc"].ToString().Trim();
                    txtTremPaymentT1.Text = dt1.Rows[0]["cm_top_id"].ToString().Trim();
                    //txtTremPaymentToT1.Text = dt1.Rows[0]["cm_payment_type"].ToString().Trim();

                    if (dt1.Rows[0]["cm_with_warehouse"].ToString().Trim() == "Y")
                    {
                        checkBoxWarehouseTidak.Checked = true;
                    }
                    else
                    {
                        checkBoxWarehouseTidak.Checked = false;
                    }

                    if (dt1.Rows[0]["cm_inv_recalc_flag"].ToString().Trim() == "Y")
                    {
                        checkBoxInvoiceTidak.Checked = true;
                    }
                    else
                    {
                        checkBoxInvoiceTidak.Checked = false;
                    }
                    if (dt1.Rows[0]["cm_npwp_flag"].ToString().Trim() == "Y")
                    {
                        checkBoxNpwpT1.Checked = true;
                        settab2aktif();

                    }
                    else
                    {
                        checkBoxNpwpT1.Checked = false;
                        settab2();
                    }
                    comboBoxPajakT1.EditValue = dt1.Rows[0]["cm_notret_pajak"].ToString().Trim();

                    string tmpLimit;
                    tmpLimit = dt1.Rows[0]["cm_ship_credit_limit"].ToString().Trim();

                    txtLimitKreditT1.Text = tmpLimit.Replace(".0000", "");

                    txtOutSandingT1.Text = dt1.Rows[0]["cm_outstanding"].ToString().Trim();
                    if (dt1.Rows[0]["cm_status_cr"].ToString().Trim() == "Y")
                    {
                        checkBoxtxtBatasLimitT1.Checked = true;
                    }
                    else
                    {
                        checkBoxtxtBatasLimitT1.Checked = false;
                    }
                    CBPembyaranFktT1.EditValue = dt1.Rows[0]["cm_payment_type"].ToString().Trim();

                    if (dt1.Rows[0]["cm_kontra_bon"].ToString().Trim() == "Y")
                    {
                        checkBoxKontraBonT1.Checked = true;
                    }
                    else
                    {
                        checkBoxKontraBonT1.Checked = false;
                    }
                    txtNoPekanAwalT1.Text = dt1.Rows[0]["cm_first_week"].ToString().Trim();
                    txtMengageniT1.Text = dt1.Rows[0]["cm_mengageni"].ToString().Trim();

                    if (dt1.Rows[0]["cm_discount_base"].ToString().Trim() == "Y")//cm_discount_base
                    {
                        checkBoxDiscountBaseT1.Checked = true;
                    }
                    else
                    {
                        checkBoxDiscountBaseT1.Checked = false;
                    }
                    if (dt1.Rows[0]["cm_top_by_cust"].ToString().Trim() == "Y")
                    {
                        checkBoxTopByCustomerT1.Checked = true;
                    }
                    else
                    {
                        checkBoxTopByCustomerT1.Checked = false;
                    }



                    txtRataRppT1.Text = dt1.Rows[0]["cm_rpp13"].ToString().Trim();
                    txtTransaksiPertamaT1.Text = dt1.Rows[0]["cpl_first_update"].ToString().Trim();
                    txtTglTransaksiAkhirT1.Text = dt1.Rows[0]["cpl_last_update"].ToString().Trim();
                    txtTglRegisT1.Text = Convert.ToDateTime(dt1.Rows[0]["cm_outlet_register"].ToString().Trim()).ToString("dd MMM yyyy");
                    CBBagunanT1.EditValue = dt1.Rows[0]["cm_bangunan"].ToString().Trim();
                    txtLastUpdateT1.Text = dt1.Rows[0]["cm_update_date"].ToString().Trim();

                    dateTimePicker1.Text = Convert.ToDateTime(dt1.Rows[0]["cm_outlet_open"].ToString()).ToString("dd MMM yyyy");
                    dateTimePicker2.Text = Convert.ToDateTime(dt1.Rows[0]["cm_birthdate"].ToString()).ToString("dd MMM yyyy");


                    //tab2
                    txtTaxNameT2.Text = dt1.Rows[0]["cm_tax_name"].ToString().Trim();
                    txtAddress1T2.Text = dt1.Rows[0]["cm_tax_address1"].ToString().Trim();
                    txtAddress2T2.Text = dt1.Rows[0]["cm_tax_address2"].ToString().Trim();

                    if (dt1.Rows[0]["cm_tax_flaq"].ToString().Trim() == "Y")//GNTI cm_tax_flaq
                    {
                        checkBoxFakPajakT2.Checked = true;
                    }
                    else
                    {
                        checkBoxFakPajakT2.Checked = false;
                    }

                    //if (dt1.Rows[0]["cm_npwp_flag"].ToString().Trim() == "Y")//GNTI cm_npwp_flag
                    //{
                    //    checkBoxFakPajakT2.Checked = true;
                    //    txtNpwpT2.Enabled = true;
                    //}
                    //else
                    //{
                    //    checkBoxFakPajakT2.Checked = false;
                    //    txtNpwpT2.Enabled = false;
                    //}

                    if (dt1.Rows[0]["cm_flag_blacklist"].ToString().Trim() == "Y")
                    {
                        CBBlacklist.Checked = true;
                        eBlackList = true;
                    }
                    else
                    {
                        CBBlacklist.Checked = false;
                        eBlackList = false;
                    }

                    txtCountryT2.Text = dt1.Rows[0]["cm_tax_country"].ToString().Trim();

                    //txtNpwpT2.Text = dt1.Rows[0]["cm_ship_npwp"].ToString().Trim();
                    txtNpwpT2Masked.Text = dt1.Rows[0]["cm_ship_npwp"].ToString().Trim();

                    txtTaxIdT2.Text = dt1.Rows[0]["cm_ship_tax_id"].ToString().Trim();
                    txtTaxIdT2To.Text = dt1.Rows[0]["t_tax"].ToString().Trim();
                    txtBankGroupT2.Text = dt1.Rows[0]["cm_bank_group"].ToString().Trim();
                    txtBankCodeT2.Text = dt1.Rows[0]["cm_bank_code"].ToString().Trim();
                    txtAccountNoT2.Text = dt1.Rows[0]["cm_bank_account"].ToString().Trim();
                    txtPostCodeT2.Text = dt1.Rows[0]["cm_tax_zip_code"].ToString().Trim();
                    txtTelephoneT2.Text = dt1.Rows[0]["cm_tax_phone"].ToString().Trim();
                    txtFaxT2.Text = dt1.Rows[0]["cm_tax_fax"].ToString().Trim();

                    popUpProvinsinCityTax1.txtPropinsiCMC.Text = dt1.Rows[0]["cm_tax_proviency"].ToString().Trim();
                    //popUpProvinsinCityTax1.txtPropinsiCMCTo.Text = dt1.Rows[0]["proviency_delv"].ToString().Trim();
                    popUpProvinsinCityTax1.txtKabupatenCMC.Text = dt1.Rows[0]["cm_tax_city"].ToString().Trim();
                    //popUpProvinsinCityTax1.txtKabupatenCMCTo.Text = dt1.Rows[0]["bill_cy_city_desc"].ToString().Trim();

                    txtNoReffSupplierT2.Text = dt1.Rows[0]["cm_no_reff_supply"].ToString().Trim();
                    txtNIK.Text = dt1.Rows[0]["cm_nik"].ToString().Trim();
                    txtBillBranch.Text = dt1.Rows[0]["cm_cust_bill_branch"].ToString().Trim();

                    //tab 3

                    txtCustNameT3.Text = dt1.Rows[0]["cm_bill_name"].ToString().Trim();
                    txtAddress1T3.Text = dt1.Rows[0]["cm_bill_adrress1"].ToString().Trim();
                    txtAddress2T3.Text = dt1.Rows[0]["cm_bill_adrress2"].ToString().Trim();
                    txtAddress3T3.Text = dt1.Rows[0]["cm_bill_adrress3"].ToString().Trim();
                    txtAddress4T3.Text = dt1.Rows[0]["cm_bill_adrress4"].ToString().Trim();
                    txtContactPersonT3.Text = dt1.Rows[0]["cm_bill_contact_person"].ToString().Trim();
                    txtCountryT3.Text = dt1.Rows[0]["cm_bill_country"].ToString().Trim();
                    txtPostCodeT3.Text = dt1.Rows[0]["cm_bill_zip_code"].ToString().Trim();
                    txtTeleponT3.Text = dt1.Rows[0]["cm_bill_phone"].ToString().Trim();
                    txtFaxT3.Text = dt1.Rows[0]["cm_bill_fax"].ToString().Trim();
                    txtEmailT3.Text = dt1.Rows[0]["cm_bill_email_name"].ToString().Trim();

                    popUpProvinsinCityBil1.txtPropinsiCMCBil.Text = dt1.Rows[0]["cm_bill_proviency"].ToString().Trim();
                    //popUpProvinsinCityBil1.txtPropinsiCMCToBil.Text = dt1.Rows[0]["bill_pv_prov_desc"].ToString().Trim();
                    popUpProvinsinCityBil1.txtKabupatenCMCBil.Text = dt1.Rows[0]["cm_bill_city"].ToString().Trim();
                    //popUpProvinsinCityBil1.txtKabupatenCMCToBil.Text = dt1.Rows[0]["bill_cy_city_desc"].ToString().Trim();

                    //tab4
                    txtCustNameT4.Text = dt1.Rows[0]["cm_delv_name"].ToString().Trim();
                    txtAddress1T4.Text = dt1.Rows[0]["cm_delv_address1"].ToString().Trim();
                    txtAddress2T4.Text = dt1.Rows[0]["cm_delv_address2"].ToString().Trim();
                    txtAddress3T4.Text = dt1.Rows[0]["cm_delv_address3"].ToString().Trim();
                    txtAddress4T4.Text = dt1.Rows[0]["cm_delv_address4"].ToString().Trim();
                    txtContactPersonT4.Text = dt1.Rows[0]["cm_delv_contact_person"].ToString().Trim();
                    txtCountryT4.Text = dt1.Rows[0]["cm_delv_country"].ToString().Trim();
                    txtPostCodeT4.Text = dt1.Rows[0]["cm_delv_zip_code"].ToString().Trim();
                    txtTelephoneT4.Text = dt1.Rows[0]["cm_delv_phone"].ToString().Trim();
                    txtFaxT4.Text = dt1.Rows[0]["cm_delv_fax"].ToString().Trim();
                    txtEmailT4.Text = dt1.Rows[0]["cm_delv_email_name"].ToString().Trim();
                    if (dt1.Rows[0]["cm_flag_pinalty_sl"].ToString() == "Y") { chkPinSL.Checked = true; } else { chkPinSL.Checked = false; }
                    if (dt1.Rows[0]["cm_dear_notification"].ToString() == "Y") { cbDearNotification.Checked = true; } else { cbDearNotification.Checked = false; }
                    txtShipArea.Text = dt1.Rows[0]["cm_ship_area"].ToString().Trim();

                    popUpProvinsinCity1.txtPropinsiCMC.Text = dt1.Rows[0]["cm_delv_proviency"].ToString().Trim();
                    //popUpProvinsinCity1.txtPropinsiCMCTo.Text = dt1.Rows[0]["proviency_delv"].ToString().Trim();
                    popUpProvinsinCity1.txtKabupatenCMC.Text = dt1.Rows[0]["cm_delv_city"].ToString().Trim();
                    //popUpProvinsinCity1.txtKabupatenCMCTo.Text = dt1.Rows[0]["city_delv_desc"].ToString().Trim();
                    popUpProvinsinCity1.txtKecamatan.Text = dt1.Rows[0]["cm_devl_kecamatan"].ToString().Trim();
                    //popUpProvinsinCity1.txtKecamatanTo.Text = dt1.Rows[0]["kc_kecamatan_desc"].ToString().Trim();
                    popUpProvinsinCity1.txtKelurahan.Text = dt1.Rows[0]["cm_devl_kelurahan"].ToString().Trim();
                    //popUpProvinsinCity1.txtKelurahanTo.Text = dt1.Rows[0]["kl_kelurahan_desc"].ToString().Trim();

                    skillidcodetb.Text = dt1.Rows[0]["cm_skillid"].ToString().Trim();
                    //update yuda 09062021//
                    if (string.IsNullOrEmpty(dt1.Rows[0]["cm_pemilik_nik"].ToString().Trim()))
                    {
                        cbPemilikNIK.ItemIndex = 0;
                    }
                    else
                    {
                        cbPemilikNIK.EditValue = dt1.Rows[0]["cm_pemilik_nik"].ToString().Trim();
                    }
                    txtSPPKP.Text = dt1.Rows[0]["cm_sppkp"].ToString().Trim();
                    if (dt1.Rows[0]["cm_monitoring_insurance"].ToString().Trim() == "Y")
                    {
                        checkBoxMonitoringInsuranceT1.Checked = true;
                        lbltglfromins.Visible = true;
                        lbltgltoins.Visible = true;
                        dtTglFromIns.Visible = true;
                        dtTglToIns.Visible = true;

                        if (!string.IsNullOrEmpty(dt1.Rows[0]["cm_insurance_date_from"].ToString()) && !string.IsNullOrEmpty(dt1.Rows[0]["cm_insurance_date_to"].ToString()))
                        {
                            dtTglFromIns.Text = Convert.ToDateTime(dt1.Rows[0]["cm_insurance_date_from"].ToString()).ToString("dd MMM yyyy");
                            dtTglToIns.Text = Convert.ToDateTime(dt1.Rows[0]["cm_insurance_date_to"].ToString()).ToString("dd MMM yyyy");
                        }
                    }
                    else
                    {
                        checkBoxMonitoringInsuranceT1.Checked = false;
                        lbltglfromins.Visible = false;
                        lbltgltoins.Visible = false;
                        dtTglFromIns.Visible = false;
                        dtTglToIns.Visible = false;
                    }
                    dtOpenHour.DateTime = Convert.ToDateTime(dt1.Rows[0]["cm_open_hour"].ToString());
                    dtCloseHour.DateTime = Convert.ToDateTime(dt1.Rows[0]["cm_close_hour"].ToString());
                    nmPriority.Value = Convert.ToInt32(dt1.Rows[0]["cm_priority"].ToString());
                    nmUnloadingTime.Value = Convert.ToInt32(dt1.Rows[0]["cm_unloading_time"].ToString());
                    // ---- end ----//

                    /** ticket 27918 **/
                    if (__isAsk27918)
                    {
                        if (string.IsNullOrEmpty(dt1.Rows[0]["cm_tax_address_choice"].ToString().Trim()))
                        {
                            if (__isFlagKAM)
                            {
                                cbAddressChoice.ItemIndex = 1;
                            }
                            else
                            {
                                cbAddressChoice.ItemIndex = 0;
                            }
                        }
                        else
                        {
                            cbAddressChoice.EditValue = dt1.Rows[0]["cm_tax_address_choice"].ToString().Trim();
                        }
                    }
                    /** end of ticket 27918 **/
                    if (__isAsk28151)
                    {
                        if (dt1.Rows[0]["cm_fullfilment"].ToString().Trim().Equals("Y"))
                        {
                            cbfullfilment.Checked = true;
                        }
                        else
                        {
                            cbfullfilment.Checked = false;
                        }

                    }


                    if (dt1.Rows[0]["cm_flag_tukar_faktur"].ToString().Trim().Equals("Y"))
                    {
                        cbFakturPajakFlag.Checked = true;
                        cbFakturPajakFlag.Text = "Ya";
                    }
                    else
                    {
                        cbFakturPajakFlag.Checked = false;
                        cbFakturPajakFlag.Text = "Tidak";
                    }


                    /** ticket 30207 **/
                    txtDelvZone.Text = dt1.Rows[0]["cm_destination_site"].ToString().Trim();
                    /** end of ticket 30207 **/

                    /** ship plant **/
                    txtShipPlant.Text = dt1.Rows[0]["cm_ship_plant"].ToString().Trim();
                    /** end ship plant **/


                    /** Long lat new,delivery,eCom **/
                    txtLatNew.Text = dt1.Rows[0]["cm_lat_new"].ToString().Trim();
                    txtLongNew.Text = dt1.Rows[0]["cm_lng_new"].ToString().Trim();
                    txtLatDelv.Text = dt1.Rows[0]["cm_lat_delv"].ToString().Trim();
                    txtLongDelv.Text = dt1.Rows[0]["cm_lng_delv"].ToString().Trim();
                    txtLatEcom.Text = dt1.Rows[0]["cm_lat_ecom"].ToString().Trim();
                    txtLongEcom.Text = dt1.Rows[0]["cm_lng_ecom"].ToString().Trim();
                    /** end long lat new,delivery,eCom **/

                    txtDistChannel.Text = dt1.Rows[0]["cm_distribution_channel"].ToString().ToString();
                    popUpCOA1.EntityId = dt1.Rows[0]["cm_gl_entity"].ToString().Trim();
                    popUpCOA1.BranchId = dt1.Rows[0]["cm_gl_branch"].ToString();
                    popUpCOA1.DivisionId = dt1.Rows[0]["cm_gl_div"].ToString();
                    popUpCOA1.DepartmentId = dt1.Rows[0]["cm_gl_dept"].ToString();
                    popUpCOA1.Major1 = dt1.Rows[0]["cm_gl_major1"].ToString();
                    popUpCOA1.Major2 = dt1.Rows[0]["cm_gl_major2"].ToString();
                    popUpCOA1.Minor = dt1.Rows[0]["cm_gl_minor"].ToString();
                    popUpCOA1.Analisys = dt1.Rows[0]["cm_gl_analisys"].ToString();
                    popUpCOA1.Filler = dt1.Rows[0]["cm_gl_filler"].ToString();

                    filltipePajak(dt1.Rows[0]["cm_hott_tax_codes"].ToString());
                    popUpCOA1.GetFormatCOA(popUpCOA1.EntityId, popUpCOA1.BranchId, popUpCOA1.DivisionId, popUpCOA1.DepartmentId, popUpCOA1.Major1, popUpCOA1.Major2, popUpCOA1.Minor, popUpCOA1.Analisys, popUpCOA1.Filler);

                    //Jenis Identitas
                    if (string.IsNullOrEmpty(dt1.Rows[0]["cm_jenis_identitas"].ToString().Trim()))
                    {
                        cbJenisIdentitas.ItemIndex = 0;
                    }
                    else
                    {
                        cbJenisIdentitas.EditValue = dt1.Rows[0]["cm_jenis_identitas"].ToString().Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void SetShipBillValue(string value)
        {
            value = (value ?? "").Trim();

            CBShipnBill.EditValue = value;

            if (CBShipnBill.ItemIndex < 0)
            {
                int idx = CBShipnBill.Properties.GetDataSourceRowIndex(
                    CBShipnBill.Properties.ValueMember,
                    value
                );

                if (idx >= 0)
                    CBShipnBill.ItemIndex = idx;
            }

            CBShipnBill.RefreshEditValue();
        }
        private void btnWHCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search WH ";
            frm.Query = " select wh_loc_id1 [WH Code 1], wh_loc_id2[WH Code 2], wh_loc_name [Warehouse Name] from IM_WH_LOC where wh_loc_type = 'N'";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtWHCode.Text = frm.ArrField[0].Trim();
                txtWHCodeTo.Text = frm.ArrField[1].Trim();
            }
        }

        private void BTNBillcust_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Bill ";
            //frm.Query = " SELECT cm_cust_code1[Code1], cm_cust_code2 [Code2],        cm_cust_name[Name], cm_delv_address1[Address]  From SO_CUST_MASTER ";
            frm.Query = GetActiveCustomerLookupQuery();
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtBilltoCustCode.Text = frm.ArrField[2].Trim();
                txtBilltoCustCodeTo.Text = frm.ArrField[3].Trim();
                txtBillBranch.Text = frm.ArrField[1].ToString();
            }
        }

        private void FillGridHirarki()
        {
            try
            {
                strSQL = @"
                            SELECT DISTINCT
                                cpl.cpl_line_code,
                                ISNULL(pl.pl_prd_line_desc, '') AS pl_prd_line_desc,
                                cpl.cpl_cust_type,
                                ISNULL(ct.ct_cust_type_desc, '') AS ct_cust_type_desc
                            FROM TBL_SD_CUSTPRDLINE cpl WITH (NOLOCK)
                            LEFT JOIN IM_PRD_LINE pl WITH (NOLOCK)
                                ON cpl.cpl_line_code = pl.pl_prd_line_code
                            LEFT JOIN SO_CUST_TYPE ct WITH (NOLOCK)
                                ON cpl.cpl_cust_type = ct.ct_cust_type
                               AND ct.ct_cluster = 'Y'
                            WHERE cpl.cpl_entity = " + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + @"
                              AND cpl.cpl_branch = " + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + @"
                              AND cpl.cpl_cust_code1 = " + FmtStr(txtCustCode.Text.Trim()) + @"
                              AND cpl.cpl_cust_code2 = " + FmtStr(txtCustCodeTo.Text.Trim()) + @"
                            ORDER BY cpl.cpl_line_code";

                DataTable dt = _clsGlobal.ExecDT(strSQL);
                DGVHIR.DataSource = dt;
                ConfigureHierarchyGridRuntime();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DGVHIR_RowPostPaint(object sender, System.Windows.Forms.DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(Color.Black))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void DGVHIR_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            var senderGrid = sender as DevExpress.XtraGrid.GridControl;
            if (e.RowIndex > -1)
            {
                if (GetFocusedColumnIndex(senderGrid) == DGVHIRView.Columns["btna"].VisibleIndex)
                {
                    frmPopUp frm = new frmPopUp();
                    frm.FrmText = "Search Prdline ";
                    frm.Query = "select pl_prd_line_code [PrdLine], pl_prd_line_desc [Description] from IM_PRD_LINE ";
                    frm.ShowDialog();
                    if (frm.ArrField != null)
                    {
                        try
                        {
                            System.Windows.Forms.DataGridViewRow dgr, dgrB;
                            string maxB = "";
                            string min = "";
                            dgrB = null;

                            dgr = GetGridRows(DGVHIR, DGVHIRView)[GetGridCurrentRow(DGVHIR, DGVHIRView).Index];

                            if (GetGridRowCount(DGVHIR, DGVHIRView) > 1)
                            {
                                if (dgr.Index != 0)
                                {
                                    dgrB = GetGridRows(DGVHIR, DGVHIRView)[GetGridCurrentRow(DGVHIR, DGVHIRView).Index - 1];
                                    maxB = dgrB.Cells["cpl_line_code"].Value.ToString();
                                }
                                else
                                {
                                    min = frm.ArrField[0].Trim();
                                    rowDuplicate = false;
                                    exitData = false;

                                    foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(DGVHIR, DGVHIRView))
                                    {
                                        if ((gRow.Cells["cpl_line_code"].Value.ToString().Trim() == min))
                                        {
                                            rowDuplicate = true;
                                            break;
                                        }
                                    }
                                    if (rowDuplicate == true)
                                    {
                                        ShowValidationError(tabPage5, "Product Line " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Product Line yang berbeda ", MessageBoxIcon.Error);
                                        if (dgr.Cells["cpl_line_code"].Value.ToString() == "")
                                        {
                                            dgr.Cells["cpl_line_code"].Value = "";
                                            dgr.Cells["pl_prd_line_desc"].Value = "";
                                            exitData = true;
                                        }
                                        else
                                        {
                                            dgr.Cells["cpl_line_code"].Value = "";
                                            dgr.Cells["pl_prd_line_desc"].Value = "";
                                            exitData = true;
                                        }
                                    }
                                    else
                                    {
                                        GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["cpl_line_code"].Value = frm.ArrField[0].Trim();
                                        GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["pl_prd_line_desc"].Value = frm.ArrField[1].Trim();
                                    }
                                }
                            }
                            else
                            {
                                dgrB = null;
                            }
                            if (dgrB != null)
                            {
                                min = frm.ArrField[0].Trim();

                                rowDuplicate = false;

                                foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(DGVHIR, DGVHIRView))
                                {
                                    if ((gRow.Cells["cpl_line_code"].Value.ToString().Trim() == min))
                                    {
                                        rowDuplicate = true;
                                        break;
                                    }
                                }

                                if (rowDuplicate == true)
                                {
                                    ShowValidationError(tabPage5, "Product Line " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Product Line yang berbeda ", MessageBoxIcon.Error);
                                    if (dgr.Cells["cpl_line_code"].Value.ToString() == "")
                                    {
                                        dgr.Cells["cpl_line_code"].Value = "";
                                        dgr.Cells["pl_prd_line_desc"].Value = "";
                                    }
                                    else
                                    {
                                        dgr.Cells["cpl_line_code"].Value = "";
                                        dgr.Cells["pl_prd_line_desc"].Value = "";
                                    }
                                }
                                else
                                {
                                    GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["cpl_line_code"].Value = frm.ArrField[0].Trim();
                                    GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["pl_prd_line_desc"].Value = frm.ArrField[1].Trim();
                                }
                            }
                            else
                            {
                                if (exitData != true)
                                {
                                    GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["cpl_line_code"].Value = frm.ArrField[0].Trim();
                                    GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["pl_prd_line_desc"].Value = frm.ArrField[1].Trim();
                                }
                            }

                        }
                        catch (Exception ex) { }
                    }
                }
                if (GetFocusedColumnIndex(senderGrid) == DGVHIRView.Columns["btnb"].VisibleIndex)
                {

                    frmPopUp frm = new frmPopUp();
                    frm.FrmText = "Search Cust ";
                    //frm.Query = "select ct_cust_type [Type] , ct_cust_type_desc [Description] from SO_CUST_TYPE ";
                    frm.Query = "select ct_cust_type [Type] , ct_cust_type_desc [Description] from SO_CUST_TYPE where ct_cluster='Y'";
                    frm.ShowDialog();
                    if (frm.ArrField != null)
                    {
                        GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["cpl_cust_type"].Value = frm.ArrField[0].Trim();
                        GetGridRows(DGVHIR, DGVHIRView)[e.RowIndex].Cells["ct_cust_type_desc"].Value = frm.ArrField[1].Trim();

                    }
                }

            }

        }

        private void DGVHIR_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Insert)
            {
                System.Windows.Forms.DataGridViewRow dgr;
                string column0, column1, column2;
                try
                {
                    dgr = GetGridCurrentRow(DGVHIR, DGVHIRView);

                    if (GetGridRows(DGVHIR, DGVHIRView).Count > 0)
                    {

                        column0 = dgr.Cells["cpl_line_code"].Value.ToString();
                        column1 = dgr.Cells["cpl_cust_type"].Value.ToString();
                        column2 = dgr.Cells["ct_cust_type_desc"].Value.ToString();
                        if (column0 != "" && dgr.Cells["pl_prd_line_desc"].Value.ToString() != "" && column1 != "" && column2 != "")
                        {
                            //if (clsGlobal.MODE_TRX == 1)
                            if (__state == StateEntry.New)
                            {
                                GetGridRows(DGVHIR, DGVHIRView).Add("", "", "", "", "", "");
                            }
                            else
                            {
                                DataTable dt = DGVHIR.DataSource as DataTable;
                                dt.Rows.Add();
                                DGVHIR.DataSource = dt;
                            }
                        }
                    }
                    else
                    {
                        //if (clsGlobal.MODE_TRX == 1)
                        if (__state == StateEntry.New)
                        {
                            GetGridRows(DGVHIR, DGVHIRView).Add("", "", "", "", "", "");
                        }
                        else
                        {
                            if (GetGridRowCount(DGVHIR, DGVHIRView) > 0)
                            {
                                GetGridRows(DGVHIR, DGVHIRView).Add("", "", "", "", "", "");
                            }
                            else
                            {
                                DataTable dt = DGVHIR.DataSource as DataTable;
                                dt.Rows.Add();
                                DGVHIR.DataSource = dt;
                            }
                        }
                    }
                }
                finally
                {
                    dgr = null;
                    column0 = null;
                }
            }
            if (e.KeyCode == Keys.Delete)
            {
                if (GetGridRowCount(DGVHIR, DGVHIRView) != 0)
                {
                    GetGridRows(DGVHIR, DGVHIRView).Remove(GetGridCurrentRow(DGVHIR, DGVHIRView));
                }
            }
        }

        private void FillGridSalesman()
        {
            try
            {
                strSQL = "";
                strSQL = "select csc_salesman_id, isnull(sgm_spgm_name, '') sgm_spgm_name, isnull(sgm_type_operasi, '') sgm_type_operasi, " +
                        "Case csc_visit WHEN '1' THEN 'Senin' WHEN '2' THEN 'Selasa' WHEN '3' THEN 'Rabu' WHEN '4' THEN 'Kamis' WHEN '5' THEN 'Jumat' WHEN '6' THEN 'Sabtu' WHEN '7' THEN 'Minggu' END csc_visit, " +
                        "(case csc_visit_week1 when 'T' then 'false' else 'True' end) csc_visit_week1, " +
                        "(case csc_visit_week2 when 'T' then 'false' else 'True' end) csc_visit_week2, " +
                        "(case csc_visit_week3 when 'T' then 'false' else 'True' end) csc_visit_week3, " +
                        "(case csc_visit_week4 when 'T' then 'false' else 'True' end) csc_visit_week4, csc_route " +
                        " From TBL_SD_CUSTCOVER " +
                        "LEFT OUTER JOIN SO_SPG_GIRL_MAN ON sgm_spgm_id = csc_salesman_id and sgm_entity_id = csc_entity and sgm_branch_id = csc_branch " +
                        "where csc_cust_code1 = '" + _prdCustCode1 + "' and csc_cust_code2 = '" + _Grade + "' and sgm_entity_id='" + _prdEntityCode + "' and sgm_branch_id='" + _prdBrandCode + "' order by csc_index ASC ";
                DataTable dt = new DataTable();

                dt.Columns.Add("csc_salesman_id");
                dt.Columns.Add("sgm_spgm_name");
                dt.Columns.Add("sgm_type_operasi");
                dt.Columns.Add("csc_visit");
                dt.Columns.Add("csc_visit_week1");
                dt.Columns.Add("csc_visit_week2");
                dt.Columns.Add("csc_visit_week3");
                dt.Columns.Add("csc_visit_week4");
                dt.Columns.Add("csc_route");
                dt.AcceptChanges();

                dt = _clsGlobal.ExecDT(strSQL);
                if (!dt.Columns.Contains("btnc"))
                    dt.Columns.Add("btnc", typeof(string));



                DVGSalesman.DataSource = dt;
                ConfigureSalesmanCoveredGridRuntime();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FillGridGroupHarga()
        {
            try
            {
                strSQL = @"
                            SELECT DISTINCT
                                cp.cp_price_group AS grp_code,
                                ISNULL(pl.ppl_price_list_description, '') AS grp_desc
                            FROM SO_CUST_MASTER_PRICE_GROUP cp WITH (NOLOCK)
                            LEFT JOIN OB_PRICE_LIST_HEADER pl WITH (NOLOCK)
                                ON pl.ppl_price_list_code = cp.cp_price_group
                            WHERE cp.cp_cust_code1 = " + FmtStr(txtCustCode.Text.Trim()) + @"
                              AND cp.cp_cust_code2 = " + FmtStr(txtCustCodeTo.Text.Trim()) + @"
                              AND cp.cp_entity = " + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + @"
                              AND cp.cp_branch = " + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + @"
                            ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                dgvGroupHarga.DataSource = dt;
                ConfigureGroupHargaGridRuntime();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FillTopByPrdLine()
        {
            StringBuilder __query = new StringBuilder();
            __query.AppendLine("SELECT");
            __query.AppendLine("ROW_NUMBER() OVER(");
            __query.AppendLine("ORDER BY MPL.smt_entity_id");
            __query.AppendLine(",MPL.smt_branch_id");
            __query.AppendLine(",MPL.smt_cust_code1");
            __query.AppendLine(",MPL.smt_cust_code2) AS no");
            __query.AppendLine(",MPL.smt_prdline_id AS prdline_code");
            __query.AppendLine(",ISNULL(IPL.pl_prd_line_desc, '') AS prdline_desc");
            __query.AppendLine(",MPL.smt_top_id AS top_code");
            __query.AppendLine(",PTC.pptc_term_desc AS top_desc");
            __query.AppendLine(",ISNULL(MPL.smt_plant_id, '') AS plant_id");
            __query.AppendLine(",'' AS plant_desc");
            __query.AppendLine(",ISNULL(MPL.smt_sales_employee_id, '') AS sales_employee_id");
            __query.AppendLine(",'' AS sales_employee_desc");
            __query.AppendLine("FROM SO_MAPPING_TOPBYPRDLINE MPL WITH(NOLOCK)");
            __query.AppendLine("OUTER APPLY (");
            __query.AppendLine("    SELECT TOP 1 pl_prd_line_desc");
            __query.AppendLine("    FROM IM_PRD_LINE IPLX WITH(NOLOCK)");
            __query.AppendLine("    WHERE IPLX.pl_sapdiv_id = MPL.smt_prdline_id");
            __query.AppendLine("       OR IPLX.pl_prd_line_code = MPL.smt_prdline_id");
            __query.AppendLine("    ORDER BY CASE WHEN IPLX.pl_sapdiv_id = MPL.smt_prdline_id THEN 0 ELSE 1 END, IPLX.pl_prd_line_code");
            __query.AppendLine(") IPL");
            __query.AppendLine("INNER JOIN PO_PAYMENT_TERM_CODES PTC WITH(NOLOCK)");
            __query.AppendLine("ON PTC.pptc_term_code=MPL.smt_top_id");
            __query.AppendFormat("WHERE smt_entity_id='{0}'\r\n", ctrlEntityCustMaster2.txtCM.Text.Trim());
            __query.AppendFormat("AND smt_branch_id='{0}'\r\n", ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim());
            __query.AppendFormat("AND smt_cust_code1='{0}'\r\n", txtCustCode.Text.Trim());
            __query.AppendFormat("AND smt_cust_code2='{0}'\r\n", txtCustCodeTo.Text.Trim());

            //DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
            DataTable __source = _clsGlobal.ExecDT(__query.ToString());
            topbyprdlinedgv.DataSource = __source;
            ConfigureTopByPrdLineGridRuntime();
            __query.Clear();
        }

        private void FillCustSchPay()
        {
            StringBuilder __query = new StringBuilder();
            DataTable __source = new DataTable();

            __query.AppendLine("SELECT");
            __query.AppendLine(" csp_date AS tgl");
            __query.AppendLine(" ,csp_desc AS ket");
            __query.AppendLine(" FROM SO_CUST_SCHPAY WITH(NOLOCK)");
            __query.AppendLine(" INNER JOIN SO_CUST_MASTER WITH(NOLOCK)");
            __query.AppendLine(" ON cm_entity=csp_entity_id");
            __query.AppendLine(" AND cm_branch=csp_branch_id");
            __query.AppendLine(" AND cm_cust_code1=csp_cust_code1");
            __query.AppendLine(" AND cm_cust_code2=csp_cust_code2");
            __query.AppendFormat(" WHERE csp_entity_id='{0}'\r\n", ctrlEntityCustMaster2.txtCM.Text.Trim());
            __query.AppendFormat(" AND csp_branch_id='{0}'\r\n", ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim());
            __query.AppendFormat(" AND csp_cust_code1='{0}'\r\n", txtCustCode.Text.Trim());
            __query.AppendFormat(" AND csp_cust_code2='{0}'\r\n", txtCustCodeTo.Text.Trim());
            __query.AppendLine(" AND csp_type='D'");

            __source = _clsGlobal.ExecDT(__query.ToString());
            dgvschpayD.DataSource = __source;

            __query.Clear();
            __query.AppendLine("SELECT");
            __query.AppendLine(" CASE WHEN csp_day = 1 THEN 'Senin' WHEN csp_day = 2 THEN 'Selasa' WHEN csp_day = 3 THEN 'Rabu'");
            __query.AppendLine(" WHEN csp_day = 4 THEN 'Kamis' WHEN csp_day = 5 THEN 'Jumat' WHEN csp_day = 6 THEN 'Sabtu'");
            __query.AppendLine(" ELSE 'Minggu' END AS hari");
            __query.AppendLine(" ,CASE WHEN csp_week = 1 THEN 'True' ELSE 'False' END AS pola1");
            __query.AppendLine(" ,CASE WHEN csp_week = 2 THEN 'True' ELSE 'False' END AS pola2");
            __query.AppendLine(" ,CASE WHEN csp_week = 3 THEN 'True' ELSE 'False' END AS pola3");
            __query.AppendLine(" ,CASE WHEN csp_week = 4 THEN 'True' ELSE 'False' END AS pola4");
            __query.AppendLine(" ,csp_desc AS ket");
            __query.AppendLine(" FROM SO_CUST_SCHPAY WITH(NOLOCK)");
            __query.AppendLine(" INNER JOIN SO_CUST_MASTER WITH(NOLOCK)");
            __query.AppendLine(" ON cm_entity=csp_entity_id");
            __query.AppendLine(" AND cm_branch=csp_branch_id");
            __query.AppendLine(" AND cm_cust_code1=csp_cust_code1");
            __query.AppendLine(" AND cm_cust_code2=csp_cust_code2");
            __query.AppendFormat(" WHERE csp_entity_id='{0}'\r\n", ctrlEntityCustMaster2.txtCM.Text.Trim());
            __query.AppendFormat(" AND csp_branch_id='{0}'\r\n", ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim());
            __query.AppendFormat(" AND csp_cust_code1='{0}'\r\n", txtCustCode.Text.Trim());
            __query.AppendFormat(" AND csp_cust_code2='{0}'\r\n", txtCustCodeTo.Text.Trim());
            __query.AppendLine(" AND csp_type='W'");

            __source = _clsGlobal.ExecDT(__query.ToString());
            dgvschpayW.DataSource = __source;

            if (GetGridRowCount(dgvschpayW, dgvschpayWView) > 0)
            {
                chkWeekly.Checked = true;
            }
            else
            {
                chkDaily.Checked = true;
            }
        }

        private void FillPrdLineBlocking()
        {
            StringBuilder __query = new StringBuilder();
            DataTable __source = new DataTable();

            __query.AppendLine("SELECT");
            __query.AppendLine("ROW_NUMBER() OVER(");
            __query.AppendLine("ORDER BY scbp_entity_id");
            __query.AppendLine(" ,scbp_branch_id");
            __query.AppendLine(" ,scbp_cust_code1");
            __query.AppendLine(" ,scbp_cust_code2) AS no");
            __query.AppendLine(" ,scbp_prd_line_code AS prdlinecode");
            __query.AppendLine(" ,pl_prd_line_desc AS prdlinedesc");
            __query.AppendLine(" ,case when scbp_sales_block = 'Y' then 'true' else 'false' end chksales");
            __query.AppendLine(" ,case when scbp_retur_block = 'Y' then 'true' else 'false' end chkretur");
            __query.AppendLine(" FROM SO_CUST_BLOCK_PRDLINE WITH(NOLOCK)");
            __query.AppendLine(" INNER JOIN SO_CUST_MASTER WITH(NOLOCK)");
            __query.AppendLine(" ON cm_entity=scbp_entity_id");
            __query.AppendLine(" AND cm_branch=scbp_branch_id");
            __query.AppendLine(" AND cm_cust_code1=scbp_cust_code1");
            __query.AppendLine(" AND cm_cust_code2=scbp_cust_code2");
            __query.AppendLine(" INNER JOIN IM_PRD_LINE WITH(NOLOCK)");
            __query.AppendLine(" ON pl_prd_line_code=scbp_prd_line_code");
            __query.AppendFormat(" WHERE scbp_entity_id='{0}'\r\n", ctrlEntityCustMaster2.txtCM.Text.Trim());
            __query.AppendFormat(" AND scbp_branch_id='{0}'\r\n", ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim());
            __query.AppendFormat(" AND scbp_cust_code1='{0}'\r\n", txtCustCode.Text.Trim());
            __query.AppendFormat(" AND scbp_cust_code2='{0}'\r\n", txtCustCodeTo.Text.Trim());

            __source = _clsGlobal.ExecDT(__query.ToString());
            dgvPrdLineBlocking.DataSource = __source;
        }

        private void DVGSalesman_Paint(object sender, PaintEventArgs e)
        {
            // DevExpress GridControl tidak memakai header painting DataGridView.
            // Header grouping sebaiknya dibuat via BandedGridView jika diperlukan.
            // Method dipertahankan supaya event lama tetap aman, tetapi logic bisnis tidak berubah.
        }

        private void DVGSalesman_RowPostPaint(object sender, System.Windows.Forms.DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(Color.Black))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void DVGSalesman_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                var senderGrid = sender as DevExpress.XtraGrid.GridControl;
                if (e.RowIndex > -1)
                {
                    if (GetFocusedColumnIndex(senderGrid) == DVGSalesmanView.Columns["btnc"].VisibleIndex)
                    {
                        frmPopUp frm = new frmPopUp();
                        frm.FrmText = "Search Salesman ";

                        //if (clsGlobal.MODE_TRX == 2)
                        if (__state == StateEntry.Edit)
                        {
                            frm.Query = "SELECT sgm_spgm_id as [csc_salesman_id],sgm_spgm_name as [sgm_spgm_name],sgm_type_operasi as [sgm_type_operasi] FROM SO_SPG_GIRL_MAN WHERE sgm_entity_id='" + _prdEntityCode + "' and sgm_branch_id='" + _prdBrandCode + "'";
                        }
                        //if (clsGlobal.MODE_TRX == 1)
                        if (__state == StateEntry.New)
                        {
                            frm.Query = "SELECT sgm_spgm_id as [csc_salesman_id],sgm_spgm_name as [sgm_spgm_name],sgm_type_operasi as [sgm_type_operasi] FROM SO_SPG_GIRL_MAN WHERE sgm_entity_id='" + ctrlEntityCustMaster2.txtCM.Text.Trim() + "' and sgm_branch_id='" + ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim() + "'";


                        }


                        //frm.Query = "select sgm_spgm_id as [csc_salesman_id], sgm_spgm_name as [sgm_spgm_name], gh_function_code as [sgm_type_operasi] " +
                        //"From SO_SPG_GIRL_MAN WITH(NOLOCK) LEFT OUTER JOIN (select gh_sequence_no, gh_function_code, gh_function_desc " +
                        //"From GS_GEN_HARDCODED WITH(NOLOCK) where gh_function_name ='TYPEOPERASISLS') TYPE ON gh_function_code = sgm_type_operasi " +
                        //"WHERE sgm_spgm_status = '2' ";

                        frm.ShowDialog();
                        if (frm.ArrField != null)
                        {
                            try
                            {
                                System.Windows.Forms.DataGridViewRow dgr, dgrB;
                                string maxB = "";
                                string min = "";
                                dgrB = null;
                                dgr = null;

                                dgr = GetGridRows(DVGSalesman, DVGSalesmanView)[GetGridCurrentRow(DVGSalesman, DVGSalesmanView).Index];

                                if (GetGridRowCount(DVGSalesman, DVGSalesmanView) > 1)
                                {
                                    if (dgr.Index != 0)
                                    {
                                        dgrB = GetGridRows(DVGSalesman, DVGSalesmanView)[GetGridCurrentRow(DVGSalesman, DVGSalesmanView).Index - 1];
                                        maxB = dgrB.Cells["csc_salesman_id"].Value.ToString();
                                    }
                                    else
                                    {
                                        min = frm.ArrField[0].Trim();
                                        rowDuplicate = false;
                                        exitData = false;

                                        foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(DVGSalesman, DVGSalesmanView))
                                        {
                                            if ((gRow.Cells["csc_salesman_id"].Value.ToString().Trim() == min))
                                            {
                                                rowDuplicate = true;
                                                break;
                                            }
                                        }

                                        if (rowDuplicate == true)
                                        {
                                            ShowValidationError(tabPage6, "Salesman Sudah ada ", MessageBoxIcon.Error);
                                            if (dgr.Cells["csc_salesman_id"].Value.ToString() == "")
                                            {
                                                dgr.Cells["csc_salesman_id"].Value = "";
                                                dgr.Cells["sgm_spgm_name"].Value = "";
                                                dgr.Cells["sgm_type_operasi"].Value = "";
                                                exitData = true;
                                            }
                                            else
                                            {
                                                dgr.Cells["csc_salesman_id"].Value = "";
                                                dgr.Cells["sgm_spgm_name"].Value = "";
                                                dgr.Cells["sgm_type_operasi"].Value = "";
                                                exitData = true;
                                            }
                                        }
                                        else
                                        {
                                            //DataTable dt = DVGSalesman.DataSource as DataTable;
                                            //dt.Rows[e.RowIndex]["csc_salesman_id"] = frm.ArrField[0].Trim();
                                            //dt.Rows[e.RowIndex]["sgm_spgm_name"] = frm.ArrField[1].Trim();
                                            //dt.Rows[e.RowIndex]["sgm_type_operasi"] = frm.ArrField[2].Trim();
                                            GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["csc_salesman_id"].Value = frm.ArrField[0].Trim();
                                            GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["sgm_spgm_name"].Value = frm.ArrField[1].Trim();
                                            GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["sgm_type_operasi"].Value = frm.ArrField[2].Trim();
                                        }
                                    }
                                }
                                else
                                {
                                    dgrB = null;
                                }
                                if (dgrB != null)
                                {
                                    min = frm.ArrField[0].Trim();
                                    rowDuplicate = false;

                                    foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(DVGSalesman, DVGSalesmanView))
                                    {
                                        if ((gRow.Cells["csc_salesman_id"].Value.ToString().Trim() == min))
                                        {
                                            rowDuplicate = true;
                                            break;
                                        }
                                    }

                                    if (rowDuplicate == true)
                                    {
                                        ShowValidationError(tabPage6, "Salesman Sudah ada ", MessageBoxIcon.Error);
                                        if (dgr.Cells["csc_salesman_id"].Value.ToString() == "")
                                        {
                                            dgr.Cells["csc_salesman_id"].Value = "";
                                            dgr.Cells["sgm_spgm_name"].Value = "";
                                            dgr.Cells["sgm_type_operasi"].Value = "";
                                        }
                                        else
                                        {
                                            dgr.Cells["csc_salesman_id"].Value = "";
                                            dgr.Cells["sgm_spgm_name"].Value = "";
                                            dgr.Cells["sgm_type_operasi"].Value = "";
                                        }
                                    }
                                    else
                                    {
                                        GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["csc_salesman_id"].Value = frm.ArrField[0].Trim();
                                        GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["sgm_spgm_name"].Value = frm.ArrField[1].Trim();
                                        GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["sgm_type_operasi"].Value = frm.ArrField[2].Trim();
                                    }
                                }
                                else
                                {
                                    if (exitData != true)
                                    {
                                        GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["csc_salesman_id"].Value = frm.ArrField[0].Trim();
                                        GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["sgm_spgm_name"].Value = frm.ArrField[1].Trim();
                                        GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["sgm_type_operasi"].Value = frm.ArrField[2].Trim();
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void CBBlacklist_CheckedChanged(object sender, EventArgs e)
        {
            if (CBBlacklist.Checked)
            {
                CBBlacklist.Text = "BlackList";
                txtReason.Visible = true;
                label18.Visible = true;
            }
            else
            {
                txtReason.Visible = true;
                label18.Visible = true;
                CBBlacklist.Text = "UnBlackList";
                txtReason.Text = "";
            }
        }

        private void DVGSalesman_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Insert)
            {
                System.Windows.Forms.DataGridViewRow dgr;
                string column0, column1, column2;
                try
                {
                    dgr = GetGridCurrentRow(DVGSalesman, DVGSalesmanView);

                    if (GetGridRows(DVGSalesman, DVGSalesmanView).Count > 0)
                    {

                        column0 = dgr.Cells["csc_salesman_id"].Value.ToString();
                        column1 = dgr.Cells["sgm_type_operasi"].Value.ToString();
                        column2 = dgr.Cells["csc_visit"].Value.ToString();

                        if (column0 != "" || dgr.Cells["sgm_spgm_name"].Value.ToString() != "" || column1 != "" || column2 != "")
                        {

                            DataTable dt = DVGSalesman.DataSource as DataTable;
                            dt.Rows.Add();
                            DVGSalesman.DataSource = dt;

                        }
                    }
                    else
                    {
                        DataTable dt = DVGSalesman.DataSource as DataTable;
                        dt.Rows.Add();
                        DVGSalesman.DataSource = dt;

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    dgr = null;
                    column0 = null;
                }
            }
            if (e.KeyCode == Keys.Delete)
            {
                if (GetGridRowCount(DVGSalesman, DVGSalesmanView) != 0)
                {
                    GetGridRows(DVGSalesman, DVGSalesmanView).Remove(GetGridCurrentRow(DVGSalesman, DVGSalesmanView));
                }
            }
        }

        private void InitializeDxValidation()
        {
            if (dxValidator != null) return;

            dxValidator = this.dxValidationProvider1;
            if (dxValidator == null)
                dxValidator = new DXValidationProvider(this.components);

            dxValidator.ValidationMode = ValidationMode.Manual;
            RegisterCustomerValidationRules();
        }

        private void RegisterCustomerValidationRules()
        {
            _registeredValidationControls.Clear();
            _validationTabMap.Clear();
            _validationRuleMap.Clear();

            // Header / data utama
            AddRequiredRule(txtCustCode, null, "Customer Code wajib diisi.");
            AddRequiredRule(txtCustCodeTo, null, "Customer Code To wajib diisi.");
            AddRequiredRule(txtCustName, null, "Customer Name wajib diisi.");
            AddRequiredRule(txtShortName, null, "Short Name wajib diisi.");
            AddRequiredRule(txtLeadtime, null, "Lead Time wajib diisi.");
            AddRequiredRule(txtBilltoCustCode, null, "Bill to Cust Code wajib diisi.");

            if (ctrlEntityCustMaster2 != null)
            {
                AddRequiredRule(ctrlEntityCustMaster2.txtCM, null, "Entity wajib diisi.");
                AddRequiredRule(ctrlEntityCustMaster2.txtBranchIdCM, null, "Branch wajib diisi.");
            }

            if (ctrlDistrikCustMaster1 != null)
            {
                AddRequiredRule(ctrlDistrikCustMaster1.txtDistrikCM, null, "Distrik wajib diisi.");
                AddRequiredRule(ctrlDistrikCustMaster1.txtBeatCM, null, "Beat wajib diisi.");
                AddRequiredRule(ctrlDistrikCustMaster1.txtSubBeatCM, null, "Sub Beat wajib diisi.");
            }

            if (ctrlGroupOutlet1 != null)
                AddRequiredRule(ctrlGroupOutlet1.txtGroupCM, null, "Group Outlet wajib diisi.");

            if (ctrlpopUpTypeOutletCustMaster1 != null)
                AddRequiredRule(ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet, null, "Type Outlet wajib diisi.");

            // Customer Attribute
            if (ctrlGroupDiscount1 != null)
                AddRequiredRule(ctrlGroupDiscount1.txtGpDiscountT1, tabPage1, "Group Discount wajib diisi.");

            AddRequiredRule(txtIndustriT1, tabPage1, "Industri wajib diisi.");

            if (ctrlKodepasarDesc != null)
                AddRequiredRule(ctrlKodepasarDesc.txtKodePasarT1, tabPage1, "Pasar wajib diisi.");

            AddRequiredRule(CBBagunanT1, tabPage1, "Bangunan wajib dipilih.");

            // Cust Property Tax
            AddMinLengthRequiredRule(txtAddress1T2, tabPage2, 20, "Address 1 Cust Property wajib diisi minimal 20 karakter.");
            AddRequiredRule(cbJenisIdentitas, tabPage2, "Jenis Identitas wajib dipilih.");

            // Billing Address
            AddMinLengthRequiredRule(txtAddress1T3, tabPage3, 20, "Address 1 Billing Address wajib diisi minimal 20 karakter.");

            // Delivery Address
            AddMinLengthRequiredRule(txtAddress1T4, tabPage4, 20, "Address 1 Delivery Address wajib diisi minimal 20 karakter.");

            if (popUpProvinsinCity1 != null)
            {
                AddRequiredRule(popUpProvinsinCity1.txtPropinsiCMC, tabPage4, "Provinsi wajib diisi.");
                AddRequiredRule(popUpProvinsinCity1.txtKabupatenCMC, tabPage4, "City wajib diisi.");
                AddRequiredRule(popUpProvinsinCity1.txtKecamatan, tabPage4, "Kecamatan wajib diisi.");
                AddRequiredRule(popUpProvinsinCity1.txtKelurahan, tabPage4, "Kelurahan wajib diisi.");
            }
        }

        private void AddRequiredRule(Control control, XtraTabPage tabPage, string errorText)
        {
            if (control == null || dxValidator == null) return;

            RequiredControlRule rule = new RequiredControlRule();
            rule.ErrorText = errorText;
            rule.ErrorType = ErrorType.Critical;
            RegisterValidationRule(control, tabPage, rule);
        }

        private void AddMinLengthRequiredRule(Control control, XtraTabPage tabPage, int minLength, string errorText)
        {
            if (control == null || dxValidator == null) return;

            RequiredMinLengthControlRule rule = new RequiredMinLengthControlRule(minLength);
            rule.ErrorText = errorText;
            rule.ErrorType = ErrorType.Critical;
            RegisterValidationRule(control, tabPage, rule);
        }

        private void RegisterValidationRule(Control control, XtraTabPage tabPage, ValidationRule rule)
        {
            dxValidator.SetValidationRule(control, rule);
            _validationRuleMap[control] = rule;

            if (!_registeredValidationControls.Contains(control))
                _registeredValidationControls.Add(control);

            if (tabPage != null)
                _validationTabMap[control] = tabPage;

            HookValidationRefreshEvent(control);
        }

        private void HookValidationRefreshEvent(Control control)
        {
            BaseEdit baseEdit = control as BaseEdit;
            if (baseEdit != null)
            {
                baseEdit.EditValueChanged -= ValidationControl_ValueChanged;
                baseEdit.EditValueChanged += ValidationControl_ValueChanged;
                return;
            }

            control.TextChanged -= ValidationControl_ValueChanged;
            control.TextChanged += ValidationControl_ValueChanged;
        }

        private void ValidationControl_ValueChanged(object sender, EventArgs e)
        {
            if (_isRefreshingValidationState) return;

            Control control = sender as Control;
            if (control == null || dxValidator == null) return;
            if (!_registeredValidationControls.Contains(control)) return;

            ValidateSingleControlAndUpdateUi(control);
            RefreshValidationTabColors(false);
        }

        private bool ValidateCustomerForm()
        {
            if (dxValidator == null)
                InitializeDxValidation();

            bool isValid = true;
            ResetValidationTabColors();

            foreach (Control control in _registeredValidationControls)
            {
                if (control == null) continue;
                if (!IsControlValidationActive(control))
                {
                    ClearControlInvalid(control);
                    continue;
                }

                if (!ValidateSingleControlAndUpdateUi(control))
                    isValid = false;
            }

            RefreshValidationTabColors(false);

            if (!isValid)
            {
                XtraTabPage firstInvalidTab = GetFirstInvalidTab();
                if (firstInvalidTab != null)
                    tabControl1.SelectedTabPage = firstInvalidTab;

                MessageBox.Show(
                    "Masih ada data yang wajib diisi atau belum sesuai. Silakan cek field berwarna merah dan tab berwarna merah.",
                    clsGlobal.APP_MSG_CAPTION,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }

            return isValid;
        }

        private bool ValidateSingleControlAndUpdateUi(Control control)
        {
            if (control == null || dxValidator == null) return true;
            if (!IsControlValidationActive(control))
            {
                ClearControlInvalid(control);
                return true;
            }

            bool controlValid = dxValidator.Validate(control);
            if (controlValid)
            {
                ClearControlInvalid(control);
            }
            else
            {
                MarkControlInvalid(control);
                XtraTabPage tabPage;
                if (_validationTabMap.TryGetValue(control, out tabPage))
                    MarkTabInvalid(tabPage);
            }

            return controlValid;
        }

        private bool IsControlValidationActive(Control control)
        {
            if (control == null) return false;

            Control current = control;
            while (current != null)
            {
                if (!current.Visible || !current.Enabled)
                    return false;

                current = current.Parent;
            }

            return true;
        }

        private void MarkControlInvalid(Control control)
        {
            if (control == null) return;

            BaseEdit baseEdit = control as BaseEdit;
            if (baseEdit != null)
            {
                if (!_validationOriginalBaseEditBackColorMap.ContainsKey(control))
                    _validationOriginalBaseEditBackColorMap[control] = baseEdit.Properties.Appearance.BackColor;

                baseEdit.Properties.Appearance.BackColor = Color.MistyRose;
                baseEdit.Properties.Appearance.Options.UseBackColor = true;
                MarkPopupHostInvalid(control);
                return;
            }

            if (!_validationOriginalBackColorMap.ContainsKey(control))
                _validationOriginalBackColorMap[control] = control.BackColor;

            control.BackColor = Color.MistyRose;
            MarkPopupHostInvalid(control);
        }

        private void ClearControlInvalid(Control control)
        {
            if (control == null) return;

            BaseEdit baseEdit = control as BaseEdit;
            if (baseEdit != null)
            {
                if (_validationOriginalBaseEditBackColorMap.ContainsKey(control))
                    baseEdit.Properties.Appearance.BackColor = _validationOriginalBaseEditBackColorMap[control];

                baseEdit.Properties.Appearance.Options.UseBackColor = false;
                ClearPopupHostInvalid(control);
                return;
            }

            if (_validationOriginalBackColorMap.ContainsKey(control))
                control.BackColor = _validationOriginalBackColorMap[control];

            ClearPopupHostInvalid(control);
        }

        private void MarkPopupHostInvalid(Control control)
        {
            Control host = FindValidationPopupHost(control);
            if (host == null) return;

            if (!_validationOriginalBackColorMap.ContainsKey(host))
                _validationOriginalBackColorMap[host] = host.BackColor;

            host.BackColor = Color.MistyRose;
        }

        private void ClearPopupHostInvalid(Control control)
        {
            Control host = FindValidationPopupHost(control);
            if (host == null) return;
            if (HasInvalidRegisteredChildInHost(host, control)) return;

            if (_validationOriginalBackColorMap.ContainsKey(host))
                host.BackColor = _validationOriginalBackColorMap[host];
        }

        private bool HasInvalidRegisteredChildInHost(Control host, Control excludedControl)
        {
            if (host == null) return false;

            foreach (Control registeredControl in _registeredValidationControls)
            {
                if (registeredControl == null || object.ReferenceEquals(registeredControl, excludedControl)) continue;
                if (!IsControlValidationActive(registeredControl)) continue;
                if (!object.ReferenceEquals(FindValidationPopupHost(registeredControl), host)) continue;
                if (!EvaluateValidationRuleOnly(registeredControl))
                    return true;
            }

            return false;
        }

        private Control FindValidationPopupHost(Control control)
        {
            if (control == null) return null;

            Control current = control.Parent;
            while (current != null)
            {
                if (current is XtraTabPage || current is Form)
                    return null;

                if (IsValidationPopupHost(current))
                    return current;

                current = current.Parent;
            }

            return null;
        }

        private static bool IsValidationPopupHost(Control control)
        {
            if (control == null) return false;

            string name = control.Name ?? string.Empty;
            string typeName = control.GetType().Name ?? string.Empty;

            return name.StartsWith("ctrl", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("popUp", StringComparison.OrdinalIgnoreCase)
                || typeName.IndexOf("PopUp", StringComparison.OrdinalIgnoreCase) >= 0
                || typeName.IndexOf("Popup", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void ResetValidationTabColors()
        {
            ResetTabHeader(tabPage1);
            ResetTabHeader(partnerFunctionPage);
            ResetTabHeader(tabPage2);
            ResetTabHeader(tabPage3);
            ResetTabHeader(tabPage4);
            ResetTabHeader(tabPage5);
            ResetTabHeader(tabPage6);
            ResetTabHeader(tabPage7);
            ResetTabHeader(tabPage8);
            ResetTabHeader(topbyprdline_page);
            ResetTabHeader(tabPage9);
            ResetTabHeader(tabPage10);
            ResetTabHeader(tabPage11);
        }

        private void RefreshValidationTabColors(bool validateControls)
        {
            if (_isRefreshingValidationState) return;

            _isRefreshingValidationState = true;
            try
            {
                ResetValidationTabColors();

                foreach (Control control in _registeredValidationControls)
                {
                    if (control == null) continue;
                    if (!IsControlValidationActive(control)) continue;

                    bool controlValid;
                    if (validateControls)
                        controlValid = dxValidator.Validate(control);
                    else
                        controlValid = EvaluateValidationRuleOnly(control);

                    if (!controlValid)
                    {
                        XtraTabPage tabPage;
                        if (_validationTabMap.TryGetValue(control, out tabPage))
                            MarkTabInvalid(tabPage);
                    }
                }
            }
            finally
            {
                _isRefreshingValidationState = false;
            }
        }

        private bool EvaluateValidationRuleOnly(Control control)
        {
            if (control == null || dxValidator == null) return true;

            ValidationRule rule;
            if (!_validationRuleMap.TryGetValue(control, out rule) || rule == null) return true;

            BaseEdit baseEdit = control as BaseEdit;
            object value = baseEdit != null ? baseEdit.EditValue : (object)control.Text;
            return rule.Validate(control, value);
        }

        private void MarkTabInvalid(XtraTabPage tabPage)
        {
            if (tabPage == null) return;

            tabPage.Appearance.Header.ForeColor = Color.Red;
            tabPage.Appearance.Header.Font = new Font(tabControl1.Font, FontStyle.Bold);
            tabPage.Appearance.Header.Options.UseForeColor = true;
            tabPage.Appearance.Header.Options.UseFont = true;
        }

        private void MarkInvalidTab(XtraTabPage tabPage)
        {
            if (tabPage == null) return;

            MarkTabInvalid(tabPage);
            if (tabControl1 != null)
                tabControl1.SelectedTabPage = tabPage;
        }

        private void MarkInvalidTabForControl(Control control)
        {
            if (control == null) return;

            XtraTabPage mappedTab;
            if (_validationTabMap.TryGetValue(control, out mappedTab))
            {
                MarkInvalidTab(mappedTab);
                return;
            }

            Control current = control;
            while (current != null)
            {
                XtraTabPage tabPage = current as XtraTabPage;
                if (tabPage != null)
                {
                    MarkInvalidTab(tabPage);
                    return;
                }

                current = current.Parent;
            }
        }

        private void ShowValidationError(Control control, string message, MessageBoxIcon icon)
        {
            if (control != null)
            {
                MarkControlInvalid(control);
                MarkInvalidTabForControl(control);
            }

            MessageBox.Show(message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, icon);

            if (control != null && control.CanSelect)
                control.Select();
        }

        private void ShowValidationError(Control control, string message)
        {
            ShowValidationError(control, message, MessageBoxIcon.Exclamation);
        }

        private void ShowValidationError(XtraTabPage tabPage, string message, MessageBoxIcon icon)
        {
            MarkInvalidTab(tabPage);
            MessageBox.Show(message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, icon);
        }

        private void ResetTabHeader(XtraTabPage tabPage)
        {
            if (tabPage == null) return;

            tabPage.Appearance.Header.Options.UseForeColor = false;
            tabPage.Appearance.Header.Options.UseFont = false;
        }

        private XtraTabPage GetFirstInvalidTab()
        {
            foreach (Control control in _registeredValidationControls)
            {
                if (control == null) continue;
                if (!IsControlValidationActive(control)) continue;

                if (!EvaluateValidationRuleOnly(control))
                {
                    XtraTabPage tabPage;
                    if (_validationTabMap.TryGetValue(control, out tabPage))
                        return tabPage;
                }
            }

            return null;
        }

        private class RequiredControlRule : ValidationRule
        {
            public override bool Validate(Control control, object value)
            {
                return eARCustMasterEntry.HasValue(control, value);
            }
        }

        private class RequiredMinLengthControlRule : ValidationRule
        {
            private readonly int _minLength;

            public RequiredMinLengthControlRule(int minLength)
            {
                _minLength = minLength;
            }

            public override bool Validate(Control control, object value)
            {
                if (!eARCustMasterEntry.HasValue(control, value)) return false;

                string text = Convert.ToString(value);
                text = text == null ? string.Empty : text.Trim();
                return text.Length >= _minLength;
            }
        }

        private static bool HasValue(Control control, object value)
        {
            if (control == null) return false;

            BaseEdit baseEdit = control as BaseEdit;
            if (baseEdit != null)
            {
                object editValue = baseEdit.EditValue;
                if (editValue == null || editValue == DBNull.Value) return false;

                string text = Convert.ToString(editValue);
                text = text == null ? string.Empty : text.Trim();

                if (text.Length == 0) return false;
                if (text.Equals("[EditValue is null]", StringComparison.OrdinalIgnoreCase)) return false;

                return true;
            }

            string plainText = control.Text;
            plainText = plainText == null ? string.Empty : plainText.Trim();
            return plainText.Length > 0;
        }

        private string FmtStr(string value_str)
        {
            if (value_str == null) value_str = string.Empty;

            string formatStr;
            formatStr = "'" + value_str.Trim().Replace("'", "''") + "'";

            return formatStr;
        }
        private string GetEditValueText(DevExpress.XtraEditors.BaseEdit editor)
        {
            if (editor == null) return string.Empty;

            object value = editor.EditValue;
            if (value == null || value == DBNull.Value) return string.Empty;

            string text = Convert.ToString(value);
            return text == null ? string.Empty : text.Trim();
        }

        private long GetEditValueInt64(DevExpress.XtraEditors.BaseEdit editor)
        {
            long result;
            if (!long.TryParse(GetEditValueText(editor), out result))
                return 0;

            return result;
        }


        private string GetActiveCustomerLookupQuery()
        {
            return @"SELECT cm_entity [Entity], cm_branch [Branch], cm_cust_code1 [Code1], cm_cust_code2 [Code2], cm_cust_name [Name], cm_delv_address1 [Address]
                     FROM SO_CUST_MASTER WITH(NOLOCK)
                     WHERE ISNULL(cm_active_flag, '') <> 'D'
                     ORDER BY cm_cust_code1, cm_cust_code2";
        }

        private string GetPodRelevantFlag()
        {
            return checkBoxPODRelevant != null && checkBoxPODRelevant.Checked ? "Y" : "N";
        }

        private void LoadPodRelevantFlag()
        {
            try
            {
                if (checkBoxPODRelevant == null) return;
                string query = "SELECT ISNULL(cm_pod_relevant_flag, 'N') AS cm_pod_relevant_flag FROM SO_CUST_MASTER WITH(NOLOCK) WHERE cm_entity = " + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) +
                    " AND cm_branch = " + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) +
                    " AND cm_cust_code1 = " + FmtStr(txtCustCode.Text.Trim()) +
                    " AND cm_cust_code2 = " + FmtStr(txtCustCodeTo.Text.Trim());
                DataTable dt = _clsGlobal.ExecDT(query);
                checkBoxPODRelevant.Checked = dt.Rows.Count > 0 && dt.Rows[0]["cm_pod_relevant_flag"].ToString().Trim().Equals("Y", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                if (checkBoxPODRelevant != null) checkBoxPODRelevant.Checked = false;
            }
        }

        private void UpdateCustomerMasterEnhancementFlagsTrans()
        {
            string sql = "IF COL_LENGTH('dbo.SO_CUST_MASTER','cm_pod_relevant_flag') IS NOT NULL " +
                "UPDATE dbo.SO_CUST_MASTER SET cm_pod_relevant_flag = " + FmtStr(GetPodRelevantFlag()) +
                " WHERE cm_entity = " + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) +
                " AND cm_branch = " + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) +
                " AND cm_cust_code1 = " + FmtStr(txtCustCode.Text.Trim()) +
                " AND cm_cust_code2 = " + FmtStr(txtCustCodeTo.Text.Trim());
            _clsGlobal.ExecuteTrans(sql);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!ValidateCustomerForm())
                return;

            if (txtDelivByDays.Text.Trim() == "0")
            {
                txtDelivByDays.Text = "";
            }
            if ((txtDelivByDays.Text.Trim() == "" && GetEditValueInt64(CBDelivByDay) == 0))
            {
                if (txtDelivByDays.Text.Trim() == "")
                {
                    txtDelivByDays.Text = "0";
                }
                ShowValidationError(txtDelivByDays, "Harus mmengisi delivery by day atau delivery by days");
                //txtDelivByDays.Select();
                return;
            }

            else if (txtDelivByDays.Text.Trim() != "" && GetEditValueInt64(CBDelivByDay) != 0)
            {
                if (txtDelivByDays.Text.Trim() == "")
                {
                    txtDelivByDays.Text = "0";
                }
                ShowValidationError(txtDelivByDays, "Harus memilih salah satu delivery by day atau delivery by days");
                //txtDelivByDays.Select();
                return;
            }

            if (txtCustCode.Text.Trim() == "")
            {
                ShowValidationError(txtCustCode, " Code Customer belum di-input ");
                return;
            }
            if (this.State == StateEntry.New && txtCustCode.Text.Trim().Length != 6)
            {
                ShowValidationError(txtCustCode, " Customer Code harus 6 digit/karakter ! ");
                return;
            }
            if (txtCustCodeTo.Text.Trim() == "")
            {
                ShowValidationError(txtCustCodeTo, " Customer Code To Does not Allow Empty ! ");
                return;
            }
            if (txtCustName.Text.Trim() == "")
            {
                ShowValidationError(txtCustName, " Nama Customer belum di-input ");
                return;
            }
            if (txtShortName.Text.Trim() == "")
            {
                ShowValidationError(txtShortName, " Short Name Code  Does not Allow Empty ! ");
                return;
            }
            //adien
            //if (ctrlEntityCustMaster1.txtCM.Text.Trim() == "")
            if (ctrlEntityCustMaster2.txtCM.Text.Trim() == "")
            {
                ShowValidationError(ctrlEntityCustMaster2.txtCM, " Entity belum di-pilih ");
                //ctrlEntityCustMaster1.txtCM.Select();
                return;
            }

            //if (ctrlBranchCustMaster1.txtBranchIdCM.Text.Trim() == "")
            if (ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim() == "")
            {
                ShowValidationError(ctrlEntityCustMaster2.txtBranchIdCM, " Branch belum di-pilih ");
                //ctrlBranchCustMaster1.txtBranchIdCM.Select();
                return;
            }
            //adien
            if (ctrlDistrikCustMaster1.txtDistrikCM.Text.Trim() == "")
            {
                ShowValidationError(ctrlDistrikCustMaster1.txtDistrikCM, " Distrik belum di-pilih ");
                return;
            }
            if (ctrlDistrikCustMaster1.txtBeatCM.Text.Trim() == "")
            {
                ShowValidationError(ctrlDistrikCustMaster1.txtBeatCM, " Beat belum di-pilih ");
                return;
            }
            if (ctrlDistrikCustMaster1.txtSubBeatCM.Text.Trim() == "")
            {
                ShowValidationError(ctrlDistrikCustMaster1.txtSubBeatCM, " SubBeat belum di-pilih ");
                return;
            }
            if (ctrlGroupDiscount1.txtGpDiscountT1.Text.Trim() == "")
            {
                ShowValidationError(ctrlGroupDiscount1.txtGpDiscountT1, " Group Discount belum di-pilih ");
                return;
            }
            if (txtIndustriT1.Text.Trim() == "")
            {
                ShowValidationError(txtIndustriT1, " Industri belum di-pilih");
                return;
            }
            if (ctrlKodepasarDesc.txtKodePasarT1.Text.Trim() == "")
            {
                ShowValidationError(ctrlKodepasarDesc.txtKodePasarT1, " Pasar belum di-pilih");
                return;
            }
            if (ctrlGroupOutlet1.txtGroupCM.Text.Trim() == "")
            {
                ShowValidationError(ctrlGroupOutlet1.txtGroupCM, " Group Outlet belum di-pilih ");
                return;
            }
            if (ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Text.Trim() == "")
            {
                ShowValidationError(ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet, " Tipe Outlet belum di-pilih");
                return;
            }
            if (txtLeadtime.Text.Trim() == "")
            {
                ShowValidationError(txtLeadtime, " Lead Time harus diisi ");
                return;
            }
            if (popUpProvinsinCity1.txtPropinsiCMC.Text.Trim() == "")
            {
                ShowValidationError(popUpProvinsinCity1.txtPropinsiCMC, " Provinsi belum di-pilih ");
                return;
            }
            if (popUpProvinsinCity1.txtKabupatenCMC.Text.Trim() == "")
            {
                ShowValidationError(popUpProvinsinCity1.txtKabupatenCMC, " City belum di-pilih ");
                return;
            }
            if (popUpProvinsinCity1.txtKecamatan.Text.Trim() == "")
            {
                ShowValidationError(popUpProvinsinCity1.txtKecamatan, " Kecamatan belum di-pilih");
                return;
            }
            if (popUpProvinsinCity1.txtKelurahan.Text.Trim() == "")
            {
                ShowValidationError(popUpProvinsinCity1.txtKelurahan, " Kelurahan belum di-pilih ");
                return;
            }
            if (txtBilltoCustCode.Text.Trim() == "")
            {
                ShowValidationError(txtBilltoCustCode, " Bill to Cust Code belum diisi! ");
                return;
            }
            if (txtAddress1T2.Text.Trim() == "")
            {
                ShowValidationError(txtAddress1T2, " Address 1 cust property belum diisi! ");
                return;
            }
            if (txtAddress1T2.Text.Length < 20)
            {
                ShowValidationError(txtAddress1T2, " Address 1 cust property diisi minimal 20 karakter ! ");
                return;
            }
            if (txtAddress1T3.Text.Trim() == "")
            {
                ShowValidationError(txtAddress1T3, " Address 1 billing address belum diisi! ");
                return;
            }

            if (txtAddress1T3.Text.Length < 20)
            {
                ShowValidationError(txtAddress1T3, " Address 1 billing address diisi minimal 20 karakter ! ");
                return;
            }
            if (txtAddress1T4.Text.Trim() == "")
            {
                ShowValidationError(txtAddress1T4, " Address 1 delivery address belum diisi! ");
                return;
            }
            if (txtAddress1T4.Text.Length < 20)
            {
                ShowValidationError(txtAddress1T4, " Address 1 delivery address diisi minimal 20 karakter ! ");
                return;
            }
            /* Validasi Telepon Yuda 09.03.2021*/
            if (!String.IsNullOrEmpty(txtTelephoneT4.Text))
            {
                try
                {
                    Double dblcheck = Convert.ToDouble(txtTelephoneT4.Text);
                }
                catch (Exception ex)
                {
                    ShowValidationError(txtTelephoneT4, " Telepon Harus berupa angka dan tidak boleh ada spasi atau karakter lain! ");
                    return;
                }
                if (txtTelephoneT4.Text.Length < 8)
                {
                    ShowValidationError(txtTelephoneT4, " Telepon Delivery Harus Minimal 8 Digit! ");
                    return;
                }
            }

            /*end*/

            /* Validasi Monitoring Insurance Yuda 02.03.2022 */
            if (checkBoxMonitoringInsuranceT1.Checked)
            {
                if (!dtTglToIns.Text.Equals(dtTglFromIns.Text))
                {
                    if (Convert.ToDateTime(dtTglToIns.DateTime) < Convert.ToDateTime(dtTglFromIns.DateTime))
                    {
                        ShowValidationError(dtTglToIns, " Tanggal Insurance From harus lebih kecil dari pada Tanggal Insurance To. ");
                        return;
                    }
                }

            }
            /* End Validasi Monitoring Insurance */

            /* Validasi LNG & LAT */
            if (string.IsNullOrEmpty(txtLong.Text) || string.IsNullOrEmpty(txtLat.Text))
            {
                ShowValidationError(string.IsNullOrEmpty(txtLong.Text) ? txtLong : txtLat, "Lat & Lng harus terisi. ");
                return;
            }

            if (!string.IsNullOrEmpty(txtLong.Text))
            {
                string[] split = txtLong.Text.Split('.');
                if (split.Length < 2)
                {
                    ShowValidationError(txtLong, "mohon isi lng sesuai format. ");
                    return;
                }
                else
                {
                    if (split[1].Length < 5)
                    {
                        ShowValidationError(txtLong, "mohon isi lng sesuai format. ");
                        return;
                    }
                }

                if (!is_valid_coordinate(txtLong.Text))
                {
                    ShowValidationError(txtLong, "Silahkan isi data Geotag dengan benar!. ");
                    return;
                }
                decimal dLat = Convert.ToDecimal(txtLong.Text);
                if (dLat < -180 || dLat > 180)
                {
                    ShowValidationError(txtLong, "Longitude tidak boleh lebih besar 180 dan tidak boleh lebih kecil dari -180 !. ");
                    return;
                }
            }

            if (!string.IsNullOrEmpty(txtLat.Text))
            {
                string[] split = txtLat.Text.Split('.');
                if (split.Length < 2)
                {
                    ShowValidationError(txtLat, "mohon isi lat sesuai format. ");
                    return;
                }
                else
                {
                    if (split[1].Length < 5)
                    {
                        ShowValidationError(txtLat, "mohon isi lat sesuai format. ");
                        return;
                    }
                }
                if (!is_valid_coordinate(txtLat.Text))
                {
                    ShowValidationError(txtLat, "Silahkan isi data Geotag dengan benar!. ");
                    return;
                }
                decimal dLat = Convert.ToDecimal(txtLat.Text);
                if (dLat < -90 || dLat > 90)
                {
                    ShowValidationError(txtLat, "Latitude tidak boleh lebih besar 90 dan tidak boleh lebih kecil dari -90 !. ");
                    return;
                }

            }

            /* end */


            /* ASK 27780 */
            //if (string.IsNullOrEmpty(txtNIK.Text.Trim()))
            //{
            //    MessageBox.Show("Silahkan isi nomor NIK dengan benar!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    this.txtNIK.Focus();
            //    return;
            //}

            string __custaddress = (txtAddress1T2.Text + " " + txtAddress2T2.Text).Trim();
            string __billaddress = (txtAddress1T3.Text + " " + txtAddress2T3.Text + " " + txtAddress3T3.Text + " " + txtAddress4T3.Text).Trim();
            string __delvaddress = (txtAddress1T4.Text + " " + txtAddress2T4.Text + " " + txtAddress3T4.Text + " " + txtAddress4T4.Text).Trim();

            //if(!is_valid_Address(__custaddress)){
            //    MessageBox.Show("Silahkan isi data Alamat Customer dengan benar!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    this.txtAddress1T2.Focus();
            //    return;
            //}

            if (!is_valid_Address(__billaddress))
            {
                ShowValidationError(txtAddress1T3, "Silahkan isi data Alamat Billing dengan benar!");
                return;
            }

            if (!is_valid_Address(__delvaddress))
            {
                ShowValidationError(txtAddress1T4, "Silahkan isi data Alamat Delivery dengan benar!");
                return;
            }

            //jenis identitas

            if (string.IsNullOrEmpty(GetEditValueText(cbJenisIdentitas).Trim()))
            {
                ShowValidationError(cbJenisIdentitas, " Jenis Identitas Mohon Dipilih !");
                return;
            }

            if (GetEditValueText(cbJenisIdentitas).Trim() == "1")
            {
                if (string.IsNullOrEmpty(this.txtNpwpT2Masked.Text.Trim()))
                {
                    ShowValidationError(txtNpwpT2Masked, "Jenis Identitas : (BADAN) : NPWP harus diisi!");
                    return;
                }
            }
            else if (GetEditValueText(cbJenisIdentitas).Trim() == "2")
            {
                if (string.IsNullOrEmpty(this.txtNIK.Text.Trim()))
                {
                    ShowValidationError(txtNIK, "Jenis Identitas : (PRIBADI) : NIK harus diisi!");
                    return;
                }
            }


            /* END OF ASK 27780 */
            if (checkBoxNpwpT1.Checked)
            {
                if (txtTaxNameT2.Text.Trim() == "")
                {
                    ShowValidationError(txtTaxNameT2, " Tax Name harus diisi! ");
                    return;
                }
                if (txtAddress1T2.Text.Trim() == "")
                {
                    ShowValidationError(txtAddress1T2, " Tax Address 1 harus diisi! ");
                    return;
                }
                if (txtAddress2T2.Text.Trim() == "")
                {
                    ShowValidationError(txtAddress2T2, " Tax Address 2 harus diisi! ");
                    return;
                }
                if (txtCountryT2.Text.Trim() == "")
                {
                    ShowValidationError(txtCountryT2, " Tax Country harus diisi! ");
                    return;
                }

                //if (txtNIK.Text.Trim() == "")
                //{
                //    MessageBox.Show(" NIK harus diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    txtNIK.Select();
                //    return;
                //}

                //if (txtNIK.Text.Length < 16)
                //{
                //    MessageBox.Show(" NIK harus 16 karakter! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    txtNIK.Select();
                //    return;
                //}

                //string __npwp = this.txtNpwpT2.Text.Trim();
                string __npwp = this.txtNpwpT2Masked.Text.Trim();
                string __sppkp = this.txtSPPKP.Text.Trim();
                if (!IsTextEditMaskCompleted(this.txtNpwpT2Masked))
                {
                    if (this.txtNpwpT2Masked.Enabled)
                    {
                        ShowValidationError(txtNpwpT2Masked, "Pengisian NPWP tidak komplit!");
                        return;
                    }
                }
                if (__npwp.IsNullOrEmptyOrWhiteSpace() || __sppkp.IsNullOrEmptyOrWhiteSpace())
                {
                    ShowValidationError(__npwp.IsNullOrEmptyOrWhiteSpace() ? txtNpwpT2Masked : txtSPPKP, "NPWP & SPPKP harus diisi!");
                    return;
                }
                else
                {
                    if (!__sppkp.IsNullOrEmptyOrWhiteSpace())
                    {
                        if (__sppkp.Length < 20)
                        {
                            ShowValidationError(txtSPPKP, "sppkp harus minimal 20 karakter!");
                            return;
                        }
                    }
                    string __nik = this.txtNIK.Text.Trim();
                    if (!__nik.IsNullOrEmptyOrWhiteSpace())
                    {
                        if (__nik.Length < 16)
                        {
                            ShowValidationError(txtNIK, "Panjang minimal NIK tidak boleh kurang dari 16");
                            return;
                        }
                        else if (!is_valid_NIK(__nik))
                        {
                            ShowValidationError(txtNIK, "NIK tidak sesuai, Mohon lakukan pengisian dengan benar!!!");
                            return;
                        }
                    }

                }

                if (__npwp.Length < 16)
                {
                    ShowValidationError(txtNpwpT2Masked, "Panjang NPWP minimal tidak boleh kurang dari 16");
                    //this.txtNpwpT2.Focus();
                    return;
                }
                else if (__npwp.Length > 16)
                {
                    ShowValidationError(txtNpwpT2Masked, "Panjang NPWP maksimal 16");
                    //this.txtNpwpT2.Focus();
                    return;
                }
                //else
                //{
                //    decimal __lNpwp = 0;
                //    decimal.TryParse(__npwp.Replace(".", "").Replace("-", ""), out __lNpwp);
                //    if (__lNpwp == 0)
                //    {
                //        MessageBox.Show("Silahkan isi nomor NPWP dengan benar", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        //this.txtNpwpT2.Focus();
                //        this.txtNpwpT2Masked.Focus();
                //        return;
                //    }
                //}
                //if (!is_valid_npwp(__npwp))
                //{
                //    MessageBox.Show("Silahkan isi nomor NPWP dengan format yang benar", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //    //this.txtNpwpT2.Focus();
                //    this.txtNpwpT2Masked.Focus();
                //    return;
                //}
            }
            else
            {
                string __nik = this.txtNIK.Text.Trim();
                //string __npwp = this.txtNpwpT2.Text.Trim();
                string __npwp = this.txtNpwpT2Masked.Text.Trim();
                if (__nik.IsNullOrEmptyOrWhiteSpace() && __npwp.IsNullOrEmptyOrWhiteSpace())
                {
                    ShowValidationError(txtNIK, "Silahkan lengkapi Nomor NPWP atau NIK");
                    return;
                }
                else if (!__nik.IsNullOrEmptyOrWhiteSpace() && __npwp.IsNullOrEmptyOrWhiteSpace())
                {
                    if (__nik.Length < 16)
                    {
                        ShowValidationError(txtNIK, "Panjang minimal NIK tidak boleh kurang dari 16");
                        return;
                    }
                    else if (__nik.Length > 16)
                    {
                        ShowValidationError(txtNIK, "Panjang maksimal NIK 16");
                        return;
                    }
                    /* 
                    * nik validasi yuda 08032021
                    */
                    else if (!is_valid_NIK(__nik))
                    {
                        ShowValidationError(txtNIK, "NIK tidak sesuai, Mohon lakukan pengisian dengan benar!!!");
                        return;
                    }
                    /* END */
                    else
                    {
                        decimal __lNik = 0;
                        decimal.TryParse(__nik, out __lNik);
                        if (__lNik == 0)
                        {
                            ShowValidationError(txtNIK, "Silahkan isi nomor NIK dengan benar");
                            return;
                        }
                    }

                }
                else if (__nik.IsNullOrEmptyOrWhiteSpace() && !__npwp.IsNullOrEmptyOrWhiteSpace())
                {
                    if (__npwp.Length < 16)
                    {
                        ShowValidationError(txtNpwpT2Masked, "Panjang NPWP minimal tidak boleh kurang dari 16");
                        return;
                    }
                    else if (__npwp.Length > 16)
                    {
                        ShowValidationError(txtNpwpT2Masked, "Panjang NPWP maksimal 16");
                        return;
                    }
                    //else
                    //{
                    //    decimal __lNpwp = 0;
                    //    decimal.TryParse(__npwp.Replace(".", "").Replace("-", ""), out __lNpwp);
                    //    if (__lNpwp == 0)
                    //    {
                    //        MessageBox.Show("Silahkan isi nomor NPWP dengan benar", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //        this.txtNpwpT2.Focus();
                    //        return;
                    //    }
                    //}
                    //if (!is_valid_npwp(__npwp))
                    //{
                    //    MessageBox.Show("Silahkan isi nomor NPWP dengan format yang benar", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    //this.txtNpwpT2.Focus();
                    //    this.txtNpwpT2Masked.Focus();
                    //    return;
                    //}
                }
                else
                {
                    if (!__nik.IsNullOrEmptyOrWhiteSpace())
                    {
                        if (__nik.Length < 16)
                        {
                            ShowValidationError(txtNIK, "Panjang minimal NIK tidak boleh kurang dari 16");
                            return;
                        }
                        else if (!is_valid_NIK(__nik))
                        {
                            ShowValidationError(txtNIK, "NIK tidak sesuai, Mohon lakukan pengisian dengan benar!!!");
                            return;
                        }
                    }
                    //if (!__npwp.IsNullOrEmptyOrWhiteSpace())
                    //{
                    //    if (!is_valid_npwp(__npwp))
                    //    {
                    //        MessageBox.Show("Silahkan isi nomor NPWP dengan format yang benar", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //        //this.txtNpwpT2.Focus();
                    //        this.txtNpwpT2Masked.Focus();
                    //        return;
                    //    }
                    //}


                    //if (__npwp.Length < 20)
                    //{
                    //    MessageBox.Show("Panjang NPMP minimal tidak boleh kurang dari 20", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    this.txtNpwpT2.Focus();
                    //    return;
                    //}
                }
            }

            //if (txtNpwpT2.Text.Trim() == "")
            //{
            //    MessageBox.Show(" NPWP harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    txtNpwpT2.Select();
            //    return;
            //}
            //if (txtTremPaymentT1.Text.Trim() == "")
            //{
            //    MessageBox.Show(" Term Of Payment belum di-pilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    txtTremPaymentT1.Select();
            //    return;
            //}

            //if (CBBlacklist.Checked == true)
            //{                
            //if(txtReason.Text.Trim() == "")
            if (txtReason.Text.Trim() == "" && eBlackList != CBBlacklist.Checked)
            {
                ShowValidationError(txtReason, " Flag BackList telah berubah \n Reason BackList harus diisi!");
                return;
            }
            //}
            // Update Yuda 06092021 //
            if (!string.IsNullOrEmpty(txtNIK.Text.Trim()) && string.IsNullOrEmpty(GetEditValueText(cbPemilikNIK).Trim()))
            {
                ShowValidationError(cbPemilikNIK, " Pemilik NIK Mohon Dipilih !");
                return;
            }
            //------end--------//

            /*** update yuda 18082022 **/
            /** Request Pak Wawan : Bill Name Harus Diisi **/
            /** Request Kang Sam : ZipCode Harus Diisi **/
            //if (string.IsNullOrEmpty(txtPostCodeT2.Text))
            //{
            //    MessageBox.Show("Post Code Customer Tidak boleh kosong!!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    txtPostCodeT2.Select();
            //    return;
            //}
            if (string.IsNullOrEmpty(txtPostCodeT3.Text))
            {
                ShowValidationError(txtPostCodeT3, "Post Code Billing Tidak boleh kosong!!! ");
                return;
            }
            if (string.IsNullOrEmpty(txtPostCodeT4.Text))
            {
                ShowValidationError(txtPostCodeT4, "Post Code Delivery Tidak boleh kosong!!! ");
                return;
            }
            if (string.IsNullOrEmpty(txtCustNameT3.Text))
            {
                ShowValidationError(txtCustNameT3, "Billing Name Tidak boleh kosong!!! ");
                return;
            }
            /** end of Update 18082022 **/

            if (GetGridRowCount(DVGSalesman, DVGSalesmanView) > 0)
            {
                System.Windows.Forms.DataGridViewRow dgr;
                try
                {
                    dgr = GetGridCurrentRow(DVGSalesman, DVGSalesmanView);
                    if (dgr != null)
                    {
                        bool rowDuplicate = false;
                        bool rowDuplicate2 = false;
                        bool rowDuplicate3 = false;
                        bool rowDuplicate4 = false;
                        foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(DVGSalesman, DVGSalesmanView))
                        {
                            if ((gRow.Cells["csc_salesman_id"].Value.ToString().Trim() == ""))
                            {
                                rowDuplicate = true;
                                break;
                            }
                            if (gRow.Cells["csc_visit"].Value.ToString() == "")
                            {
                                rowDuplicate3 = true;
                                break;
                            }
                            if ((gRow.Cells["csc_visit_week1"].Value.ToString() == "false" || gRow.Cells["csc_visit_week1"].Value.ToString() == "") && (gRow.Cells["csc_visit_week2"].Value.ToString() == "false" || gRow.Cells["csc_visit_week2"].Value.ToString() == "") && (gRow.Cells["csc_visit_week3"].Value.ToString() == "false" || gRow.Cells["csc_visit_week3"].Value.ToString() == "") && (gRow.Cells["csc_visit_week4"].Value.ToString() == "false" || gRow.Cells["csc_visit_week4"].Value.ToString() == ""))
                            {
                                rowDuplicate4 = true;
                                break;
                            }
                            if (gRow.Cells["csc_route"].Value.ToString().Trim() == "")
                            {
                                rowDuplicate2 = true;
                                break;
                            }
                        }
                        if (rowDuplicate == true)
                        {
                            ShowValidationError(tabPage6, "Salesman belum dipilih", MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate2 == true)
                        {
                            ShowValidationError(tabPage6, "Rute Belum diinput", MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate3 == true)
                        {
                            ShowValidationError(tabPage6, "Hari Kunjungan belum dipilih", MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate4 == true)
                        {
                            ShowValidationError(tabPage6, "Pola kunjungan belum dipilih", MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                finally
                {
                    dgr = null;
                }
            }

            if (GetGridRowCount(DGVHIR, DGVHIRView) > 0)
            {
                System.Windows.Forms.DataGridViewRow dgr;
                try
                {
                    dgr = GetGridCurrentRow(DGVHIR, DGVHIRView);

                    if (dgr != null)
                    {
                        bool rowDuplicate11 = false;
                        bool rowDuplicate2 = false;
                        foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(DGVHIR, DGVHIRView))
                        {
                            if ((gRow.Cells["cpl_line_code"].Value.ToString().Trim() == ""))
                            {
                                rowDuplicate11 = true;
                                break;
                            }
                            if (gRow.Cells["cpl_cust_type"].Value.ToString() == "")
                            {
                                rowDuplicate2 = true;
                                break;
                            }
                        }
                        if (rowDuplicate11 == true)
                        {
                            ShowValidationError(tabPage5, "Produk Line belum dipilih", MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate2 == true)
                        {
                            ShowValidationError(tabPage5, "Type Outlet belum dipilih", MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                finally
                {
                    dgr = null;
                }
            }
            else
            {
                //if (clsGlobal.MODE_TRX == 1)
                if (__state == StateEntry.New)
                {
                    ShowValidationError(tabPage5, "Data Cluster kosong, silahkan isi terlebih dahulu sebelum disimpan data", MessageBoxIcon.Error);
                    return;
                }
            }

            // checking bangunan
            if (CBBagunanT1.EditValue == null)
            {
                ShowValidationError(CBBagunanT1, "Jenis bangunan belum diisi, silahkan isi terlebih dahulu sebelum disimpan data", MessageBoxIcon.Error);
                return;
            }
            FlushAllGridEditors();
            if (!isValidGroupPrice())
            {
                ShowValidationError(tabPage8, "Data Group Price STD (standart) harus ada, silahkan isi terlebih dahulu sebelum simpan data", MessageBoxIcon.Error);
                return;
            }

            if (txtShipArea.Text.Trim() == string.Empty)
            {
                ShowValidationError(txtShipArea, " Shipment Area harus diisi!");
                return;
            }

            if (__isFlagDC && skillidcodetb.Text.IsNullOrEmptyOrWhiteSpace())
            {
                ShowValidationError(skillidcodetb, " Skill ID harus diisi!");
                return;
            }

            if (__isFlagDC && txtTelephoneT4.Text.IsNullOrEmptyOrWhiteSpace())
            {
                ShowValidationError(txtTelephoneT4, " Telephone Delivery Address harus diisi!");
                return;
            }

            if (CBPembyaranFktT1.EditValue == null)
            {
                ShowValidationError(CBPembyaranFktT1, "Pembayaran Faktur belum diisi, silahkan isi terlebih dahulu sebelum disimpan data", MessageBoxIcon.Error);
                return;
            }
            // PENGECEKAN NIK JIKA BERISI 1111 1111 1111 1111
            // MAKA PEMBAYARANG AUTO TUNAI
            // BASED ASK : 11060
            __isAutoTunai = false;
            string __znik = this.txtNIK.Text.Trim();
            string __ztypepemb = this.GetEditValueText(CBPembyaranFktT1);

            if (__nik_blocked.DATA.Any(x => x.Code.CompareC(__znik)))
            {
                if (!__ztypepemb.CompareC("T"))
                {
                    DialogResult dres = MessageBox.Show("NIK ter-block.\r\nJika anda melanjutkan proses ini, Tipe pembayaran akan dirubah menjadi Tunai secara otomatis.\r\nApakah anda ingin melanjutkan proses ini?"
                        , clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dres == DialogResult.Yes)
                        __isAutoTunai = true;
                    else
                        return;
                }
            }
            // END 11060

            if (__isAsk11214)
            {
                // PENGECEKAN HIRARKY PEMERINTAHAN
                // BASED ASK : 11214
                string code_kelurahan = popUpProvinsinCity1.txtKelurahan.Text;
                string code_kecamatan = popUpProvinsinCity1.txtKecamatan.Text;
                string code_kabupaten = popUpProvinsinCity1.txtKabupatenCMC.Text;
                string code_provinsi = popUpProvinsinCity1.txtPropinsiCMC.Text;
                try
                {
                    using (var hirarki = new TIRAData<TIRAItem>(string.Format(QUERY_HIRARKI, code_kelurahan, code_kecamatan, code_kabupaten, code_provinsi), 0))
                    {
                        hirarki.Execute();
                        if (hirarki.DATA.Count < 1)
                        {
                            ShowValidationError(popUpProvinsinCity1.txtKelurahan, "Hirarky Pemerintahaan tidak sesuai, Pastikan Hirarky pemerintahaan benar.", MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                catch (Exception exc)
                {
                    ShowValidationError(popUpProvinsinCity1.txtKelurahan, string.Format("Error Pengecekan Hirarki: {0}", exc.Message), MessageBoxIcon.Error);
                    return;
                }
                // END 11214
            }

            //// PENGECEKAN FOTO
            //// MAKA PEMBAYARANG AUTO TUNAI
            //// BASED ASK : 21256
            //var ph = photoExist.DATA.FirstOrDefault() ?? new TiraPhoto();
            //bool isPkp = checkBoxNpwpT1.Checked;
            //if (isPkp && !ph.IsNPWP)
            //{
            //    MessageBox.Show("Untuk outlet dengan pajak standart, wajib upload foto NPWP.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}
            //else if (!isPkp && !ph.IsNIK)
            //{
            //    MessageBox.Show("Untuk outlet dengan pajak gabungan, wajib upload foto NIK.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}
            //// END 21256

            //if (__isAsk30724)
            //{
            FlushAllGridEditors();
            if (!ValidatePartnerFunctionDefaults())
                return;

            string statusCode = GetStatusCode();
            if (string.IsNullOrWhiteSpace(statusCode))
            {
                ShowValidationError(cbstatus, " Status harus diisi!");
                return;
            }
            //}

            //if (clsGlobal.MODE_TRX == 1)
            if (__state == StateEntry.New)
            {
                SaveNew();
                SaveBackList();
            }
            else
            {
                SaveEdit();
                SaveBackList();
            }

            if (__isFlagSync)
            {
                string __urllog = "";
                string __payloadlog = "";
                string __responselog = "";
                string __errorcodelog = "";
                string __errorlog = "";
                string __sellingpoint = "";
                string __outlet1 = "";
                List<CreateCollectionModel> listcoll = new List<CreateCollectionModel>();
                strSQL = "";
                strSQL += " SELECT * FROM ( SELECT DISTINCT  cp_branch+cp_cust_code1 outlet,cp_cust_code1,br_branch_ud1,cp_index,cp_price_group,br_branch_ud2 ";
                strSQL += " FROM SO_CUST_MASTER ";
                strSQL += " INNER JOIN SO_CUST_MASTER_PRICE_GROUP ON cm_cust_code1=cp_cust_code1 and cm_cust_code2=cp_cust_code2 and cm_entity = cp_entity and cm_branch = cp_branch ";
                strSQL += " INNER JOIN OB_PRICE_LIST_DETAIL ON cp_price_group=opl_price_code ";
                strSQL += " INNER JOIN GS_BRANCH ON br_gl_entity_initial = cp_entity and br_branch_id = cp_branch  ";
                strSQL += " WHERE cm_entity = '" + ctrlEntityCustMaster2.txtCM.Text + "' and cm_branch = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "' AND cm_cust_code1  = '" + txtCustCode.Text + "'AND cm_cust_code2  = '" + txtCustCodeTo.Text + "'  ";
                strSQL += " ) A order by CASE WHEN cp_price_group = 'STD' then 1 else 0 end, case when cp_price_group = br_branch_ud2 then 1 else 0 end,cp_index asc ";

                DataTable data = _clsGlobal.ExecDT(strSQL);
                if (data.Rows.Count > 0)
                {

                    /** Update Group Harga **/
                    __sellingpoint = data.Rows[0]["br_branch_ud1"].ToString();
                    __outlet1 = data.Rows[0]["cp_cust_code1"].ToString();
                    string __outlet = data.Rows[0]["outlet"].ToString();
                    for (int i = 0; i < data.Rows.Count; i++)
                    {
                        string __grpharga = data.Rows[i]["cp_price_group"].ToString();
                        string __index = (i + 1).ToString();

                        CreateCollectionModel __list = new CreateCollectionModel();
                        __list.groupPriceRef = __grpharga;
                        __list.priority = Convert.ToInt32(__index);
                        listcoll.Add(__list);

                    }
                    using (var wclient = new WebClient())
                    {
                        var __url = string.Empty;
                        var encodedJson = JsonConvert.SerializeObject(listcoll);
                        __payloadlog = encodedJson;
                        wclient.Headers.Add("Content-Type:application/json");
                        try
                        {
                            StringBuilder query = new StringBuilder();
                            query.AppendLine(" SELECT H.gh_function_code [Code], H.gh_function_desc [Description]");
                            query.AppendLine(" FROM[dbo].[GS_GEN_HARDCODED] H WITH(NOLOCK)");
                            query.AppendLine(" WHERE H.gh_function_name IN ('TIRA_API_SDM1')");
                            query.AppendLine(" AND H.gh_function_code IN ('ENDPOINT_UPDATE_GROUPPRICE','API_X_KEY')");
                            var __auths = from x in _clsGlobal.ExecDT(query.ToString())
                                            .Rows.Cast<DataRow>()
                                          select new
                                          {
                                              Key = x["Code"].ToString(),
                                              Value = x["Description"].ToString(),
                                          };
                            if (__auths.Count() < 2)
                            {
                                MessageBox.Show("Parameter authentication tidak ada.\r\nMohon cek parameter user dan password TIRA API."
                                    , clsGlobal.APP_MSG_CAPTION
                                    , MessageBoxButtons.OK
                                    , MessageBoxIcon.Information);
                                return;
                            }

                            var __default = new { Key = string.Empty, Value = string.Empty };
                            __url = (__auths.FirstOrDefault(x => x.Key.CompareC("ENDPOINT_UPDATE_GROUPPRICE")) ?? __default).Value + __outlet;
                            __urllog = __url;
                            var __apikey = (__auths.FirstOrDefault(x => x.Key.CompareC("API_X_KEY")) ?? __default).Value;

                            //wclient.Credentials = new NetworkCredential(__user, __pass);
                            wclient.Headers.Add("x-api-key", __apikey);

                            var response = wclient.UploadString(__url, encodedJson);
                            __responselog = response;

                            string __sql = "EXEC  [DB_LOG_API].dbo.[IP_INSERT_LOG_OUTLET_GROUP_PRICE] ";
                            __sql += "'" + __sellingpoint + "', ";
                            __sql += "'" + __outlet1 + "', ";
                            __sql += "'" + __payloadlog + "', ";
                            __sql += "'SINBAD', ";
                            __sql += "1, ";
                            __sql += "'" + __responselog.ToString() + "' ";

                            _clsGlobal.Execute(__sql);

                            //updatelogAPI(__urllog, __payloadlog, 1, __responselog, "", "");
                            MessageBox.Show("Proses Sync Group Price ke API berhasil", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (WebException ex)
                        {

                            if (ex.Response == null)
                            {
                                insertToLOGAPI(__urllog, __payloadlog, clsLogin.USERID);
                                MessageBox.Show("Koneksi ke TIRA API terputus");
                            }
                            else
                            {
                                var resp = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                                dynamic obj = JsonConvert.DeserializeObject(resp);

                                /** Insert Ke DB_LOG_API **/
                                string __sql = "EXEC  [DB_LOG_API].dbo.[IP_INSERT_LOG_OUTLET_GROUP_PRICE] ";
                                __sql += "'" + __sellingpoint + "', ";
                                __sql += "'" + __outlet1 + "', ";
                                __sql += "'" + __payloadlog + "', ";
                                __sql += "'SINBAD', ";
                                __sql += "0, ";
                                __sql += "'" + resp.ToString() + "' ";

                                _clsGlobal.Execute(__sql);

                                /** End Insert Ke DB_LOG_API **/

                                HttpWebResponse webResp = (HttpWebResponse)ex.Response;

                                switch (webResp.StatusCode)
                                {
                                    case HttpStatusCode.NotFound: // 404
                                        MessageBox.Show("Error API: URL API tidak ditemukan");
                                        break;
                                    case HttpStatusCode.InternalServerError: // 500
                                        MessageBox.Show("Error API : Internal Server Error");
                                        break;
                                    default:
                                        MessageBox.Show(ex.Message);
                                        break;
                                }
                            }
                        }
                    }
                }

            }
        }
        public void insertToLOGAPI(string url, string payload, string userid)
        {
            _clsGlobal.Execute("EXEC IP_INSERT_LOG_API '" + url + "', '" + payload + "', 0, '" + userid + "'");
        }
        private void SaveNew()
        {
            //adien
            //strSQL = " select cm_cust_code1, cm_entity , cm_branch, cm_cust_code2 from SO_CUST_MASTER where cm_entity = '" + ctrlEntityCustMaster1.txtCM.Text + "' and cm_branch = '" + ctrlBranchCustMaster1.txtBranchIdCM.Text + "' AND cm_cust_code1  = '" + txtCustCode.Text + "'AND cm_cust_code2  = '" + txtCustCodeTo.Text + "'";
            strSQL = " select cm_cust_code1, cm_entity , cm_branch, cm_cust_code2 from SO_CUST_MASTER WITH (NOLOCK) where cm_entity = '" + ctrlEntityCustMaster2.txtCM.Text + "' and cm_branch = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "' AND cm_cust_code1  = '" + txtCustCode.Text + "'AND cm_cust_code2  = '" + txtCustCodeTo.Text + "'";
            //adien
            DataTable dt = _clsGlobal.ExecDT(strSQL);

            if (dt.Rows.Count > 0)
            {
                ShowValidationError(txtCustCode, "Kode Cust ini sudah terdapat dalam Data Base", MessageBoxIcon.Error);
                return;
            }
            else
            {
                //save SP_AR_CUST_MASTER
                try
                {
                    string cb1, cb2, cb3, cb4, cb5, cb6, cb7, cb8, cb9, cb10, cb11, tmpCMTOPID, tmpCMTOPIDDesc = "", __cbfullfilment;
                    if (checkBoxDiscountBaseT1.Checked)
                    {
                        cb1 = "Y";
                    }
                    else
                    {
                        cb1 = "N";
                    }
                    if (checkBoxtxtBatasLimitT1.Checked)
                    {
                        cb2 = "Y";
                    }
                    else
                    {
                        cb2 = "N";
                    }
                    if (checkBoxKontraBonT1.Checked)
                    {
                        cb3 = "Y";
                    }
                    else
                    {
                        cb3 = "N";
                    }
                    if (checkBoxWarehouseTidak.Checked)
                    {
                        cb4 = "Y";
                    }
                    else
                    {
                        cb4 = "N";
                    }
                    if (checkBoxFakPajakT2.Checked)
                    {
                        cb5 = "Y";
                    }
                    else
                    {
                        cb5 = "N";
                    }
                    if (checkBoxNpwpT1.Checked)
                    {
                        cb6 = "Y";
                    }
                    else
                    {
                        cb6 = "N";
                    }
                    if (checkBoxTopByCustomerT1.Checked)
                    {
                        cb7 = "Y";
                    }
                    else
                    {
                        cb7 = "N";
                    }
                    if (checkBoxInvoiceTidak.Checked)
                    {
                        cb8 = "Y";
                    }
                    else
                    {
                        cb8 = "N";
                    }
                    if (CBBlacklist.Checked)
                    {
                        cb9 = "Y";
                    }
                    else
                    {
                        cb9 = "N";
                    }
                    if (chkPinSL.Checked)
                    {
                        cb10 = "Y";
                    }
                    else
                    {
                        cb10 = "N";
                    }

                    if (checkBoxMonitoringInsuranceT1.Checked)
                    {
                        cb11 = "Y";
                    }
                    else
                    {
                        cb11 = "N";
                    }

                    tmpCMTOPID = txtTremPaymentT1.Text.Trim();

                    if (tmpCMTOPID != "")
                    {
                        if (tmpCMTOPID.Substring(0, 1) == "P")
                        {
                            tmpCMTOPIDDesc = FmtStr(txtTremPaymentT1.Text.Trim().Replace("P", ""));
                        }
                        else
                        {
                            tmpCMTOPIDDesc = FmtStr(txtTremPaymentT1.Text.Trim());
                        }
                    }
                    else
                    {
                        tmpCMTOPIDDesc = "0";
                    }

                    if (__isAsk28151)
                    {
                        if (cbfullfilment.Checked)
                        {
                            __cbfullfilment = "Y";
                        }
                        else
                        {
                            __cbfullfilment = "N";
                        }
                    }

                    //david 17 Mei 2018
                    string mulaiush = _clsGlobal.DateToDB(dateTimePicker1.DateTime.ToString("MM/dd/yyyy"));
                    string birthday = _clsGlobal.DateToDB(dateTimePicker2.DateTime.ToString("MM/dd/yyyy"));
                    string insfrom = _clsGlobal.DateToDB(dtTglFromIns.DateTime.ToString("MM/dd/yyyy"));
                    string insto = _clsGlobal.DateToDB(dtTglToIns.DateTime.ToString("MM/dd/yyyy"));
                    string updtdt1 = "";
                    if (txtLastUpdateT1.Text != "")
                    {
                        DateTime updtdt = Convert.ToDateTime(txtLastUpdateT1.Text);
                        updtdt1 = updtdt.ToString("MM/dd/yyyy");
                    }
                    else
                    {
                        updtdt1 = "";
                    }

                    if (__isAutoTunai)
                        CBPembyaranFktT1.EditValue = "T";

                    string skillidcode = skillidcodetb.Text;

                    string statusCode = GetStatusCode();
                    if (string.IsNullOrWhiteSpace(statusCode))
                        statusCode = "O";

                    strSQL = "EXEC SP_AR_CUST_MASTER '1', " +
                        "" + FmtStr(txtCustCode.Text.Trim()) + ", " +
                        "" + FmtStr(txtCustCodeTo.Text.Trim()) + ", " +
                        //adien
                        //"" + FmtStr(ctrlEntityCustMaster1.txtCM.Text.Trim()) + ", " +
                        //"" + FmtStr(ctrlBranchCustMaster1.txtBranchIdCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + ", " +
                        //adien
                        "" + FmtStr(ctrlDistrikCustMaster1.txtDistrikCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlDistrikCustMaster1.txtBeatCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlDistrikCustMaster1.txtSubBeatCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlGroupOutlet1.txtGroupCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlCounterTipeMC1.txtCounterType.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlKalsifikasiCustMaster1.txtKlasifikasiCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlKategoriCustMaster1.txtKategoriCM.Text.Trim()) + ", " +
                        "" + FmtStr(txtCustName.Text.Trim()) + ", " +
                        "" + FmtStr(txtShortName.Text.Trim()) + ", " +
                        "" + FmtStr(txtCustNameT3.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress1T3.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress2T3.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress3T3.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress4T3.Text.Trim()) + ", " +
                        "" + FmtStr(txtPostCodeT3.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCityBil1.txtKabupatenCMCBil.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCityBil1.txtPropinsiCMCBil.Text.Trim()) + ", " +
                        "" + FmtStr(txtCountryT3.Text.Trim()) + ", " +
                        "" + FmtStr(txtTeleponT3.Text.Trim()) + ", " +
                        "" + FmtStr(txtFaxT3.Text.Trim()) + ", " +
                        "" + FmtStr(txtContactPersonT3.Text.Trim()) + ", " +
                        "" + FmtStr(txtCustNameT4.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress1T4.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress2T4.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress3T4.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress4T4.Text.Trim()) + ", " +
                        "" + FmtStr(txtPostCodeT4.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCity1.txtKabupatenCMC.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCity1.txtPropinsiCMC.Text.Trim()) + ", " +
                        "" + FmtStr(txtCountryT4.Text.Trim()) + ", " +
                        "" + FmtStr(txtTelephoneT4.Text.Trim()) + ", " +
                        "" + FmtStr(txtFaxT4.Text.Trim()) + ", " +
                        "" + FmtStr(txtContactPersonT4.Text.Trim()) + ", " +
                        "" + FmtStr(GetEditValueText(CBShipnBill)) + ", " +
                        "" + FmtStr(txtShiptoCustCode.Text.Trim()) + ", " +
                        "" + FmtStr(txtShiptoCustCodeTo.Text.Trim()) + ", " +
                        "" + FmtStr(txtBilltoCustCode.Text.Trim()) + ", " +
                        "" + FmtStr(txtBilltoCustCodeTo.Text.Trim()) + ", " +
                        " '', " +
                        "" + FmtStr(txtTaxIdT2.Text.Trim()) + ", " +
                        //"" + ((txtLimitKreditT1.Text != "") ? "" + txtLimitKreditT1.Text + "" : "0") + ", " +
                        "" + removeFormat(txtLimitKreditT1.Text.Trim()) + ", " +
                        "" + tmpCMTOPIDDesc + ", " + //cm_top 
                        "" + FmtStr(txtGroupHargaT1.Text.Trim()) + ", " +
                        "" + FmtStr(updtdt1) + ", " +//date
                        "'" + clsLogin.USERID + "', " +//USER
                        "" + FmtStr(txtWHCode.Text.Trim()) + ", " +
                        "" + FmtStr(txtWHCodeTo.Text.Trim()) + ", " +
                        "" + FmtStr(cb4.Trim()) + ", " + //cm_wh_created                  
                                                         //"" + FmtStr(txtNpwpT2.Text.Trim()) + ", " +
                        "" + FmtStr(txtNpwpT2Masked.Text.Trim()) + ", " +
                        "" + FmtStr(txtEmailT3.Text.Trim()) + ", " +
                        "" + FmtStr(txtEmailT4.Text.Trim()) + ", " +
                        "" + FmtStr(txtTremPaymentT1.Text.Trim()) + ", " + //cm_top_id  
                        "" + FmtStr(cb4.Trim()) + ", " +
                        "" + FmtStr(txtTaxNameT2.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress1T2.Text.Trim()) + ", " +
                        "" + FmtStr(txtAddress2T2.Text.Trim()) + ", " +
                        "" + FmtStr(txtPostCodeT2.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCityTax1.txtKabupatenCMC.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCityTax1.txtPropinsiCMC.Text.Trim()) + ", " +
                        "" + FmtStr(txtCountryT2.Text.Trim()) + ", " +
                        "" + FmtStr(txtTelephoneT2.Text.Trim()) + ", " +
                        "" + FmtStr(txtFaxT2.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCity1.txtKecamatan.Text.Trim()) + ", " +
                        "" + FmtStr(popUpProvinsinCity1.txtKelurahan.Text.Trim()) + ", " +
                        "" + FmtStr(txtBarcode.Text.Trim()) + ", " +
                        "" + FmtStr(cb2.Trim()) + ", " +
                        "" + FmtStr(GetEditValueText(CBPembyaranFktT1)) + ", " +   // cm_payment_type                                                  
                        "" + FmtStr(cb3.Trim()) + ", " +
                        //"'" + _clsGlobal.DateToDB(dateTimePicker1.Text) + "', " + //cm_outlet_open
                        "'" + mulaiush + "', " + //cm_outlet_open
                        "" + FmtStr(GetEditValueText(CBBagunanT1)) + ", " +
                        //"'" + _clsGlobal.DateToDB(dateTimePicker2.Text) + "', " + //cm_birthdate
                        "'" + birthday + "', " + //cm_birthdate
                        "" + FmtStr(txtMengageniT1.Text.Trim()) + ", " +
                        "" + FmtStr(txtBankCodeT2.Text.Trim()) + ", " +
                        "" + FmtStr(txtNoReffSupplierT2.Text.Trim()) + ", " +
                        "" + FmtStr(txtNIK.Text.Trim()) + ", " +
                        "" + FmtStr(txtBankGroupT2.Text.Trim()) + ", " +
                        "" + FmtStr(txtAccountNoT2.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlGroupDiscount1.txtGroupDiscountT1.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlLokasiDesc1.txtLokasiT1.Text.Trim()) + ", " +//cm_location
                        "" + FmtStr(txtIndustriT1.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlKodepasarDesc.txtKodePasarT1.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlGroupPLU1.txtGroupPLUT1.Text.Trim()) + ", " +
                        "" + FmtStr(txtGroupKonversiT1.Text.Trim()) + ", " +
                        "" + FmtStr(cb1.Trim()) + "," +
                        "" + FmtStr(cb5.Trim()) + "," +//cm_invoice_flaq                           
                        "" + FmtStr(txtPoCustCode.Text.Trim()) + ", " +
                        "" + ((txtLeadtime.Text != "") ? "'" + txtLeadtime.Text + "'" : "0") + ", " +
                        "" + FmtStr(cb8.Trim()) + ", " + //cm_inv_recalc_flag
                        "" + FmtStr(cb6.Trim()) + ", " +
                        "" + FmtStr(GetEditValueText(comboBoxPajakT1)) + ", " +
                        "" + FmtStr(cb7.Trim()) + ", " + // cm_top_by_cust
                        "" + FmtStr(txtMoidCode.Text.Trim()) + ", " +//100
                        "" + FmtStr(txtMerchanid.Text.Trim()) + ", " +//cm_merchant_id 

                        //"" + FmtStr("O") + ", " + // @SPcm_active_flag
                        "" + FmtStr(statusCode) + ", " + // @SPcm_active_flag

                        " '00', " +
                        "" + FmtStr(txtNoPekanAwalT1.Text.Trim()) + ", " +
                        "" + FmtStr(_clsGlobal.DateToDB(txtTglRegisT1.Text.Trim())) + ", " +
                        " null," +
                        "" + FmtStr(txtLat.Text.Trim()) + ", " +
                        "" + FmtStr(txtLong.Text.Trim()) + ",  " +
                        "" + FmtStr(cb9.Trim()) + ",  " +
                        "" + FmtStr(txtShipArea.Text.Trim()) + ",  " +
                        //"" + FmtStr(cb10.Trim()) + "  ";

                        "" + FmtStr(cb10.Trim()) + ",  " +
                        "" + ((txtDelivByDays.Text.Trim() != "") ? "'" + txtDelivByDays.Text.Trim() + "'" : "0") + ", " +
                        "" + CBDelivByDay.EditValue + ", " +
                        "" + FmtStr(skillidcode) + ", " +
                        "" + FmtStr(GetEditValueText(cbPemilikNIK)) + ", " +
                        "" + FmtStr(txtBillBranch.Text.ToString()) + ", " +
                        "" + FmtStr(txtSPPKP.Text.ToString().Trim()) + ", " +
                        "" + FmtStr(cb11.Trim()) + ", " +
                        "'" + insfrom + "', " +
                        "'" + insto + "', " +
                        "'" + dtOpenHour.DateTime.ToString("HH:mm:ss") + "', " +
                        "'" + dtCloseHour.DateTime.ToString("HH:mm:ss") + "', " +
                        "'" + nmPriority.Value.ToString() + "', " +
                        "'" + nmUnloadingTime.Value.ToString() + "'";
                    if (__isAsk27918)
                    {
                        strSQL += "," + FmtStr(GetEditValueText(cbAddressChoice)) + " ";
                    }
                    if (__isAsk28151)
                    {
                        strSQL += "," + FmtStr(__cbfullfilment) + " ";
                    }
                    strSQL += "," + FmtStr(txtDelvZone.Text) + " ";
                    strSQL += "," + FmtStr(txtShipPlant.Text) + " ";
                    strSQL += "," + FmtStr(txtLatNew.Text) + " ";
                    strSQL += "," + FmtStr(txtLongNew.Text) + " ";
                    strSQL += "," + FmtStr(txtLatEcom.Text) + " ";
                    strSQL += "," + FmtStr(txtLongEcom.Text) + " , ";
                    strSQL += "'" + popUpCOA1.EntityId + "', "; //gl entity
                    strSQL += "'" + popUpCOA1.BranchId + "', "; //gl branch
                    strSQL += "'" + popUpCOA1.DivisionId + "', "; // gl division
                    strSQL += "'" + popUpCOA1.DepartmentId + "', "; //gl departmen
                    strSQL += "'" + popUpCOA1.Major1 + "', "; //gl major1
                    strSQL += "'" + popUpCOA1.Major2 + "', "; //gl major2
                    strSQL += "'" + popUpCOA1.Minor + "', "; //gl minor
                    strSQL += "'" + popUpCOA1.Analisys + "', "; //gl analysis
                    strSQL += "'" + popUpCOA1.Filler + "', ";  //gl filler
                    strSQL += "'" + getHottTaxCodes() + "', ";
                    if (!string.IsNullOrEmpty(GetEditValueText(cmbSource)))
                    {
                        strSQL += "'" + GetEditValueText(cmbSource) + "', ";  //cm_source
                    }
                    else
                    {
                        strSQL = strSQL + "'',";  //cm_source
                    }

                    if (cbFakturPajakFlag.Checked)
                    {
                        strSQL = strSQL + "'Y',"; //cm_flag_tukar_faktur
                    }
                    else
                    {
                        strSQL = strSQL + "'N',"; //cm_flag_tukar_faktur
                    }
                    strSQL = strSQL + "'" + GetEditValueText(cbJenisIdentitas) + "',"; //cm_jenis_identitas
                    if (cbDearNotification.Checked)
                    {
                        strSQL = strSQL + "'Y'"; //cm_dear_notification
                    }
                    else
                    {
                        strSQL = strSQL + "'N'"; //cm_dear_notification
                    }

                    if (!string.IsNullOrWhiteSpace(txtDurationOfDays.Text.Trim()))
                    {
                        strSQL = strSQL + "," + txtDurationOfDays.Text.Trim();
                    }
                    _clsGlobal.BeginTrans();
                    _clsGlobal.ExecuteTrans(strSQL);
                    UpdateCustomerMasterEnhancementFlagsTrans();
                    _clsGlobal.CommitTrans();

                }
                catch (Exception ex)
                {
                    _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //save hirarki
                try
                {
                    SaveHierarchyByPrdLine();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                try
                {
                    SaveSalesmanCoverage();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // price group
                try
                {
                    SavePriceGroup();
                }
                catch (Exception ex)
                {
                    ShowValidationError(tabPage8, ex.Message, MessageBoxIcon.Error);
                    return;
                }

                if (!SavePartnerFunction())
                    return;

                if (__isFlagKAM)
                {
                    bool __isTrnsTOPByPrdLine = false;
                    try
                    {
                        DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
                        if (__source == null) throw new Exception("Source mapping top by product line is null.");
                        if (__source.Rows.Count < 1) throw new Exception("Mapping top by product line tidak boleh kosong.");

                        IEnumerable<DataRow> __rows = __source.Rows.Cast<DataRow>();
                        if (__rows.Any(x => !x["prdline_code"].ToString().IsNullOrEmptyOrWhiteSpace() && x["top_code"].ToString().IsNullOrEmptyOrWhiteSpace()))
                            throw new Exception("Silakan Masukan nilai TOP.\r\nJika anda menambahkan PRD Line maka wajib untuk mengisi TOP.");

                        DataTable __tvp = new DataTable();
                        __tvp.Columns.Add("cust_code1", typeof(string));
                        __tvp.Columns.Add("cust_code2", typeof(string));
                        __tvp.Columns.Add("prdline_id", typeof(string));
                        __tvp.Columns.Add("top_id", typeof(string));

                        string __code1 = txtCustCode.Text.Trim();
                        string __code2 = txtCustCodeTo.Text.Trim();
                        DataRow rnewheader = __tvp.NewRow();
                        rnewheader["cust_code1"] = __code1;
                        rnewheader["cust_code2"] = __code2;
                        rnewheader["prdline_id"] = string.Empty;
                        __tvp.Rows.Add(rnewheader);

                        foreach (DataRow r in __rows
                            .Where(x => !x["prdline_code"].ToString().IsNullOrEmptyOrWhiteSpace() && !x["top_code"].ToString().IsNullOrEmptyOrWhiteSpace()))
                        {
                            DataRow rnew = __tvp.NewRow();
                            rnew["cust_code1"] = __code1;
                            rnew["cust_code2"] = __code2;
                            rnew["prdline_id"] = r["prdline_code"].ToString();
                            rnew["top_id"] = r["top_code"].ToString();
                            __tvp.Rows.Add(rnew);
                        }

                        _clsGlobal.BeginTrans();
                        __isTrnsTOPByPrdLine = true;

                        DBSPCommand __command = new DBSPCommand(_clsGlobal.Connect, _clsGlobal.tr);
                        __command.NewParam(new DBSPParam("@ENTITYID", ctrlEntityCustMaster2.txtCM.Text.Trim()));
                        __command.NewParam(new DBSPParam("@BRANCHID", ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()));
                        __command.NewParam(new DBSPParam("@USERID", clsLogin.USERID));
                        __command.NewParamTVP(new DBSPParamTVP("@DATA", __tvp));
                        __command.CommandParameter("SP_SDM1_SAVE_TOPBYPEDLINE");
                        UpdateCustomerByDivisionEnhancementColumnsTrans(__rows, __code1, __code2);

                        _clsGlobal.CommitTrans();
                        __isTrnsTOPByPrdLine = false;
                    }
                    catch (Exception ex)
                    {
                        if (__isTrnsTOPByPrdLine)
                            _clsGlobal.RollbackTrans();
                        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                } // end of __isFlagKAM

                try //jadwal bayar
                {
                    string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
                    string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
                    string __code1 = txtCustCode.Text.Trim();
                    string __code2 = txtCustCodeTo.Text.Trim();
                    string __type = (chkDaily.Checked ? "D" : "W");
                    int __date = 0;
                    int __day = 0;
                    int __week = 0;
                    DataTable __source = (chkDaily.Checked ? (DataTable)dgvschpayD.DataSource : (DataTable)dgvschpayW.DataSource);
                    IEnumerable<DataRow> __rows = __source.Rows.Cast<DataRow>();

                    DataTable __tvp = new DataTable();
                    __tvp.Columns.Add("entity", typeof(string));
                    __tvp.Columns.Add("branch", typeof(string));
                    __tvp.Columns.Add("cust_code1", typeof(string));
                    __tvp.Columns.Add("cust_code2", typeof(string));
                    __tvp.Columns.Add("type", typeof(string));
                    __tvp.Columns.Add("date", typeof(int));
                    __tvp.Columns.Add("day", typeof(int));
                    __tvp.Columns.Add("week", typeof(int));
                    __tvp.Columns.Add("desc", typeof(string));
                    __tvp.Columns.Add("userid", typeof(string));

                    foreach (DataRow r in __rows)
                    {
                        if (__type == "D")
                        {
                            __date = Convert.ToInt32(r["tgl"]);
                        }
                        else
                        {
                            __day = (r["hari"].ToString() == "Senin" ? 1 :
                                r["hari"].ToString() == "Selasa" ? 2 :
                                r["hari"].ToString() == "Rabu" ? 3 :
                                r["hari"].ToString() == "Kamis" ? 4 :
                                r["hari"].ToString() == "Jumat" ? 5 :
                                r["hari"].ToString() == "Sabtu" ? 6 : 7);

                            __week = (r["pola1"].ToString() == "True" ? 1 :
                                r["pola2"].ToString() == "True" ? 2 :
                                r["pola3"].ToString() == "True" ? 3 : 4);
                        }

                        DataRow rnew = __tvp.NewRow();
                        rnew["entity"] = entity;
                        rnew["branch"] = branch;
                        rnew["cust_code1"] = __code1;
                        rnew["cust_code2"] = __code2;
                        rnew["type"] = __type;
                        rnew["date"] = __date;
                        rnew["day"] = __day;
                        rnew["week"] = __week;
                        rnew["desc"] = r["ket"].ToString();
                        rnew["userid"] = clsLogin.USERID;
                        __tvp.Rows.Add(rnew);
                    }

                    _clsGlobal.BeginTrans();

                    DBSPCommand __command = new DBSPCommand(_clsGlobal.Connect, _clsGlobal.tr);
                    __command.NewParamTVP(new DBSPParamTVP("@DATA", __tvp));
                    __command.CommandParameter("IP_INSERT_CUST_SCHPAY");

                    _clsGlobal.CommitTrans();
                }
                catch (Exception ex)
                {
                    _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } //jadwal bayar

                try //Product Line Blocking
                {
                    string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
                    string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
                    string __code1 = txtCustCode.Text.Trim();
                    string __code2 = txtCustCodeTo.Text.Trim();

                    DataTable __source = (DataTable)dgvPrdLineBlocking.DataSource;
                    if (__source == null) throw new Exception("Product line tidak boleh kosong.");
                    //if (__source.Rows.Count < 1) throw new Exception("Product line tidak boleh kosong.");

                    IEnumerable<DataRow> __rows = __source.Rows.Cast<DataRow>();
                    if (__rows.Any(x => !x["prdlinecode"].ToString().IsNullOrEmptyOrWhiteSpace()))
                        throw new Exception("Product line tidak boleh kosong.");

                    DataTable __tvp = new DataTable();
                    __tvp.Columns.Add("entity", typeof(string));
                    __tvp.Columns.Add("branch", typeof(string));
                    __tvp.Columns.Add("cust_code1", typeof(string));
                    __tvp.Columns.Add("cust_code2", typeof(string));
                    __tvp.Columns.Add("prdline_code", typeof(string));
                    __tvp.Columns.Add("chk_sales", typeof(string));
                    __tvp.Columns.Add("chk_retur", typeof(string));

                    foreach (DataRow r in __rows
                        .Where(x => !x["prdlinecode"].ToString().IsNullOrEmptyOrWhiteSpace()))
                    {
                        DataRow rnew = __tvp.NewRow();
                        rnew["entity"] = entity;
                        rnew["branch"] = branch;
                        rnew["cust_code1"] = __code1;
                        rnew["cust_code2"] = __code2;
                        rnew["prdline_code"] = r["prdlinecode"].ToString();
                        rnew["chk_sales"] = (bool)r["chksales"] ? "Y" : "N";
                        rnew["chk_retur"] = (bool)r["chkretur"] ? "Y" : "N";
                        __tvp.Rows.Add(rnew);
                    }

                    _clsGlobal.BeginTrans();

                    DBSPCommand __command = new DBSPCommand(_clsGlobal.Connect, _clsGlobal.tr);
                    __command.NewParamTVP(new DBSPParamTVP("@DATA", __tvp));
                    __command.CommandParameter("IP_INSERT_CUST_PRDLINE_BLOCKING");

                    _clsGlobal.CommitTrans();
                }
                catch (Exception ex)
                {
                    _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                } //Product Line Blocking

                MessageBox.Show("Save Data successfully  !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
        }

        private void SaveEdit()
        {
            //save SP_AR_CUST_MASTER
            try
            {
                string cb1, cb2, cb3, cb4, cb5, cb6, cb7, cb8, cb9, cb10, cb11, tmpCMTOPID, tmpCMTOPIDDesc = "", __cbfullfilment;
                if (checkBoxDiscountBaseT1.Checked)
                {
                    cb1 = "Y";
                }
                else
                {
                    cb1 = "N";
                }
                if (checkBoxtxtBatasLimitT1.Checked)
                {
                    cb2 = "Y";
                }
                else
                {
                    cb2 = "N";
                }
                if (checkBoxKontraBonT1.Checked)
                {
                    cb3 = "Y";
                }
                else
                {
                    cb3 = "N";
                }
                if (checkBoxWarehouseTidak.Checked)
                {
                    cb4 = "Y";
                }
                else
                {
                    cb4 = "N";
                }
                if (checkBoxFakPajakT2.Checked)
                {
                    cb5 = "Y";
                }
                else
                {
                    cb5 = "N";
                }
                if (checkBoxNpwpT1.Checked)
                {
                    cb6 = "Y";
                }
                else
                {
                    cb6 = "N";
                }
                if (checkBoxTopByCustomerT1.Checked)
                {
                    cb7 = "Y";
                }
                else
                {
                    cb7 = "N";
                }
                if (checkBoxInvoiceTidak.Checked)
                {
                    cb8 = "Y";
                }
                else
                {
                    cb8 = "N";
                }
                if (CBBlacklist.Checked)
                {
                    cb9 = "Y";
                }
                else
                {
                    cb9 = "N";
                }
                if (chkPinSL.Checked)
                {
                    cb10 = "Y";
                }
                else
                {
                    cb10 = "N";
                }
                if (checkBoxMonitoringInsuranceT1.Checked)
                {
                    cb11 = "Y";
                }
                else
                {
                    cb11 = "N";
                }

                tmpCMTOPID = txtTremPaymentT1.Text.Trim();

                if (tmpCMTOPID != "")
                {
                    if (tmpCMTOPID.Substring(0, 1) == "P")
                    {
                        tmpCMTOPIDDesc = FmtStr(txtTremPaymentT1.Text.Trim().Replace("P", ""));
                    }
                    else
                    {
                        tmpCMTOPIDDesc = FmtStr(txtTremPaymentT1.Text.Trim());
                    }
                }
                else
                {
                    tmpCMTOPIDDesc = "0";
                }

                if (__isAsk28151)
                {
                    if (cbfullfilment.Checked)
                    {
                        __cbfullfilment = "Y";
                    }
                    else
                    {
                        __cbfullfilment = "N";
                    }
                }

                //david 17 Mei 2018
                string mulaiush = _clsGlobal.DateToDB(dateTimePicker1.DateTime.ToString("MM/dd/yyyy"));
                string birthday = _clsGlobal.DateToDB(dateTimePicker2.DateTime.ToString("MM/dd/yyyy"));
                string insfrom = _clsGlobal.DateToDB(dtTglFromIns.DateTime.ToString("MM/dd/yyyy"));
                string insto = _clsGlobal.DateToDB(dtTglToIns.DateTime.ToString("MM/dd/yyyy"));
                string updtdt1 = "";
                if (txtLastUpdateT1.Text != "")
                {
                    DateTime updtdt = Convert.ToDateTime(txtLastUpdateT1.Text);
                    updtdt1 = updtdt.ToString("MM/dd/yyyy");
                }
                else
                {
                    updtdt1 = "";
                }

                if (__isAutoTunai)
                    CBPembyaranFktT1.EditValue = "T";

                string skillidcode = skillidcodetb.Text;

                string statusCode = GetStatusCode();
                if (string.IsNullOrWhiteSpace(statusCode))
                    statusCode = "O";

                strSQL = "EXEC SP_AR_CUST_MASTER '2', " +
                       "" + FmtStr(txtCustCode.Text.Trim()) + ", " +
                       "" + FmtStr(txtCustCodeTo.Text.Trim()) + ", " +
                       //adien
                       //"" + FmtStr(ctrlEntityCustMaster1.txtCM.Text.Trim()) + ", " +
                       //"" + FmtStr(ctrlBranchCustMaster1.txtBranchIdCM.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + ", " +
                       //adien
                       "" + FmtStr(ctrlDistrikCustMaster1.txtDistrikCM.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlDistrikCustMaster1.txtBeatCM.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlDistrikCustMaster1.txtSubBeatCM.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlGroupOutlet1.txtGroupCM.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlCounterTipeMC1.txtCounterType.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlKalsifikasiCustMaster1.txtKlasifikasiCM.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlKategoriCustMaster1.txtKategoriCM.Text.Trim()) + ", " +
                       "" + FmtStr(txtCustName.Text.Trim()) + ", " +
                       "" + FmtStr(txtShortName.Text.Trim()) + ", " +
                       "" + FmtStr(txtCustNameT3.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress1T3.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress2T3.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress3T3.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress4T3.Text.Trim()) + ", " +
                       "" + FmtStr(txtPostCodeT3.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCityBil1.txtKabupatenCMCBil.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCityBil1.txtPropinsiCMCBil.Text.Trim()) + ", " +
                       "" + FmtStr(txtCountryT3.Text.Trim()) + ", " +
                       "" + FmtStr(txtTeleponT3.Text.Trim()) + ", " +
                       "" + FmtStr(txtFaxT3.Text.Trim()) + ", " +
                       "" + FmtStr(txtContactPersonT3.Text.Trim()) + ", " +
                       "" + FmtStr(txtCustNameT4.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress1T4.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress2T4.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress3T4.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress4T4.Text.Trim()) + ", " +
                       "" + FmtStr(txtPostCodeT4.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCity1.txtKabupatenCMC.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCity1.txtPropinsiCMC.Text.Trim()) + ", " +
                       "" + FmtStr(txtCountryT4.Text.Trim()) + ", " +
                       "" + FmtStr(txtTelephoneT4.Text.Trim()) + ", " +
                       "" + FmtStr(txtFaxT4.Text.Trim()) + ", " +
                       "" + FmtStr(txtContactPersonT4.Text.Trim()) + ", " +
                       "" + FmtStr(GetEditValueText(CBShipnBill)) + ", " +
                       "" + FmtStr(txtShiptoCustCode.Text.Trim()) + ", " +
                       "" + FmtStr(txtShiptoCustCodeTo.Text.Trim()) + ", " +
                       "" + FmtStr(txtBilltoCustCode.Text.Trim()) + ", " +
                       "" + FmtStr(txtBilltoCustCodeTo.Text.Trim()) + ", " +
                       " '', " +
                       "" + FmtStr(txtTaxIdT2.Text.Trim()) + ", " +
                       //"" + ((txtLimitKreditT1.Text != "") ? "" + txtLimitKreditT1.Text + "" : "0") + ", " +
                       "" + removeFormat(txtLimitKreditT1.Text.Trim()) + ", " +
                       "" + tmpCMTOPIDDesc + ", " + //cm_top 
                       "" + FmtStr(txtGroupHargaT1.Text.Trim()) + ", " +
                       //"" + FmtStr(txtLastUpdateT1.Text.Trim()) + ", " +//date
                       "" + FmtStr(updtdt1) + ", " +//date
                       "'" + clsLogin.USERID + "', " +//USER
                       "" + FmtStr(txtWHCode.Text.Trim()) + ", " +
                       "" + FmtStr(txtWHCodeTo.Text.Trim()) + ", " +
                       "" + FmtStr(cb4.Trim()) + ", " + //cm_wh_created                  
                                                        //"" + FmtStr(txtNpwpT2.Text.Trim()) + ", " +
                       "" + FmtStr(txtNpwpT2Masked.Text.Trim()) + ", " +
                       "" + FmtStr(txtEmailT3.Text.Trim()) + ", " +
                       "" + FmtStr(txtEmailT4.Text.Trim()) + ", " +
                       "" + FmtStr(txtTremPaymentT1.Text.Trim()) + ", " + //cm_top_id  
                       "" + FmtStr(cb4.Trim()) + ", " +
                       "" + FmtStr(txtTaxNameT2.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress1T2.Text.Trim()) + ", " +
                       "" + FmtStr(txtAddress2T2.Text.Trim()) + ", " +
                       "" + FmtStr(txtPostCodeT2.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCityTax1.txtKabupatenCMC.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCityTax1.txtPropinsiCMC.Text.Trim()) + ", " +
                       "" + FmtStr(txtCountryT2.Text.Trim()) + ", " +
                       "" + FmtStr(txtTelephoneT2.Text.Trim()) + ", " +
                       "" + FmtStr(txtFaxT2.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCity1.txtKecamatan.Text.Trim()) + ", " +
                       "" + FmtStr(popUpProvinsinCity1.txtKelurahan.Text.Trim()) + ", " +
                       "" + FmtStr(txtBarcode.Text.Trim()) + ", " +
                       "" + FmtStr(cb2.Trim()) + ", " +
                       "" + FmtStr(GetEditValueText(CBPembyaranFktT1)) + ", " +   // cm_payment_type                                                  
                       "" + FmtStr(cb3.Trim()) + ", " +
                       //"'" + _clsGlobal.DateToDB(dateTimepicker1.text) + "', " + //cm_outlet_open
                       "'" + mulaiush + "', " + //cm_outlet_open
                       "" + FmtStr(GetEditValueText(CBBagunanT1)) + ", " +
                       //"'" + _clsGlobal.DateToDB(dateTimePicker2.Text) + "', " + //cm_birthdate
                       "'" + birthday + "', " + //cm_birthdate
                       "" + FmtStr(txtMengageniT1.Text.Trim()) + ", " +
                       "" + FmtStr(txtBankCodeT2.Text.Trim()) + ", " +
                       "" + FmtStr(txtNoReffSupplierT2.Text.Trim()) + ", " +
                       "" + FmtStr(txtNIK.Text.Trim()) + ", " +
                       "" + FmtStr(txtBankGroupT2.Text.Trim()) + ", " +
                       "" + FmtStr(txtAccountNoT2.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlGroupDiscount1.txtGroupDiscountT1.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlLokasiDesc1.txtLokasiT1.Text.Trim()) + ", " +//cm_location
                       "" + FmtStr(txtIndustriT1.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlKodepasarDesc.txtKodePasarT1.Text.Trim()) + ", " +
                       "" + FmtStr(ctrlGroupPLU1.txtGroupPLUT1.Text.Trim()) + ", " +
                       "" + FmtStr(txtGroupKonversiT1.Text.Trim()) + ", " +
                       "" + FmtStr(cb1.Trim()) + "," +
                       "" + FmtStr(cb5.Trim()) + "," +//cm_invoice_flaq                           
                       "" + FmtStr(txtPoCustCode.Text.Trim()) + ", " +
                       "" + ((txtLeadtime.Text != "") ? "'" + txtLeadtime.Text + "'" : "0") + ", " +
                       "" + FmtStr(cb8.Trim()) + ", " + //cm_inv_recalc_flag
                       "" + FmtStr(cb6.Trim()) + ", " +
                       "" + FmtStr(GetEditValueText(comboBoxPajakT1)) + ", " +
                       "" + FmtStr(cb7.Trim()) + ", " + // cm_top_by_cust
                       "" + FmtStr(txtMoidCode.Text.Trim()) + ", " +//100
                       "" + FmtStr(txtMerchanid.Text.Trim()) + ", " +//cm_merchant_id 

                       //"" + FmtStr("O") + ", " + // @SPcm_active_flag
                       "" + FmtStr(statusCode) + ", " + // @SPcm_active_flag

                       " '00', " +
                       "" + FmtStr(txtNoPekanAwalT1.Text.Trim()) + ", " +
                       "" + FmtStr(txtTglRegisT1.Text.Trim()) + ", " +
                       " null," +
                       "" + FmtStr(txtLat.Text.Trim()) + ", " +
                       "" + FmtStr(txtLong.Text.Trim()) + ",  " +
                       "" + FmtStr(cb9.Trim()) + ",  " +
                       "" + FmtStr(txtShipArea.Text.Trim()) + ",  " +
                       //"" + FmtStr(cb10.Trim()) + "  ";

                       "" + FmtStr(cb10.Trim()) + ",  " +
                       "" + ((txtDelivByDays.Text.Trim() != "") ? "'" + txtDelivByDays.Text.Trim() + "'" : "0") + ", " +
                       "" + CBDelivByDay.EditValue + ", " +
                       "" + FmtStr(skillidcode) + ", " +
                       "" + FmtStr(GetEditValueText(cbPemilikNIK)) + ", " +
                       "" + FmtStr(txtBillBranch.Text) + ", " +
                       "" + FmtStr(txtSPPKP.Text) + ", " +
                       "" + FmtStr(cb11.Trim()) + ", " +
                       "'" + insfrom + "', " +
                       "'" + insto + "', " +
                       "'" + dtOpenHour.DateTime.ToString("HH:mm:ss") + "', " +
                       "'" + dtCloseHour.DateTime.ToString("HH:mm:ss") + "', " +
                       "'" + nmPriority.Value.ToString() + "', " +
                       "'" + nmUnloadingTime.Value.ToString() + "' ";
                if (__isAsk27918)
                {
                    strSQL += "," + FmtStr(GetEditValueText(cbAddressChoice)) + " ";
                }
                if (__isAsk28151)
                {
                    strSQL += "," + FmtStr(__cbfullfilment) + " ";
                }
                strSQL += "," + FmtStr(txtDelvZone.Text) + " ";
                strSQL += "," + FmtStr(txtShipPlant.Text) + " ";
                strSQL += "," + FmtStr(txtLatNew.Text) + " ";
                strSQL += "," + FmtStr(txtLongNew.Text) + " ";
                strSQL += "," + FmtStr(txtLatEcom.Text) + " ";
                strSQL += "," + FmtStr(txtLongEcom.Text) + ", ";
                strSQL += "'" + popUpCOA1.EntityId + "', "; //gl entity
                strSQL += "'" + popUpCOA1.BranchId + "', "; //gl branch
                strSQL += "'" + popUpCOA1.DivisionId + "', "; // gl division
                strSQL += "'" + popUpCOA1.DepartmentId + "', "; //gl departmen
                strSQL += "'" + popUpCOA1.Major1 + "', "; //gl major1
                strSQL += "'" + popUpCOA1.Major2 + "', "; //gl major2
                strSQL += "'" + popUpCOA1.Minor + "', "; //gl minor
                strSQL += "'" + popUpCOA1.Analisys + "', "; //gl analysis
                strSQL += "'" + popUpCOA1.Filler + "', ";  //gl filler
                strSQL += "'" + getHottTaxCodes() + "', ";
                if (cbFakturPajakFlag.Checked)
                {
                    strSQL = strSQL + "'','Y',"; //cm_source,cm_flag_tukar_faktur
                }
                else
                {
                    strSQL = strSQL + "'','N',"; //cm_source,cm_flag_tukar_faktur
                }
                strSQL = strSQL + "'" + GetEditValueText(cbJenisIdentitas) + "',";
                if (cbDearNotification.Checked)
                {
                    strSQL = strSQL + "'Y'"; //cm_dear_notification
                }
                else
                {
                    strSQL = strSQL + "'N'"; //cm_dear_notification
                }

                if (!string.IsNullOrWhiteSpace(txtDurationOfDays.Text.Trim()))
                {
                    strSQL = strSQL + "," + txtDurationOfDays.Text.Trim();
                }

                _clsGlobal.BeginTrans();
                _clsGlobal.ExecuteTrans(strSQL);
                UpdateCustomerMasterEnhancementFlagsTrans();
                _clsGlobal.CommitTrans();
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //save hirarki
            try
            {
                SaveHierarchyByPrdLine();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            try
            {
                SaveSalesmanCoverage();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // price group
            try
            {
                SavePriceGroup();
            }
            catch (Exception ex)
            {
                ShowValidationError(tabPage8, ex.Message, MessageBoxIcon.Error);
                return;
            }

            if (!SavePartnerFunction())
                return;

            if (__isFlagKAM)
            {
                bool __isTrnsTOPByPrdLine = false;
                try
                {
                    DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
                    if (__source == null) throw new Exception("Source mapping top by product line is null.");
                    if (__source.Rows.Count < 1) throw new Exception("Mapping top by product line tidak boleh kosong.");

                    IEnumerable<DataRow> __rows = __source.Rows.Cast<DataRow>();
                    if (__rows.Any(x => !x["prdline_code"].ToString().IsNullOrEmptyOrWhiteSpace() && x["top_code"].ToString().IsNullOrEmptyOrWhiteSpace()))
                        throw new Exception("Silakan Masukan nilai TOP.\r\nJika anda menambahkan PRD Line maka wajib untuk mengisi TOP.");

                    DataTable __tvp = new DataTable();
                    __tvp.Columns.Add("cust_code1", typeof(string));
                    __tvp.Columns.Add("cust_code2", typeof(string));
                    __tvp.Columns.Add("prdline_id", typeof(string));
                    __tvp.Columns.Add("top_id", typeof(string));

                    string __code1 = txtCustCode.Text.Trim();
                    string __code2 = txtCustCodeTo.Text.Trim();
                    DataRow rnewheader = __tvp.NewRow();
                    rnewheader["cust_code1"] = __code1;
                    rnewheader["cust_code2"] = __code2;
                    rnewheader["prdline_id"] = string.Empty;
                    __tvp.Rows.Add(rnewheader);

                    foreach (DataRow r in __rows
                        .Where(x => !x["prdline_code"].ToString().IsNullOrEmptyOrWhiteSpace() && !x["top_code"].ToString().IsNullOrEmptyOrWhiteSpace()))
                    {
                        DataRow rnew = __tvp.NewRow();
                        rnew["cust_code1"] = __code1;
                        rnew["cust_code2"] = __code2;
                        rnew["prdline_id"] = r["prdline_code"].ToString();
                        rnew["top_id"] = r["top_code"].ToString();
                        __tvp.Rows.Add(rnew);
                    }

                    _clsGlobal.BeginTrans();
                    __isTrnsTOPByPrdLine = true;

                    DBSPCommand __command = new DBSPCommand(_clsGlobal.Connect, _clsGlobal.tr);
                    __command.NewParam(new DBSPParam("@ENTITYID", ctrlEntityCustMaster2.txtCM.Text.Trim()));
                    __command.NewParam(new DBSPParam("@BRANCHID", ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()));
                    __command.NewParam(new DBSPParam("@USERID", clsLogin.USERID));
                    __command.NewParamTVP(new DBSPParamTVP("@DATA", __tvp));
                    __command.CommandParameter("SP_SDM1_SAVE_TOPBYPEDLINE");
                    UpdateCustomerByDivisionEnhancementColumnsTrans(__rows, __code1, __code2);

                    _clsGlobal.CommitTrans();
                    __isTrnsTOPByPrdLine = false;
                }
                catch (Exception ex)
                {
                    if (__isTrnsTOPByPrdLine)
                        _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            } // end of __isFlagKAM

            try //jadwal bayar
            {
                string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
                string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
                string __code1 = txtCustCode.Text.Trim();
                string __code2 = txtCustCodeTo.Text.Trim();
                string __type = (chkDaily.Checked ? "D" : "W");
                int __date = 0;
                int __day = 0;
                int __week = 0;
                DataTable __source = (chkDaily.Checked ? (DataTable)dgvschpayD.DataSource : (DataTable)dgvschpayW.DataSource);
                IEnumerable<DataRow> __rows = __source.Rows.Cast<DataRow>();

                DataTable __tvp = new DataTable();
                __tvp.Columns.Add("entity", typeof(string));
                __tvp.Columns.Add("branch", typeof(string));
                __tvp.Columns.Add("cust_code1", typeof(string));
                __tvp.Columns.Add("cust_code2", typeof(string));
                __tvp.Columns.Add("type", typeof(string));
                __tvp.Columns.Add("date", typeof(int));
                __tvp.Columns.Add("day", typeof(int));
                __tvp.Columns.Add("week", typeof(int));
                __tvp.Columns.Add("desc", typeof(string));
                __tvp.Columns.Add("userid", typeof(string));

                foreach (DataRow r in __rows)
                {
                    if (__type == "D")
                    {
                        __date = Convert.ToInt32(r["tgl"]);
                    }
                    else
                    {
                        __day = (r["hari"].ToString() == "Senin" ? 1 :
                            r["hari"].ToString() == "Selasa" ? 2 :
                            r["hari"].ToString() == "Rabu" ? 3 :
                            r["hari"].ToString() == "Kamis" ? 4 :
                            r["hari"].ToString() == "Jumat" ? 5 :
                            r["hari"].ToString() == "Sabtu" ? 6 : 7);

                        __week = (r["pola1"].ToString() == "True" ? 1 :
                            r["pola2"].ToString() == "True" ? 2 :
                            r["pola3"].ToString() == "True" ? 3 : 4);
                    }

                    DataRow rnew = __tvp.NewRow();
                    rnew["entity"] = entity;
                    rnew["branch"] = branch;
                    rnew["cust_code1"] = __code1;
                    rnew["cust_code2"] = __code2;
                    rnew["type"] = __type;
                    rnew["date"] = __date;
                    rnew["day"] = __day;
                    rnew["week"] = __week;
                    rnew["desc"] = r["ket"].ToString();
                    rnew["userid"] = clsLogin.USERID;
                    __tvp.Rows.Add(rnew);
                }

                _clsGlobal.BeginTrans();

                DBSPCommand __command = new DBSPCommand(_clsGlobal.Connect, _clsGlobal.tr);
                __command.NewParamTVP(new DBSPParamTVP("@DATA", __tvp));
                __command.CommandParameter("IP_INSERT_CUST_SCHPAY");

                _clsGlobal.CommitTrans();
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } //jadwal bayar

            try //Product Line Blocking
            {
                string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
                string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
                string __code1 = txtCustCode.Text.Trim();
                string __code2 = txtCustCodeTo.Text.Trim();

                DataTable __source = (DataTable)dgvPrdLineBlocking.DataSource;
                if (__source == null) throw new Exception("Product line tidak boleh kosong.");
                //if (__source.Rows.Count < 1) throw new Exception("Product line tidak boleh kosong.");

                IEnumerable<DataRow> __rows = __source.Rows.Cast<DataRow>();
                if (__rows.Any(x => x["prdlinecode"].ToString().IsNullOrEmptyOrWhiteSpace()))
                    throw new Exception("Product line tidak boleh kosong.");

                DataTable __tvp = new DataTable();
                __tvp.Columns.Add("entity", typeof(string));
                __tvp.Columns.Add("branch", typeof(string));
                __tvp.Columns.Add("cust_code1", typeof(string));
                __tvp.Columns.Add("cust_code2", typeof(string));
                __tvp.Columns.Add("prdline_code", typeof(string));
                __tvp.Columns.Add("chk_sales", typeof(string));
                __tvp.Columns.Add("chk_retur", typeof(string));

                foreach (DataRow r in __rows)
                {
                    DataRow rnew = __tvp.NewRow();
                    rnew["entity"] = entity;
                    rnew["branch"] = branch;
                    rnew["cust_code1"] = __code1;
                    rnew["cust_code2"] = __code2;
                    rnew["prdline_code"] = r["prdlinecode"].ToString();
                    rnew["chk_sales"] = r["chksales"].ToString().ToUpper().Equals("TRUE") ? "Y" : "N";
                    rnew["chk_retur"] = r["chkretur"].ToString().ToUpper().Equals("TRUE") ? "Y" : "N";
                    __tvp.Rows.Add(rnew);
                }

                _clsGlobal.BeginTrans();

                DBSPCommand __command = new DBSPCommand(_clsGlobal.Connect, _clsGlobal.tr);
                __command.NewParamTVP(new DBSPParamTVP("@DATA", __tvp));
                __command.CommandParameter("IP_INSERT_CUST_PRDLINE_BLOCKING");

                _clsGlobal.CommitTrans();
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } //Product Line Blocking

            MessageBox.Show("Data successfully update !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private bool isValidGroupPrice()
        {
            try
            {
                if (dgvGroupHargaView != null)
                {
                    dgvGroupHargaView.PostEditor();
                    dgvGroupHargaView.UpdateCurrentRow();
                    dgvGroupHargaView.CloseEditor();

                    for (int i = 0; i < dgvGroupHargaView.DataRowCount; i++)
                    {
                        string grpCode = Convert.ToString(
                            dgvGroupHargaView.GetRowCellValue(i, "grp_code")
                        ).Trim();

                        if (grpCode.Equals("STD", StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }

                DataTable dt = dgvGroupHarga.DataSource as DataTable;
                if (dt != null)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row.RowState == DataRowState.Deleted) continue;

                        string grpCode = Convert.ToString(row["grp_code"]).Trim();

                        if (grpCode.Equals("STD", StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }
            catch
            {
                // kalau ada error, anggap tidak valid supaya tidak salah save
            }

            return false;
        }

        private void checkBoxFakPajakT2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxFakPajakT2.Checked)
            {
                checkBoxFakPajakT2.Text = "Standart";

                if (checkBoxNpwpT1.Checked)
                {
                    //txtNpwpT2.Enabled = true;
                    txtNpwpT2Masked.Enabled = true;
                }
            }
            else
            {
                checkBoxFakPajakT2.Text = "Sederhana";
                //txtNpwpT2.Enabled = false;
                //txtNpwpT2.Text = "";
                txtNpwpT2Masked.Enabled = false;
                txtNpwpT2Masked.Text = "000000000000000";
            }
        }

        private void BTNTaxid_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Tax  ";
            frm.Query = "select t_tax_id [Tax], t_tax [Desc]from IM_TAX ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtTaxIdT2.Text = frm.ArrField[0].Trim();
                txtTaxIdT2To.Text = frm.ArrField[1].Trim();
            }
        }

        private void BTNBank_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Bank Group  ";
            frm.Query = "select bg_group_id, bg_group from CB_BANK_GROUP";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtBankGroupT2.Text = frm.ArrField[0].Trim();
            }
        }

        private void BTNbankcode_Click(object sender, EventArgs e)
        {
            loadBankc(txtBankGroupT2.Text);
        }

        void loadBankc(string bsCode)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Bank ";
            frm.Query = "select a.bm_bank_id  as Bank,a.bm_bank as Description " +
                         " from CB_BANK_MASTER as a " +
                         " where bm_bank_id = '" + bsCode + "'" +
                         " order by bm_bank_id asc";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtBankCodeT2.Text = frm.ArrField[0].Trim();
            }
        }

        private void txtBarcode_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtLimitKreditT1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtNoPekanAwalT1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void DVGSalesman_EditingControlShowing(object sender, System.Windows.Forms.DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column1_KeyPress);
            if (GetGridCurrentCell(DVGSalesman, DVGSalesmanView).ColumnIndex == 9) //Desired Column
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column1_KeyPress);
                }
            }
        }

        private void Column1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtGroupHargaT1_Leave(object sender, EventArgs e)
        {

            try
            {
                strSQL = "";
                strSQL = " select ppl_price_list_code , ppl_price_list_description   from OB_PRICE_LIST_HEADER as a where a.ppl_price_list_code = '" + txtGroupHargaT1.Text.Trim() + "' order by a.ppl_price_list_code asc ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    txtGpHargaToT1.Text = dt.Rows[0]["ppl_price_list_description"].ToString().Trim();
                }
                else
                {
                    txtGroupHargaT1.Text = string.Empty;
                    txtGpHargaToT1.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtIndustriT1_Leave(object sender, EventArgs e)
        {
            try
            {
                strSQL = "";
                strSQL = " select ind_grp_disc  , ind_sort_desc from TBL_SD_INDUSTRY as a where a.ind_grp_disc = '" + txtIndustriT1.Text.Trim() + "' order by a.ind_grp_disc asc ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    txtIndustriToT1.Text = dt.Rows[0]["ind_sort_desc"].ToString().Trim();
                }
                else
                {
                    txtIndustriT1.Text = string.Empty;
                    txtIndustriToT1.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtGroupKonversiT1_Leave(object sender, EventArgs e)
        {
            try
            {
                strSQL = "";
                strSQL = " select grc_grp_convertion_id  , grc_sort_desc   from TBL_SD_GRPCONVERTION  as a where a.grc_grp_convertion_id = '" + txtGroupKonversiT1.Text.Trim() + "' order by a.grc_grp_convertion_id asc ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    txtGroupKonversiToT1.Text = dt.Rows[0]["grc_sort_desc"].ToString().Trim();
                }
                else
                {
                    txtGroupKonversiToT1.Text = string.Empty;
                    txtGroupKonversiT1.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnPopUpGroupKonversiT1_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Group Konversi ";
            frm.Query = "select grc_grp_convertion_id [Konversi], grc_sort_desc [Description] from TBL_SD_GRPCONVERTION  ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtGroupKonversiT1.Text = frm.ArrField[0].Trim();
                txtGroupKonversiToT1.Text = frm.ArrField[1].Trim();
            }
        }

        private void txtTremPaymentT1_Leave(object sender, EventArgs e)
        {
            try
            {
                strSQL = "";
                strSQL = " select pptc_term_code  , pptc_term_desc  , pptc_no_of_days from PO_PAYMENT_TERM_CODES   as a where a.pptc_term_code = '" + txtTremPaymentT1.Text.Trim() + "' order by a.pptc_term_code asc ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    txtTremPaymentToT1.Text = dt.Rows[0]["pptc_term_desc"].ToString().Trim();
                }
                else
                {
                    txtTremPaymentT1.Text = string.Empty;
                    txtTremPaymentToT1.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveBackList()
        {
            if (txtReason.Text != "")
            {
                string cb9;
                if (CBBlacklist.Checked)
                {
                    cb9 = "Y";
                }
                else
                {
                    cb9 = "N";
                }
                //save BlackList
                try
                {
                    strSQL = " select * from TBL_CUST_BLACKLIST_HIS where cbh_cust_code1 =" + FmtStr(txtCustCode.Text.Trim()) + " and cbh_cust_code2 =" + FmtStr(txtCustCodeTo.Text.Trim()) + " and cbh_flag='Y' ";

                    if (_clsGlobal.ExecDT(strSQL).Rows.Count > 0)
                    {

                        strSQL = "EXEC SP_BLACKLIST " +
                        "" + FmtStr(txtCustCode.Text.Trim()) + ", " +
                        "" + FmtStr(txtCustCodeTo.Text.Trim()) + ", " +
                        "" + FmtStr(cb9.Trim()) + ", " +
                        "" + FmtStr(txtReason.Text.Trim()) + ", " +
                        "'" + clsLogin.USERID + "' ";

                        _clsGlobal.BeginTrans();
                        _clsGlobal.ExecuteTrans(strSQL);
                        _clsGlobal.CommitTrans();
                    }
                    else
                    {
                        if (CBBlacklist.Checked)
                        {
                            strSQL = "EXEC SP_BLACKLIST " +
                            "" + FmtStr(txtCustCode.Text.Trim()) + ", " +
                            "" + FmtStr(txtCustCodeTo.Text.Trim()) + ", " +
                            "" + FmtStr(cb9.Trim()) + ", " +
                            "" + FmtStr(txtReason.Text.Trim()) + ", " +
                            "'" + clsLogin.USERID + "' ";

                            _clsGlobal.BeginTrans();
                            _clsGlobal.ExecuteTrans(strSQL);
                            _clsGlobal.CommitTrans();
                        }
                    }

                }
                catch (Exception ex)
                {
                    _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void FillGridBackList()
        {
            try
            {
                strSQL = " select cbh_reason , cbh_last_update , cbh_update_by  from TBL_CUST_BLACKLIST_HIS WHERE cbh_flag = 'Y' AND  cbh_cust_code1 ='" + txtCustCode.Text + "' AND cbh_cust_code2 ='" + txtCustCodeTo.Text + "'";

                DGVBackList.DataSource = _clsGlobal.ExecDT(strSQL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FillGridUNBackList()
        {
            try
            {
                strSQL = " select cbh_reason as cbh_reason2, cbh_last_update as cbh_last_update2, cbh_update_by as cbh_update_by2 from TBL_CUST_BLACKLIST_HIS WHERE cbh_flag = 'N' AND  cbh_cust_code1 ='" + txtCustCode.Text + "' AND cbh_cust_code2 ='" + txtCustCodeTo.Text + "'";

                DGVUnBackList.DataSource = _clsGlobal.ExecDT(strSQL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DVGSalesman_CellValueChanged(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                System.Windows.Forms.DataGridViewRow dgr, dgrB;
                string column0, column1;
                string maxB = "";
                string min = "";

                try
                {
                    dgr = GetGridRows(DVGSalesman, DVGSalesmanView)[GetGridCurrentRow(DVGSalesman, DVGSalesmanView).Index];

                    if (GetGridRows(DVGSalesman, DVGSalesmanView)[e.RowIndex].Cells["csc_salesman_id"].Value.ToString() != "")
                    {
                        column0 = dgr.Cells["csc_salesman_id"].Value.ToString();

                        //if (clsGlobal.MODE_TRX == 2)
                        if (__state == StateEntry.Edit)
                        {
                            strSQL = "SELECT isnull(sgm_spgm_name, '') sgm_spgm_name,isnull(sgm_type_operasi, '') sgm_type_operasi FROM SO_SPG_GIRL_MAN WHERE sgm_entity_id='" + _prdEntityCode + "' and sgm_branch_id='" + _prdBrandCode + "' and sgm_spgm_id='" + column0 + "'";
                        }
                        //if (clsGlobal.MODE_TRX == 1)
                        if (__state == StateEntry.New)
                        {
                            strSQL = "SELECT isnull(sgm_spgm_name, '') sgm_spgm_name,isnull(sgm_type_operasi, '') sgm_type_operasi FROM SO_SPG_GIRL_MAN WHERE sgm_entity_id='" + ctrlEntityCustMaster2.txtCM.Text.Trim() + "' and sgm_branch_id='" + ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim() + "' and sgm_spgm_id='" + column0 + "'";
                        }
                        DataTable dt = _clsGlobal.ExecDT(strSQL);
                        if (dt.Rows.Count > 0)
                        {
                            try
                            {
                                if (GetGridRowCount(DVGSalesman, DVGSalesmanView) > 1)
                                {
                                    dgrB = GetGridRows(DVGSalesman, DVGSalesmanView)[GetGridCurrentRow(DVGSalesman, DVGSalesmanView).Index - 1];
                                    maxB = dgrB.Cells["csc_salesman_id"].Value.ToString();
                                }
                                else
                                {
                                    dgrB = null;
                                }
                                if (dgrB != null)
                                {
                                    if (CheckSalesmanDuplicate())
                                    {
                                        dgr.Cells["csc_salesman_id"].Value = "";
                                        dgr.Cells["sgm_spgm_name"].Value = "";
                                        dgr.Cells["sgm_type_operasi"].Value = "";

                                        ShowValidationError(tabPage6, "Salesman Sudah ada ", MessageBoxIcon.Error);
                                        if (dgr.Cells["csc_salesman_id"].Value.ToString() == "")
                                        {
                                            dgr.Cells["csc_salesman_id"].Value = "";
                                            dgr.Cells["sgm_spgm_name"].Value = "";
                                            dgr.Cells["sgm_type_operasi"].Value = "";
                                        }
                                        dgr.Cells["csc_salesman_id"].Value = "";
                                    }
                                    else
                                    {
                                        dgr.Cells["sgm_spgm_name"].Value = dt.Rows[0][0].ToString();
                                        dgr.Cells["sgm_type_operasi"].Value = dt.Rows[0][1].ToString();
                                    }
                                }
                                else
                                {
                                    dgr.Cells["sgm_spgm_name"].Value = dt.Rows[0][0].ToString();
                                    dgr.Cells["sgm_type_operasi"].Value = dt.Rows[0][1].ToString();
                                }

                            }
                            catch (Exception ex) { }
                        }
                        else
                        {
                            dgr.Cells["csc_salesman_id"].Value = "";
                            dgr.Cells["sgm_spgm_name"].Value = "";
                            dgr.Cells["sgm_type_operasi"].Value = "";
                        }
                    }
                    else
                    {
                        dgr.Cells["sgm_spgm_name"].Value = string.Empty;
                        dgr.Cells["sgm_type_operasi"].Value = string.Empty;
                    }
                }
                finally
                {
                    dgr = null;
                    column0 = null;
                }
            }
        }

        private bool CheckSalesmanDuplicate()
        {
            bool isDuplicate = false;
            int i, x;

            for (i = 0; i < GetGridRows(DVGSalesman, DVGSalesmanView).Count; i++)
            {
                //cek nilai yg null atau kosong
                if ((GetGridRows(DVGSalesman, DVGSalesmanView)[i].Cells["csc_salesman_id"].Value == null) || (GetGridRows(DVGSalesman, DVGSalesmanView)[i].Cells["sgm_spgm_name"].Value == null) || (GetGridRows(DVGSalesman, DVGSalesmanView)[i].Cells["sgm_type_operasi"].Value == null))
                {
                    isDuplicate = true;
                    break;
                }

                string empEnt = GetGridRows(DVGSalesman, DVGSalesmanView)[i].Cells["csc_salesman_id"].Value.ToString().Trim();
                string empBrn = GetGridRows(DVGSalesman, DVGSalesmanView)[i].Cells["sgm_spgm_name"].Value.ToString().Trim();
                string empDiv = GetGridRows(DVGSalesman, DVGSalesmanView)[i].Cells["sgm_type_operasi"].Value.ToString().Trim();

                //cek duplicate data
                for (x = 0; x < GetGridRows(DVGSalesman, DVGSalesmanView).Count; x++)
                {
                    if (i != x)//data yg dipilih tidak usah di compare
                    {
                        if (((empEnt == GetGridRows(DVGSalesman, DVGSalesmanView)[x].Cells["csc_salesman_id"].Value.ToString().Trim()) && (empBrn == GetGridRows(DVGSalesman, DVGSalesmanView)[x].Cells["sgm_spgm_name"].Value.ToString().Trim()) && (empDiv == GetGridRows(DVGSalesman, DVGSalesmanView)[x].Cells["sgm_type_operasi"].Value.ToString().Trim())) == true)
                        {
                            isDuplicate = true;
                            break;
                        }
                    }
                }
            }
            return isDuplicate;
        }

        private void txtGroupHargaT1_TextChanged(object sender, EventArgs e)
        {

            getGroupHarga(txtGroupHargaT1);
        }

        private void getGroupHarga(System.Windows.Forms.Control objText)
        {
            try
            {
                strSQL = " select ppl_price_list_code , ppl_price_list_description  from OB_PRICE_LIST_HEADER where ppl_price_list_code = '" + objText.Text.Trim() + "'";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    txtGpHargaToT1.Text = dt.Rows[0]["ppl_price_list_description"].ToString().Trim();
                }
                else
                {
                    txtGpHargaToT1.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DVGSalesman_Scroll(object sender, ScrollEventArgs e)
        {
            Rectangle rtHeader = this.DVGSalesman.DisplayRectangle;
            this.DVGSalesman.Invalidate(rtHeader);
        }

        private void DVGSalesman_CellPainting(object sender, System.Windows.Forms.DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex > -1)
            {
                Rectangle r2 = e.CellBounds;
                r2.Y += e.CellBounds.Height / 2;
                r2.Height = e.CellBounds.Height / 2;
                e.PaintBackground(r2, true);
                e.PaintContent(r2);
                e.Handled = true;
            }
        }

        private void DVGSalesman_ColumnWidthChanged(object sender, EventArgs e)
        {
            Rectangle rtHeader = this.DVGSalesman.DisplayRectangle;
            this.DVGSalesman.Invalidate(rtHeader);
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            // Setelah tab diganti ke DevExpress.XtraTab.XtraTabControl, event bridge lama
            // membuat TabControlEventArgs dengan TabPage = null. Karena itu jangan langsung
            // membaca e.TabPage.Name; ambil halaman aktif dari XtraTabControl sebagai fallback.
            string selectedPageName = null;

            if (e != null && e.TabPage != null)
            {
                selectedPageName = e.TabPage.Name;
            }
            else if (tabControl1 != null && tabControl1.SelectedTabPage != null)
            {
                selectedPageName = tabControl1.SelectedTabPage.Name;
            }

            if (string.IsNullOrEmpty(selectedPageName))
                return;

            if (selectedPageName == tabPage1.Name ||
                selectedPageName == tabPage2.Name ||
                selectedPageName == tabPage3.Name ||
                selectedPageName == tabPage4.Name ||
                selectedPageName == tabPage5.Name ||
                selectedPageName == tabPage6.Name)
            {
                txtReason.Visible = false;
                label18.Visible = false;
            }
        }

        private void tabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            tabControl1_Selected(sender, null);
        }

        private void DGVBackList_RowPostPaint(object sender, System.Windows.Forms.DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(Color.Black))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void DGVUnBackList_RowPostPaint(object sender, System.Windows.Forms.DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(Color.Black))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void txtCustCode_TextChanged(object sender, EventArgs e)
        {
            IsiTextBill(CBShipnBill.ItemIndex, txtCustCode.Text);
        }

        private void txtLeadtime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtTaxIdT2_TextChanged(object sender, EventArgs e)
        {
            getTaxID(txtTaxIdT2);
        }

        private void getTaxID(System.Windows.Forms.Control objText)
        {
            try
            {
                strSQL = " select t_tax_id , t_tax from IM_TAX as a where a.t_tax_id = '" + objText.Text.Trim() + "' order by a.t_tax_id asc ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    txtTaxIdT2To.Text = dt.Rows[0]["t_tax"].ToString().Trim();
                }
                else
                {

                    txtTaxIdT2To.Text = string.Empty;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtTaxIdT2_Leave(object sender, EventArgs e)
        {
            if (txtTaxIdT2To.Text.Trim() == "")
            {
                txtTaxIdT2.Text = string.Empty;
            }
        }

        private void txtTremPaymentT1_TextChanged(object sender, EventArgs e)
        {
            LoadTOP(txtTremPaymentT1);
        }

        void LoadTOP(System.Windows.Forms.Control objText)
        {
            try
            {
                strSQL = "";
                strSQL = " select pptc_term_code  , pptc_term_desc  , pptc_no_of_days from PO_PAYMENT_TERM_CODES   as a where a.pptc_term_code = '" + objText.Text.Trim() + "' order by a.pptc_term_code asc ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count > 0)
                {
                    txtTremPaymentToT1.Text = dt.Rows[0]["pptc_term_desc"].ToString().Trim();
                }
                else
                {
                    txtTremPaymentT1.Text = string.Empty;
                    txtTremPaymentToT1.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void eARCustMasterEntry_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void txtNIK_KeyPress(object sender, KeyPressEventArgs e)
        {
            _clsGlobal.NumericOnly(sender, e);
        }

        private void tbNumeric_Enter(object sender, EventArgs e)
        {
            string __data = ((System.Windows.Forms.Control)sender).Text;
            ((System.Windows.Forms.Control)sender).Text = string.Format("{0:#.##}", removeFormat(__data));
        }
        private void tbNumeric_Leave(object sender, EventArgs e)
        {
            string __data = ((System.Windows.Forms.Control)sender).Text.Trim();
            try
            {
                decimal __value = 0;
                decimal.TryParse(__data, out __value);
                __data = string.Format("{0:#,##0.00}", __value);
            }
            catch (Exception) { }
            ((System.Windows.Forms.Control)sender).Text = __data;
        }

        private object removeFormat(string __text)
        {
            try
            {
                decimal __value = 0;
                decimal.TryParse(__text, out __value);
                return __value;
            }
            catch (Exception) { }
            return 0;
        }

        private void dgvGroupHarga_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Insert)
            {
                System.Windows.Forms.DataGridViewRow dgr;
                string column0, column1;
                try
                {
                    dgr = GetGridCurrentRow(dgvGroupHarga, dgvGroupHargaView);

                    if (GetGridRows(dgvGroupHarga, dgvGroupHargaView).Count > 0)
                    {

                        column0 = dgr.Cells["grp_code"].Value.ToString();
                        column1 = dgr.Cells["grp_desc"].Value.ToString();

                        if (column0 != "")
                        {
                            //if (clsGlobal.MODE_TRX == 1)
                            if (__state == StateEntry.New)
                            {
                                GetGridRows(dgvGroupHarga, dgvGroupHargaView).Add("", "", "", "", "", "");
                            }
                            else
                            {
                                DataTable dt = dgvGroupHarga.DataSource as DataTable;
                                dt.Rows.Add();
                                dgvGroupHarga.DataSource = dt;
                            }
                        }
                    }
                    else
                    {
                        //if (clsGlobal.MODE_TRX == 1)
                        if (__state == StateEntry.New)
                        {
                            GetGridRows(dgvGroupHarga, dgvGroupHargaView).Add("", "", "", "", "", "");
                        }
                        else
                        {
                            if (GetGridRowCount(dgvGroupHarga, dgvGroupHargaView) > 0)
                            {
                                GetGridRows(dgvGroupHarga, dgvGroupHargaView).Add("", "", "", "", "", "");
                            }
                            else
                            {
                                DataTable dt = dgvGroupHarga.DataSource as DataTable;
                                dt.Rows.Add();
                                dgvGroupHarga.DataSource = dt;
                            }
                        }
                    }
                }
                finally
                {
                    dgr = null;
                    column0 = null;
                }
            }
            if (e.KeyCode == Keys.Delete)
            {
                if (GetGridRowCount(dgvGroupHarga, dgvGroupHargaView) != 0)
                {
                    GetGridRows(dgvGroupHarga, dgvGroupHargaView).Remove(GetGridCurrentRow(dgvGroupHarga, dgvGroupHargaView));
                }
            }
        }

        private void dgvGroupHarga_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var senderGrid = sender as DevExpress.XtraGrid.GridControl;
            if (GetFocusedColumnIndex(senderGrid) == dgvGroupHargaView.Columns["btnGroup"].VisibleIndex)
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Group Price ";
                frm.Query = "select ppl_price_list_code [Group Price Code], ppl_price_list_description [Description] from OB_PRICE_LIST_HEADER  ";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    try
                    {
                        System.Windows.Forms.DataGridViewRow dgr, dgrB;
                        string maxB = "";
                        string min = "";
                        dgrB = null;

                        dgr = GetGridRows(dgvGroupHarga, dgvGroupHargaView)[GetGridCurrentRow(dgvGroupHarga, dgvGroupHargaView).Index];

                        if (GetGridRowCount(dgvGroupHarga, dgvGroupHargaView) > 1)
                        {
                            if (dgr.Index != 0)
                            {
                                dgrB = GetGridRows(dgvGroupHarga, dgvGroupHargaView)[GetGridCurrentRow(dgvGroupHarga, dgvGroupHargaView).Index - 1];
                                maxB = dgrB.Cells["grp_code"].Value.ToString();
                            }
                            else
                            {
                                min = frm.ArrField[0].Trim();
                                rowDuplicate = false;
                                exitData = false;

                                foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(dgvGroupHarga, dgvGroupHargaView))
                                {
                                    if ((gRow.Cells["grp_code"].Value.ToString().Trim() == min))
                                    {
                                        rowDuplicate = true;
                                        break;
                                    }
                                }
                                if (rowDuplicate == true)
                                {
                                    ShowValidationError(tabPage8, "Group Code " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Group Price yang berbeda ", MessageBoxIcon.Error);
                                    if (dgr.Cells["grp_code"].Value.ToString() == "")
                                    {
                                        dgr.Cells["grp_code"].Value = "";
                                        dgr.Cells["grp_desc"].Value = "";
                                        exitData = true;
                                    }
                                    else
                                    {
                                        dgr.Cells["grp_code"].Value = "";
                                        dgr.Cells["grp_desc"].Value = "";
                                        exitData = true;
                                    }
                                }
                                else
                                {
                                    GetGridRows(dgvGroupHarga, dgvGroupHargaView)[e.RowIndex].Cells["grp_code"].Value = frm.ArrField[0].Trim();
                                    GetGridRows(dgvGroupHarga, dgvGroupHargaView)[e.RowIndex].Cells["grp_desc"].Value = frm.ArrField[1].Trim();
                                }
                            }
                        }
                        else
                        {
                            dgrB = null;
                        }
                        if (dgrB != null)
                        {
                            min = frm.ArrField[0].Trim();

                            rowDuplicate = false;

                            foreach (System.Windows.Forms.DataGridViewRow gRow in GetGridRows(dgvGroupHarga, dgvGroupHargaView))
                            {
                                if ((gRow.Cells["grp_code"].Value.ToString().Trim() == min))
                                {
                                    rowDuplicate = true;
                                    break;
                                }
                            }

                            if (rowDuplicate == true)
                            {
                                ShowValidationError(tabPage8, "Group Price " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Group Prie yang berbeda ", MessageBoxIcon.Error);
                                if (dgr.Cells["grp_code"].Value.ToString() == "")
                                {
                                    dgr.Cells["grp_code"].Value = "";
                                    dgr.Cells["grp_desc"].Value = "";
                                }
                                else
                                {
                                    dgr.Cells["grp_code"].Value = "";
                                    dgr.Cells["grp_desc"].Value = "";
                                }
                            }
                            else
                            {
                                GetGridRows(dgvGroupHarga, dgvGroupHargaView)[e.RowIndex].Cells["grp_code"].Value = frm.ArrField[0].Trim();
                                GetGridRows(dgvGroupHarga, dgvGroupHargaView)[e.RowIndex].Cells["grp_desc"].Value = frm.ArrField[1].Trim();
                            }
                        }
                        else
                        {
                            if (exitData != true)
                            {
                                GetGridRows(dgvGroupHarga, dgvGroupHargaView)[e.RowIndex].Cells["grp_code"].Value = frm.ArrField[0].Trim();
                                GetGridRows(dgvGroupHarga, dgvGroupHargaView)[e.RowIndex].Cells["grp_desc"].Value = frm.ArrField[1].Trim();
                            }
                        }

                    }
                    catch (Exception ex) { }
                }
            }
        }

        private void btnShipArea_Click(object sender, EventArgs e)
        {
            //frmPopUp frm = new frmPopUp();
            //frm.FrmText = "Search Shipment Area";
            //frm.Query = " select msa_area_code [Kode Area], msa_long_desc [Description] from TBL_SD_MST_SHIP_AREA";
            ////frm.Query = " select msa_area_code [Kode Area], msa_long_desc [Description] from TBL_SD_MST_SHIP_AREA where msa_entity_id = '" + ctrlEntityCustMaster2.txtCM.Text + "' and msa_branch_id = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "'";
            //frm.ShowDialog();
            //if (frm.ArrField != null)
            //{
            //    txtShipArea.Text = frm.ArrField[0].Trim();
            //    txtShipAreaDesc.Text = frm.ArrField[1].Trim();
            //}
            DialogResult dr;
            dr = MessageBox.Show("apakah anda ingin melakukan syncronize dengan table ship area detail ?", "Konfirmasi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Text))
                {
                    strSQL = "";
                    strSQL = "select gh_function_code from GS_GEN_HARDCODED WITH (NOLOCK) where gh_function_name = 'EXCLUDE_SYNC_OUTLET_TYPE_SHIP_AREA' and gh_sys = 'H' and gh_function_code = '" + ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Text + "'";
                    DataTable dtCheckSyncEx = _clsGlobal.ExecDT(strSQL);
                    if (dtCheckSyncEx.Rows.Count > 0)
                    {
                        ShowValidationError(ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet, "Data Shipment area tidak bisa di Sync mohon pilih mode manual.", MessageBoxIcon.Information);
                        return;
                    }
                }

                strSQL = "";
                strSQL = "SELECT * FROM VW_SHIP_AREA_DETAIL_LOOKUP where msd_entity_id = '" + ctrlEntityCustMaster2.txtCM.Text + "' and msd_branch_id = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "' and kl_prov_code = '" + popUpProvinsinCity1.txtPropinsiCMC.Text + "' and kl_city_code = '" + popUpProvinsinCity1.txtKabupatenCMC.Text + "' and kl_kecamatan_code = '" + popUpProvinsinCity1.txtKecamatan.Text + "' and kl_kelurahan_code = '" + popUpProvinsinCity1.txtKelurahan.Text + "' ";
                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count == 0)
                {
                    ShowValidationError(txtShipArea, "Data Shipment area tidak ditemukan untuk lokasi " + popUpProvinsinCity1.txtPropinsiCMC.Text + popUpProvinsinCity1.txtKabupatenCMC.Text + popUpProvinsinCity1.txtKecamatan.Text + popUpProvinsinCity1.txtKelurahan.Text + "", MessageBoxIcon.Information);
                    txtShipArea.Text = "";
                    txtShipAreaDesc.Text = "";
                }
                else
                {
                    txtShipArea.Text = dt.Rows[0]["msd_area_code"].ToString();
                    txtShipAreaDesc.Text = dt.Rows[0]["msd_long_desc"].ToString();
                }

            }
            else
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Shipment Area";
                frm.Query = " select msa_area_code [Kode Area], msa_long_desc [Description] from TBL_SD_MST_SHIP_AREA";
                //frm.Query = " select msa_area_code [Kode Area], msa_long_desc [Description] from TBL_SD_MST_SHIP_AREA where msa_entity_id = '" + ctrlEntityCustMaster2.txtCM.Text + "' and msa_branch_id = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "'";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    txtShipArea.Text = frm.ArrField[0].Trim();
                    txtShipAreaDesc.Text = frm.ArrField[1].Trim();
                }
            }
        }

        private void txtShipArea_TextChanged(object sender, EventArgs e)
        {
            getShipArea(txtShipArea);
        }

        private void getShipArea(System.Windows.Forms.Control objText)
        {
            try
            {
                strSQL = " select msa_area_code, msa_long_desc from TBL_SD_MST_SHIP_AREA where msa_area_code = '" + objText.Text.Trim() + "'";
                //strSQL = " select msa_area_code, msa_long_desc from TBL_SD_MST_SHIP_AREA where msa_area_code = '" + objText.Text.Trim() + "' and  msa_entity_id = '" + ctrlEntityCustMaster2.txtCM.Text + "' and msa_branch_id = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "' ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    txtShipAreaDesc.Text = dt.Rows[0]["msa_long_desc"].ToString().Trim();
                }
                else
                {
                    txtShipAreaDesc.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtShipArea_Leave(object sender, EventArgs e)
        {
            if (txtShipAreaDesc.Text.Trim() == "")
            {
                txtShipArea.Text = string.Empty;
            }
        }

        private void BTNShipcust_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Bill ";
            frm.Query = GetActiveCustomerLookupQuery();
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtShiptoCustCode.Text = frm.ArrField[2].Trim();
                txtShiptoCustCodeTo.Text = frm.ArrField[3].Trim();
            }
        }

        private void numberTopByPrdLine()
        {
            try
            {
                DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
                int i = 1;
                foreach (DataRow r in __source.Rows)
                {
                    r["no"] = i;
                    i++;
                }
            }
            catch (Exception ex) { throw ex; }
        }
        private void createTopByPrdLine()
        {
            GridColumn col0 = new GridColumn()
            {
                Name = "nocol",
                Caption = "No.",
                FieldName = "no",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col1 = new GridColumn()
            {
                Name = "prdlinecol",
                Caption = "Division",
                FieldName = "prdline_code",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,

            };
            GridColumn col1b = new GridColumn()
            {
                Name = "prdlinebtncol",
                Caption = "...",
                Width = 40,
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col2 = new GridColumn()
            {
                Name = "desccol",
                Caption = "Nama Division",
                FieldName = "prdline_desc",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            //GridColumn col3 = new GridColumn()
            //{
            //    Name = "topcol",
            //    Caption = "TOP",
            //            //            //    FieldName = "top",
            //            //    SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            //            //    DisplayMember = "Nama",
            //    ValueMember = "Kode",
            //            //};
            GridColumn col3 = new GridColumn()
            {
                Name = "topcol",
                Caption = "TOP",
                FieldName = "top_code",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,

            };
            GridColumn col3b = new GridColumn()
            {
                Name = "topbtncol",
                Caption = "...",
                Width = 40,
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col4 = new GridColumn()
            {
                Name = "topdesccol",
                Caption = "Nama TOP",
                FieldName = "top_desc",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col5 = new GridColumn()
            {
                Name = "plantcol",
                Caption = "Plant",
                FieldName = "plant_id",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col5b = new GridColumn()
            {
                Name = "plantbtncol",
                Caption = "",
                Width = 40,
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col6 = new GridColumn()
            {
                Name = "salesemployeecol",
                Caption = "Sales employee",
                FieldName = "sales_employee_id",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col6b = new GridColumn()
            {
                Name = "salesemployeebtncol",
                Caption = "",
                Width = 40,
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            topbyprdlinedgvView.Columns.AddRange(new GridColumn[] { col0, col1, col1b, col2, col3, col3b, col4, col5, col5b, col6, col6b });

            DataTable __source = new DataTable();
            foreach (GridColumn c in topbyprdlinedgvView.Columns.Cast<GridColumn>().Where(x => !string.IsNullOrEmpty(x.FieldName)))
                if (!__source.Columns.Contains(c.FieldName))
                    __source.Columns.Add(c.FieldName, GetGridColumnType(c.FieldName));
            EnsureTopByDivisionDataColumns(__source);
            __source.Rows.Add(__source.NewRow());
            topbyprdlinedgv.DataSource = __source;
            numberTopByPrdLine();
            ConfigureTopByPrdLineGridRuntime();
        }

        private void topbyprdlinedgv_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Insert || e.KeyCode == Keys.F10)
                {
                    InsertTopByDivisionRow();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Delete)
                {
                    DataTable __source = topbyprdlinedgv.DataSource as DataTable;
                    if (__source == null) return;

                    int rowHandle = topbyprdlinedgvView == null ? -1 : topbyprdlinedgvView.FocusedRowHandle;
                    if (rowHandle < 0) return;

                    DataRow row = topbyprdlinedgvView.GetDataRow(rowHandle);
                    if (row == null) return;

                    __source.Rows.Remove(row);
                    numberTopByPrdLine();
                    topbyprdlinedgv.RefreshDataSource();
                    if (topbyprdlinedgvView != null) topbyprdlinedgvView.RefreshData();

                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.F4)
                {
                    if (GetGridCurrentCell(topbyprdlinedgv, topbyprdlinedgvView) == null) return;
                    topbyprdlinedgv_CellContentClick(sender, new System.Windows.Forms.DataGridViewCellEventArgs(2, GetGridCurrentCell(topbyprdlinedgv, topbyprdlinedgvView).RowIndex));
                }
            }
            catch (Exception) { }
        }

        private void InsertTopByDivisionRow()
        {
            DataTable __source = topbyprdlinedgv.DataSource as DataTable;
            if (__source == null)
            {
                __source = CreateTopByDivisionDataTable();
                topbyprdlinedgv.DataSource = __source;
            }

            EnsureTopByDivisionDataColumns(__source);

            DataRow row = __source.NewRow();
            row["no"] = __source.Rows.Count + 1;
            row["prdline_code"] = string.Empty;
            row["prdline_btn"] = string.Empty;
            row["prdline_desc"] = string.Empty;
            row["top_code"] = string.Empty;
            row["top_btn"] = string.Empty;
            row["top_desc"] = string.Empty;
            row["plant_id"] = string.Empty;
            row["plant_btn"] = string.Empty;
            row["plant_desc"] = string.Empty;
            row["sales_employee_id"] = string.Empty;
            row["sales_employee_btn"] = string.Empty;
            row["sales_employee_desc"] = string.Empty;
            __source.Rows.Add(row);

            numberTopByPrdLine();
            ConfigureTopByPrdLineGridRuntime();
            topbyprdlinedgv.RefreshDataSource();
            topbyprdlinedgvView.RefreshData();

            int rowHandle = topbyprdlinedgvView.GetRowHandle(__source.Rows.Count - 1);
            if (rowHandle >= 0)
            {
                topbyprdlinedgvView.FocusedRowHandle = rowHandle;
                if (topbyprdlinedgvView.Columns["prdlinebtncol"] != null)
                    topbyprdlinedgvView.FocusedColumn = topbyprdlinedgvView.Columns["prdlinebtncol"];
            }
        }

        private DataTable CreateTopByDivisionDataTable()
        {
            DataTable table = new DataTable();
            EnsureTopByDivisionDataColumns(table);
            return table;
        }

        private void EnsureTopByDivisionDataColumns(DataTable table)
        {
            if (table == null) return;
            if (!table.Columns.Contains("no")) table.Columns.Add("no", typeof(int));
            if (!table.Columns.Contains("prdline_code")) table.Columns.Add("prdline_code", typeof(string));
            if (!table.Columns.Contains("prdline_btn")) table.Columns.Add("prdline_btn", typeof(string));
            if (!table.Columns.Contains("prdline_desc")) table.Columns.Add("prdline_desc", typeof(string));
            if (!table.Columns.Contains("top_code")) table.Columns.Add("top_code", typeof(string));
            if (!table.Columns.Contains("top_btn")) table.Columns.Add("top_btn", typeof(string));
            if (!table.Columns.Contains("top_desc")) table.Columns.Add("top_desc", typeof(string));
            if (!table.Columns.Contains("plant_id")) table.Columns.Add("plant_id", typeof(string));
            if (!table.Columns.Contains("plant_btn")) table.Columns.Add("plant_btn", typeof(string));
            if (!table.Columns.Contains("plant_desc")) table.Columns.Add("plant_desc", typeof(string));
            if (!table.Columns.Contains("sales_employee_id")) table.Columns.Add("sales_employee_id", typeof(string));
            if (!table.Columns.Contains("sales_employee_btn")) table.Columns.Add("sales_employee_btn", typeof(string));
            if (!table.Columns.Contains("sales_employee_desc")) table.Columns.Add("sales_employee_desc", typeof(string));
        }
        private void topbyprdlinedgv_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                string name = topbyprdlinedgvView.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinebtncol"))
                {
                    string key = GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "prdlinecol", e.RowIndex).Value.ToString();
                    __prdlines.ActionSelector(key, true, new IndexGrid(e));
                }
                else if (name.CompareC("topbtncol"))
                {
                    string key = GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "topcol", e.RowIndex).Value.ToString();
                    __top.ActionSelector(key, true, new IndexGrid(e));
                }
                else if (name.CompareC("plantbtncol"))
                {
                    ShowTopByDivisionRepositoryPopup("plant_btn");
                }
                else if (name.CompareC("salesemployeebtncol"))
                {
                    ShowTopByDivisionRepositoryPopup("sales_employee_btn");
                }
            }
            catch (Exception) { }
        }
        private bool checkExistCode(object value, object key)
        {
            if (value == null) return false;
            if (value.ToString().IsNullOrEmptyOrWhiteSpace()) return false;
            return value.Equals(key);
        }

        //static int indexRow = 0;
        private void topbyprdlinedgv_CellValueChanged(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                string name = topbyprdlinedgvView.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinecol"))
                {
                    string key = GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, e.ColumnIndex, e.RowIndex).Value.ToString();
                    __prdlines.ActionSelector(key, false, new IndexGrid(e));
                }
                else if (name.CompareC("topcol"))
                {
                    string key = GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, e.ColumnIndex, e.RowIndex).Value.ToString();
                    __top.ActionSelector(key, false, new IndexGrid(e));
                }
            }
            catch (Exception) { }
        }
        private void _EndPrdLine(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    IndexGrid e = (IndexGrid)__result.State;
                    if (__result.IsError)
                    {
                        GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "prdlinecol", e.Row).Value = string.Empty;
                        GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "desccol", e.Row).Value = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (__item == null) return;

                        DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
                        if (__source == null) return;
                        if (__source.Rows.Cast<DataRow>().Count(x => checkExistCode(x["prdline_code"], __item.Kode)) > 1)
                        {
                            GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "prdlinecol", e.Row).Value = string.Empty;
                            GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "desccol", e.Row).Value = string.Empty;
                        }
                        else
                        {
                            GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "prdlinecol", e.Row).Value = __item.Kode;
                            GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "desccol", e.Row).Value = __item.Nama;
                        }
                    }
                }));
            }
            catch (Exception) { }
        }
        private void _EndPrdLineBlocking(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    IndexGrid e = (IndexGrid)__result.State;
                    if (__result.IsError)
                    {
                        GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, "prdlinecol", e.Row).Value = string.Empty;
                        GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, "prdlinedesccol", e.Row).Value = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (__item == null) return;

                        DataTable __source = (DataTable)dgvPrdLineBlocking.DataSource;
                        if (__source == null) return;
                        if (__source.Rows.Cast<DataRow>().Count(x => checkExistCode(x["prdlinecode"], __item.Kode)) > 1)
                        {
                            GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, "prdlinecol", e.Row).Value = string.Empty;
                            GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, "prdlinedesccol", e.Row).Value = string.Empty;
                        }
                        else
                        {
                            GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, "prdlinecol", e.Row).Value = __item.Kode;
                            GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, "prdlinedesccol", e.Row).Value = __item.Nama;
                        }
                    }
                }));
            }
            catch (Exception) { }
        }
        private void _EndTOP(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    IndexGrid e = (IndexGrid)__result.State;
                    if (__result.IsError)
                    {
                        GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "topcol", e.Row).Value = string.Empty;
                        GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "topdesccol", e.Row).Value = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (__item == null) return;
                        GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "topcol", e.Row).Value = __item.Kode;
                        GetGridCell(topbyprdlinedgv, topbyprdlinedgvView, "topdesccol", e.Row).Value = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        private void _EndSkillID(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                        this.skilliddesctb.Text = string.Empty;
                    else
                    {
                        __skill.IsPostBack = __result.IsClick;
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (__result.Index == 0)
                        {
                            this.skillidcodetb.Text = __item.Kode;
                            this.skilliddesctb.Text = __item.Nama;
                        }
                        __skill.IsPostBack = false;
                    }
                }));
            }
            catch (Exception) { }
        }

        private void topbyprdlinedgv_EditingControlShowing(object sender, EventArgs e)
        {
            if (GetGridCurrentCell(topbyprdlinedgv, topbyprdlinedgvView) == null) return;
            if (!topbyprdlinedgvView.Columns[GetGridCurrentCell(topbyprdlinedgv, topbyprdlinedgvView).ColumnIndex].Name.CompareC("prdlinecol")) return;
            if (sender is TextBox)
                ((TextBox)(sender)).CharacterCasing = CharacterCasing.Upper;
        }

        private void topbyprdlinedgv_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            ((DevExpress.XtraGrid.Views.Grid.GridView)topbyprdlinedgv.MainView).PostEditor();
        }

        private void txtDelivByDays_TextChanged(object sender, EventArgs e)
        {
            int i = 0;
            int.TryParse(txtDelivByDays.Text.Trim(), out i);
            if (i > 0) CBDelivByDay.EditValue = 0;
        }

        private void CBDelivByDay_SelectedValueChanged(object sender, EventArgs e)
        {
            int i = 0;
            int.TryParse(GetEditValueText(CBDelivByDay).Trim(), out i);
            if (i > 0) txtDelivByDays.Text = "0";
        }

        string angka = "0123456789";
        private void txtDelivByDays_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            //    e.Handled = true;
            //e.Handled = (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'));
            e.Handled = !angka.Any(x => char.IsControl(e.KeyChar) || e.KeyChar.Equals(x));
        }


        private bool is_valid_NIK(string nik)
        {
            if (string.IsNullOrEmpty(nik) || nik.Trim().ToString() == "1111111111111111" || nik.Trim().ToString() == "0000000000000001" || nik.Trim().ToString() == "9999999999999999" || nik.Trim().ToString() == "1111111111111110" || nik.Trim().ToString() == "0000000000000001")
            {
                return false;
            }
            // untuk perhitungan valid ga validnya 
            string datenik = nik.Substring(6, 6);
            DateTime now = DateTime.Now;
            int yearnow = Convert.ToInt32(now.ToString("yy"));
            int yearbirth = Convert.ToInt32(datenik.Substring(4, 2));
            string prefixyear = "20";
            if (yearbirth > yearnow)
            {
                prefixyear = "19";
            }

            int date = Convert.ToInt32(datenik.Substring(0, 2));
            if ((date - 40) > 0)
            {
                date = date - 40;
            }

            string tgl = prefixyear + yearbirth + "-" + datenik.Substring(2, 2).ToString() + "-" + date;

            try
            {
                DateTime dtBirth = Convert.ToDateTime(tgl);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private void NormalizeConvertedDevExpressUi()
        {
            NormalizeConvertedDevExpressUi(this);
        }

        private void NormalizeConvertedDevExpressUi(Control parent)
        {
            if (parent == null) return;

            BaseEdit editor = parent as BaseEdit;
            if (editor != null)
                NormalizeConvertedDevExpressEditor(editor);

            foreach (Control child in parent.Controls)
                NormalizeConvertedDevExpressUi(child);
        }

        private void NormalizeConvertedDevExpressEditor(BaseEdit editor)
        {
            if (editor == null || editor.Properties == null) return;

            SetRepositoryStringProperty(editor.Properties, "NullText", string.Empty);
            SetRepositoryStringProperty(editor.Properties, "NullValuePrompt", string.Empty);

            Color backColor = editor.Properties.Appearance.BackColor;
            if (backColor.IsEmpty || backColor == Color.Transparent)
                backColor = editor.BackColor;

            Color foreColor = IsDarkUiColor(backColor) ? Color.White : SystemColors.ControlText;
            editor.Properties.Appearance.ForeColor = foreColor;
            editor.Properties.Appearance.Options.UseForeColor = true;
            editor.Properties.AppearanceDisabled.ForeColor = foreColor;
            editor.Properties.AppearanceDisabled.Options.UseForeColor = true;
            editor.Properties.AppearanceReadOnly.ForeColor = foreColor;
            editor.Properties.AppearanceReadOnly.Options.UseForeColor = true;
        }

        private void NormalizeTaxTabLayout()
        {
            if (tabPage2 == null) return;

            try
            {
                tabPage2.AutoScroll = true;
                tabPage2.AutoScrollMinSize = new Size(0, 425);

                if (txtAccountNoT2 != null)
                    txtAccountNoT2.Size = new Size(324, 20);

                if (label109 != null)
                    label109.Location = new Point(5, 261);

                if (popUpCOA1 != null)
                {
                    popUpCOA1.Location = new Point(79, 255);
                    popUpCOA1.Size = new Size(324, 23);
                    popUpCOA1.Anchor = AnchorStyles.Left | AnchorStyles.Top;
                }

                if (tabControl2 != null)
                {
                    tabControl2.Visible = true;
                    tabControl2.Location = new Point(7, 286);
                    tabControl2.Size = new Size(371, 132);
                    tabControl2.Anchor = AnchorStyles.Left | AnchorStyles.Top;
                }

                if (cbJenisIdentitas != null)
                    cbJenisIdentitas.Size = new Size(253, 20);
            }
            catch
            {
                // Perapihan layout tidak boleh mengganggu load form.
            }
        }

        private static void SetRepositoryStringProperty(object repositoryItem, string propertyName, string value)
        {
            if (repositoryItem == null || string.IsNullOrWhiteSpace(propertyName)) return;

            try
            {
                System.Reflection.PropertyInfo property = repositoryItem.GetType().GetProperty(propertyName);
                if (property != null && property.CanWrite && property.PropertyType == typeof(string))
                    property.SetValue(repositoryItem, value, null);
            }
            catch
            {
                // Perbedaan tipe repository DevExpress tidak boleh mengganggu load form.
            }
        }

        private static bool IsDarkUiColor(Color color)
        {
            if (color.IsEmpty || color == Color.Transparent)
                return false;

            return color.GetBrightness() < 0.45f
                || color.ToArgb() == SystemColors.GrayText.ToArgb()
                || color.ToArgb() == Color.Gray.ToArgb()
                || color.ToArgb() == Color.DimGray.ToArgb()
                || color.ToArgb() == Color.DarkGray.ToArgb();
        }

        private bool is_valid_Address(string address)
        {
            bool result = false;
            string __add = address.ToUpper();
            if (
                __add.Contains("NO")
                || __add.Contains("NOMOR")
                || __add.Contains("BLK")
                || __add.Contains("BLOK")
                || __add.Contains("LT")
                || __add.Contains("LANTAI")
                )
            {
                result = true;
            }

            Regex reg = new Regex(@"\d");
            if (reg.IsMatch(__add))
            {
                if (
                  __add.Contains("NO")
                  || __add.Contains("NOMOR")
                  || __add.Contains("BLK")
                  || __add.Contains("BLOK")
                  || __add.Contains("LT")
                  || __add.Contains("LANTAI")
                  )
                {
                    result = true;
                    //if (__add.Contains("NO") || __add.Contains("NOMOR"))
                    //{
                    Regex reg2 = new Regex(@"(NO|NOMOR)( |.|  |. )([0-9]|[0-9][1-9]|[0-9][0-9][0-9]|[0-9][0-9][0-9][0-9])");
                    Regex reg3 = new Regex(@"(NO|NOMOR)(.| |. | )[#]");
                    if (reg2.IsMatch(__add))
                    {
                        result = true;
                    }
                    else if (reg3.IsMatch(__add))
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                    //}
                }
                else
                {
                    result = false;
                }
            }
            else
            {
                result = false;
            }


            if (__add.Contains("#"))
            {
                if (
                    __add.Contains("NO")
                    || __add.Contains("NOMOR")
                   )
                {
                    Regex reg3 = new Regex(@"(NO|NOMOR)(.| |. | )[#]");
                    if (reg3.IsMatch(__add))
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }
                else
                {
                    result = false;
                }
            }

            return result;
        }

        private bool is_valid_coordinate(string coordinate)
        {
            bool result = false;
            Regex reg = new Regex(@"^-?[0-9]{1,3}(?:\.[0-9]{1,15})?$");
            if (reg.IsMatch(coordinate))
            {
                result = true;
            }

            return result;
        }
        private bool is_valid_npwp(string npwp)
        {
            bool result = false;
            Regex reg = new Regex(@"^([\d]{2})[.]([\d]{3})[.]([\d]{3})[.][\d][-]([\d]{3})[.]([\d]{3})$");
            if (reg.IsMatch(npwp))
            {
                result = true;
            }
            return result;
        }
        private void SetGridViewSchPayD()
        {
            dgvschpayDView.Columns.Clear();
            int[] tgl = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31 };

            GridColumn col0 = new GridColumn()
            {
                Name = "tglcol",
                Caption = "Tanggal",
                FieldName = "tgl",
            };
            GridColumn col1 = new GridColumn()
            {
                Name = "ketcol",
                Caption = "Keterangan",
                FieldName = "ket",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            dgvschpayDView.Columns.AddRange(new GridColumn[] { col0, col1 });

            DataTable __source = new DataTable();
            foreach (GridColumn c in dgvschpayDView.Columns.Cast<GridColumn>().Where(x => !string.IsNullOrEmpty(x.FieldName)))
                __source.Columns.Add(c.FieldName, GetGridColumnType(c.FieldName));
            __source.Columns["tgl"].DefaultValue = 1;
            //__source.Rows.Add(__source.NewRow());
            dgvschpayD.DataSource = __source;
        }

        private void SetGridViewSchPayW()
        {
            dgvschpayWView.Columns.Clear();
            string[] hari = { "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu" };
            //List<Hari> hari = new List<Hari>();
            //int i = 0;

            //foreach (string nama in namahari)
            //    hari.Add(new Hari()
            //    {
            //        Code = i++,
            //        Name = nama,
            //    });

            GridColumn col0 = new GridColumn()
            {
                Name = "haricol",
                Caption = "Hari",
                FieldName = "hari",
                //ValueMember = "Code",
                //DisplayMember = "Name",
            };
            GridColumn col1 = new GridColumn()
            {
                Name = "pola1col",
                Caption = "Pola.1",
                FieldName = "pola1",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col2 = new GridColumn()
            {
                Name = "pola2col",
                Caption = "Pola.2",
                FieldName = "pola2",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col3 = new GridColumn()
            {
                Name = "pola3col",
                Caption = "Pola.3",
                FieldName = "pola3",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col4 = new GridColumn()
            {
                Name = "pola4col",
                Caption = "Pola.4",
                FieldName = "pola4",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col5 = new GridColumn()
            {
                Name = "ketcol",
                Caption = "Keterangan",
                FieldName = "ket",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            dgvschpayWView.Columns.AddRange(new GridColumn[] { col0, col1, col2, col3, col4, col5 });

            DataTable __source = new DataTable();
            foreach (GridColumn c in dgvschpayWView.Columns.Cast<GridColumn>().Where(x => !string.IsNullOrEmpty(x.FieldName)))
                __source.Columns.Add(c.FieldName, GetGridColumnType(c.FieldName));
            __source.Columns["hari"].DefaultValue = "Senin";
            __source.Columns["pola1"].DefaultValue = true;
            //__source.Rows.Add(__source.NewRow());
            dgvschpayW.DataSource = __source;
        }

        private void chkDaily_CheckedChanged(object sender, EventArgs e)
        {
            if (chkDaily.Checked)
            {
                chkWeekly.Checked = false;
                DataTable __source = (DataTable)dgvschpayW.DataSource;
                __source.Rows.Clear();
            }
        }

        private void chkWeekly_CheckedChanged(object sender, EventArgs e)
        {
            if (chkWeekly.Checked)
            {
                chkDaily.Checked = false;
                DataTable __source = (DataTable)dgvschpayD.DataSource;
                __source.Rows.Clear();
            }
        }

        private void dgvschpayD_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (chkDaily.Checked)
                {
                    DataTable __source = EnsureGridDataTable(dgvschpayD, "tgl", "ket");
                    if (e.KeyCode == Keys.Insert)
                    {
                        __source.Rows.Add(__source.NewRow());
                        dgvschpayD.RefreshDataSource();
                        dgvschpayDView.FocusedRowHandle = dgvschpayDView.GetRowHandle(__source.Rows.Count - 1);
                        e.Handled = true;
                    }
                    else if (e.KeyCode == Keys.Delete)
                    {
                        DataRow row = GetFocusedDataRow(dgvschpayDView);
                        if (row != null)
                        {
                            __source.Rows.Remove(row);
                            dgvschpayD.RefreshDataSource();
                            e.Handled = true;
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        private void dgvschpayW_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (chkWeekly.Checked)
                {
                    DataTable __source = EnsureGridDataTable(dgvschpayW, "hari", "pola1", "pola2", "pola3", "pola4", "ket");
                    if (e.KeyCode == Keys.Insert)
                    {
                        __source.Rows.Add(__source.NewRow());
                        dgvschpayW.RefreshDataSource();
                        dgvschpayWView.FocusedRowHandle = dgvschpayWView.GetRowHandle(__source.Rows.Count - 1);
                        e.Handled = true;
                    }
                    else if (e.KeyCode == Keys.Delete)
                    {
                        DataRow row = GetFocusedDataRow(dgvschpayWView);
                        if (row != null)
                        {
                            __source.Rows.Remove(row);
                            dgvschpayW.RefreshDataSource();
                            e.Handled = true;
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        private void dgvschpayW_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                if (GetGridCell(dgvschpayW, dgvschpayWView, "pola1col", e.RowIndex).Value.ToString() == "True"
                    || GetGridCell(dgvschpayW, dgvschpayWView, "pola2col", e.RowIndex).Value.ToString() == "True"
                    || GetGridCell(dgvschpayW, dgvschpayWView, "pola3col", e.RowIndex).Value.ToString() == "True"
                    || GetGridCell(dgvschpayW, dgvschpayWView, "pola4col", e.RowIndex).Value.ToString() == "True")
                {
                    GetGridCell(dgvschpayW, dgvschpayWView, "pola1col", e.RowIndex).Value = false;
                    GetGridCell(dgvschpayW, dgvschpayWView, "pola2col", e.RowIndex).Value = false;
                    GetGridCell(dgvschpayW, dgvschpayWView, "pola3col", e.RowIndex).Value = false;
                    GetGridCell(dgvschpayW, dgvschpayWView, "pola4col", e.RowIndex).Value = false;

                    GetGridCell(dgvschpayW, dgvschpayWView, e.ColumnIndex, e.RowIndex).Value = true;
                }
            }
        }

        private void viewnikbtn_Click(object sender, EventArgs e)
        {
            try
            {
                var ph = photoExist.DATA.FirstOrDefault() ?? new TiraPhoto();
                nik = photoView(ph.IsNIK, PhotoMode.NIK);
                if (nik != null)
                    photoCheck();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void viewnpwpbtn_Click(object sender, EventArgs e)
        {
            try
            {
                var ph = photoExist.DATA.FirstOrDefault() ?? new TiraPhoto();
                npwp = photoView(ph.IsNPWP, PhotoMode.NPWP);
                if (npwp != null)
                    photoCheck();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private PhotoItem photoView(bool isavail, PhotoMode mode)
        {
            try
            {
                string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
                string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
                string code1 = txtCustCode.Text.Trim();
                string code2 = txtCustCodeTo.Text.Trim();

                if (entity.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception(" Entity belum di-pilih ");

                if (branch.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception(" Branch belum di-pilih ");

                if (code1.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception("Code Customer belum di-input.");

                if (this.State == StateEntry.New && code1.Length != 6)
                    throw new Exception("Customer Code harus 6 digit/karakter ! ");

                if (code2.IsNullOrEmptyOrWhiteSpace())
                    throw new Exception(" Customer Code To Does not Allow Empty ! ");

                using (eARCustPhotoView view = new eARCustPhotoView(entity, branch, code1, code2, isavail, mode))
                    if (view.ShowDialog() == DialogResult.OK)
                        return view.Result;
                return null;

            }
            catch (Exception exc)
            { throw exc; }
        }

        private void SetGridViewPrdLineBlocking()
        {
            dgvPrdLineBlockingView.Columns.Clear();
            GridColumn col0 = new GridColumn()
            {
                Name = "nocol",
                Caption = "No.",
                FieldName = "no",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col1 = new GridColumn()
            {
                Name = "prdlinecol",
                Caption = "Product Line",
                FieldName = "prdlinecode",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,

            };
            GridColumn col2 = new GridColumn()
            {
                Name = "prdlinebtncol",
                Caption = "...",
                Width = 40,
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col3 = new GridColumn()
            {
                Name = "prdlinedesccol",
                Caption = "Nama Product Line",
                FieldName = "prdlinedesc",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col4 = new GridColumn()
            {
                Name = "chksalescol",
                Caption = "Sales",
                FieldName = "chksales",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            GridColumn col5 = new GridColumn()
            {
                Name = "chkreturcol",
                Caption = "Retur",
                FieldName = "chkretur",
                SortMode = DevExpress.XtraGrid.ColumnSortMode.Default,
            };
            dgvPrdLineBlockingView.Columns.AddRange(new GridColumn[] { col0, col1, col2, col3, col4, col5 });

            DataTable __source = new DataTable();
            foreach (GridColumn c in dgvPrdLineBlockingView.Columns.Cast<GridColumn>().Where(x => !string.IsNullOrEmpty(x.FieldName)))
                __source.Columns.Add(c.FieldName, GetGridColumnType(c.FieldName));
            __source.Columns["chksales"].DefaultValue = false;
            __source.Columns["chkretur"].DefaultValue = false;
            __source.Rows.Add(__source.NewRow());
            dgvPrdLineBlocking.DataSource = __source;
            numberPrdLineBlocking();
        }

        private void numberPrdLineBlocking()
        {
            try
            {
                DataTable __source = (DataTable)dgvPrdLineBlocking.DataSource;
                int i = 1;
                foreach (DataRow r in __source.Rows)
                {
                    r["no"] = i;
                    i++;
                }
            }
            catch (Exception ex) { throw ex; }
        }

        private void dgvPrdLineBlocking_CellValueChanged(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                string name = dgvPrdLineBlockingView.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinecol"))
                {
                    string key = GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, e.ColumnIndex, e.RowIndex).Value.ToString();
                    __prdlineblocking.ActionSelector(key, false, new IndexGrid(e));
                }
            }
            catch (Exception) { }
        }

        private void dgvPrdLineBlocking_EditingControlShowing(object sender, EventArgs e)
        {
            if (GetGridCurrentCell(dgvPrdLineBlocking, dgvPrdLineBlockingView) == null) return;
            if (!dgvPrdLineBlockingView.Columns[GetGridCurrentCell(dgvPrdLineBlocking, dgvPrdLineBlockingView).ColumnIndex].Name.CompareC("prdlinecol")) return;
            if (sender is TextBox)
                ((TextBox)(sender)).CharacterCasing = CharacterCasing.Upper;
        }

        private void dgvPrdLineBlocking_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //((DevExpress.XtraGrid.Views.Grid.GridView)topbyprdlinedgv.MainView).PostEditor();
        }

        private void dgvPrdLineBlocking_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                DataTable __source = EnsureGridDataTable(dgvPrdLineBlocking, "no", "prdlinecode", "prdlinedesc", "chksales", "chkretur");
                if (e.KeyCode == Keys.Insert)
                {
                    __source.Rows.Add(__source.NewRow());
                    numberPrdLineBlocking();
                    dgvPrdLineBlocking.RefreshDataSource();
                    dgvPrdLineBlockingView.FocusedRowHandle = dgvPrdLineBlockingView.GetRowHandle(__source.Rows.Count - 1);
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.Delete)
                {
                    DataRow row = GetFocusedDataRow(dgvPrdLineBlockingView);
                    if (row != null)
                    {
                        __source.Rows.Remove(row);
                        numberPrdLineBlocking();
                        dgvPrdLineBlocking.RefreshDataSource();
                        e.Handled = true;
                    }
                }
                else if (e.KeyCode == Keys.F4)
                {
                    DataRow row = GetFocusedDataRow(dgvPrdLineBlockingView);
                    if (row == null) return;
                    string key = Convert.ToString(row["prdlinecode"]).Trim();
                    __prdlineblocking.ActionSelector(key, true, new IndexGrid(2, dgvPrdLineBlockingView.FocusedRowHandle));
                    e.Handled = true;
                }
            }
            catch (Exception) { }
        }

        private void dgvPrdLineBlocking_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                string name = dgvPrdLineBlockingView.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinebtncol"))
                {
                    string key = GetGridCell(dgvPrdLineBlocking, dgvPrdLineBlockingView, "prdlinecol", e.RowIndex).Value.ToString();
                    __prdlineblocking.ActionSelector(key, true, new IndexGrid(e));
                }
            }
            catch (Exception) { }
        }

        private void checkBoxMonitoringInsuranceT1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMonitoringInsuranceT1.Checked)
            {
                checkBoxMonitoringInsuranceT1.Text = "Ya";
                lbltglfromins.Visible = true;
                lbltgltoins.Visible = true;
                dtTglFromIns.Visible = true;
                dtTglToIns.Visible = true;
            }
            else
            {
                checkBoxMonitoringInsuranceT1.Text = "Tidak";
                lbltglfromins.Visible = false;
                lbltgltoins.Visible = false;
                dtTglFromIns.Visible = false;
                dtTglToIns.Visible = false;
            }
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

        private void btnGeoTag_Click(object sender, EventArgs e)
        {
            eARGeoTag frm = new eARGeoTag();
            frm.FrmText = "Geotag Verification ";
            frm.CustCode = txtCustCode.Text;
            frm.Query = @"SELECT ROW_NUMBER() OVER (ORDER BY [Trx. Date]) AS [No.],*,'Open Maps' AS Map2 
                            FROM(
                                    SELECT DISTINCT b.[sgm_spgm_name] AS[Salesman], a.[pcc_doc_date] AS [Trx. Date], a.[ppc_lattitude] AS[Lat], a.[ppc_longitude] AS[Lng],
                                        'https://www.google.com/maps?q=' + a.ppc_lattitude + ',' + a.[ppc_longitude] AS[Maps]

                                    FROM TBL_PDA_CUST_CARD a JOIN SO_SPG_GIRL_MAN b ON a.pcc_spgm_id = b.sgm_spgm_id

                                    WHERE a.[ppc_lattitude] <> '' and a.[pcc_cust_id]='" + frm.CustCode + "' and 1 = 2) x";
            frm.Query2 = @"SELECT TOP 1000 ROW_NUMBER() OVER (ORDER BY [Trx. Date]) AS [No.],*,'Open Maps' AS Map2
                            FROM(
                                    SELECT DISTINCT b.[sgm_spgm_name] AS[Salesman], a.[pcc_doc_date] AS [Trx. Date], a.[ppc_lattitude] AS[Lat], a.[ppc_longitude] AS[Lng],
                                        'https://www.google.com/maps?q=' + a.ppc_lattitude + ',' + a.[ppc_longitude] AS[Maps],
                                        a.[pcc_spgm_id],a.[pcc_doc_date]

                                    FROM TBL_PDA_CUST_CARD a JOIN SO_SPG_GIRL_MAN b ON a.pcc_spgm_id = b.sgm_spgm_id

                                    WHERE a.[ppc_lattitude] <> '' and a.[pcc_cust_id]='" + frm.CustCode + "' ) x ";

            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtLat.Text = frm.ArrField[3].Trim();
                txtLong.Text = frm.ArrField[4].Trim();
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cbDearNotification_CheckedChanged(object sender, EventArgs e)
        {
            if (cbDearNotification.Checked)
            {
                cbDearNotification.Text = "Ya";
            }
            else
            {
                cbDearNotification.Text = "Tidak";
            }
        }
        private void btnDelvZone_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Delivery Site ";
            frm.Query = " select Distinct itzr_destination_site [Dest Zone],itzr_desc_destination_site [Desc] from IM_TRANS_ZONE_ROUTE where itzr_entity = '" + ctrlEntityCustMaster2.txtCM.Text + "' and itzr_branch = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "' ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtDelvZone.Text = frm.ArrField[0].Trim();
            }
        }


        private void txtDurationOfDays_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !angka.Any(x => char.IsControl(e.KeyChar) || e.KeyChar.Equals(x));
        }

        private void txtDelvZone_Leave(object sender, EventArgs e)
        {
            strSQL = "";
            strSQL = " select Distinct itzr_destination_site [Dest Zone],itzr_desc_destination_site [Desc] from IM_TRANS_ZONE_ROUTE where itzr_entity = '" + ctrlEntityCustMaster2.txtCM.Text + "' and itzr_branch = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "' and itzr_destination_site = '" + txtDelvZone.Text + "' ";

            DataTable dt = _clsGlobal.ExecDT(strSQL);
            if (dt.Rows.Count > 0)
            {
                //txtTremPaymentToT1.Text = dt.Rows[0]["pptc_term_desc"].ToString().Trim();
            }
            else
            {
                txtDelvZone.Text = string.Empty;
            }
        }

        private void txtShipPlant_TextChanged(object sender, EventArgs e)
        {
            strSQL = "";
            strSQL = "select Distinct br_branch_ud2 [Ship Plant],br_branch_desc [Desc] from GS_BRANCH where br_branch_ud2 = '" + txtShipPlant.Text + "'  ";

            DataTable dt = _clsGlobal.ExecDT(strSQL);
            if (dt.Rows.Count > 0)
            {
                //txtTremPaymentToT1.Text = dt.Rows[0]["pptc_term_desc"].ToString().Trim();
            }
            else
            {
                txtShipPlant.Text = string.Empty;
            }
        }

        private void BtnShipPlant_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Delivery Site ";
            frm.Query = " select Distinct br_branch_ud2 [Ship Plant],br_branch_desc [Desc] from GS_BRANCH ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtShipPlant.Text = frm.ArrField[0].Trim();
            }
        }

        private void gridPajak_CellContentClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    gridPajakView.FocusedRowHandle = e.RowIndex;
                    if (e.ColumnIndex >= 0 && e.ColumnIndex < gridPajakView.Columns.Count)
                        gridPajakView.FocusedColumn = gridPajakView.Columns[e.ColumnIndex];
                    ShowTaxCodePopupForFocusedRow();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridPajak_CellEndEdit(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex > -1)
                    ValidateTaxGridRow(e.RowIndex);
            }
            catch (Exception ex)
            {

            }
        }
        int max = 0;
        private void gridPajak_RowPostPaint(object sender, System.Windows.Forms.DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(Color.Black))
            {
                SizeF sif = e.Graphics.MeasureString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font);
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
                max = Math.Max(max, (int)sif.Width + 20);
            }
        }

        private void gridPajak_KeyDown(object sender, KeyEventArgs e)
        {
            gridPajakView_KeyDown(sender, e);
        }

        private string getHottTaxCodes()
        {
            FlushGridEditor(gridPajak, gridPajakView);
            List<string> taxCodes = new List<string>();
            DataTable table = EnsureGridDataTable(gridPajak, "code", "btn", "desc");
            foreach (DataRow row in table.Rows)
            {
                object codeValue = row["code"];
                if (codeValue != null && !string.IsNullOrEmpty(codeValue.ToString())) taxCodes.Add(codeValue.ToString());
            }
            return string.Join(",", taxCodes);
        }

        private void filltipePajak(string tax_codes)
        {
            try
            {
                if (clsGlobal.MODE_TRX != 2) return;

                tax_codes = tax_codes.Replace(" ", "");
                strSQL = "select stt_tax_code,stt_tax_desc from so_tax_type where stt_tax_code in " + clsGlobal.FmtStrWhereIn(tax_codes);
                DataTable dt = new DataTable();
                dt = _clsGlobal.ExecDT(strSQL);
                DataTable table = EnsureGridDataTable(gridPajak, "code", "btn", "desc");
                table.Rows.Clear();
                if (dt.Rows.Count > 0)
                {
                    //kalo ada data detail
                    if (dt.Rows[0]["stt_tax_code"].ToString().Trim() != "")
                    {

                        foreach (DataRow rw in dt.Rows)
                        {
                            DataRow newRow = table.NewRow();
                            newRow["code"] = rw["stt_tax_code"];
                            newRow["btn"] = string.Empty;
                            newRow["desc"] = rw["stt_tax_desc"];
                            table.Rows.Add(newRow);
                        }
                    }
                }
                gridPajak.RefreshDataSource();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool pCheckValue(string sValue)
        {
            return IsTaxCodeUnique(sValue, GetFocusedDataRow(gridPajakView));
        }

        private void ctrlEntityCustMaster2_Leave(object sender, EventArgs e)
        {
            if (ctrlEntityCustMaster2.txtCM.Text != "")
            {
                popUpCOA1.EntityId = ctrlEntityCustMaster2.txtCM.Text.Trim();
                popUpCOA1.GetFormatCOA(popUpCOA1.EntityId, popUpCOA1.BranchId, popUpCOA1.DivisionId, popUpCOA1.DepartmentId, popUpCOA1.Major1, popUpCOA1.Major2, popUpCOA1.Minor, popUpCOA1.Analisys, popUpCOA1.Filler);
            }
            else if (btnCancel.Focus())
            { }
            else
            {
                ShowValidationError(ctrlEntityCustMaster2.txtCM, "Please Select Branch !", MessageBoxIcon.Information);
                //popUpEntityBranch1.txtBranch.Select();
                return;
            }
        }

        private void panel1_ControlAdded(object sender, ControlEventArgs e)
        {

        }

        private void getSource()
        {
            //Tipe Operasi
            string strSQL = "";
            strSQL = " select gh_function_code, gh_function_desc AS descr from dbo.GS_GEN_HARDCODED WITH (NOLOCK) where gh_sys = 'H' AND gh_function_name = 'SOURCE_CLNT'  ORDER BY CASE WHEN gh_function_code = 'N' THEN 0 END desc";

            cmbSource.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
            cmbSource.Properties.ValueMember = "gh_function_code";
            cmbSource.Properties.DisplayMember = "descr";

        }

        private void cbFakturPajakFlag_CheckedChanged(object sender, EventArgs e)
        {
            if (cbFakturPajakFlag.Checked)
            {
                cbFakturPajakFlag.Text = "Ya";
            }
            else
            {
                cbFakturPajakFlag.Text = "Tidak";
            }
        }

        private void SaveHierarchyByPrdLine()
        {
            FlushAllGridEditors();

            DataTable dt = DGVHIR.DataSource as DataTable;

            // Penting: kalau datasource kosong, jangan delete data lama
            if (dt == null || dt.Rows.Count == 0)
                return;

            List<DataRow> validRows = dt.Rows
                .Cast<DataRow>()
                .Where(r => r.RowState != DataRowState.Deleted)
                .Where(r =>
                    !string.IsNullOrWhiteSpace(Convert.ToString(r["cpl_line_code"])) &&
                    !string.IsNullOrWhiteSpace(Convert.ToString(r["cpl_cust_type"])))
                .ToList();

            // Penting: kalau tidak ada row valid, jangan delete data lama
            if (validRows.Count == 0)
                return;

            string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
            string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
            string custCode1 = txtCustCode.Text.Trim();
            string custCode2 = txtCustCodeTo.Text.Trim();

            _clsGlobal.BeginTrans();

            try
            {
                strSQL = "EXEC SP_AR_CUST_MASTER_CUSTPRDLINE '3', " +
                         FmtStr(entity) + ", " +
                         FmtStr(branch) + ", " +
                         FmtStr(custCode1) + ", " +
                         FmtStr(custCode2) + ", NULL, NULL, NULL";

                _clsGlobal.ExecuteTrans(strSQL);

                for (int i = 0; i < validRows.Count; i++)
                {
                    string lineCode = Convert.ToString(validRows[i]["cpl_line_code"]).Trim();
                    string custType = Convert.ToString(validRows[i]["cpl_cust_type"]).Trim();

                    strSQL = "EXEC SP_AR_CUST_MASTER_CUSTPRDLINE '1', " +
                             FmtStr(entity) + ", " +
                             FmtStr(branch) + ", " +
                             FmtStr(custCode1) + ", " +
                             FmtStr(custCode2) + ", " +
                             FmtStr(lineCode) + ", " +
                             (i + 1).ToString() + ", " +
                             FmtStr(custType);

                    _clsGlobal.ExecuteTrans(strSQL);
                }

                _clsGlobal.CommitTrans();
            }
            catch
            {
                _clsGlobal.RollbackTrans();
                throw;
            }
        }

        private void SaveSalesmanCoverage()
        {
            FlushAllGridEditors();

            DataTable dt = DVGSalesman.DataSource as DataTable;

            if (dt == null || dt.Rows.Count == 0)
                return; // jangan delete data lama kalau grid kosong

            List<DataRow> validRows = dt.Rows
                .Cast<DataRow>()
                .Where(r => r.RowState != DataRowState.Deleted)
                .Where(r =>
                    !string.IsNullOrWhiteSpace(Convert.ToString(r["csc_salesman_id"])) &&
                    !string.IsNullOrWhiteSpace(Convert.ToString(r["csc_visit"])) &&
                    !string.IsNullOrWhiteSpace(Convert.ToString(r["csc_route"])))
                .ToList();

            if (validRows.Count == 0)
                return; // jangan delete data lama

            string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
            string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
            string custCode1 = txtCustCode.Text.Trim();
            string custCode2 = txtCustCodeTo.Text.Trim();

            _clsGlobal.BeginTrans();

            try
            {
                strSQL = "EXEC SP_AR_CUST_MASTER_CUSTCOVER '3', " +
                         FmtStr(entity) + ", " +
                         FmtStr(branch) + ", " +
                         FmtStr(custCode1) + ", " +
                         FmtStr(custCode2) + ", " +
                         "NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL";

                _clsGlobal.ExecuteTrans(strSQL);

                for (int i = 0; i < validRows.Count; i++)
                {
                    string visit = Convert.ToString(validRows[i]["csc_visit"]).Trim();

                    int visitDay =
                        visit == "Senin" ? 1 :
                        visit == "Selasa" ? 2 :
                        visit == "Rabu" ? 3 :
                        visit == "Kamis" ? 4 :
                        visit == "Jumat" ? 5 :
                        visit == "Sabtu" ? 6 :
                        visit == "Minggu" ? 7 : 0;

                    strSQL = "EXEC SP_AR_CUST_MASTER_CUSTCOVER '1', " +
                             FmtStr(entity) + ", " +
                             FmtStr(branch) + ", " +
                             FmtStr(custCode1) + ", " +
                             FmtStr(custCode2) + ", " +
                             FmtStr(Convert.ToString(validRows[i]["csc_salesman_id"]).Trim()) + ", " +
                             visitDay + ", " +
                             (i + 1) + ", " +
                             FmtStr(IsTrueValue(validRows[i]["csc_visit_week1"]) ? "Y" : "T") + ", " +
                             FmtStr(IsTrueValue(validRows[i]["csc_visit_week2"]) ? "Y" : "T") + ", " +
                             FmtStr(IsTrueValue(validRows[i]["csc_visit_week3"]) ? "Y" : "T") + ", " +
                             FmtStr(IsTrueValue(validRows[i]["csc_visit_week4"]) ? "Y" : "T") + ", " +
                             FmtStr(Convert.ToString(validRows[i]["csc_route"]).Trim()) + ", " +
                             FmtStr(clsLogin.USERID);

                    _clsGlobal.ExecuteTrans(strSQL);
                }

                _clsGlobal.CommitTrans();
            }
            catch
            {
                _clsGlobal.RollbackTrans();
                throw;
            }
        }

        private void SavePriceGroup()
        {
            FlushAllGridEditors();

            DataTable dt = dgvGroupHarga.DataSource as DataTable;

            if (dt == null || dt.Rows.Count == 0)
                return; // jangan delete data lama kalau grid kosong

            List<DataRow> validRows = dt.Rows
                .Cast<DataRow>()
                .Where(r => r.RowState != DataRowState.Deleted)
                .Where(r => !string.IsNullOrWhiteSpace(Convert.ToString(r["grp_code"])))
                .ToList();

            if (validRows.Count == 0)
                return; // jangan delete data lama

            bool hasStd = validRows.Any(r =>
                Convert.ToString(r["grp_code"])
                    .Trim()
                    .Equals("STD", StringComparison.OrdinalIgnoreCase));

            if (!hasStd)
                throw new Exception("Data Group Price STD (standart) harus ada, silahkan isi terlebih dahulu sebelum simpan data");

            string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
            string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
            string custCode1 = txtCustCode.Text.Trim();
            string custCode2 = txtCustCodeTo.Text.Trim();

            _clsGlobal.BeginTrans();

            try
            {
                strSQL = "EXEC IP_INSERT_CUST_PRICE_GROUP " +
                         "1, " +
                         FmtStr(entity) + ", " +
                         FmtStr(branch) + ", " +
                         FmtStr(custCode1) + ", " +
                         FmtStr(custCode2) + ", " +
                         "NULL, NULL, NULL";

                _clsGlobal.ExecuteTrans(strSQL);

                for (int i = 0; i < validRows.Count; i++)
                {
                    string grpCode = Convert.ToString(validRows[i]["grp_code"]).Trim();

                    strSQL = "EXEC IP_INSERT_CUST_PRICE_GROUP " +
                             "0, " +
                             FmtStr(entity) + ", " +
                             FmtStr(branch) + ", " +
                             FmtStr(custCode1) + ", " +
                             FmtStr(custCode2) + ", " +
                             (i + 1).ToString() + ", " +
                             FmtStr(grpCode) + ", " +
                             FmtStr(clsLogin.USERID);

                    _clsGlobal.ExecuteTrans(strSQL);
                }

                _clsGlobal.CommitTrans();
            }
            catch
            {
                _clsGlobal.RollbackTrans();
                throw;
            }
        }
        private bool IsTrueValue(object value)
        {
            string text = Convert.ToString(value).Trim();
            return text.Equals("True", StringComparison.OrdinalIgnoreCase)
                || text.Equals("true", StringComparison.OrdinalIgnoreCase)
                || text.Equals("Y", StringComparison.OrdinalIgnoreCase)
                || text.Equals("1", StringComparison.OrdinalIgnoreCase);
        }
        #region Native DevExpress helper methods
        private static int GetFocusedColumnIndex(DevExpress.XtraGrid.GridControl grid)
        {
            var view = grid == null ? null : grid.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
            return view == null || view.FocusedColumn == null ? -1 : view.FocusedColumn.VisibleIndex;
        }

        private static int GetFocusedRowIndex(DevExpress.XtraGrid.GridControl grid)
        {
            var view = grid == null ? null : grid.MainView as DevExpress.XtraGrid.Views.Grid.GridView;
            return view == null ? -1 : view.FocusedRowHandle;
        }

        private static bool IsTextEditMaskCompleted(DevExpress.XtraEditors.TextEdit editor)
        {
            return editor != null && !string.IsNullOrWhiteSpace(editor.Text);
        }

        private static Type GetGridColumnType(string fieldName)
        {
            string f = (fieldName ?? string.Empty).ToLowerInvariant();
            if (f == "no" || f == "tgl") return typeof(int);
            if (f.StartsWith("pola") || f.StartsWith("chk")) return typeof(bool);
            return typeof(string);
        }

        private static System.Windows.Forms.DataGridView GetHiddenGridAdapter(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            var dgv = grid == null ? null : grid.Tag as System.Windows.Forms.DataGridView;
            if (dgv == null)
            {
                dgv = new System.Windows.Forms.DataGridView();
                dgv.AutoGenerateColumns = true;
                dgv.AllowUserToAddRows = false;
                if (grid != null) grid.Tag = dgv;
            }

            EnsureHiddenGridColumns(dgv, grid, view);

            if (grid != null && !object.ReferenceEquals(dgv.DataSource, grid.DataSource))
            {
                dgv.DataSource = grid.DataSource;
                EnsureHiddenGridColumns(dgv, grid, view);
            }

            return dgv;
        }

        private static void EnsureHiddenGridColumns(System.Windows.Forms.DataGridView dgv, DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            if (dgv == null) return;

            DataTable sourceTable = grid == null ? null : grid.DataSource as DataTable;
            if (sourceTable != null)
            {
                foreach (DataColumn dc in sourceTable.Columns)
                {
                    string colName = dc.ColumnName;
                    if (!dgv.Columns.Contains(colName))
                    {
                        var col = new System.Windows.Forms.DataGridViewTextBoxColumn();
                        col.Name = colName;
                        col.DataPropertyName = colName;
                        col.HeaderText = colName;
                        dgv.Columns.Add(col);
                    }
                }
            }

            if (view != null)
            {
                foreach (DevExpress.XtraGrid.Columns.GridColumn gridColumn in view.Columns)
                {
                    string colName = !string.IsNullOrWhiteSpace(gridColumn.FieldName) ? gridColumn.FieldName : gridColumn.Name;
                    if (string.IsNullOrWhiteSpace(colName)) continue;

                    if (!dgv.Columns.Contains(colName))
                    {
                        System.Windows.Forms.DataGridViewColumn col;
                        Type valueType = GetGridColumnType(colName);
                        if (valueType == typeof(bool))
                        {
                            col = new System.Windows.Forms.DataGridViewCheckBoxColumn();
                        }
                        else
                        {
                            col = new System.Windows.Forms.DataGridViewTextBoxColumn();
                        }

                        col.Name = colName;
                        col.DataPropertyName = colName;
                        col.HeaderText = string.IsNullOrWhiteSpace(gridColumn.Caption) ? colName : gridColumn.Caption;
                        dgv.Columns.Add(col);
                    }
                }
            }

            // Beberapa logic lama melakukan Rows.Add() pada adapter tersembunyi.
            // DataGridView akan error jika Rows.Add() dipanggil sebelum ada kolom.
            if (dgv.Columns.Count == 0)
            {
                dgv.Columns.Add(new System.Windows.Forms.DataGridViewTextBoxColumn()
                {
                    Name = "col0",
                    HeaderText = "col0"
                });
            }
        }

        private static System.Windows.Forms.DataGridViewRowCollection GetGridRows(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            return GetHiddenGridAdapter(grid, view).Rows;
        }

        private static System.Windows.Forms.DataGridViewRow GetGridCurrentRow(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            var dgv = GetHiddenGridAdapter(grid, view);
            if (dgv.CurrentRow != null) return dgv.CurrentRow;
            int rowHandle = view == null ? -1 : view.FocusedRowHandle;
            if (rowHandle >= 0 && rowHandle < dgv.Rows.Count) return dgv.Rows[rowHandle];
            return dgv.Rows.Count > 0 ? dgv.Rows[0] : null;
        }

        private static System.Windows.Forms.DataGridViewCell GetGridCurrentCell(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            var dgv = GetHiddenGridAdapter(grid, view);
            if (dgv.CurrentCell != null) return dgv.CurrentCell;
            int rowHandle = view == null ? 0 : Math.Max(0, view.FocusedRowHandle);
            int colIndex = view == null || view.FocusedColumn == null ? 0 : Math.Max(0, view.FocusedColumn.VisibleIndex);
            if (dgv.Rows.Count > rowHandle && dgv.Columns.Count > colIndex) return dgv[colIndex, rowHandle];
            return null;
        }

        private static int GetGridRowCount(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            if (view != null) return view.RowCount;
            var dt = grid == null ? null : grid.DataSource as DataTable;
            return dt == null ? 0 : dt.Rows.Count;
        }

        private static System.Windows.Forms.DataGridViewCell GetGridCell(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view, int columnIndex, int rowIndex)
        {
            return GetHiddenGridAdapter(grid, view)[columnIndex, rowIndex];
        }

        private static System.Windows.Forms.DataGridViewCell GetGridCell(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view, string columnName, int rowIndex)
        {
            return GetHiddenGridAdapter(grid, view)[columnName, rowIndex];
        }

        private DataTable EnsureGridDataTable(DevExpress.XtraGrid.GridControl grid, params string[] fieldNames)
        {
            DataTable table = grid == null ? null : grid.DataSource as DataTable;
            if (table == null)
            {
                table = new DataTable();
                if (grid != null) grid.DataSource = table;
            }

            foreach (string fieldName in fieldNames)
            {
                if (string.IsNullOrWhiteSpace(fieldName)) continue;
                if (!table.Columns.Contains(fieldName))
                    table.Columns.Add(fieldName, GetGridColumnType(fieldName));
            }

            return table;
        }

        private static DataRow GetFocusedDataRow(DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            if (view == null || view.FocusedRowHandle < 0) return null;
            return view.GetDataRow(view.FocusedRowHandle);
        }

        private static int GetFocusedDataRowIndex(DevExpress.XtraGrid.Views.Grid.GridView view, DataTable table)
        {
            DataRow row = GetFocusedDataRow(view);
            if (row == null || table == null) return -1;
            return table.Rows.IndexOf(row);
        }

        private void FlushAllGridEditors()
        {
            FlushGridEditor(gridPajak, gridPajakView);
            FlushGridEditor(DGVHIR, DGVHIRView);
            FlushGridEditor(DVGSalesman, DVGSalesmanView);
            FlushGridEditor(dgvGroupHarga, dgvGroupHargaView);
            FlushGridEditor(partnerFunctionGrid, partnerFunctionView);
            FlushGridEditor(topbyprdlinedgv, topbyprdlinedgvView);
            FlushGridEditor(dgvschpayD, dgvschpayDView);
            FlushGridEditor(dgvschpayW, dgvschpayWView);
            FlushGridEditor(dgvPrdLineBlocking, dgvPrdLineBlockingView);
            FlushGridEditor(DGVBackList, DGVBackListView);
            FlushGridEditor(DGVUnBackList, DGVUnBackListView);
        }

        private static void FlushGridEditor(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            try
            {
                if (view != null)
                {
                    view.PostEditor();
                    view.UpdateCurrentRow();
                    view.RefreshData();
                }
                if (grid != null)
                    grid.RefreshDataSource();
            }
            catch
            {
                // Flush grid editor tidak boleh menghentikan flow save lama.
            }
        }

        private string GetStatusCode()
        {
            return GetEditValueText(cbstatus);
        }

        private void ConfigureTaxGridRuntime()
        {
            if (gridPajak == null || gridPajakView == null) return;

            try
            {
                EnsureGridDataTable(gridPajak, "code", "btn", "desc");

                if (gridPajak.MainView != gridPajakView)
                    gridPajak.MainView = gridPajakView;

                PrepareGridColumn(gridPajakView, "code", "code", "Code", 80, 0, true);
                PrepareGridColumn(gridPajakView, "btn", "btn", "", 28, 1, true);
                PrepareGridColumn(gridPajakView, "desc", "desc", "Desc", 220, 2, false);

                if (_taxCodeButtonRepository == null)
                {
                    _taxCodeButtonRepository = CreateButtonRepository();
                    _taxCodeButtonRepository.ButtonClick -= TaxCodeButtonRepository_ButtonClick;
                    _taxCodeButtonRepository.ButtonClick += TaxCodeButtonRepository_ButtonClick;
                    if (!gridPajak.RepositoryItems.Contains(_taxCodeButtonRepository))
                        gridPajak.RepositoryItems.Add(_taxCodeButtonRepository);
                }

                ConfigureSearchButtonColumn(GetGridColumnByNameOrField(gridPajakView, "btn", "btn"), _taxCodeButtonRepository, "Search Tipe Pajak");

                gridPajakView.OptionsView.ShowGroupPanel = false;
                gridPajakView.OptionsView.ColumnAutoWidth = false;
                gridPajakView.OptionsBehavior.Editable = true;

                gridPajak.KeyDown -= gridPajakView_KeyDown;
                gridPajak.KeyDown += gridPajakView_KeyDown;
                gridPajakView.KeyDown -= gridPajakView_KeyDown;
                gridPajakView.KeyDown += gridPajakView_KeyDown;
                gridPajakView.RowCellClick -= gridPajakView_RowCellClick;
                gridPajakView.RowCellClick += gridPajakView_RowCellClick;
                gridPajakView.CellValueChanged -= gridPajakView_CellValueChanged;
                gridPajakView.CellValueChanged += gridPajakView_CellValueChanged;

                gridPajak.RefreshDataSource();
                gridPajakView.RefreshData();
            }
            catch
            {
                // Konfigurasi tampilan tidak boleh mengganggu logic lama.
            }
        }

        private void ConfigureSchedulePaymentGridRuntime()
        {
            ConfigureGridMainView(dgvschpayD, dgvschpayDView);
            ConfigureGridMainView(dgvschpayW, dgvschpayWView);

            dgvschpayD.KeyDown -= dgvschpayD_KeyDown;
            dgvschpayD.KeyDown += dgvschpayD_KeyDown;
            dgvschpayDView.KeyDown -= dgvschpayD_KeyDown;
            dgvschpayDView.KeyDown += dgvschpayD_KeyDown;

            dgvschpayW.KeyDown -= dgvschpayW_KeyDown;
            dgvschpayW.KeyDown += dgvschpayW_KeyDown;
            dgvschpayWView.KeyDown -= dgvschpayW_KeyDown;
            dgvschpayWView.KeyDown += dgvschpayW_KeyDown;
            dgvschpayWView.CellValueChanged -= dgvschpayWView_CellValueChanged;
            dgvschpayWView.CellValueChanged += dgvschpayWView_CellValueChanged;
        }

        private void ConfigurePrdLineBlockingGridRuntime()
        {
            ConfigureGridMainView(dgvPrdLineBlocking, dgvPrdLineBlockingView);

            if (_gridSearchButtonRepository == null)
                _gridSearchButtonRepository = CreateButtonRepository();
            if (dgvPrdLineBlocking != null && !dgvPrdLineBlocking.RepositoryItems.Contains(_gridSearchButtonRepository))
                dgvPrdLineBlocking.RepositoryItems.Add(_gridSearchButtonRepository);
            ConfigureSearchButtonColumn(GetGridColumnByNameOrField(dgvPrdLineBlockingView, "prdlinebtncol", string.Empty), _gridSearchButtonRepository, "Search Product Line");

            dgvPrdLineBlocking.KeyDown -= dgvPrdLineBlocking_KeyDown;
            dgvPrdLineBlocking.KeyDown += dgvPrdLineBlocking_KeyDown;
            dgvPrdLineBlockingView.KeyDown -= dgvPrdLineBlocking_KeyDown;
            dgvPrdLineBlockingView.KeyDown += dgvPrdLineBlocking_KeyDown;
            dgvPrdLineBlockingView.RowCellClick -= dgvPrdLineBlockingView_RowCellClick;
            dgvPrdLineBlockingView.RowCellClick += dgvPrdLineBlockingView_RowCellClick;
            dgvPrdLineBlockingView.CellValueChanged -= dgvPrdLineBlockingView_CellValueChanged;
            dgvPrdLineBlockingView.CellValueChanged += dgvPrdLineBlockingView_CellValueChanged;
        }

        private static void ConfigureGridMainView(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            if (grid == null || view == null) return;
            if (grid.MainView != view)
                grid.MainView = view;
            view.OptionsView.ShowGroupPanel = false;
            view.OptionsView.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowAlways;
            view.OptionsBehavior.Editable = true;
        }

        private void TaxCodeButtonRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ShowTaxCodePopupForFocusedRow();
        }

        private void gridPajakView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null) return;
            if (!IsGridButtonColumn(e.Column, "btn", "btn")) return;

            gridPajakView.FocusedRowHandle = e.RowHandle;
            gridPajakView.FocusedColumn = e.Column;
            ShowTaxCodePopupForFocusedRow();
        }

        private void gridPajakView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null) return;
            if (!string.Equals(e.Column.FieldName, "code", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(e.Column.FieldName, "btn", StringComparison.OrdinalIgnoreCase))
                return;

            ValidateTaxGridRow(e.RowHandle);
        }

        private void gridPajakView_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                DataTable table = EnsureGridDataTable(gridPajak, "code", "btn", "desc");
                if (e.KeyCode == Keys.Insert || e.KeyCode == Keys.Down)
                {
                    if (table.Rows.Count < LIMIT_COUNT_HOTT_TAX_CODES)
                    {
                        DataRow row = table.NewRow();
                        table.Rows.Add(row);
                        gridPajak.RefreshDataSource();
                        gridPajakView.FocusedRowHandle = gridPajakView.GetRowHandle(table.Rows.Count - 1);
                        gridPajakView.FocusedColumn = gridPajakView.Columns["code"];
                        e.Handled = true;
                    }
                    else
                    {
                        ShowValidationError(tabPage2, "Maksmial " + LIMIT_COUNT_HOTT_TAX_CODES + " Tipe Pajak !!", MessageBoxIcon.Information);
                    }
                }
                else if (e.KeyCode == Keys.Delete)
                {
                    DataRow row = GetFocusedDataRow(gridPajakView);
                    if (row != null)
                    {
                        table.Rows.Remove(row);
                        gridPajak.RefreshDataSource();
                        e.Handled = true;
                    }
                }
                else if (e.KeyCode == Keys.F4)
                {
                    ShowTaxCodePopupForFocusedRow();
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowTaxCodePopupForFocusedRow()
        {
            try
            {
                if (gridPajakView == null || gridPajakView.FocusedRowHandle < 0) return;

                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Tipe Pajak";
                frm.Query = "select stt_tax_code,stt_tax_desc from so_tax_type ";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    ApplyTaxCodeToRow(gridPajakView.FocusedRowHandle, frm.ArrField[0].Trim(), frm.ArrField[1].Trim());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyTaxCodeToRow(int rowHandle, string codeValue, string descValue)
        {
            DataTable table = EnsureGridDataTable(gridPajak, "code", "btn", "desc");
            DataRow row = gridPajakView.GetDataRow(rowHandle);
            if (row == null)
            {
                row = table.NewRow();
                table.Rows.Add(row);
                rowHandle = gridPajakView.GetRowHandle(table.Rows.Count - 1);
            }

            if (!IsTaxCodeUnique(codeValue, row))
            {
                ShowValidationError(tabPage2, "Data Sudah Ada!", MessageBoxIcon.Information);
                row["code"] = string.Empty;
                row["desc"] = string.Empty;
                row["btn"] = string.Empty;
                return;
            }

            row["code"] = codeValue;
            row["btn"] = string.Empty;
            row["desc"] = descValue;
            gridPajak.RefreshDataSource();
            if (rowHandle >= 0) gridPajakView.FocusedRowHandle = rowHandle;
        }

        private void ValidateTaxGridRow(int rowHandle)
        {
            DataTable table = EnsureGridDataTable(gridPajak, "code", "btn", "desc");
            DataRow row = gridPajakView.GetDataRow(rowHandle);
            if (row == null) return;

            string condID = Convert.ToString(row["code"]).Trim();
            if (string.IsNullOrEmpty(condID))
            {
                row["desc"] = string.Empty;
                return;
            }

            strSQL = "select stt_tax_code,stt_tax_desc from so_tax_type where stt_tax_code = '" + condID.Replace("'", "''") + "' ";
            DataTable dtDesc = _clsGlobal.ExecDT(strSQL);
            if (dtDesc.Rows.Count > 0)
            {
                string code = dtDesc.Rows[0]["stt_tax_code"].ToString().Trim();
                if (IsTaxCodeUnique(code, row))
                {
                    row["code"] = code;
                    row["desc"] = dtDesc.Rows[0]["stt_tax_desc"].ToString().Trim();
                }
                else
                {
                    ShowValidationError(tabPage2, "Data Sudah Ada!", MessageBoxIcon.Information);
                    row["code"] = string.Empty;
                    row["desc"] = string.Empty;
                }
            }
            else
            {
                row["desc"] = string.Empty;
            }

            gridPajak.RefreshDataSource();
        }

        private bool IsTaxCodeUnique(string codeValue, DataRow currentRow)
        {
            if (string.IsNullOrWhiteSpace(codeValue)) return true;

            DataTable table = EnsureGridDataTable(gridPajak, "code", "btn", "desc");
            foreach (DataRow row in table.Rows)
            {
                if (object.ReferenceEquals(row, currentRow)) continue;
                string value = Convert.ToString(row["code"]).Trim();
                if (string.Equals(value, codeValue.Trim(), StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }

        private void dgvschpayWView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null) return;
            string fieldName = e.Column.FieldName;
            if (fieldName != "pola1" && fieldName != "pola2" && fieldName != "pola3" && fieldName != "pola4") return;
            if (!Convert.ToBoolean(e.Value)) return;

            string[] polaFields = new string[] { "pola1", "pola2", "pola3", "pola4" };
            foreach (string polaField in polaFields)
            {
                if (polaField == fieldName) continue;
                dgvschpayWView.SetRowCellValue(e.RowHandle, polaField, false);
            }
        }

        private void dgvPrdLineBlockingView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null) return;
            if (!IsGridButtonColumn(e.Column, "prdlinecol", "prdlinecode")) return;

            string key = Convert.ToString(e.Value).Trim();
            __prdlineblocking.ActionSelector(key, false, new IndexGrid(e.Column.VisibleIndex, e.RowHandle));
        }

        private void dgvPrdLineBlockingView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null) return;
            if (!IsGridButtonColumn(e.Column, "prdlinebtncol", string.Empty)) return;

            DataRow row = dgvPrdLineBlockingView.GetDataRow(e.RowHandle);
            string key = row == null ? string.Empty : Convert.ToString(row["prdlinecode"]).Trim();
            __prdlineblocking.ActionSelector(key, true, new IndexGrid(e.Column.VisibleIndex, e.RowHandle));
        }



        private void ConfigureHierarchyGridRuntime()
        {
            // Pastikan kolom penampung nilai pilihan (btna/btnb) ada di DataTable
            DataTable hirSource = DGVHIR.DataSource as DataTable;
            if (hirSource != null)
            {
                if (!hirSource.Columns.Contains("btna")) hirSource.Columns.Add("btna", typeof(string));
                if (!hirSource.Columns.Contains("btnb")) hirSource.Columns.Add("btnb", typeof(string));
            }

            // === SearchLookUpEdit: Product Line (BTNA) ===
            if (_hirPrdLineSearchLookupRepository == null)
            {
                _hirPrdLineSearchLookupRepository = CreateSearchButtonRepository();
                _hirPrdLineSearchLookupRepository.DisplayMember = "Kode";
                _hirPrdLineSearchLookupRepository.ValueMember = "Kode";
                _hirPrdLineSearchLookupRepository.ButtonClick += HirPrdLineSearchLookupRepository_ButtonClick;
                if (!DGVHIR.RepositoryItems.Contains(_hirPrdLineSearchLookupRepository))
                    DGVHIR.RepositoryItems.Add(_hirPrdLineSearchLookupRepository);
            }

            // === SearchLookUpEdit: Tipe Outlet (BTNB) ===
            if (_hirCustTypeSearchLookupRepository == null)
            {
                _hirCustTypeSearchLookupRepository = CreateSearchButtonRepository();
                _hirCustTypeSearchLookupRepository.DisplayMember = "Kode";
                _hirCustTypeSearchLookupRepository.ValueMember = "Kode";
                _hirCustTypeSearchLookupRepository.ButtonClick += HirCustTypeSearchLookupRepository_ButtonClick;
                if (!DGVHIR.RepositoryItems.Contains(_hirCustTypeSearchLookupRepository))
                    DGVHIR.RepositoryItems.Add(_hirCustTypeSearchLookupRepository);
            }

            _hirPrdLineSearchLookupRepository.DataSource = GetHirPrdLineLookupDataSource();
            ConfigureLookupView(_hirPrdLineSearchLookupRepository,
                new string[] { "Kode", "Nama" },
                new string[] { "PrdLine", "Nama Produk Line" },
                new int[] { 90, 220 });

            _hirCustTypeSearchLookupRepository.DataSource = GetHirCustTypeLookupDataSource();
            ConfigureLookupView(_hirCustTypeSearchLookupRepository,
                new string[] { "Kode", "Nama" },
                new string[] { "Tipe Outlet", "Nama Tipe Outlet" },
                new int[] { 110, 220 });

            ConfigureSearchButtonColumn(GetGridColumnByNameOrField(DGVHIRView, "BTNA", "btna"), _hirPrdLineSearchLookupRepository, "Search Product Line");
            ConfigureSearchButtonColumn(GetGridColumnByNameOrField(DGVHIRView, "BTNB", "btnb"), _hirCustTypeSearchLookupRepository, "Search Tipe Outlet");

            // Pakai popup SearchLookUpEdit, BUKAN form search lama.
            // (Hapus wiring RowCellClick ke handler lama supaya form lama tidak ikut terbuka.)
            DGVHIRView.ShownEditor -= DGVHIRView_ShownEditor_SearchLookup;
            DGVHIRView.ShownEditor += DGVHIRView_ShownEditor_SearchLookup;
            DGVHIRView.CellValueChanged -= DGVHIRView_CellValueChanged_SearchLookup;
            DGVHIRView.CellValueChanged += DGVHIRView_CellValueChanged_SearchLookup;
        }



        private void ConfigureGroupHargaGridRuntime()
        {
            if (dgvGroupHarga == null || dgvGroupHargaView == null) return;

            try
            {
                DataTable source = dgvGroupHarga.DataSource as DataTable;
                if (source != null && !source.Columns.Contains("btnGroup"))
                    source.Columns.Add("btnGroup", typeof(string));

                if (dgvGroupHarga.MainView != dgvGroupHargaView)
                    dgvGroupHarga.MainView = dgvGroupHargaView;

                bool viewRegistered = false;
                foreach (DevExpress.XtraGrid.Views.Base.BaseView v in dgvGroupHarga.ViewCollection)
                {
                    if (object.ReferenceEquals(v, dgvGroupHargaView))
                    {
                        viewRegistered = true;
                        break;
                    }
                }
                if (!viewRegistered)
                    dgvGroupHarga.ViewCollection.Add(dgvGroupHargaView);

                if (_groupHargaSearchLookupRepository == null)
                {
                    _groupHargaSearchLookupRepository = CreateSearchButtonRepository();
                    _groupHargaSearchLookupRepository.DisplayMember = "Kode";
                    _groupHargaSearchLookupRepository.ValueMember = "Kode";
                    _groupHargaSearchLookupRepository.ButtonClick += GroupHargaSearchLookupRepository_ButtonClick;
                    if (!dgvGroupHarga.RepositoryItems.Contains(_groupHargaSearchLookupRepository))
                        dgvGroupHarga.RepositoryItems.Add(_groupHargaSearchLookupRepository);
                }

                ConfigureSearchLookupButton(_groupHargaSearchLookupRepository, "Search Group Price");
                _groupHargaSearchLookupRepository.DataSource = GetGroupHargaLookupDataSource();
                ConfigureLookupView(_groupHargaSearchLookupRepository,
                    new string[] { "Kode", "Nama" },
                    new string[] { "Group Price Code", "Group Price Desc" },
                    new int[] { 130, 280 });

                PrepareGridColumn(dgvGroupHargaView, "grp_code", "grp_code", "Group Price Code", 150, 0, false);
                PrepareGridColumn(dgvGroupHargaView, "btnGroup", "btnGroup", "", 28, 1, true);
                PrepareGridColumn(dgvGroupHargaView, "grp_desc", "grp_desc", "Group Price Desc", 280, 2, false);

                if (dgvGroupHargaView.Columns["btnGroup"] != null)
                {
                    ConfigureSearchButtonColumn(dgvGroupHargaView.Columns["btnGroup"], _groupHargaSearchLookupRepository, "Search Group Price");
                }

                dgvGroupHargaView.OptionsView.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowAlways;
                dgvGroupHargaView.OptionsView.ShowGroupPanel = false;
                dgvGroupHargaView.OptionsView.ColumnAutoWidth = false;
                dgvGroupHargaView.OptionsBehavior.Editable = true;

                dgvGroupHargaView.CellValueChanged -= dgvGroupHargaView_CellValueChanged_SearchLookup;
                dgvGroupHargaView.CellValueChanged += dgvGroupHargaView_CellValueChanged_SearchLookup;
                dgvGroupHargaView.ShownEditor -= dgvGroupHargaView_ShownEditor_SearchLookup;
                dgvGroupHargaView.ShownEditor += dgvGroupHargaView_ShownEditor_SearchLookup;
                dgvGroupHargaView.RowCellClick -= dgvGroupHargaView_RowCellClick_SearchLookup;
                dgvGroupHargaView.RowCellClick += dgvGroupHargaView_RowCellClick_SearchLookup;

                dgvGroupHarga.RefreshDataSource();
                dgvGroupHargaView.RefreshData();
                dgvGroupHargaView.LayoutChanged();
            }
            catch
            {
                // Tampilan search lookup tidak boleh mengganggu logic edit/save.
            }
        }

        private void SetEditableGridView()
        {
            // Group Harga
            if (dgvGroupHarga != null && dgvGroupHargaView != null)
            {
                dgvGroupHarga.Enabled = true;
                dgvGroupHargaView.OptionsBehavior.Editable = true;
                dgvGroupHargaView.OptionsBehavior.ReadOnly = false;
                dgvGroupHargaView.OptionsView.NewItemRowPosition =
                    DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;

                foreach (GridColumn col in dgvGroupHargaView.Columns)
                {
                    col.OptionsColumn.AllowEdit = true;
                    col.OptionsColumn.ReadOnly = false;
                }
            }

            // Salesman Coverage
            if (DVGSalesman != null && DVGSalesmanView != null)
            {
                DVGSalesman.Enabled = true;
                DVGSalesmanView.OptionsBehavior.Editable = true;
                DVGSalesmanView.OptionsBehavior.ReadOnly = false;
                DVGSalesmanView.OptionsView.NewItemRowPosition =
                    DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;

                foreach (GridColumn col in DVGSalesmanView.Columns)
                {
                    col.OptionsColumn.AllowEdit = true;
                    col.OptionsColumn.ReadOnly = false;
                }
            }
        }

        private void EnsurePartnerFunctionTabRuntime()
        {
            if (tabControl1 == null) return;

            if (partnerFunctionPage == null)
            {
                partnerFunctionPage = new XtraTabPage();
                partnerFunctionPage.Name = "partnerFunctionPage";
                partnerFunctionPage.Text = "Partner Function";
                partnerFunctionPage.BackColor = System.Drawing.Color.Gainsboro;
            }

            if (partnerFunctionGrid == null)
            {
                partnerFunctionGrid = new GridControl();
                partnerFunctionGrid.Name = "partnerFunctionGrid";
                partnerFunctionGrid.Dock = DockStyle.Fill;
            }

            if (partnerFunctionView == null)
            {
                partnerFunctionView = new GridView(partnerFunctionGrid);
                partnerFunctionView.Name = "partnerFunctionView";
                partnerFunctionGrid.MainView = partnerFunctionView;
                partnerFunctionGrid.ViewCollection.Add(partnerFunctionView);
            }

            if (!partnerFunctionPage.Controls.Contains(partnerFunctionGrid))
                partnerFunctionPage.Controls.Add(partnerFunctionGrid);

            if (!tabControl1.TabPages.Contains(partnerFunctionPage))
            {
                int index = tabControl1.TabPages.Contains(tabPage1) ? tabControl1.TabPages.IndexOf(tabPage1) + 1 : 1;
                tabControl1.TabPages.Insert(index, partnerFunctionPage);
            }

            if (partnerFunctionGrid.DataSource == null)
                partnerFunctionGrid.DataSource = CreatePartnerFunctionTable();
        }

        private void ConfigurePartnerFunctionGridRuntime()
        {
            EnsurePartnerFunctionTabRuntime();
            if (partnerFunctionGrid == null || partnerFunctionView == null) return;

            try
            {
                DataTable source = partnerFunctionGrid.DataSource as DataTable;
                if (source == null)
                {
                    source = CreatePartnerFunctionTable();
                    partnerFunctionGrid.DataSource = source;
                }
                EnsurePartnerFunctionColumns(source);

                if (_partnerShipToSearchLookupRepository == null)
                {
                    _partnerShipToSearchLookupRepository = CreateSearchButtonRepository();
                    _partnerShipToSearchLookupRepository.DisplayMember = "ShipCode1";
                    _partnerShipToSearchLookupRepository.ValueMember = "LookupKey";
                    _partnerShipToSearchLookupRepository.ButtonClick += PartnerShipToSearchLookupRepository_ButtonClick;
                    partnerFunctionGrid.RepositoryItems.Add(_partnerShipToSearchLookupRepository);
                }

                // Pastikan repository selalu terdaftar di grid (kalau grid sempat di-recreate)
                if (!partnerFunctionGrid.RepositoryItems.Contains(_partnerShipToSearchLookupRepository))
                    partnerFunctionGrid.RepositoryItems.Add(_partnerShipToSearchLookupRepository);

                ConfigureSearchLookupButton(_partnerShipToSearchLookupRepository, "Search Ship To");
                _partnerShipToSearchLookupRepository.DataSource = GetPartnerShipToLookupDataSource();
                ConfigureLookupView(_partnerShipToSearchLookupRepository,
                    new string[] { "BranchCode", "ShipCode1", "ShipDescription" },
                    new string[] { "Branch Code", "Ship To Code", "Ship To Description" },
                    new int[] { 110, 130, 440 });
                string[] partnerLookupVisibleFields = new string[] { "BranchCode", "ShipCode1", "ShipDescription" };
                foreach (DevExpress.XtraGrid.Columns.GridColumn column in _partnerShipToSearchLookupRepository.View.Columns)
                    column.Visible = partnerLookupVisibleFields.Contains(column.FieldName);

                if (_partnerDefaultCheckRepository == null)
                {
                    _partnerDefaultCheckRepository = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
                    _partnerDefaultCheckRepository.ValueChecked = true;
                    _partnerDefaultCheckRepository.ValueUnchecked = false;
                    partnerFunctionGrid.RepositoryItems.Add(_partnerDefaultCheckRepository);
                }

                partnerFunctionView.BeginUpdate();   
                try
                {
                    PrepareGridColumn(partnerFunctionView, "partnerNoCol", "no", "Nomor", 55, 0, false);
                    PrepareGridColumn(partnerFunctionView, "partnerShipToCode1Col", "msh_shipto_code1", "msh_shipto_code1", 120, 1, false);
                    PrepareGridColumn(partnerFunctionView, "partnerShipToButtonCol", "shipto_btn", "", 28, 2, true);   // <<< button index 2
                    PrepareGridColumn(partnerFunctionView, "partnerShipToDescCol", "shipto_desc", "shipto description", 430, 3, false);
                    PrepareGridColumn(partnerFunctionView, "partnerShipToCode2Col", "msh_shipto_code2", "msh_shipto_code2", 120, 4, false); // pindah ke index 4, lalu di-hide
                    PrepareGridColumn(partnerFunctionView, "partnerDefaultCol", "msh_default", "default", 70, 5, true);

                   
                    DevExpress.XtraGrid.Columns.GridColumn shipToButtonColumn =
                        GetGridColumnByNameOrField(partnerFunctionView, "partnerShipToButtonCol", "shipto_btn");
                    ConfigureSearchButtonColumn(shipToButtonColumn, _partnerShipToSearchLookupRepository, "Search Ship To");

                    DevExpress.XtraGrid.Columns.GridColumn shipToCode2Column =
                        GetGridColumnByNameOrField(partnerFunctionView, "partnerShipToCode2Col", "msh_shipto_code2");
                    if (shipToCode2Column != null)
                        shipToCode2Column.Visible = false;

                    if (partnerFunctionView.Columns["msh_default"] != null)
                        partnerFunctionView.Columns["msh_default"].ColumnEdit = _partnerDefaultCheckRepository;

                    partnerFunctionView.OptionsView.ShowGroupPanel = false;
                    partnerFunctionView.OptionsView.ColumnAutoWidth = false;   
                    partnerFunctionView.OptionsBehavior.Editable = true;
                    partnerFunctionView.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
                    partnerFunctionView.OptionsView.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowAlways;
                    partnerFunctionView.OptionsSelection.EnableAppearanceFocusedCell = true;   
                }
                finally
                {
                    partnerFunctionView.EndUpdate();
                }

                partnerFunctionView.CellValueChanged -= PartnerFunctionView_CellValueChanged;
                partnerFunctionView.CellValueChanged += PartnerFunctionView_CellValueChanged;
                partnerFunctionView.ShownEditor -= PartnerFunctionView_ShownEditor;
                partnerFunctionView.ShownEditor += PartnerFunctionView_ShownEditor;
                partnerFunctionView.RowCellClick -= PartnerFunctionView_RowCellClick;
                partnerFunctionView.RowCellClick += PartnerFunctionView_RowCellClick;
                partnerFunctionView.KeyDown -= PartnerFunctionView_KeyDown;
                partnerFunctionView.KeyDown += PartnerFunctionView_KeyDown;
                partnerFunctionGrid.KeyDown -= PartnerFunctionView_KeyDown;
                partnerFunctionGrid.KeyDown += PartnerFunctionView_KeyDown;

                RenumberPartnerFunctionRows();
                partnerFunctionGrid.RefreshDataSource();
                partnerFunctionView.RefreshData();
                partnerFunctionView.LayoutChanged();  
            }
            catch
            {
                // Konfigurasi UI tab partner tidak boleh memutus flow utama form.
            }
        }

        private DataTable CreatePartnerFunctionTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("no", typeof(int));
            table.Columns.Add("msh_shipto_code1", typeof(string));
            table.Columns.Add("msh_shipto_code2", typeof(string));
            table.Columns.Add("shipto_btn", typeof(string));
            table.Columns.Add("shipto_desc", typeof(string));
            table.Columns.Add("msh_default", typeof(bool));
            return table;
        }

        private void EnsurePartnerFunctionColumns(DataTable table)
        {
            if (table == null) return;
            if (!table.Columns.Contains("no")) table.Columns.Add("no", typeof(int));
            if (!table.Columns.Contains("msh_shipto_code1")) table.Columns.Add("msh_shipto_code1", typeof(string));
            if (!table.Columns.Contains("msh_shipto_code2")) table.Columns.Add("msh_shipto_code2", typeof(string));
            if (!table.Columns.Contains("shipto_btn")) table.Columns.Add("shipto_btn", typeof(string));
            if (!table.Columns.Contains("shipto_desc")) table.Columns.Add("shipto_desc", typeof(string));
            if (!table.Columns.Contains("msh_default")) table.Columns.Add("msh_default", typeof(bool));
        }

        private string GetPartnerFunctionEntity()
        {
            string entity = __state == StateEntry.Edit ? _prdEntityCode : ctrlEntityCustMaster2.txtCM.Text.Trim();
            if (string.IsNullOrWhiteSpace(entity))
                entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
            return entity;
        }

        private string GetPartnerFunctionBranch()
        {
            string branch = __state == StateEntry.Edit ? _prdBrandCode : ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
            if (string.IsNullOrWhiteSpace(branch))
                branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();
            return branch;
        }

        private DataTable GetPartnerShipToLookupDataSource()
        {
            string entity = GetPartnerFunctionEntity();
            string branch = GetPartnerFunctionBranch();
            string query = @"
SELECT
    LTRIM(RTRIM(cm.cm_cust_code1)) + '|' + LTRIM(RTRIM(cm.cm_cust_code2)) AS [LookupKey],
    LTRIM(RTRIM(cm.cm_branch)) AS [BranchCode],
    LTRIM(RTRIM(cm.cm_cust_code1)) AS [ShipCode1],
    LTRIM(RTRIM(cm.cm_cust_code2)) AS [Kode2],
    LTRIM(RTRIM(cm.cm_cust_name)) AS [ShipDescription]
FROM SO_CUST_MASTER cm WITH(NOLOCK)
WHERE cm.cm_entity = '" + entity.Replace("'", "''") + @"'
  AND cm.cm_branch = '" + branch.Replace("'", "''") + @"'
ORDER BY cm.cm_branch, cm.cm_cust_code1, cm.cm_cust_code2";

            return _clsGlobal.ExecDT(query);
        }

        private void PartnerShipToSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen)
                editor.ShowPopup();
        }

        private void PartnerFunctionView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || partnerFunctionView == null) return;

            if (string.Equals(e.Column.FieldName, "shipto_btn", StringComparison.OrdinalIgnoreCase))
            {
                ApplyPartnerShipToLookupValue(e.RowHandle, e.Value == null ? string.Empty : e.Value.ToString());
                return;
            }

            if (string.Equals(e.Column.FieldName, "msh_default", StringComparison.OrdinalIgnoreCase))
                ApplySinglePartnerDefault(e.RowHandle, e.Value);
        }

        private void PartnerFunctionView_ShownEditor(object sender, EventArgs e)
        {
            if (partnerFunctionView == null || partnerFunctionView.FocusedColumn == null) return;
            if (!string.Equals(partnerFunctionView.FocusedColumn.FieldName, "shipto_btn", StringComparison.OrdinalIgnoreCase)) return;

            SearchLookUpEdit editor = partnerFunctionView.ActiveEditor as SearchLookUpEdit;
            if (editor == null) return;

            editor.EditValueChanged -= PartnerShipToSearchLookupEditor_EditValueChanged;
            editor.EditValueChanged += PartnerShipToSearchLookupEditor_EditValueChanged;
        }

        private void PartnerShipToSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            SearchLookUpEdit editor = sender as SearchLookUpEdit;
            if (editor == null || partnerFunctionView == null) return;

            ApplyPartnerShipToLookupValue(partnerFunctionView.FocusedRowHandle, editor.EditValue == null ? string.Empty : editor.EditValue.ToString());
        }

        private void PartnerFunctionView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null || partnerFunctionView == null) return;
            if (!string.Equals(e.Column.FieldName, "shipto_btn", StringComparison.OrdinalIgnoreCase)) return;

            partnerFunctionView.FocusedRowHandle = e.RowHandle;
            partnerFunctionView.FocusedColumn = e.Column;
            partnerFunctionView.ShowEditor();

            SearchLookUpEdit editor = partnerFunctionView.ActiveEditor as SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen)
                editor.ShowPopup();
        }

        private void PartnerFunctionView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e == null) return;
            DataTable table = partnerFunctionGrid == null ? null : partnerFunctionGrid.DataSource as DataTable;
            if (table == null) return;

            if (e.KeyCode == Keys.Insert)
            {
                DataRow row = table.NewRow();
                row["msh_default"] = false;
                table.Rows.Add(row);
                RenumberPartnerFunctionRows();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Delete && partnerFunctionView != null && partnerFunctionView.FocusedRowHandle >= 0)
            {
                partnerFunctionView.DeleteRow(partnerFunctionView.FocusedRowHandle);
                RenumberPartnerFunctionRows();
                e.Handled = true;
            }
        }

        private void ApplyPartnerShipToLookupValue(int rowHandle, string key)
        {
            if (partnerFunctionView == null || string.IsNullOrWhiteSpace(key)) return;

            DataRow row = partnerFunctionView.GetDataRow(rowHandle);
            if (row == null) return;

            DataTable lookup = _partnerShipToSearchLookupRepository == null ? null : _partnerShipToSearchLookupRepository.DataSource as DataTable;
            DataRow selected = null;
            if (lookup != null)
            {
                selected = lookup.Rows.Cast<DataRow>()
                    .FirstOrDefault(r => string.Equals(r["LookupKey"].ToString(), key, StringComparison.OrdinalIgnoreCase));
            }

            if (selected == null) return;

            partnerFunctionView.SetRowCellValue(rowHandle, "msh_shipto_code1", selected["ShipCode1"].ToString());
            partnerFunctionView.SetRowCellValue(rowHandle, "msh_shipto_code2", selected["Kode2"].ToString());
            partnerFunctionView.SetRowCellValue(rowHandle, "shipto_desc", selected["ShipDescription"].ToString());
            partnerFunctionView.SetRowCellValue(rowHandle, "shipto_btn", string.Empty);
            partnerFunctionView.PostEditor();
            partnerFunctionView.UpdateCurrentRow();
            RenumberPartnerFunctionRows();
        }

        private void ApplySinglePartnerDefault(int rowHandle, object value)
        {
            if (_isUpdatingPartnerDefault || !IsCheckedValue(value)) return;

            DataTable table = partnerFunctionGrid == null ? null : partnerFunctionGrid.DataSource as DataTable;
            if (table == null) return;

            DataRow current = partnerFunctionView == null ? null : partnerFunctionView.GetDataRow(rowHandle);
            if (current == null) return;

            try
            {
                _isUpdatingPartnerDefault = true;
                foreach (DataRow row in table.Rows)
                {
                    if (row.RowState == DataRowState.Deleted || object.ReferenceEquals(row, current)) continue;
                    row["msh_default"] = false;
                }
            }
            finally
            {
                _isUpdatingPartnerDefault = false;
            }
        }

        private void RenumberPartnerFunctionRows()
        {
            DataTable table = partnerFunctionGrid == null ? null : partnerFunctionGrid.DataSource as DataTable;
            if (table == null) return;

            int number = 1;
            foreach (DataRow row in table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                row["no"] = number++;
            }
        }

        private IEnumerable<DataRow> GetValidPartnerFunctionRows()
        {
            DataTable table = partnerFunctionGrid == null ? null : partnerFunctionGrid.DataSource as DataTable;
            if (table == null)
                return Enumerable.Empty<DataRow>();

            return table.Rows.Cast<DataRow>()
                .Where(row => row.RowState != DataRowState.Deleted)
                .Where(row => !string.IsNullOrWhiteSpace(row["msh_shipto_code1"].ToString()) ||
                              !string.IsNullOrWhiteSpace(row["msh_shipto_code2"].ToString()))
                .ToList();
        }

        private bool ValidatePartnerFunctionDefaults()
        {
            FlushGridEditor(partnerFunctionGrid, partnerFunctionView);

            List<DataRow> rows = GetValidPartnerFunctionRows().ToList();
            if (rows.Count == 0)
            {
                FocusInvalidPartnerFunctionTab();
                ShowValidationError(partnerFunctionPage, "Partner Function wajib diisi minimal satu Ship To.", MessageBoxIcon.Error);
                return false;
            }

            if (rows.Any(row => string.IsNullOrWhiteSpace(row["msh_shipto_code1"].ToString()) ||
                                string.IsNullOrWhiteSpace(row["msh_shipto_code2"].ToString())))
            {
                FocusInvalidPartnerFunctionTab();
                ShowValidationError(partnerFunctionPage, "Shipto Code 1 dan Shipto Code 2 pada Partner Function wajib lengkap.", MessageBoxIcon.Error);
                return false;
            }

            int defaultCount = rows.Count(row => IsCheckedValue(row["msh_default"]));
            if (defaultCount != 1)
            {
                FocusInvalidPartnerFunctionTab();
                ShowValidationError(partnerFunctionPage, "Partner Function harus memiliki tepat satu Ship To yang menjadi default. Pilih satu checkbox default saja.", MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void FocusInvalidPartnerFunctionTab()
        {
            if (partnerFunctionPage != null)
                MarkInvalidTab(partnerFunctionPage);

            if (partnerFunctionView != null)
                partnerFunctionView.Focus();
        }

        private static bool IsCheckedValue(object value)
        {
            if (value == null || value == DBNull.Value) return false;
            if (value is bool) return (bool)value;
            string text = value.ToString().Trim();
            return string.Equals(text, "Y", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "1", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "True", StringComparison.OrdinalIgnoreCase);
        }

        private bool SavePartnerFunction()
        {
            if (!ValidatePartnerFunctionDefaults())
                return false;

            string entity = GetPartnerFunctionEntity();
            string branch = GetPartnerFunctionBranch();
            string custCode1 = txtCustCode.Text.Trim();
            string custCode2 = txtCustCodeTo.Text.Trim();
            List<DataRow> rows = GetValidPartnerFunctionRows().ToList();

            try
            {
                _clsGlobal.BeginTrans();

                strSQL = "DELETE FROM SD_MULTI_SHIPTO " +
                         "WHERE msh_entity_id = " + FmtStr(entity) + " " +
                         "AND msh_branch_id = " + FmtStr(branch) + " " +
                         "AND msh_cust_code1 = " + FmtStr(custCode1) + " " +
                         "AND msh_cust_code2 = " + FmtStr(custCode2);
                _clsGlobal.ExecuteTrans(strSQL);

                foreach (DataRow row in rows)
                {
                    string shipToCode1 = row["msh_shipto_code1"].ToString().Trim();
                    string shipToCode2 = row["msh_shipto_code2"].ToString().Trim();
                    string defaultFlag = IsCheckedValue(row["msh_default"]) ? "Y" : "N";

                    strSQL = "INSERT INTO SD_MULTI_SHIPTO " +
                             "(msh_entity_id, msh_branch_id, msh_cust_code1, msh_cust_code2, msh_shipto_code1, msh_shipto_code2, msh_default, msh_created_date, msh_created_by, msh_last_update, msh_update_by) VALUES (" +
                             FmtStr(entity) + ", " +
                             FmtStr(branch) + ", " +
                             FmtStr(custCode1) + ", " +
                             FmtStr(custCode2) + ", " +
                             FmtStr(shipToCode1) + ", " +
                             FmtStr(shipToCode2) + ", " +
                             FmtStr(defaultFlag) + ", " +
                             "GETDATE(), " +
                             FmtStr(clsLogin.USERID) + ", " +
                             "GETDATE(), " +
                             FmtStr(clsLogin.USERID) + ")";
                    _clsGlobal.ExecuteTrans(strSQL);
                }

                _clsGlobal.CommitTrans();
                return true;
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void DGVHIRView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null) return;
            if (!IsGridButtonColumn(e.Column, "BTNA", "btna") && !IsGridButtonColumn(e.Column, "BTNB", "btnb")) return;

            DGVHIRView.FocusedRowHandle = e.RowHandle;
            DGVHIRView.FocusedColumn = e.Column;
            DGVHIR_CellContentClick(DGVHIR, new System.Windows.Forms.DataGridViewCellEventArgs(e.Column.VisibleIndex, e.RowHandle));
        }

        private static bool IsGridButtonColumn(DevExpress.XtraGrid.Columns.GridColumn column, string columnName, string fieldName)
        {
            if (column == null) return false;
            return string.Equals(column.Name, columnName, StringComparison.OrdinalIgnoreCase)
                || (!string.IsNullOrWhiteSpace(fieldName) && string.Equals(column.FieldName, fieldName, StringComparison.OrdinalIgnoreCase));
        }

        private void ConfigureTopByPrdLineGridRuntime()
        {
            if (topbyprdlinedgv == null || topbyprdlinedgvView == null) return;

            try
            {
                DataTable source = topbyprdlinedgv.DataSource as DataTable;
                if (source != null)
                {
                    EnsureTopByDivisionDataColumns(source);
                }

                if (topbyprdlinedgv.MainView != topbyprdlinedgvView)
                    topbyprdlinedgv.MainView = topbyprdlinedgvView;

                bool viewRegistered = false;
                foreach (DevExpress.XtraGrid.Views.Base.BaseView v in topbyprdlinedgv.ViewCollection)
                {
                    if (object.ReferenceEquals(v, topbyprdlinedgvView))
                    {
                        viewRegistered = true;
                        break;
                    }
                }
                if (!viewRegistered)
                    topbyprdlinedgv.ViewCollection.Add(topbyprdlinedgvView);

                if (_topPrdLineSearchLookupRepository == null)
                {
                    _topPrdLineSearchLookupRepository = CreateSearchButtonRepository();
                    _topPrdLineSearchLookupRepository.DisplayMember = "Kode";
                    _topPrdLineSearchLookupRepository.ValueMember = "Kode";
                    _topPrdLineSearchLookupRepository.ButtonClick += TopPrdLineSearchLookupRepository_ButtonClick;
                    if (!topbyprdlinedgv.RepositoryItems.Contains(_topPrdLineSearchLookupRepository))
                        topbyprdlinedgv.RepositoryItems.Add(_topPrdLineSearchLookupRepository);
                }

                if (_topPaymentSearchLookupRepository == null)
                {
                    _topPaymentSearchLookupRepository = CreateSearchButtonRepository();
                    _topPaymentSearchLookupRepository.DisplayMember = "Kode";
                    _topPaymentSearchLookupRepository.ValueMember = "Kode";
                    _topPaymentSearchLookupRepository.ButtonClick += TopPaymentSearchLookupRepository_ButtonClick;
                    if (!topbyprdlinedgv.RepositoryItems.Contains(_topPaymentSearchLookupRepository))
                        topbyprdlinedgv.RepositoryItems.Add(_topPaymentSearchLookupRepository);
                }

                if (_topPlantSearchLookupRepository == null)
                {
                    _topPlantSearchLookupRepository = CreateSearchButtonRepository();
                    _topPlantSearchLookupRepository.DisplayMember = "Kode";
                    _topPlantSearchLookupRepository.ValueMember = "Kode";
                    _topPlantSearchLookupRepository.ButtonClick += TopPlantSearchLookupRepository_ButtonClick;
                    if (!topbyprdlinedgv.RepositoryItems.Contains(_topPlantSearchLookupRepository))
                        topbyprdlinedgv.RepositoryItems.Add(_topPlantSearchLookupRepository);
                }

                if (_topSalesEmployeeSearchLookupRepository == null)
                {
                    _topSalesEmployeeSearchLookupRepository = CreateSearchButtonRepository();
                    _topSalesEmployeeSearchLookupRepository.DisplayMember = "Kode";
                    _topSalesEmployeeSearchLookupRepository.ValueMember = "Kode";
                    _topSalesEmployeeSearchLookupRepository.ButtonClick += TopSalesEmployeeSearchLookupRepository_ButtonClick;
                    if (!topbyprdlinedgv.RepositoryItems.Contains(_topSalesEmployeeSearchLookupRepository))
                        topbyprdlinedgv.RepositoryItems.Add(_topSalesEmployeeSearchLookupRepository);
                }

                _topPrdLineSearchLookupRepository.DataSource = GetPrdLineLookupDataSource();
                ConfigureLookupView(_topPrdLineSearchLookupRepository,
                    new string[] { "Kode", "Nama" },
                    new string[] { "Division", "Nama Division" },
                    new int[] { 120, 260 });

                _topPaymentSearchLookupRepository.DataSource = GetTopLookupDataSource();
                ConfigureLookupView(_topPaymentSearchLookupRepository,
                    new string[] { "Kode", "Nama" },
                    new string[] { "TOP", "Nama TOP" },
                    new int[] { 120, 260 });

                _topPlantSearchLookupRepository.DataSource = GetPlantLookupDataSource();
                ConfigureLookupView(_topPlantSearchLookupRepository,
                    new string[] { "Kode", "Nama" },
                    new string[] { "Plant", "Description" },
                    new int[] { 120, 260 });

                _topSalesEmployeeSearchLookupRepository.DataSource = GetSalesEmployeeLookupDataSource();
                ConfigureLookupView(_topSalesEmployeeSearchLookupRepository,
                    new string[] { "Kode", "Nama" },
                    new string[] { "Sales employee", "Nama Salesman" },
                    new int[] { 130, 260 });

                FillMissingTopByDivisionDescriptions(source);

                PrepareGridColumn(topbyprdlinedgvView, "nocol", "no", "No.", 45, 0, false);
                PrepareGridColumn(topbyprdlinedgvView, "prdlinecol", "prdline_code", "Division", 90, 1, false);
                PrepareGridColumn(topbyprdlinedgvView, "prdlinebtncol", "prdline_btn", "", 28, 2, true);
                PrepareGridColumn(topbyprdlinedgvView, "desccol", "prdline_desc", "Nama Division", 220, 3, false);
                PrepareGridColumn(topbyprdlinedgvView, "topcol", "top_code", "TOP", 80, 4, false);
                PrepareGridColumn(topbyprdlinedgvView, "topbtncol", "top_btn", "", 28, 5, true);
                PrepareGridColumn(topbyprdlinedgvView, "topdesccol", "top_desc", "Nama TOP", 220, 6, false);
                PrepareGridColumn(topbyprdlinedgvView, "plantcol", "plant_id", "Plant", 90, 7, false);
                PrepareGridColumn(topbyprdlinedgvView, "plantbtncol", "plant_btn", "", 28, 8, true);
                PrepareGridColumn(topbyprdlinedgvView, "plantdesccol", "plant_desc", "Plant Description", 180, 9, false);
                PrepareGridColumn(topbyprdlinedgvView, "salesemployeecol", "sales_employee_id", "Sales employee", 120, 10, false);
                PrepareGridColumn(topbyprdlinedgvView, "salesemployeebtncol", "sales_employee_btn", "", 28, 11, true);
                PrepareGridColumn(topbyprdlinedgvView, "salesemployeedesccol", "sales_employee_desc", "Sales Employee Description", 220, 12, false);

                ConfigureSearchButtonColumn(GetGridColumnByNameOrField(topbyprdlinedgvView, "prdlinebtncol", "prdline_btn"), _topPrdLineSearchLookupRepository, "Search Division");
                ConfigureSearchButtonColumn(GetGridColumnByNameOrField(topbyprdlinedgvView, "topbtncol", "top_btn"), _topPaymentSearchLookupRepository, "Search TOP");
                ConfigureSearchButtonColumn(GetGridColumnByNameOrField(topbyprdlinedgvView, "plantbtncol", "plant_btn"), _topPlantSearchLookupRepository, "Search Plant");
                ConfigureSearchButtonColumn(GetGridColumnByNameOrField(topbyprdlinedgvView, "salesemployeebtncol", "sales_employee_btn"), _topSalesEmployeeSearchLookupRepository, "Search Sales Employee");

                topbyprdlinedgvView.OptionsView.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowAlways;
                topbyprdlinedgvView.OptionsView.ShowGroupPanel = false;
                topbyprdlinedgvView.OptionsView.ColumnAutoWidth = false;
                topbyprdlinedgvView.OptionsBehavior.Editable = true;

                topbyprdlinedgvView.CellValueChanged -= topbyprdlinedgvView_CellValueChanged_SearchLookup;
                topbyprdlinedgvView.CellValueChanged += topbyprdlinedgvView_CellValueChanged_SearchLookup;
                topbyprdlinedgvView.ShownEditor -= topbyprdlinedgvView_ShownEditor_SearchLookup;
                topbyprdlinedgvView.ShownEditor += topbyprdlinedgvView_ShownEditor_SearchLookup;

                topbyprdlinedgv.KeyDown -= topbyprdlinedgv_KeyDown;
                topbyprdlinedgv.KeyDown += topbyprdlinedgv_KeyDown;
                topbyprdlinedgvView.KeyDown -= topbyprdlinedgv_KeyDown;
                topbyprdlinedgvView.KeyDown += topbyprdlinedgv_KeyDown;

                topbyprdlinedgv.RefreshDataSource();
                topbyprdlinedgvView.RefreshData();
                topbyprdlinedgvView.LayoutChanged();
            }
            catch
            {
                // Tampilan search lookup tidak boleh mengganggu logic edit/save.
            }
        }

        private static DevExpress.XtraGrid.Columns.GridColumn GetGridColumnByNameOrField(DevExpress.XtraGrid.Views.Grid.GridView view, string columnName, string fieldName)
        {
            if (view == null) return null;

            DevExpress.XtraGrid.Columns.GridColumn column = null;

            if (!string.IsNullOrWhiteSpace(columnName))
                column = view.Columns.ColumnByName(columnName);

            if (column == null && !string.IsNullOrWhiteSpace(fieldName))
                column = view.Columns.ColumnByFieldName(fieldName);

            if (column == null && !string.IsNullOrWhiteSpace(fieldName))
                column = view.Columns[fieldName];

            return column;
        }

        private static DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit CreateSearchButtonRepository()
        {
            DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit repository = new DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit();
            ConfigureSearchLookupButton(repository, "Search");
            return repository;
        }

        private static void ConfigureSearchLookupButton(DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit repository, string tooltip)
        {
            if (repository == null) return;

            repository.NullText = "";
            repository.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
            repository.ShowDropDown = DevExpress.XtraEditors.Controls.ShowDropDown.SingleClick;
            repository.PopupFormSize = new System.Drawing.Size(650, 350);
            repository.Buttons.Clear();
            repository.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Search) { ToolTip = tooltip });
        }

        private static DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit CreateButtonRepository()
        {
            DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit repository = new DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit();
            repository.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
            repository.Buttons.Clear();
            repository.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Search) {  ToolTip = "Search" });
            return repository;
        }

        private static void ConfigureSearchButtonColumn(DevExpress.XtraGrid.Columns.GridColumn column, DevExpress.XtraEditors.Repository.RepositoryItem repository, string tooltip)
        {
            if (column == null) return;

            column.ColumnEdit = repository;
            column.Caption = string.Empty;
            column.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            column.AppearanceCell.Options.UseTextOptions = true;
            column.OptionsColumn.ShowCaption = false;
            column.OptionsColumn.AllowEdit = true;
            column.OptionsColumn.ReadOnly = false;
            column.OptionsColumn.FixedWidth = true;
            column.OptionsColumn.AllowFocus = true;
            column.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowAlways;
            column.Width = 28;
            column.MinWidth = 28;
            column.MaxWidth = 28;
            column.ToolTip = tooltip;
        }

        private static void ConfigureLookupView(DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit repository, string[] fieldNames, string[] captions, int[] widths)
        {
            if (repository == null || repository.View == null) return;

            repository.View.OptionsView.ShowGroupPanel = false;
            repository.View.OptionsView.ColumnAutoWidth = false;
            repository.View.Columns.Clear();

            for (int i = 0; i < fieldNames.Length; i++)
            {
                DevExpress.XtraGrid.Columns.GridColumn column = repository.View.Columns.AddField(fieldNames[i]);
                if (column == null) continue;
                column.Caption = captions.Length > i ? captions[i] : fieldNames[i];
                if (widths.Length > i) column.Width = widths[i];
                column.Visible = true;
                column.VisibleIndex = i;
            }
        }

        private DataTable GetGroupHargaLookupDataSource()
        {
            // Ambil data lookup group harga secara DISTINCT supaya SearchLookUpEdit tidak menampilkan baris dobel.
            // Alias dibuat sederhana (Kode/Nama) agar kolom popup tidak tampil sebagai ppl_price_list_code yang terpotong.
            string query = @"
                SELECT
                    ppl_price_list_code AS [Kode],
                    MAX(ppl_price_list_description) AS [Nama]
                FROM OB_PRICE_LIST_HEADER WITH(NOLOCK)
                WHERE ISNULL(LTRIM(RTRIM(ppl_price_list_code)), '') <> ''
                GROUP BY ppl_price_list_code
                ORDER BY ppl_price_list_code";

            return _clsGlobal.ExecDT(query);
        }

        private DataTable GetPrdLineLookupDataSource()
        {
            return _clsGlobal.ExecDT(@"SELECT pl_sapdiv_id AS [Kode], MAX(pl_prd_line_desc) AS [Nama]
                FROM IM_PRD_LINE WITH(NOLOCK)
                WHERE ISNULL(LTRIM(RTRIM(pl_sapdiv_id)), '') <> ''
                GROUP BY pl_sapdiv_id
                ORDER BY pl_sapdiv_id");
        }

        private DataTable GetPlantLookupDataSource()
        {
            try
            {
                // Plant diambil dari GS_BRANCH: kode = br_branch_ud2, deskripsi = br_branch_desc
                return _clsGlobal.ExecDT(@"SELECT LTRIM(RTRIM(br_branch_ud2)) AS [Kode],
                                           MAX(LTRIM(RTRIM(br_branch_desc))) AS [Nama]
                                           FROM GS_BRANCH WITH(NOLOCK)
                                           WHERE ISNULL(LTRIM(RTRIM(br_branch_ud2)), '') <> ''
                                           GROUP BY LTRIM(RTRIM(br_branch_ud2))
                                           ORDER BY [Kode]");
            }
            catch
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Kode", typeof(string));
                dt.Columns.Add("Nama", typeof(string));
                return dt;
            }
        }

        private DataTable GetSalesEmployeeLookupDataSource()
        {
            try
            {
                string entity = __state == StateEntry.Edit ? _prdEntityCode : ctrlEntityCustMaster2.txtCM.Text.Trim();
                string branch = __state == StateEntry.Edit ? _prdBrandCode : ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();

                if (string.IsNullOrWhiteSpace(entity)) entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
                if (string.IsNullOrWhiteSpace(branch)) branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();

                string query =
                    "SELECT DISTINCT LTRIM(RTRIM(sgm_spgm_id)) AS [Kode], " +
                    "LTRIM(RTRIM(sgm_spgm_name)) AS [Nama] " +
                    "FROM SO_SPG_GIRL_MAN WITH(NOLOCK) " +
                    "WHERE sgm_entity_id='" + entity.Replace("'", "''") + "' " +
                    "AND sgm_branch_id='" + branch.Replace("'", "''") + "' " +
                    "AND ISNULL(LTRIM(RTRIM(sgm_spgm_id)), '') <> '' " +
                    "ORDER BY [Kode]";

                return _clsGlobal.ExecDT(query);
            }
            catch
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Kode", typeof(string));
                dt.Columns.Add("Nama", typeof(string));
                return dt;
            }
        }

        private DataTable GetTopLookupDataSource()
        {
            return _clsGlobal.ExecDT("SELECT pptc_term_code AS [Kode], pptc_term_code+' - '+pptc_term_desc AS [Nama] FROM PO_PAYMENT_TERM_CODES WITH(NOLOCK) ORDER BY pptc_no_of_days");
        }

        private void GroupHargaSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void GroupHargaButtonRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            ShowGroupHargaPopupForFocusedRow();
        }

        private void TopPrdLineSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void TopPaymentSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void TopPlantSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void TopSalesEmployeeSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void dgvGroupHargaView_ShownEditor_SearchLookup(object sender, EventArgs e)
        {
            if (dgvGroupHargaView == null || dgvGroupHargaView.FocusedColumn == null || dgvGroupHargaView.FocusedColumn.FieldName != "btnGroup") return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = dgvGroupHargaView.ActiveEditor as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            editor.EditValueChanged -= GroupHargaSearchLookupEditor_EditValueChanged;
            editor.EditValueChanged += GroupHargaSearchLookupEditor_EditValueChanged;
        }

        private void dgvGroupHargaView_RowCellClick_SearchLookup(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e == null || e.RowHandle < 0 || e.Column == null || e.Column.FieldName != "btnGroup") return;

            dgvGroupHargaView.FocusedRowHandle = e.RowHandle;
            dgvGroupHargaView.FocusedColumn = e.Column;
            ShowGroupHargaSearchLookupForFocusedRow();
        }

        private void ShowGroupHargaSearchLookupForFocusedRow()
        {
            if (dgvGroupHargaView == null || dgvGroupHargaView.FocusedRowHandle < 0) return;

            DevExpress.XtraGrid.Columns.GridColumn buttonColumn = GetGridColumnByNameOrField(dgvGroupHargaView, "btnGroup", "btnGroup");
            if (buttonColumn == null) return;

            dgvGroupHargaView.FocusedColumn = buttonColumn;
            dgvGroupHargaView.ShowEditor();

            DevExpress.XtraEditors.SearchLookUpEdit editor = dgvGroupHargaView.ActiveEditor as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen)
                editor.ShowPopup();
        }

        private void ShowGroupHargaPopupForFocusedRow()
        {
            if (dgvGroupHargaView == null || dgvGroupHargaView.FocusedRowHandle < 0) return;
            DevExpress.XtraGrid.Columns.GridColumn buttonColumn = dgvGroupHargaView.Columns["btnGroup"];
            if (buttonColumn == null) return;
            if (_isOpeningGroupHargaPopup) return;

            try
            {
                _isOpeningGroupHargaPopup = true;
                dgvGroupHargaView.FocusedColumn = buttonColumn;
                dgvGroupHarga_CellContentClick(dgvGroupHarga, new System.Windows.Forms.DataGridViewCellEventArgs(buttonColumn.VisibleIndex, dgvGroupHargaView.FocusedRowHandle));
            }
            finally
            {
                _isOpeningGroupHargaPopup = false;
            }
        }

        private void HirPrdLineSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void HirCustTypeSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void DGVHIRView_ShownEditor_SearchLookup(object sender, EventArgs e)
        {
            if (DGVHIRView == null || DGVHIRView.FocusedColumn == null) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = DGVHIRView.ActiveEditor as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;

            if (DGVHIRView.FocusedColumn.FieldName == "btna")
            {
                editor.EditValueChanged -= HirPrdLineSearchLookupEditor_EditValueChanged;
                editor.EditValueChanged += HirPrdLineSearchLookupEditor_EditValueChanged;
            }
            else if (DGVHIRView.FocusedColumn.FieldName == "btnb")
            {
                editor.EditValueChanged -= HirCustTypeSearchLookupEditor_EditValueChanged;
                editor.EditValueChanged += HirCustTypeSearchLookupEditor_EditValueChanged;
            }
        }

        private void HirPrdLineSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingHirFromSearchLookup) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            ApplySelectedHirPrdLineToRow(editor.EditValue, DGVHIRView.FocusedRowHandle);
        }

        private void HirCustTypeSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingHirFromSearchLookup) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            ApplySelectedHirCustTypeToRow(editor.EditValue, DGVHIRView.FocusedRowHandle);
        }

        private void DGVHIRView_CellValueChanged_SearchLookup(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (_isUpdatingHirFromSearchLookup) return;
            if (e == null || e.Column == null) return;
            if (e.Column.FieldName == "btna") ApplySelectedHirPrdLineToRow(e.Value, e.RowHandle);
            else if (e.Column.FieldName == "btnb") ApplySelectedHirCustTypeToRow(e.Value, e.RowHandle);
        }

        private void ApplySelectedHirPrdLineToRow(object value, int rowHandle)
        {
            if (value == null || rowHandle < 0 || DGVHIRView == null) return;
            string code = value.ToString().Trim();
            if (string.IsNullOrWhiteSpace(code)) return;

            DataTable lookup = _hirPrdLineSearchLookupRepository == null ? null : _hirPrdLineSearchLookupRepository.DataSource as DataTable;
            if (lookup == null) return;
            DataRow selected = lookup.Rows.Cast<DataRow>().FirstOrDefault(r => r["Kode"] != DBNull.Value && r["Kode"].ToString().Trim() == code);
            if (selected == null) return;

            try
            {
                _isUpdatingHirFromSearchLookup = true;
                DGVHIRView.SetRowCellValue(rowHandle, "btna", code);
                DGVHIRView.SetRowCellValue(rowHandle, "cpl_line_code", code);
                DGVHIRView.SetRowCellValue(rowHandle, "pl_prd_line_desc", selected["Nama"] == DBNull.Value ? string.Empty : selected["Nama"].ToString().Trim());
                DGVHIRView.PostEditor();
                DGVHIRView.UpdateCurrentRow();
            }
            finally { _isUpdatingHirFromSearchLookup = false; }
        }

        private void ApplySelectedHirCustTypeToRow(object value, int rowHandle)
        {
            if (value == null || rowHandle < 0 || DGVHIRView == null) return;
            string code = value.ToString().Trim();
            if (string.IsNullOrWhiteSpace(code)) return;

            DataTable lookup = _hirCustTypeSearchLookupRepository == null ? null : _hirCustTypeSearchLookupRepository.DataSource as DataTable;
            if (lookup == null) return;
            DataRow selected = lookup.Rows.Cast<DataRow>().FirstOrDefault(r => r["Kode"] != DBNull.Value && r["Kode"].ToString().Trim() == code);
            if (selected == null) return;

            try
            {
                _isUpdatingHirFromSearchLookup = true;
                DGVHIRView.SetRowCellValue(rowHandle, "btnb", code);
                DGVHIRView.SetRowCellValue(rowHandle, "cpl_cust_type", code);
                DGVHIRView.SetRowCellValue(rowHandle, "ct_cust_type_desc", selected["Nama"] == DBNull.Value ? string.Empty : selected["Nama"].ToString().Trim());
                DGVHIRView.PostEditor();
                DGVHIRView.UpdateCurrentRow();
            }
            finally { _isUpdatingHirFromSearchLookup = false; }
        }

        private DataTable GetHirPrdLineLookupDataSource()
        {
            try
            {
                return _clsGlobal.ExecDT(@"SELECT pl_prd_line_code AS [Kode], MAX(pl_prd_line_desc) AS [Nama]
            FROM IM_PRD_LINE WITH(NOLOCK)
            WHERE ISNULL(LTRIM(RTRIM(pl_prd_line_code)), '') <> ''
            GROUP BY pl_prd_line_code
            ORDER BY pl_prd_line_code");
            }
            catch
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Kode", typeof(string));
                dt.Columns.Add("Nama", typeof(string));
                return dt;
            }
        }

        private DataTable GetHirCustTypeLookupDataSource()
        {
            try
            {
                return _clsGlobal.ExecDT(@"SELECT ct_cust_type AS [Kode], MAX(ct_cust_type_desc) AS [Nama]
            FROM SO_CUST_TYPE WITH(NOLOCK)
            WHERE ct_cluster='Y' AND ISNULL(LTRIM(RTRIM(ct_cust_type)), '') <> ''
            GROUP BY ct_cust_type
            ORDER BY ct_cust_type");
            }
            catch
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Kode", typeof(string));
                dt.Columns.Add("Nama", typeof(string));
                return dt;
            }
        }

        private void FillGridPartnerFunction()
        {
            try
            {
                EnsurePartnerFunctionTabRuntime();

                string entity = GetPartnerFunctionEntity();
                string branch = GetPartnerFunctionBranch();
                string custCode1 = txtCustCode.Text.Trim();
                string custCode2 = txtCustCodeTo.Text.Trim();

                DataTable source;
                if (string.IsNullOrWhiteSpace(entity) || string.IsNullOrWhiteSpace(branch) ||
                    string.IsNullOrWhiteSpace(custCode1) || string.IsNullOrWhiteSpace(custCode2))
                {
                    source = CreatePartnerFunctionTable();
                }
                else
                {
                    strSQL = @"
                    SELECT
                        ROW_NUMBER() OVER(ORDER BY m.msh_shipto_code1, m.msh_shipto_code2) AS [no],
                        LTRIM(RTRIM(m.msh_shipto_code1)) AS msh_shipto_code1,
                        LTRIM(RTRIM(m.msh_shipto_code2)) AS msh_shipto_code2,
                        '' AS shipto_btn,
                        ISNULL(cm.cm_cust_name, '') AS shipto_desc,
                        CASE WHEN ISNULL(m.msh_default, 'N') = 'Y' THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS msh_default
                    FROM SD_MULTI_SHIPTO m WITH(NOLOCK)
                    LEFT JOIN SO_CUST_MASTER cm WITH(NOLOCK)
                        ON cm.cm_entity = m.msh_entity_id
                        AND cm.cm_branch = m.msh_branch_id
                        AND cm.cm_cust_code1 = m.msh_shipto_code1
                        AND cm.cm_cust_code2 = m.msh_shipto_code2
                    WHERE m.msh_entity_id = '" + entity.Replace("'", "''") + @"'
                      AND m.msh_branch_id = '" + branch.Replace("'", "''") + @"'
                      AND m.msh_cust_code1 = '" + custCode1.Replace("'", "''") + @"'
                      AND m.msh_cust_code2 = '" + custCode2.Replace("'", "''") + @"'
                    ORDER BY m.msh_shipto_code1, m.msh_shipto_code2";

                    source = _clsGlobal.ExecDT(strSQL);
                    EnsurePartnerFunctionColumns(source);
                }

                partnerFunctionGrid.DataSource = source;
                ConfigurePartnerFunctionGridRuntime();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GroupHargaSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingGroupHargaFromSearchLookup) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            ApplySelectedGroupHargaToRow(editor.EditValue, dgvGroupHargaView.FocusedRowHandle);
        }

        private void dgvGroupHargaView_CellValueChanged_SearchLookup(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (_isUpdatingGroupHargaFromSearchLookup) return;
            if (e == null || e.Column == null || e.Column.FieldName != "btnGroup") return;
            ApplySelectedGroupHargaToRow(e.Value, e.RowHandle);
        }

        private void ApplySelectedGroupHargaToRow(object groupCodeValue, int rowHandle)
        {
            if (groupCodeValue == null || rowHandle < 0 || dgvGroupHargaView == null) return;
            string groupCode = groupCodeValue.ToString().Trim();
            if (string.IsNullOrWhiteSpace(groupCode)) return;

            DataTable lookup = _groupHargaSearchLookupRepository == null ? null : _groupHargaSearchLookupRepository.DataSource as DataTable;
            if (lookup == null) return;

            DataRow selected = lookup.Rows.Cast<DataRow>().FirstOrDefault(r => r["Kode"] != DBNull.Value && r["Kode"].ToString().Trim() == groupCode);
            if (selected == null) return;

            try
            {
                _isUpdatingGroupHargaFromSearchLookup = true;
                dgvGroupHargaView.SetRowCellValue(rowHandle, "btnGroup", groupCode);
                dgvGroupHargaView.SetRowCellValue(rowHandle, "grp_code", groupCode);
                dgvGroupHargaView.SetRowCellValue(rowHandle, "grp_desc", selected["Nama"] == DBNull.Value ? string.Empty : selected["Nama"].ToString().Trim());
                dgvGroupHargaView.PostEditor();
                dgvGroupHargaView.UpdateCurrentRow();
            }
            finally
            {
                _isUpdatingGroupHargaFromSearchLookup = false;
            }
        }

        private void topbyprdlinedgvView_ShownEditor_SearchLookup(object sender, EventArgs e)
        {
            if (topbyprdlinedgvView == null || topbyprdlinedgvView.FocusedColumn == null) return;

            DevExpress.XtraEditors.SearchLookUpEdit editor = topbyprdlinedgvView.ActiveEditor as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;

            if (topbyprdlinedgvView.FocusedColumn.FieldName == "prdline_btn")
            {
                editor.EditValueChanged -= TopPrdLineSearchLookupEditor_EditValueChanged;
                editor.EditValueChanged += TopPrdLineSearchLookupEditor_EditValueChanged;
            }
            else if (topbyprdlinedgvView.FocusedColumn.FieldName == "top_btn")
            {
                editor.EditValueChanged -= TopPaymentSearchLookupEditor_EditValueChanged;
                editor.EditValueChanged += TopPaymentSearchLookupEditor_EditValueChanged;
            }
            else if (topbyprdlinedgvView.FocusedColumn.FieldName == "plant_btn")
            {
                editor.EditValueChanged -= TopPlantSearchLookupEditor_EditValueChanged;
                editor.EditValueChanged += TopPlantSearchLookupEditor_EditValueChanged;
            }
            else if (topbyprdlinedgvView.FocusedColumn.FieldName == "sales_employee_btn")
            {
                editor.EditValueChanged -= TopSalesEmployeeSearchLookupEditor_EditValueChanged;
                editor.EditValueChanged += TopSalesEmployeeSearchLookupEditor_EditValueChanged;
            }
        }

        private void TopPrdLineSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingTopByPrdLineFromSearchLookup) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            ApplySelectedTopPrdLineToRow(editor.EditValue, topbyprdlinedgvView.FocusedRowHandle);
        }

        private void TopPaymentSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingTopByPrdLineFromSearchLookup) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            ApplySelectedTopPaymentToRow(editor.EditValue, topbyprdlinedgvView.FocusedRowHandle);
        }

        private void TopPlantSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingTopByPrdLineFromSearchLookup) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            ApplySelectedTopPlantToRow(editor.EditValue, topbyprdlinedgvView.FocusedRowHandle);
        }

        private void TopSalesEmployeeSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingTopByPrdLineFromSearchLookup) return;
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;
            ApplySelectedTopSalesEmployeeToRow(editor.EditValue, topbyprdlinedgvView.FocusedRowHandle);
        }

        private void topbyprdlinedgvView_CellValueChanged_SearchLookup(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (_isUpdatingTopByPrdLineFromSearchLookup) return;
            if (e == null || e.Column == null) return;
            if (e.Column.FieldName == "prdline_btn") ApplySelectedTopPrdLineToRow(e.Value, e.RowHandle);
            else if (e.Column.FieldName == "top_btn") ApplySelectedTopPaymentToRow(e.Value, e.RowHandle);
            else if (e.Column.FieldName == "plant_btn") ApplySelectedTopPlantToRow(e.Value, e.RowHandle);
            else if (e.Column.FieldName == "sales_employee_btn") ApplySelectedTopSalesEmployeeToRow(e.Value, e.RowHandle);
        }

        private void ApplySelectedTopPrdLineToRow(object prdLineValue, int rowHandle)
        {
            if (prdLineValue == null || rowHandle < 0 || topbyprdlinedgvView == null) return;
            string prdLine = prdLineValue.ToString().Trim();
            if (string.IsNullOrWhiteSpace(prdLine)) return;

            DataTable lookup = _topPrdLineSearchLookupRepository == null ? null : _topPrdLineSearchLookupRepository.DataSource as DataTable;
            if (lookup == null) return;

            DataRow selected = lookup.Rows.Cast<DataRow>().FirstOrDefault(r => r["Kode"] != DBNull.Value && r["Kode"].ToString().Trim() == prdLine);
            if (selected == null) return;

            try
            {
                _isUpdatingTopByPrdLineFromSearchLookup = true;
                topbyprdlinedgvView.SetRowCellValue(rowHandle, "prdline_btn", prdLine);
                topbyprdlinedgvView.SetRowCellValue(rowHandle, "prdline_code", prdLine);
                topbyprdlinedgvView.SetRowCellValue(rowHandle, "prdline_desc", selected["Nama"] == DBNull.Value ? string.Empty : selected["Nama"].ToString().Trim());
                topbyprdlinedgvView.PostEditor();
                topbyprdlinedgvView.UpdateCurrentRow();
            }
            finally
            {
                _isUpdatingTopByPrdLineFromSearchLookup = false;
            }
        }

        private void ApplySelectedTopPaymentToRow(object topValue, int rowHandle)
        {
            if (topValue == null || rowHandle < 0 || topbyprdlinedgvView == null) return;
            string topCode = topValue.ToString().Trim();
            if (string.IsNullOrWhiteSpace(topCode)) return;

            DataTable lookup = _topPaymentSearchLookupRepository == null ? null : _topPaymentSearchLookupRepository.DataSource as DataTable;
            if (lookup == null) return;

            DataRow selected = lookup.Rows.Cast<DataRow>().FirstOrDefault(r => r["Kode"] != DBNull.Value && r["Kode"].ToString().Trim() == topCode);
            if (selected == null) return;

            try
            {
                _isUpdatingTopByPrdLineFromSearchLookup = true;
                topbyprdlinedgvView.SetRowCellValue(rowHandle, "top_btn", topCode);
                topbyprdlinedgvView.SetRowCellValue(rowHandle, "top_code", topCode);
                topbyprdlinedgvView.SetRowCellValue(rowHandle, "top_desc", selected["Nama"] == DBNull.Value ? string.Empty : selected["Nama"].ToString().Trim());
                topbyprdlinedgvView.PostEditor();
                topbyprdlinedgvView.UpdateCurrentRow();
            }
            finally
            {
                _isUpdatingTopByPrdLineFromSearchLookup = false;
            }
        }

        private void UpdateCustomerByDivisionEnhancementColumnsTrans(IEnumerable<DataRow> rows, string custCode1, string custCode2)
        {
            if (rows == null) return;

            string entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
            string branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();

            foreach (DataRow r in rows.Where(x => x.Table.Columns.Contains("prdline_code") && !x["prdline_code"].ToString().IsNullOrEmptyOrWhiteSpace()))
            {
                string divisionCode = r["prdline_code"].ToString().Trim();
                string topCode = r.Table.Columns.Contains("top_code") ? r["top_code"].ToString().Trim() : string.Empty;
                string plantId = r.Table.Columns.Contains("plant_id") ? r["plant_id"].ToString().Trim() : string.Empty;
                string salesEmployeeId = r.Table.Columns.Contains("sales_employee_id") ? r["sales_employee_id"].ToString().Trim() : string.Empty;

                if (string.IsNullOrWhiteSpace(divisionCode) || string.IsNullOrWhiteSpace(topCode))
                    continue;

                // SP_SDM1_SAVE_TOPBYPEDLINE tetap menjadi logic utama insert/update core mapping.
                // Blok berikut memastikan field enhancement ikut terisi untuk row existing maupun row hasil F10/extend.
                string sql = @"
                IF EXISTS (
                    SELECT 1
                    FROM dbo.SO_MAPPING_TOPBYPRDLINE WITH (UPDLOCK, HOLDLOCK)
                    WHERE smt_entity_id = {0}
                      AND smt_branch_id = {1}
                      AND smt_cust_code1 = {2}
                      AND smt_cust_code2 = {3}
                      AND smt_prdline_id = {4}
                )
                BEGIN
                    UPDATE dbo.SO_MAPPING_TOPBYPRDLINE
                       SET smt_top_id = {5}
                " +
                ((true) ? @"
                         , smt_plant_id = CASE WHEN COL_LENGTH('dbo.SO_MAPPING_TOPBYPRDLINE','smt_plant_id') IS NOT NULL THEN {6} ELSE smt_plant_id END
                         , smt_sales_employee_id = CASE WHEN COL_LENGTH('dbo.SO_MAPPING_TOPBYPRDLINE','smt_sales_employee_id') IS NOT NULL THEN {7} ELSE smt_sales_employee_id END
                " : string.Empty) +
                @"     WHERE smt_entity_id = {0}
                       AND smt_branch_id = {1}
                       AND smt_cust_code1 = {2}
                       AND smt_cust_code2 = {3}
                       AND smt_prdline_id = {4};
                END
                ELSE
                BEGIN
                    INSERT INTO dbo.SO_MAPPING_TOPBYPRDLINE
                    (smt_entity_id, smt_branch_id, smt_cust_code1, smt_cust_code2, smt_prdline_id, smt_top_id)
                    VALUES ({0}, {1}, {2}, {3}, {4}, {5});
                
                    IF COL_LENGTH('dbo.SO_MAPPING_TOPBYPRDLINE','smt_plant_id') IS NOT NULL
                    BEGIN
                        UPDATE dbo.SO_MAPPING_TOPBYPRDLINE
                           SET smt_plant_id = {6}
                         WHERE smt_entity_id = {0}
                           AND smt_branch_id = {1}
                           AND smt_cust_code1 = {2}
                           AND smt_cust_code2 = {3}
                           AND smt_prdline_id = {4};
                    END
                
                    IF COL_LENGTH('dbo.SO_MAPPING_TOPBYPRDLINE','smt_sales_employee_id') IS NOT NULL
                    BEGIN
                        UPDATE dbo.SO_MAPPING_TOPBYPRDLINE
                           SET smt_sales_employee_id = {7}
                         WHERE smt_entity_id = {0}
                           AND smt_branch_id = {1}
                           AND smt_cust_code1 = {2}
                           AND smt_cust_code2 = {3}
                           AND smt_prdline_id = {4};
                    END
                END";

                sql = string.Format(sql,
                    FmtStr(entity),
                    FmtStr(branch),
                    FmtStr(custCode1),
                    FmtStr(custCode2),
                    FmtStr(divisionCode),
                    FmtStr(topCode),
                    FmtStr(plantId),
                    FmtStr(salesEmployeeId));

                _clsGlobal.ExecuteTrans(sql);
            }
        }

        private void ApplySelectedTopPlantToRow(object plantValue, int rowHandle)
        {
            ApplySelectedSimpleLookupToTopByDivisionRow(plantValue, rowHandle, _topPlantSearchLookupRepository, "plant_btn", "plant_id", "plant_desc");
        }

        private void ApplySelectedTopSalesEmployeeToRow(object salesEmployeeValue, int rowHandle)
        {
            ApplySelectedSimpleLookupToTopByDivisionRow(salesEmployeeValue, rowHandle, _topSalesEmployeeSearchLookupRepository, "sales_employee_btn", "sales_employee_id", "sales_employee_desc");
        }

        private void ApplySelectedSimpleLookupToTopByDivisionRow(object value, int rowHandle, DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit repository, string buttonField, string targetField, string descField)
        {
            if (value == null || rowHandle < 0 || topbyprdlinedgvView == null) return;
            string code = value.ToString().Trim();
            if (string.IsNullOrWhiteSpace(code)) return;

            string description = GetTopByDivisionLookupDescription(repository, code);

            try
            {
                _isUpdatingTopByPrdLineFromSearchLookup = true;
                topbyprdlinedgvView.SetRowCellValue(rowHandle, buttonField, code);
                topbyprdlinedgvView.SetRowCellValue(rowHandle, targetField, code);
                if (!string.IsNullOrWhiteSpace(descField))
                    topbyprdlinedgvView.SetRowCellValue(rowHandle, descField, description);
                topbyprdlinedgvView.PostEditor();
                topbyprdlinedgvView.UpdateCurrentRow();
            }
            finally
            {
                _isUpdatingTopByPrdLineFromSearchLookup = false;
            }
        }

        private string GetTopByDivisionLookupDescription(DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit repository, string code)
        {
            DataTable lookup = repository == null ? null : repository.DataSource as DataTable;
            if (lookup == null || string.IsNullOrWhiteSpace(code) || !lookup.Columns.Contains("Kode") || !lookup.Columns.Contains("Nama"))
                return string.Empty;

            DataRow selected = lookup.Rows.Cast<DataRow>()
                .FirstOrDefault(r => r["Kode"] != DBNull.Value && string.Equals(r["Kode"].ToString().Trim(), code, StringComparison.OrdinalIgnoreCase));
            return selected == null || selected["Nama"] == DBNull.Value ? string.Empty : selected["Nama"].ToString().Trim();
        }

        private void FillMissingTopByDivisionDescriptions(DataTable source)
        {
            if (source == null) return;
            EnsureTopByDivisionDataColumns(source);

            foreach (DataRow row in source.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                string plantId = row["plant_id"] == DBNull.Value ? string.Empty : row["plant_id"].ToString().Trim();
                if (!string.IsNullOrWhiteSpace(plantId) && string.IsNullOrWhiteSpace(row["plant_desc"].ToString()))
                    row["plant_desc"] = GetTopByDivisionLookupDescription(_topPlantSearchLookupRepository, plantId);

                string salesEmployeeId = row["sales_employee_id"] == DBNull.Value ? string.Empty : row["sales_employee_id"].ToString().Trim();
                if (!string.IsNullOrWhiteSpace(salesEmployeeId) && string.IsNullOrWhiteSpace(row["sales_employee_desc"].ToString()))
                    row["sales_employee_desc"] = GetTopByDivisionLookupDescription(_topSalesEmployeeSearchLookupRepository, salesEmployeeId);
            }
        }

        private void ShowTopByDivisionRepositoryPopup(string fieldName)
        {
            if (topbyprdlinedgvView == null || string.IsNullOrEmpty(fieldName)) return;
            var column = topbyprdlinedgvView.Columns.ColumnByFieldName(fieldName);
            if (column == null) return;
            topbyprdlinedgvView.FocusedColumn = column;
            topbyprdlinedgvView.ShowEditor();
            DevExpress.XtraEditors.SearchLookUpEdit editor = topbyprdlinedgvView.ActiveEditor as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen) editor.ShowPopup();
        }

        private void ConfigureSalesmanCoveredGridRuntime()
        {
            if (DVGSalesman == null || DVGSalesmanView == null) return;

            try
            {
                DataTable source = DVGSalesman.DataSource as DataTable;
                if (source != null && !source.Columns.Contains("btnc"))
                    source.Columns.Add("btnc", typeof(string));

                EnsureSalesmanBandedGridView();

                DevExpress.XtraGrid.Views.BandedGrid.BandedGridView bandedView =
                    DVGSalesmanView as DevExpress.XtraGrid.Views.BandedGrid.BandedGridView;
                if (bandedView == null) return;

                if (_salesmanSearchLookupRepository == null)
                {
                    _salesmanSearchLookupRepository = new DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit();
                    _salesmanSearchLookupRepository.NullText = "";
                    _salesmanSearchLookupRepository.DisplayMember = "sgm_spgm_id";
                    _salesmanSearchLookupRepository.ValueMember = "sgm_spgm_id";
                    _salesmanSearchLookupRepository.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
                    _salesmanSearchLookupRepository.ShowDropDown = DevExpress.XtraEditors.Controls.ShowDropDown.SingleClick;
                    _salesmanSearchLookupRepository.PopupFormSize = new System.Drawing.Size(650, 350);
                    _salesmanSearchLookupRepository.Buttons.Clear();
                    _salesmanSearchLookupRepository.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Search));
                    _salesmanSearchLookupRepository.ButtonClick -= SalesmanSearchLookupRepository_ButtonClick;
                    _salesmanSearchLookupRepository.ButtonClick += SalesmanSearchLookupRepository_ButtonClick;
                    if (!DVGSalesman.RepositoryItems.Contains(_salesmanSearchLookupRepository))
                        DVGSalesman.RepositoryItems.Add(_salesmanSearchLookupRepository);
                }

                if (_salesmanWeekCheckRepository == null)
                {
                    _salesmanWeekCheckRepository = new DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit();
                    _salesmanWeekCheckRepository.ValueChecked = "True";
                    _salesmanWeekCheckRepository.ValueUnchecked = "False";
                    _salesmanWeekCheckRepository.ValueGrayed = "False";
                    _salesmanWeekCheckRepository.NullStyle = DevExpress.XtraEditors.Controls.StyleIndeterminate.Unchecked;
                    if (!DVGSalesman.RepositoryItems.Contains(_salesmanWeekCheckRepository))
                        DVGSalesman.RepositoryItems.Add(_salesmanWeekCheckRepository);
                }

                if (_salesmanDayComboRepository == null)
                {
                    _salesmanDayComboRepository = new DevExpress.XtraEditors.Repository.RepositoryItemComboBox();
                    _salesmanDayComboRepository.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
                    _salesmanDayComboRepository.Items.AddRange(new object[] { "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu" });
                    if (!DVGSalesman.RepositoryItems.Contains(_salesmanDayComboRepository))
                        DVGSalesman.RepositoryItems.Add(_salesmanDayComboRepository);
                }

                _salesmanSearchLookupRepository.DataSource = GetSalesmanLookupDataSource();
                ConfigureSalesmanLookupView(_salesmanSearchLookupRepository);

                ConfigureSalesmanBandedColumns(bandedView);

                DVGSalesmanView.OptionsView.ShowButtonMode = DevExpress.XtraGrid.Views.Base.ShowButtonModeEnum.ShowAlways;
                DVGSalesmanView.OptionsView.ShowGroupPanel = false;
                DVGSalesmanView.OptionsView.ColumnAutoWidth = false;
                DVGSalesmanView.OptionsSelection.EnableAppearanceFocusedCell = true;
                DVGSalesmanView.OptionsBehavior.Editable = true;
                DVGSalesmanView.OptionsView.ShowColumnHeaders = true;

                DVGSalesmanView.CellValueChanged -= DVGSalesmanView_CellValueChanged;
                DVGSalesmanView.CellValueChanged += DVGSalesmanView_CellValueChanged;
                DVGSalesmanView.ShownEditor -= DVGSalesmanView_ShownEditor;
                DVGSalesmanView.ShownEditor += DVGSalesmanView_ShownEditor;

                DVGSalesman.RefreshDataSource();
                DVGSalesmanView.RefreshData();
                DVGSalesmanView.LayoutChanged();
            }
            catch
            {
                // Jangan ganggu proses edit/save jika hanya konfigurasi tampilan grid salesman yang gagal.
            }
        }

        private void EnsureSalesmanBandedGridView()
        {
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridView bandedView =
                DVGSalesmanView as DevExpress.XtraGrid.Views.BandedGrid.BandedGridView;
            if (bandedView != null)
            {
                if (DVGSalesman.MainView != bandedView)
                    DVGSalesman.MainView = bandedView;
                return;
            }

            bandedView = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridView(DVGSalesman);
            bandedView.Name = "DVGSalesmanView";
            bandedView.GridControl = DVGSalesman;
            bandedView.IndicatorWidth = DVGSalesmanView.IndicatorWidth > 0 ? DVGSalesmanView.IndicatorWidth : 40;
            bandedView.OptionsView.ShowGroupPanel = false;
            bandedView.OptionsView.ColumnAutoWidth = false;
            bandedView.OptionsView.ShowBands = true;
            bandedView.OptionsView.ShowColumnHeaders = true;
            bandedView.OptionsSelection.EnableAppearanceFocusedCell = true;
            bandedView.OptionsBehavior.Editable = true;

            DVGSalesman.ViewCollection.Add(bandedView);

            DVGSalesman.MainView = bandedView;
            DVGSalesmanView = bandedView;
        }

        private void ConfigureSalesmanBandedColumns(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView view)
        {
            if (view == null) return;

            view.BeginUpdate();
            try
            {
                view.Bands.Clear();
                view.Columns.Clear();

                DevExpress.XtraGrid.Views.BandedGrid.GridBand bandSalesman = CreateSalesmanBand(view, "Salesman", 0);
                DevExpress.XtraGrid.Views.BandedGrid.GridBand bandHari = CreateSalesmanBand(view, "Hari", 1);
                DevExpress.XtraGrid.Views.BandedGrid.GridBand bandPola = CreateSalesmanBand(view, "Pola", 2);
                DevExpress.XtraGrid.Views.BandedGrid.GridBand bandRoute = CreateSalesmanBand(view, "Rute", 3);

                CreateSalesmanBandedColumn(view, bandSalesman, "csc_salesman_id", "csc_salesman_id", "Kode", 90, 0, false, null);
                DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn searchColumn = CreateSalesmanBandedColumn(view, bandSalesman, "btnc", "btnc", "", 28, 1, true, _salesmanSearchLookupRepository);
                btnc = searchColumn;
                CreateSalesmanBandedColumn(view, bandSalesman, "sgm_spgm_name", "sgm_spgm_name", "Nama Salesman", 190, 2, false, null);
                CreateSalesmanBandedColumn(view, bandSalesman, "sgm_type_operasi", "sgm_type_operasi", "Type", 80, 3, false, null);
                CreateSalesmanBandedColumn(view, bandHari, "csc_visit", "csc_visit", "Hari", 80, 4, true, _salesmanDayComboRepository);
                CreateSalesmanBandedColumn(view, bandPola, "csc_visit_week1", "csc_visit_week1", "1", 45, 5, true, _salesmanWeekCheckRepository);
                CreateSalesmanBandedColumn(view, bandPola, "csc_visit_week2", "csc_visit_week2", "2", 45, 6, true, _salesmanWeekCheckRepository);
                CreateSalesmanBandedColumn(view, bandPola, "csc_visit_week3", "csc_visit_week3", "3", 45, 7, true, _salesmanWeekCheckRepository);
                CreateSalesmanBandedColumn(view, bandPola, "csc_visit_week4", "csc_visit_week4", "4", 45, 8, true, _salesmanWeekCheckRepository);
                CreateSalesmanBandedColumn(view, bandRoute, "csc_route", "csc_route", "Route", 60, 9, true, null);

                foreach (DevExpress.XtraGrid.Views.BandedGrid.GridBand band in view.Bands)
                {
                    band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    band.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    band.AppearanceHeader.Options.UseTextOptions = true;
                }

                foreach (DevExpress.XtraGrid.Columns.GridColumn column in view.Columns)
                {
                    column.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    column.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    column.AppearanceHeader.Options.UseTextOptions = true;
                }
            }
            finally
            {
                view.EndUpdate();
            }
        }

        private static DevExpress.XtraGrid.Views.BandedGrid.GridBand CreateSalesmanBand(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView view, string caption, int visibleIndex)
        {
            DevExpress.XtraGrid.Views.BandedGrid.GridBand band = view.Bands.AddBand(caption);
            band.VisibleIndex = visibleIndex;
            band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            band.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            band.AppearanceHeader.Options.UseTextOptions = true;
            return band;
        }

        private static DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn CreateSalesmanBandedColumn(
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridView view,
            DevExpress.XtraGrid.Views.BandedGrid.GridBand band,
            string name,
            string fieldName,
            string caption,
            int width,
            int visibleIndex,
            bool allowEdit,
            DevExpress.XtraEditors.Repository.RepositoryItem repositoryItem)
        {
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn column = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            column.Name = name;
            column.FieldName = fieldName;
            column.Caption = caption;
            column.Visible = true;
            column.VisibleIndex = visibleIndex;
            column.Width = width;
            column.OptionsColumn.AllowEdit = allowEdit;
            column.OptionsColumn.ReadOnly = !allowEdit;
            column.OptionsColumn.FixedWidth = (fieldName == "btnc");
            if (repositoryItem != null)
                column.ColumnEdit = repositoryItem;

            if (fieldName == "btnc")
            {
                column.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                column.AppearanceCell.Options.UseTextOptions = true;
                column.OptionsColumn.ShowCaption = false;
                column.ToolTip = "Search Salesman";
                column.MinWidth = width;
                column.MaxWidth = width;
            }

            view.Columns.Add(column);
            band.Columns.Add(column);
            return column;
        }

        private static void PrepareGridColumn(DevExpress.XtraGrid.Views.Grid.GridView view, string columnName, string fieldName, string caption, int width, int visibleIndex, bool allowEdit)
        {
            if (view == null) return;

            DevExpress.XtraGrid.Columns.GridColumn column = view.Columns.ColumnByName(columnName);
            if (column == null && !string.IsNullOrWhiteSpace(fieldName))
                column = view.Columns[fieldName];
            if (column == null)
            {
                column = view.Columns.AddField(fieldName);
                column.Name = columnName;
            }

            column.FieldName = fieldName;
            column.Caption = caption;
            column.Visible = true;
            column.VisibleIndex = visibleIndex;
            column.Width = width;
            column.OptionsColumn.AllowEdit = allowEdit;
            column.OptionsColumn.ReadOnly = !allowEdit;
        }

        private DataTable GetSalesmanLookupDataSource()
        {
            string entity = __state == StateEntry.Edit ? _prdEntityCode : ctrlEntityCustMaster2.txtCM.Text.Trim();
            string branch = __state == StateEntry.Edit ? _prdBrandCode : ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();

            if (string.IsNullOrWhiteSpace(entity)) entity = ctrlEntityCustMaster2.txtCM.Text.Trim();
            if (string.IsNullOrWhiteSpace(branch)) branch = ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim();

            string query = "SELECT sgm_spgm_id, sgm_spgm_name, sgm_type_operasi " +
                           "FROM SO_SPG_GIRL_MAN WITH(NOLOCK) " +
                           "WHERE sgm_entity_id='" + entity.Replace("'", "''") + "' " +
                           "AND sgm_branch_id='" + branch.Replace("'", "''") + "' " +
                           "ORDER BY sgm_spgm_id";

            return _clsGlobal.ExecDT(query);
        }

        private void SalesmanSearchLookupRepository_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor != null && !editor.IsPopupOpen)
                editor.ShowPopup();
        }

        private static void ConfigureSalesmanLookupView(DevExpress.XtraEditors.Repository.RepositoryItemSearchLookUpEdit repository)
        {
            if (repository == null || repository.View == null) return;

            repository.View.OptionsView.ShowGroupPanel = false;
            repository.View.OptionsView.ColumnAutoWidth = false;
            repository.View.PopulateColumns();

            if (repository.View.Columns["sgm_spgm_id"] != null)
            {
                repository.View.Columns["sgm_spgm_id"].Caption = "Kode";
                repository.View.Columns["sgm_spgm_id"].Width = 100;
            }
            if (repository.View.Columns["sgm_spgm_name"] != null)
            {
                repository.View.Columns["sgm_spgm_name"].Caption = "Nama Salesman";
                repository.View.Columns["sgm_spgm_name"].Width = 250;
            }
            if (repository.View.Columns["sgm_type_operasi"] != null)
            {
                repository.View.Columns["sgm_type_operasi"].Caption = "Type";
                repository.View.Columns["sgm_type_operasi"].Width = 80;
            }
        }

        private void DVGSalesmanView_ShownEditor(object sender, EventArgs e)
        {
            if (DVGSalesmanView == null || DVGSalesmanView.FocusedColumn == null || DVGSalesmanView.FocusedColumn.FieldName != "btnc") return;

            DevExpress.XtraEditors.SearchLookUpEdit editor = DVGSalesmanView.ActiveEditor as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;

            editor.EditValueChanged -= SalesmanSearchLookupEditor_EditValueChanged;
            editor.EditValueChanged += SalesmanSearchLookupEditor_EditValueChanged;
        }

        private void SalesmanSearchLookupEditor_EditValueChanged(object sender, EventArgs e)
        {
            if (_isUpdatingSalesmanFromSearchLookup) return;

            DevExpress.XtraEditors.SearchLookUpEdit editor = sender as DevExpress.XtraEditors.SearchLookUpEdit;
            if (editor == null) return;

            ApplySelectedSalesmanToRow(editor.EditValue, DVGSalesmanView.FocusedRowHandle);
        }

        private void DVGSalesmanView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (_isUpdatingSalesmanFromSearchLookup) return;
            if (e == null || e.Column == null) return;

            if (e.Column.FieldName == "btnc")
                ApplySelectedSalesmanToRow(e.Value, e.RowHandle);
        }

        private void ApplySelectedSalesmanToRow(object salesmanIdValue, int rowHandle)
        {
            if (salesmanIdValue == null || rowHandle < 0 || DVGSalesmanView == null) return;

            string salesmanId = salesmanIdValue.ToString().Trim();
            if (string.IsNullOrWhiteSpace(salesmanId)) return;

            DataTable lookup = _salesmanSearchLookupRepository == null ? null : _salesmanSearchLookupRepository.DataSource as DataTable;
            if (lookup == null) return;

            DataRow selected = null;
            foreach (DataRow row in lookup.Rows)
            {
                if (row["sgm_spgm_id"] != DBNull.Value && row["sgm_spgm_id"].ToString().Trim() == salesmanId)
                {
                    selected = row;
                    break;
                }
            }
            if (selected == null) return;

            try
            {
                _isUpdatingSalesmanFromSearchLookup = true;
                DVGSalesmanView.SetRowCellValue(rowHandle, "btnc", salesmanId);
                DVGSalesmanView.SetRowCellValue(rowHandle, "csc_salesman_id", salesmanId);
                DVGSalesmanView.SetRowCellValue(rowHandle, "sgm_spgm_name", selected["sgm_spgm_name"] == DBNull.Value ? string.Empty : selected["sgm_spgm_name"].ToString().Trim());
                DVGSalesmanView.SetRowCellValue(rowHandle, "sgm_type_operasi", selected["sgm_type_operasi"] == DBNull.Value ? string.Empty : selected["sgm_type_operasi"].ToString().Trim());
                DVGSalesmanView.PostEditor();
                DVGSalesmanView.UpdateCurrentRow();
            }
            finally
            {
                _isUpdatingSalesmanFromSearchLookup = false;
            }
        }

        private void EnsureAllDevExpressGridColumnsVisible()
        {
            EnsureDevExpressGridColumnsVisible(gridPajak, gridPajakView);
            EnsureDevExpressGridColumnsVisible(DGVHIR, DGVHIRView);
            EnsureDevExpressGridColumnsVisible(DVGSalesman, DVGSalesmanView);
            EnsureDevExpressGridColumnsVisible(DGVBackList, DGVBackListView);
            EnsureDevExpressGridColumnsVisible(DGVUnBackList, DGVUnBackListView);
            EnsureDevExpressGridColumnsVisible(dgvGroupHarga, dgvGroupHargaView);
            EnsureDevExpressGridColumnsVisible(topbyprdlinedgv, topbyprdlinedgvView);
            EnsureDevExpressGridColumnsVisible(dgvschpayD, dgvschpayDView);
            EnsureDevExpressGridColumnsVisible(dgvschpayW, dgvschpayWView);
            EnsureDevExpressGridColumnsVisible(dgvPrdLineBlocking, dgvPrdLineBlockingView);
        }

        private static void EnsureDevExpressGridColumnsVisible(DevExpress.XtraGrid.GridControl grid, DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            if (grid == null || view == null) return;

            try
            {
                if (grid.MainView != view)
                    grid.MainView = view;

                bool viewAlreadyRegistered = false;
                foreach (DevExpress.XtraGrid.Views.Base.BaseView registeredView in grid.ViewCollection)
                {
                    if (object.ReferenceEquals(registeredView, view))
                    {
                        viewAlreadyRegistered = true;
                        break;
                    }
                }
                if (!viewAlreadyRegistered)
                    grid.ViewCollection.Add(view);

                if (view.Columns.Count == 0 && grid.DataSource != null)
                    view.PopulateColumns();

                for (int i = 0; i < view.Columns.Count; i++)
                {
                    DevExpress.XtraGrid.Columns.GridColumn column = view.Columns[i];
                    column.Visible = true;
                    if (column.VisibleIndex < 0)
                        column.VisibleIndex = i;
                }

                view.OptionsView.ShowColumnHeaders = true;
                view.OptionsView.ShowGroupPanel = false;
                view.OptionsBehavior.Editable = true;
                view.BestFitColumns();
                view.LayoutChanged();
                grid.RefreshDataSource();
            }
            catch
            {
                // Jangan mengganggu logic save/edit jika hanya refresh tampilan grid yang gagal.
            }
        }

        private static System.Windows.Forms.TextBox CreateHiddenTextBox(DevExpress.XtraEditors.TextEdit source)
        {
            var tb = new System.Windows.Forms.TextBox();
            if (source != null)
            {
                bool updating = false;
                tb.Text = source.Text;
                tb.TextChanged += (s, e) =>
                {
                    if (updating) return;
                    updating = true;
                    try
                    {
                        source.Text = tb.Text;
                    }
                    finally
                    {
                        updating = false;
                    }
                };
                source.EditValueChanged += (s, e) =>
                {
                    if (updating) return;
                    updating = true;
                    try
                    {
                        tb.Text = source.Text;
                    }
                    finally
                    {
                        updating = false;
                    }
                };
            }
            return tb;
        }
        #endregion
    }

    public class GsGenHardcodedItem
    {
        public decimal No { get; set; }
        public string Code { get; set; }
    }
    public class TIRAItem
    {
        public string Kode { get; set; }
        public string Nama { get; set; }
        public TIRAItem() { }
        public TIRAItem(string __kode, string __nama)
        {
            this.Kode = __kode;
            this.Nama = __nama;
        }
    }
    public class TiraPhoto
    {
        public TiraPhoto()
        {
            Code = string.Empty;
            IsNIK = false;
            IsNPWP = false;
        }
        public string Code { get; set; }
        public bool IsNIK { get; set; }
        public bool IsNPWP { get; set; }
    }
    public class IndexGrid
    {
        public int Column { get; set; }
        public int Row { get; set; }
        public IndexGrid(int column, int row)
        {
            this.Column = column;
            this.Row = row;
        }
        public IndexGrid(System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            this.Column = e.ColumnIndex;
            this.Row = e.RowIndex;
        }
    }
    public class Hari
    {
        public int Code { get; set; }
        public string Name { get; set; }
    }


    /** SYNC **/
    public class CreateCollectionModel
    {
        public string groupPriceRef { get; set; }
        public int priority { get; set; }
    }

    /** END OF SYNC **/


}
