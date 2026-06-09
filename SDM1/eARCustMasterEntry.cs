using StackedHeader;
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
    public partial class eARCustMasterEntry : Form
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
            InitializeComponent();
            __COLOR = Color.FromArgb(245, 245, 255);
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
                comboBoxPajakT1.SelectedIndex = 1;
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

            this.DVGSalesman.ColumnHeadersHeight = this.DVGSalesman.ColumnHeadersHeight * 2;
            this.DVGSalesman.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            this.DVGSalesman.CellPainting += new DataGridViewCellPaintingEventHandler(DVGSalesman_CellPainting);
            this.DVGSalesman.Paint += new PaintEventHandler(DVGSalesman_Paint);
            this.DVGSalesman.Scroll += new ScrollEventHandler(DVGSalesman_Scroll);
            this.DVGSalesman.ColumnWidthChanged += new DataGridViewColumnEventHandler(DVGSalesman_ColumnWidthChanged);

            foreach (DataGridViewColumn col in DVGSalesman.Columns)
            {
                if (col.Index == 4 || col.Index == 9)
                {
                    col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
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
            //__cursor = new LineHighlight(this.panel1);
            //foreach (Control __child in this.panel1.Controls)
            //{
            //    if (__child.TabStop) __cursor.Add(__child);
            //}
            //panel1.Focus();
            getSource();

            panel1.AutoScrollMinSize = new Size(0, 500);

            TabPage t = tabControl1.TabPages[3];
            tabControl1.SelectedTab = t;

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
                __skill.AddRangeTextBox(skillidcodetb);
                __skill.Initialize();

                //if (__isAsk30724)
                //{
                __outlet_status = new TIRAData<TIRAItem>(__state == StateEntry.New ?
                    QUERY_OUTLET_STATUS_NEW :
                    QUERY_OUTLET_STATUS, 0);
                __outlet_status.Execute();

                cbstatus.DisplayMember = "Nama";
                cbstatus.ValueMember = "Kode";
                cbstatus.DataSource = __outlet_status.DATA;
                //cbstatus.SelectedItem = __outlet_status.DATA.FirstOrDefault();
                cbstatus.SelectedItem = null;

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

                if (!__isFlagKAM)
                    tabControl1.TabPages.Remove(topbyprdline_page);

                SetGridViewSchPayD();
                SetGridViewSchPayW();

                SetGridViewPrdLineBlocking();

                /*ZipCode Inisiator agar bisa langsung diisi zipcodenya ketika pilih city */
                popUpProvinsinCity1.Needzipcode(txtPostCodeT4);
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

            tabControl1.Selected += new TabControlEventHandler(tabControl1_Selected);

            foreach (DataGridViewColumn column in DVGSalesman.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            foreach (DataGridViewColumn column in DGVHIR.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            _redForeColorMandatory();

            //untuk combo box pilihan
            LoadDropDown();
            comboBoxPajakT1.Enabled = false;
            comboBoxPajakT1.SelectedIndex = 1;
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

                FillGridBackList();
                FillGridUNBackList();

                if (__isFlagKAM)
                    FillTopByPrdLine();

                FillCustSchPay();
                FillPrdLineBlocking();

                settab2aktif();
                //added 09042019
                DGVHIR.Enabled = false;

                strSQL = "select ISNULL(cm_delivery_days_no,0) cm_delivery_days_no,ISNULL(cm_delivery_day,0) cm_delivery_day,cm_duration_days, ISNULL(cm_sfa_code,'') cm_sfa_code from SO_CUST_MASTER WITH(NOLOCK) where cm_cust_code1 =";
                strSQL += "'" + custcode1 + "' And cm_cust_code2 ='" + custcode2 + "'";

                DataTable dts = new DataTable();

                dts = _clsGlobal.ExecDT(strSQL);
                if (dts.Rows.Count > 0)
                {
                    txtDelivByDays.Text = dts.Rows[0]["cm_delivery_days_no"].ToString().Trim();
                    CBDelivByDay.SelectedValue = dts.Rows[0]["cm_delivery_day"].ToString().Trim();
                    txtDurationOfDays.Text= dts.Rows[0]["cm_duration_days"].ToString().Trim();

                    if (!__isFlagKAM)
                    {
                        txtSFACustCode.Text = dts.Rows[0]["cm_sfa_code"].ToString().Trim();
                    }
                }
                else
                {
                    CBDelivByDay.SelectedIndex = 0;
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

            txtOutSandingT1.ReadOnly = true;
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
                comboBoxPajakT1.DataSource = _clsGlobal.ExecDT(strSQL);
                comboBoxPajakT1.ValueMember = "gh_sequence_no";
                comboBoxPajakT1.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            try
            {
                //Size Group Code
                strSQL = "select ghc_sequence_no, CONCAT(ghc_sequence_no,' ~ ',ghc_function_desc) as ghc_function_desc from GS_HARD_CODED WITH(NOLOCK) where ghc_sys ='H' AND ghc_function_name ='SHIPBILL'";
                CBShipnBill.DataSource = _clsGlobal.ExecDT(strSQL);
                CBShipnBill.ValueMember = "ghc_sequence_no";
                CBShipnBill.DisplayMember = "ghc_function_desc";
                //if (clsGlobal.MODE_TRX == 1)
                if (__state == StateEntry.New)
                {
                    CBShipnBill.SelectedIndex = 2;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                strSQL = "select gh_function_code, CONCAT(gh_function_code,' ~ ',gh_function_desc) as gh_function_desc from GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name = 'CUSTPAYTYPE'";
                CBPembyaranFktT1.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPembyaranFktT1.ValueMember = "gh_function_code";
                CBPembyaranFktT1.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                strSQL = "select gh_function_code, CONCAT(gh_function_code,' ~ ',gh_function_desc) as gh_function_desc from GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_function_name = 'BANGUNAN'";
                CBBagunanT1.DataSource = _clsGlobal.ExecDT(strSQL);
                CBBagunanT1.ValueMember = "gh_function_code";
                CBBagunanT1.DisplayMember = "gh_function_desc";
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
                CBDelivByDay.DataSource = _clsGlobal.ExecDT(strSQL);
                CBDelivByDay.ValueMember = "gh_function_code";
                CBDelivByDay.DisplayMember = "gh_function_desc";
                CBDelivByDay.SelectedIndex = 0;
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
                cbPemilikNIK.DataSource = _clsGlobal.ExecDT(strSQL);
                cbPemilikNIK.ValueMember = "gh_function_code";
                cbPemilikNIK.DisplayMember = "gh_function_desc";
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
                cbJenisIdentitas.DataSource = _clsGlobal.ExecDT(strSQL);
                cbJenisIdentitas.ValueMember = "gh_function_code";
                cbJenisIdentitas.DisplayMember = "gh_function_desc";
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
                    cbAddressChoice.DataSource = _clsGlobal.ExecDT(strSQL);
                    cbAddressChoice.ValueMember = "gh_function_code";
                    cbAddressChoice.DisplayMember = "gh_function_desc";

                    if (__isFlagKAM)
                    {
                        cbAddressChoice.SelectedIndex = 1;
                    }
                    else
                    {
                        cbAddressChoice.SelectedIndex = 0;
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
            int a = CBShipnBill.SelectedIndex;

            IsiTextBill(a, txtCustCode.Text);

        }

        void IsiTextBill_old(int a, string val)
        {
            string strVal = (val == "" ? "0" : val);

            if (a == 0)
            {
                txtShiptoCustCode.Visible = true;
                txtShiptoCustCode.ReadOnly = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");
                txtShiptoCustCode.Text = (__state == StateEntry.New ? "" : "");

                txtShiptoCustCodeTo.Visible = true;
                txtShiptoCustCodeTo.ReadOnly = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");//
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "" : "");//

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); // billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); // billConde1
                txtBilltoCustCode.ReadOnly = true;


                txtBilltoCustCodeTo.ReadOnly = true;
                txtBilltoCustCodeTo.Visible = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2);
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2);

                BTNBillcust.Visible = false;
                txtWHCode.Enabled = false;
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 1)
            {
                txtShiptoCustCode.Visible = true;
                txtShiptoCustCode.ReadOnly = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.ReadOnly = true;

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCode.Text = (__state == StateEntry.New ? "" : ""); //
                txtBilltoCustCode.ReadOnly = true;

                txtBilltoCustCodeTo.ReadOnly = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "" : ""); //
                txtBilltoCustCodeTo.Visible = true;

                BTNBillcust.Visible = true;
                txtWHCode.Enabled = false;
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 2)
            {
                //txtShiptoCustCode.Visible = false;
                txtShiptoCustCode.Visible = true;
                txtShiptoCustCode.ReadOnly = false;
                //txtShiptoCustCodeTo.Visible = false;
                txtShiptoCustCodeTo.Visible = true;
                txtShiptoCustCodeTo.ReadOnly = true;
                //txtBilltoCustCode.Visible = false;
                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCodeTo.Visible = false;
                txtBilltoCustCodeTo.Visible = true;
                txtBilltoCustCode.ReadOnly = true;
                //txtBilltoCustCode.Visible = false;
                txtBilltoCustCode.Visible = true;
                txtBilltoCustCode.ReadOnly = true;
                txtBilltoCustCodeTo.ReadOnly = true;
                //BTNBillcust.Visible = false;
                BTNBillcust.Visible = true;
                txtWHCode.Enabled = false;
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
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
                txtShiptoCustCode.ReadOnly = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);                
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.ReadOnly = true;

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //billConde1
                txtBilltoCustCode.ReadOnly = true;

                txtBilltoCustCodeTo.ReadOnly = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
                txtBilltoCustCodeTo.Visible = true;

                BTNBillcust.Visible = false;
                txtWHCode.Enabled = false;
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
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
                txtShiptoCustCode.ReadOnly = false;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal);

                txtShiptoCustCodeTo.Visible = true;
                txtShiptoCustCodeTo.ReadOnly = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : "");//
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);//


                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); // billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); // billConde1
                txtBilltoCustCode.ReadOnly = true;


                txtBilltoCustCodeTo.ReadOnly = true;
                txtBilltoCustCodeTo.Visible = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2);
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2);

                BTNShipcust.Visible = true;
                BTNShipcust.Enabled = true;
                BTNBillcust.Visible = false;
                BTNBillcust.Enabled = false;
                txtWHCode.Enabled = false;
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 1)
            {
                txtShiptoCustCode.Visible = true;
                txtShiptoCustCode.ReadOnly = false;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.ReadOnly = true;

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCode.Text = (__state == StateEntry.New ? "" : ""); //
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //
                txtBilltoCustCode.ReadOnly = true;

                txtBilltoCustCodeTo.ReadOnly = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "" : ""); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
                txtBilltoCustCodeTo.Visible = true;

                BTNShipcust.Visible = false;
                BTNShipcust.Enabled = false;
                BTNBillcust.Visible = true;
                BTNBillcust.Enabled = true;
                txtWHCode.Enabled = false;
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
                btnWHCode.Enabled = true;
                label16.Visible = true;
                label17.Visible = true;
            }
            else if (a == 2)
            {
                //txtShiptoCustCode.Visible = false;
                txtShiptoCustCode.Visible = true;
                txtShiptoCustCode.ReadOnly = true;
                //txtShiptoCustCodeTo.Visible = false;
                txtShiptoCustCodeTo.Visible = true;
                txtShiptoCustCodeTo.ReadOnly = true;
                //txtBilltoCustCode.Visible = false;
                txtBilltoCustCode.Visible = true;
                txtBilltoCustCode.ReadOnly = true;
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
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
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
                txtShiptoCustCode.ReadOnly = true;
                //txtShiptoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //shipCode1
                txtShiptoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //shipCode1

                txtShiptoCustCodeTo.Visible = true;
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : shipCode2);
                //txtShiptoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : txtCustCodeTo.Text);                
                txtShiptoCustCodeTo.Text = (__state == StateEntry.New ? "000" : txtCustCodeTo.Text);
                txtShiptoCustCodeTo.ReadOnly = true;

                txtBilltoCustCode.Visible = true;
                //txtBilltoCustCode.Text = (clsGlobal.MODE_TRX == 1 ? strVal : strVal); //billConde1
                txtBilltoCustCode.Text = (__state == StateEntry.New ? strVal : strVal); //billConde1
                txtBilltoCustCode.ReadOnly = true;

                txtBilltoCustCodeTo.ReadOnly = true;
                //txtBilltoCustCodeTo.Text = (clsGlobal.MODE_TRX == 1 ? "000" : billConde2); //
                txtBilltoCustCodeTo.Text = (__state == StateEntry.New ? "000" : billConde2); //
                txtBilltoCustCodeTo.Visible = true;

                BTNBillcust.Visible = false;
                txtWHCode.Enabled = false;
                txtWHCode.ReadOnly = false;
                txtWHCodeTo.Enabled = false;
                txtWHCodeTo.ReadOnly = false;
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
                    CBShipnBill.SelectedValue = dt1.Rows[0]["cm_ship_bill_to_flag"].ToString().Trim();

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
                    cbstatus.SelectedItem = status;

                    if (dt1.Rows[0]["cm_source"] != DBNull.Value)
                    {
                        cmbSource.SelectedValue = dt1.Rows[0]["cm_source"].ToString();
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
                    comboBoxPajakT1.SelectedValue = dt1.Rows[0]["cm_notret_pajak"].ToString().Trim();

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
                    CBPembyaranFktT1.SelectedValue = dt1.Rows[0]["cm_payment_type"].ToString().Trim();

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
                    CBBagunanT1.SelectedValue = dt1.Rows[0]["cm_bangunan"].ToString().Trim();
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
                        cbPemilikNIK.SelectedIndex = 0;
                    }
                    else
                    {
                        cbPemilikNIK.SelectedValue = dt1.Rows[0]["cm_pemilik_nik"].ToString().Trim();
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
                    dtOpenHour.Value = Convert.ToDateTime(dt1.Rows[0]["cm_open_hour"].ToString());
                    dtCloseHour.Value = Convert.ToDateTime(dt1.Rows[0]["cm_close_hour"].ToString());
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
                                cbAddressChoice.SelectedIndex = 1;
                            }
                            else
                            {
                                cbAddressChoice.SelectedIndex = 0;
                            }
                        }
                        else
                        {
                            cbAddressChoice.SelectedValue = dt1.Rows[0]["cm_tax_address_choice"].ToString().Trim();
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
                        cbJenisIdentitas.SelectedIndex = 0;
                    }
                    else
                    {
                        cbJenisIdentitas.SelectedValue = dt1.Rows[0]["cm_jenis_identitas"].ToString().Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

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
            frm.Query = " SELECT cm_entity[Entity], cm_branch[Branch], cm_cust_code1[Code1], cm_cust_code2 [Code2],        cm_cust_name[Name], cm_delv_address1[Address]  From SO_CUST_MASTER ";
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


                strSQL = "select distinct cpl_line_code, pl_prd_line_desc, cpl_cust_type, ct_cust_type_desc from TBL_SD_CUSTPRDLINE left join IM_PRD_LINE on cpl_line_code = pl_prd_line_code left join SO_CUST_TYPE on cpl_cust_type = ct_cust_type and ct_cluster='Y' where cpl_cust_code1 = '" + _prdCustCode1 + "'and cpl_cust_code2 = '" + _Grade + "' and cpl_entity='" + _prdEntityCode + "' and cpl_branch='" + _prdBrandCode + "'";

                DataTable dt = _clsGlobal.ExecDT(strSQL);
                DGVHIR.DataSource = dt;


                if (DGVHIR.Rows.Count > 0)
                {
                    //DataGridViewCellEventArgs DataGridViewCellEventArgs = new DataGridViewCellEventArgs(DGVHIR.Columns["cpl_line_code"].Index, DGVHIR.CurrentRow.Index);
                    //this.DGVHIR_CellClick(DGVHIR.CurrentRow.Index, DataGridViewCellEventArgs);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DGVHIR_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(DGVHIR.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void DGVHIR_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            if (e.RowIndex > -1)
            {
                if (senderGrid.CurrentCell.ColumnIndex == DGVHIR.Columns["btna"].Index)
                {
                    frmPopUp frm = new frmPopUp();
                    frm.FrmText = "Search Prdline ";
                    frm.Query = "select pl_prd_line_code [PrdLine], pl_prd_line_desc [Description] from IM_PRD_LINE ";
                    frm.ShowDialog();
                    if (frm.ArrField != null)
                    {
                        try
                        {
                            DataGridViewRow dgr, dgrB;
                            string maxB = "";
                            string min = "";
                            dgrB = null;

                            dgr = DGVHIR.Rows[DGVHIR.CurrentRow.Index];

                            if (DGVHIR.RowCount > 1)
                            {
                                if (dgr.Index != 0)
                                {
                                    dgrB = DGVHIR.Rows[DGVHIR.CurrentRow.Index - 1];
                                    maxB = dgrB.Cells["cpl_line_code"].Value.ToString();
                                }
                                else
                                {
                                    min = frm.ArrField[0].Trim();
                                    rowDuplicate = false;
                                    exitData = false;

                                    foreach (DataGridViewRow gRow in DGVHIR.Rows)
                                    {
                                        if ((gRow.Cells["cpl_line_code"].Value.ToString().Trim() == min))
                                        {
                                            rowDuplicate = true;
                                            break;
                                        }
                                    }
                                    if (rowDuplicate == true)
                                    {
                                        MessageBox.Show("Product Line " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Product Line yang berbeda ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                        DGVHIR.Rows[e.RowIndex].Cells["cpl_line_code"].Value = frm.ArrField[0].Trim();
                                        DGVHIR.Rows[e.RowIndex].Cells["pl_prd_line_desc"].Value = frm.ArrField[1].Trim();
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

                                foreach (DataGridViewRow gRow in DGVHIR.Rows)
                                {
                                    if ((gRow.Cells["cpl_line_code"].Value.ToString().Trim() == min))
                                    {
                                        rowDuplicate = true;
                                        break;
                                    }
                                }

                                if (rowDuplicate == true)
                                {
                                    MessageBox.Show("Product Line " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Product Line yang berbeda ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                    DGVHIR.Rows[e.RowIndex].Cells["cpl_line_code"].Value = frm.ArrField[0].Trim();
                                    DGVHIR.Rows[e.RowIndex].Cells["pl_prd_line_desc"].Value = frm.ArrField[1].Trim();
                                }
                            }
                            else
                            {
                                if (exitData != true)
                                {
                                    DGVHIR.Rows[e.RowIndex].Cells["cpl_line_code"].Value = frm.ArrField[0].Trim();
                                    DGVHIR.Rows[e.RowIndex].Cells["pl_prd_line_desc"].Value = frm.ArrField[1].Trim();
                                }
                            }

                        }
                        catch (Exception ex) { }
                    }
                }
                if (senderGrid.CurrentCell.ColumnIndex == DGVHIR.Columns["btnb"].Index)
                {

                    frmPopUp frm = new frmPopUp();
                    frm.FrmText = "Search Cust ";
                    //frm.Query = "select ct_cust_type [Type] , ct_cust_type_desc [Description] from SO_CUST_TYPE ";
                    frm.Query = "select ct_cust_type [Type] , ct_cust_type_desc [Description] from SO_CUST_TYPE where ct_cluster='Y'";
                    frm.ShowDialog();
                    if (frm.ArrField != null)
                    {
                        DGVHIR.Rows[e.RowIndex].Cells["cpl_cust_type"].Value = frm.ArrField[0].Trim();
                        DGVHIR.Rows[e.RowIndex].Cells["ct_cust_type_desc"].Value = frm.ArrField[1].Trim();

                    }
                }

            }

        }

        private void DGVHIR_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Insert)
            {
                DataGridViewRow dgr;
                string column0, column1, column2;
                try
                {
                    dgr = DGVHIR.CurrentRow;

                    if (DGVHIR.Rows.Count > 0)
                    {

                        column0 = dgr.Cells["cpl_line_code"].Value.ToString();
                        column1 = dgr.Cells["cpl_cust_type"].Value.ToString();
                        column2 = dgr.Cells["ct_cust_type_desc"].Value.ToString();
                        if (column0 != "" && dgr.Cells["pl_prd_line_desc"].Value.ToString() != "" && column1 != "" && column2 != "")
                        {
                            //if (clsGlobal.MODE_TRX == 1)
                            if (__state == StateEntry.New)
                            {
                                DGVHIR.Rows.Add("", "", "", "", "", "");
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
                            DGVHIR.Rows.Add("", "", "", "", "", "");
                        }
                        else
                        {
                            if (DGVHIR.RowCount > 0)
                            {
                                DGVHIR.Rows.Add("", "", "", "", "", "");
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
                if (DGVHIR.RowCount != 0)
                {
                    DGVHIR.Rows.Remove(DGVHIR.CurrentRow);
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



                DVGSalesman.DataSource = dt;
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
                strSQL = "";
                strSQL = "select cp_price_group as grp_code,ppl_price_list_description grp_desc from SO_CUST_MASTER_PRICE_GROUP INNER JOIN OB_PRICE_LIST_HEADER ON cp_price_group=ppl_price_list_code " +
                        "where cp_cust_code1 = " + FmtStr(txtCustCode.Text.Trim()) + "" +
                        "and cp_cust_code2 = " + FmtStr(txtCustCodeTo.Text.Trim()) + "" +
                        "and cp_entity = " + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + "" +
                        "and cp_branch = " + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + "";
                DataTable dt = new DataTable();

                //dt.Columns.Add("grp_code");
                //dt.Columns.Add("grp_desc");
                //dt.AcceptChanges();

                dt = _clsGlobal.ExecDT(strSQL);

                dgvGroupHarga.DataSource = dt;

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
            __query.AppendLine(",IPL.pl_prd_line_desc AS prdline_desc");
            __query.AppendLine(",MPL.smt_top_id AS top_code");
            __query.AppendLine(",PTC.pptc_term_desc AS top_desc");
            __query.AppendLine("FROM SO_MAPPING_TOPBYPRDLINE MPL WITH(NOLOCK)");
            __query.AppendLine("INNER JOIN IM_PRD_LINE IPL WITH(NOLOCK)");
            __query.AppendLine("ON IPL.pl_prd_line_code=MPL.smt_prdline_id");
            __query.AppendLine("INNER JOIN PO_PAYMENT_TERM_CODES PTC WITH(NOLOCK)");
            __query.AppendLine("ON PTC.pptc_term_code=MPL.smt_top_id");
            __query.AppendFormat("WHERE smt_entity_id='{0}'\r\n", ctrlEntityCustMaster2.txtCM.Text.Trim());
            __query.AppendFormat("AND smt_branch_id='{0}'\r\n", ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim());
            __query.AppendFormat("AND smt_cust_code1='{0}'\r\n", txtCustCode.Text.Trim());
            __query.AppendFormat("AND smt_cust_code2='{0}'\r\n", txtCustCodeTo.Text.Trim());

            //DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
            DataTable __source = _clsGlobal.ExecDT(__query.ToString());
            topbyprdlinedgv.DataSource = __source;
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

            if (dgvschpayW.RowCount > 0)
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
            try
            {
                //----------------------------------- SALESMAN ------------------------------------- 
                Rectangle r24 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["csc_salesman_id"].Index, -1, true);
                int r25 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["btnc"].Index, -1, true).Width;
                int r26 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["sgm_spgm_name"].Index, -1, true).Width;
                int r27 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["sgm_type_operasi"].Index, -1, true).Width;

                r24.X += 1;
                r24.Y += 1;
                r24.Width += (r25 + r26 + r27) - 2;
                r24.Height = r24.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(DVGSalesman.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r24);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.DVGSalesman.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Salesman", DVGSalesman.ColumnHeadersDefaultCellStyle.Font, br, r24, sf);
                }

                //----------------------------------- Hari ------------------------------------- 
                Rectangle r1 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["csc_visit_week1"].Index, -1, true);
                int r2 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["csc_visit_week2"].Index, -1, true).Width;
                int r3 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["csc_visit_week3"].Index, -1, true).Width;
                int r4 = DVGSalesman.GetCellDisplayRectangle(DVGSalesman.Columns["csc_visit_week4"].Index, -1, true).Width;

                r1.X += 1;
                r1.Y += 1;
                r1.Width += (r2 + r3 + r4) - 2;
                r1.Height = r1.Height / 2 - 2;

                //draw box
                using (SolidBrush br = new SolidBrush(DVGSalesman.ColumnHeadersDefaultCellStyle.BackColor))
                {
                    e.Graphics.FillRectangle(br, r1);
                }

                //draw text
                using (SolidBrush br = new SolidBrush(this.DVGSalesman.ColumnHeadersDefaultCellStyle.ForeColor))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Pola", DVGSalesman.ColumnHeadersDefaultCellStyle.Font, br, r1, sf);
                }
            }
            finally
            {
            }
        }

        private void DVGSalesman_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(DVGSalesman.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void DVGSalesman_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var senderGrid = (DataGridView)sender;
                if (e.RowIndex > -1)
                {
                    if (senderGrid.CurrentCell.ColumnIndex == DVGSalesman.Columns["btnc"].Index)
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
                                DataGridViewRow dgr, dgrB;
                                string maxB = "";
                                string min = "";
                                dgrB = null;
                                dgr = null;

                                dgr = DVGSalesman.Rows[DVGSalesman.CurrentRow.Index];

                                if (DVGSalesman.RowCount > 1)
                                {
                                    if (dgr.Index != 0)
                                    {
                                        dgrB = DVGSalesman.Rows[DVGSalesman.CurrentRow.Index - 1];
                                        maxB = dgrB.Cells["csc_salesman_id"].Value.ToString();
                                    }
                                    else
                                    {
                                        min = frm.ArrField[0].Trim();
                                        rowDuplicate = false;
                                        exitData = false;

                                        foreach (DataGridViewRow gRow in DVGSalesman.Rows)
                                        {
                                            if ((gRow.Cells["csc_salesman_id"].Value.ToString().Trim() == min))
                                            {
                                                rowDuplicate = true;
                                                break;
                                            }
                                        }

                                        if (rowDuplicate == true)
                                        {
                                            MessageBox.Show("Salesman Sudah ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK);
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
                                            DVGSalesman.Rows[e.RowIndex].Cells["csc_salesman_id"].Value = frm.ArrField[0].Trim();
                                            DVGSalesman.Rows[e.RowIndex].Cells["sgm_spgm_name"].Value = frm.ArrField[1].Trim();
                                            DVGSalesman.Rows[e.RowIndex].Cells["sgm_type_operasi"].Value = frm.ArrField[2].Trim();
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

                                    foreach (DataGridViewRow gRow in DVGSalesman.Rows)
                                    {
                                        if ((gRow.Cells["csc_salesman_id"].Value.ToString().Trim() == min))
                                        {
                                            rowDuplicate = true;
                                            break;
                                        }
                                    }

                                    if (rowDuplicate == true)
                                    {
                                        MessageBox.Show("Salesman Sudah ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK);
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
                                        DVGSalesman.Rows[e.RowIndex].Cells["csc_salesman_id"].Value = frm.ArrField[0].Trim();
                                        DVGSalesman.Rows[e.RowIndex].Cells["sgm_spgm_name"].Value = frm.ArrField[1].Trim();
                                        DVGSalesman.Rows[e.RowIndex].Cells["sgm_type_operasi"].Value = frm.ArrField[2].Trim();
                                    }
                                }
                                else
                                {
                                    if (exitData != true)
                                    {
                                        DVGSalesman.Rows[e.RowIndex].Cells["csc_salesman_id"].Value = frm.ArrField[0].Trim();
                                        DVGSalesman.Rows[e.RowIndex].Cells["sgm_spgm_name"].Value = frm.ArrField[1].Trim();
                                        DVGSalesman.Rows[e.RowIndex].Cells["sgm_type_operasi"].Value = frm.ArrField[2].Trim();
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
                DataGridViewRow dgr;
                string column0, column1, column2;
                try
                {
                    dgr = DVGSalesman.CurrentRow;

                    if (DVGSalesman.Rows.Count > 0)
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
                if (DVGSalesman.RowCount != 0)
                {
                    DVGSalesman.Rows.Remove(DVGSalesman.CurrentRow);
                }
            }
        }

        private string FmtStr(string value_str)
        {
            string formatStr;
            formatStr = "'" + value_str.Trim().Replace("'", "''") + "'";

            return formatStr;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (txtDelivByDays.Text.Trim() == "0")
            {
                txtDelivByDays.Text = "";
            }
            if ((txtDelivByDays.Text.Trim() == "" && Convert.ToInt64(CBDelivByDay.SelectedValue) == 0))
            {
                if (txtDelivByDays.Text.Trim() == "")
                {
                    txtDelivByDays.Text = "0";
                }
                MessageBox.Show("Harus mmengisi delivery by day atau delivery by days", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //txtDelivByDays.Select();
                return;
            }

            else if (txtDelivByDays.Text.Trim() != "" && Convert.ToInt64(CBDelivByDay.SelectedValue) != 0)
            {
                if (txtDelivByDays.Text.Trim() == "")
                {
                    txtDelivByDays.Text = "0";
                }
                MessageBox.Show("Harus memilih salah satu delivery by day atau delivery by days", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //txtDelivByDays.Select();
                return;
            }

            if (txtCustCode.Text.Trim() == "")
            {
                MessageBox.Show(" Code Customer belum di-input ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCustCode.Select();
                return;
            }
            if (this.State == StateEntry.New && txtCustCode.Text.Trim().Length != 6)
            {
                MessageBox.Show(" Customer Code harus 6 digit/karakter ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCustCode.Select();
                return;
            }
            if (txtCustCodeTo.Text.Trim() == "")
            {
                MessageBox.Show(" Customer Code To Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCustCodeTo.Select();
                return;
            }
            if (txtCustName.Text.Trim() == "")
            {
                MessageBox.Show(" Nama Customer belum di-input ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCustName.Select();
                return;
            }
            if (txtShortName.Text.Trim() == "")
            {
                MessageBox.Show(" Short Name Code  Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtShortName.Select();
                return;
            }
            //adien
            //if (ctrlEntityCustMaster1.txtCM.Text.Trim() == "")
            if (ctrlEntityCustMaster2.txtCM.Text.Trim() == "")
            {
                MessageBox.Show(" Entity belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //ctrlEntityCustMaster1.txtCM.Select();
                ctrlEntityCustMaster2.txtCM.Select();
                return;
            }

            //if (ctrlBranchCustMaster1.txtBranchIdCM.Text.Trim() == "")
            if (ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim() == "")
            {
                MessageBox.Show(" Branch belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //ctrlBranchCustMaster1.txtBranchIdCM.Select();
                ctrlEntityCustMaster2.txtBranchIdCM.Select();
                return;
            }
            //adien
            if (ctrlDistrikCustMaster1.txtDistrikCM.Text.Trim() == "")
            {
                MessageBox.Show(" Distrik belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ctrlDistrikCustMaster1.txtDistrikCM.Select();
                return;
            }
            if (ctrlDistrikCustMaster1.txtBeatCM.Text.Trim() == "")
            {
                MessageBox.Show(" Beat belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ctrlDistrikCustMaster1.txtBeatCM.Select();
                return;
            }
            if (ctrlDistrikCustMaster1.txtSubBeatCM.Text.Trim() == "")
            {
                MessageBox.Show(" SubBeat belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ctrlDistrikCustMaster1.txtSubBeatCM.Select();
                return;
            }
            if (ctrlGroupDiscount1.txtGpDiscountT1.Text.Trim() == "")
            {
                MessageBox.Show(" Group Discount belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ctrlGroupDiscount1.txtGpDiscountT1.Select();
                return;
            }
            if (txtIndustriT1.Text.Trim() == "")
            {
                MessageBox.Show(" Industri belum di-pilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtIndustriT1.Select();
                return;
            }
            if (ctrlKodepasarDesc.txtKodePasarT1.Text.Trim() == "")
            {
                MessageBox.Show(" Pasar belum di-pilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ctrlKodepasarDesc.txtKodePasarT1.Select();
                return;
            }
            if (ctrlGroupOutlet1.txtGroupCM.Text.Trim() == "")
            {
                MessageBox.Show(" Group Outlet belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ctrlGroupOutlet1.txtGroupCM.Select();
                return;
            }
            if (ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Text.Trim() == "")
            {
                MessageBox.Show(" Tipe Outlet belum di-pilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ctrlpopUpTypeOutletCustMaster1.txtTypeOutlet.Select();
                return;
            }
            if (txtLeadtime.Text.Trim() == "")
            {
                MessageBox.Show(" Lead Time harus diisi ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtLeadtime.Select();
                return;
            }
            if (popUpProvinsinCity1.txtPropinsiCMC.Text.Trim() == "")
            {
                MessageBox.Show(" Provinsi belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                popUpProvinsinCity1.txtPropinsiCMC.Select();
                return;
            }
            if (popUpProvinsinCity1.txtKabupatenCMC.Text.Trim() == "")
            {
                MessageBox.Show(" City belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                popUpProvinsinCity1.txtKabupatenCMC.Select();
                return;
            }
            if (popUpProvinsinCity1.txtKecamatan.Text.Trim() == "")
            {
                MessageBox.Show(" Kecamatan belum di-pilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                popUpProvinsinCity1.txtKecamatan.Select();
                return;
            }
            if (popUpProvinsinCity1.txtKelurahan.Text.Trim() == "")
            {
                MessageBox.Show(" Kelurahan belum di-pilih ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                popUpProvinsinCity1.txtKelurahan.Select();
                return;
            }
            if (txtBilltoCustCode.Text.Trim() == "")
            {
                MessageBox.Show(" Bill to Cust Code belum diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtBilltoCustCode.Select();
                return;
            }
            if (txtAddress1T2.Text.Trim() == "")
            {
                MessageBox.Show(" Address 1 cust property belum diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAddress1T2.Select();
                return;
            }
            if (txtAddress1T2.Text.Length < 20)
            {
                MessageBox.Show(" Address 1 cust property diisi minimal 20 karakter ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAddress1T2.Select();
                return;
            }
            if (txtAddress1T3.Text.Trim() == "")
            {
                MessageBox.Show(" Address 1 billing address belum diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAddress1T3.Select();
                return;
            }

            if (txtAddress1T3.Text.Length < 20)
            {
                MessageBox.Show(" Address 1 billing address diisi minimal 20 karakter ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAddress1T3.Select();
                return;
            }
            if (txtAddress1T4.Text.Trim() == "")
            {
                MessageBox.Show(" Address 1 delivery address belum diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAddress1T4.Select();
                return;
            }
            if (txtAddress1T4.Text.Length < 20)
            {
                MessageBox.Show(" Address 1 delivery address diisi minimal 20 karakter ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAddress1T4.Select();
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
                    MessageBox.Show(" Telepon Harus berupa angka dan tidak boleh ada spasi atau karakter lain! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtTelephoneT4.Select();
                    return;
                }
                if (txtTelephoneT4.Text.Length < 8)
                {
                    MessageBox.Show(" Telepon Delivery Harus Minimal 8 Digit! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtTelephoneT4.Select();
                    return;
                }
            }

            /*end*/

            /* Validasi Monitoring Insurance Yuda 02.03.2022 */
            if (checkBoxMonitoringInsuranceT1.Checked)
            {
                if (!dtTglToIns.Text.Equals(dtTglFromIns.Text))
                {
                    if (Convert.ToDateTime(dtTglToIns.Value) < Convert.ToDateTime(dtTglFromIns.Value))
                    {
                        dtTglToIns.Select();
                        MessageBox.Show(" Tanggal Insurance From harus lebih kecil dari pada Tanggal Insurance To. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

            }
            /* End Validasi Monitoring Insurance */

            /* Validasi LNG & LAT */
            if (string.IsNullOrEmpty(txtLong.Text) || string.IsNullOrEmpty(txtLat.Text))
            {
                MessageBox.Show("Lat & Lng harus terisi. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (string.IsNullOrEmpty(txtLong.Text))
                {
                    txtLong.Select();
                }
                else
                {
                    txtLat.Select();
                }
                return;
            }

            if (!string.IsNullOrEmpty(txtLong.Text))
            {
                string[] split = txtLong.Text.Split('.');
                if (split.Length < 2)
                {
                    MessageBox.Show("mohon isi lng sesuai format. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtLong.Select();
                    return;
                }
                else
                {
                    if (split[1].Length < 5)
                    {
                        MessageBox.Show("mohon isi lng sesuai format. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtLong.Select();
                        return;
                    }
                }

                if (!is_valid_coordinate(txtLong.Text))
                {
                    MessageBox.Show("Silahkan isi data Geotag dengan benar!. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtLong.Select();
                    return;
                }
                decimal dLat = Convert.ToDecimal(txtLong.Text);
                if (dLat < -180 || dLat > 180)
                {
                    MessageBox.Show("Longitude tidak boleh lebih besar 180 dan tidak boleh lebih kecil dari -180 !. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtLong.Select();
                    return;
                }
            }

            if (!string.IsNullOrEmpty(txtLat.Text))
            {
                string[] split = txtLat.Text.Split('.');
                if (split.Length < 2)
                {
                    MessageBox.Show("mohon isi lat sesuai format. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtLat.Select();
                    return;
                }
                else
                {
                    if (split[1].Length < 5)
                    {
                        MessageBox.Show("mohon isi lat sesuai format. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtLat.Select();
                        return;
                    }
                }
                if (!is_valid_coordinate(txtLat.Text))
                {
                    MessageBox.Show("Silahkan isi data Geotag dengan benar!. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtLat.Select();
                    return;
                }
                decimal dLat = Convert.ToDecimal(txtLat.Text);
                if (dLat < -90 || dLat > 90)
                {
                    MessageBox.Show("Latitude tidak boleh lebih besar 90 dan tidak boleh lebih kecil dari -90 !. ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtLat.Select();
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
                MessageBox.Show("Silahkan isi data Alamat Billing dengan benar!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.txtAddress1T3.Focus();
                return;
            }

            if (!is_valid_Address(__delvaddress))
            {
                MessageBox.Show("Silahkan isi data Alamat Delivery dengan benar!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.txtAddress1T4.Focus();
                return;
            }

            //jenis identitas

            if (string.IsNullOrEmpty(cbJenisIdentitas.SelectedValue.ToString().Trim()))
            {
                MessageBox.Show(" Jenis Identitas Mohon Dipilih !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cbJenisIdentitas.Select();
                return;
            }

            if (cbJenisIdentitas.SelectedValue.ToString().Trim() == "1")
            {
                if (string.IsNullOrEmpty(this.txtNpwpT2Masked.Text.Trim()))
                {
                    MessageBox.Show("Jenis Identitas : (BADAN) : NPWP harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }
            else if (cbJenisIdentitas.SelectedValue.ToString().Trim() == "2")
            {
                if (string.IsNullOrEmpty(this.txtNIK.Text.Trim()))
                {
                    MessageBox.Show("Jenis Identitas : (PRIBADI) : NIK harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }


            /* END OF ASK 27780 */
            if (checkBoxNpwpT1.Checked)
            {
                if (txtTaxNameT2.Text.Trim() == "")
                {
                    MessageBox.Show(" Tax Name harus diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtTaxNameT2.Select();
                    return;
                }
                if (txtAddress1T2.Text.Trim() == "")
                {
                    MessageBox.Show(" Tax Address 1 harus diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAddress1T2.Select();
                    return;
                }
                if (txtAddress2T2.Text.Trim() == "")
                {
                    MessageBox.Show(" Tax Address 2 harus diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtAddress2T2.Select();
                    return;
                }
                if (txtCountryT2.Text.Trim() == "")
                {
                    MessageBox.Show(" Tax Country harus diisi! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtCountryT2.Select();
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
                if (!this.txtNpwpT2Masked.MaskCompleted)
                {
                    if (this.txtNpwpT2Masked.Enabled)
                    {
                        MessageBox.Show("Pengisian NPWP tidak komplit!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                if (__npwp.IsNullOrEmptyOrWhiteSpace() || __sppkp.IsNullOrEmptyOrWhiteSpace())
                {
                    MessageBox.Show("NPWP & SPPKP harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    if (!__sppkp.IsNullOrEmptyOrWhiteSpace())
                    {
                        if (__sppkp.Length < 20)
                        {
                            MessageBox.Show("sppkp harus minimal 20 karakter!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                    }
                    string __nik = this.txtNIK.Text.Trim();
                    if (!__nik.IsNullOrEmptyOrWhiteSpace())
                    {
                        if (__nik.Length < 16)
                        {
                            MessageBox.Show("Panjang minimal NIK tidak boleh kurang dari 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            this.txtNIK.Focus();
                            return;
                        }
                        else if (!is_valid_NIK(__nik))
                        {
                            MessageBox.Show("NIK tidak sesuai, Mohon lakukan pengisian dengan benar!!!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            this.txtNIK.Focus();
                            return;
                        }
                    }

                }

                if (__npwp.Length < 16)
                {
                    MessageBox.Show("Panjang NPWP minimal tidak boleh kurang dari 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //this.txtNpwpT2.Focus();
                    this.txtNpwpT2Masked.Focus();
                    return;
                }
                else if (__npwp.Length > 16)
                {
                    MessageBox.Show("Panjang NPWP maksimal 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //this.txtNpwpT2.Focus();
                    this.txtNpwpT2Masked.Focus();
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
                    MessageBox.Show("Silahkan lengkapi Nomor NPWP atau NIK", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.txtNIK.Focus();
                    return;
                }
                else if (!__nik.IsNullOrEmptyOrWhiteSpace() && __npwp.IsNullOrEmptyOrWhiteSpace())
                {
                    if (__nik.Length < 16)
                    {
                        MessageBox.Show("Panjang minimal NIK tidak boleh kurang dari 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.txtNIK.Focus();
                        return;
                    }
                    else if (__nik.Length > 16)
                    {
                        MessageBox.Show("Panjang maksimal NIK 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.txtNIK.Focus();
                        return;
                    }
                    /* 
                    * nik validasi yuda 08032021
                    */
                    else if (!is_valid_NIK(__nik))
                    {
                        MessageBox.Show("NIK tidak sesuai, Mohon lakukan pengisian dengan benar!!!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.txtNIK.Focus();
                        return;
                    }
                    /* END */
                    else
                    {
                        decimal __lNik = 0;
                        decimal.TryParse(__nik, out __lNik);
                        if (__lNik == 0)
                        {
                            MessageBox.Show("Silahkan isi nomor NIK dengan benar", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            this.txtNIK.Focus();
                            return;
                        }
                    }

                }
                else if (__nik.IsNullOrEmptyOrWhiteSpace() && !__npwp.IsNullOrEmptyOrWhiteSpace())
                {
                    if (__npwp.Length < 16)
                    {
                        MessageBox.Show("Panjang NPWP minimal tidak boleh kurang dari 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.txtNpwpT2.Focus();
                        return;
                    }
                    else if (__npwp.Length > 16)
                    {
                        MessageBox.Show("Panjang NPWP maksimal 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.txtNpwpT2.Focus();
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
                            MessageBox.Show("Panjang minimal NIK tidak boleh kurang dari 16", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            this.txtNIK.Focus();
                            return;
                        }
                        else if (!is_valid_NIK(__nik))
                        {
                            MessageBox.Show("NIK tidak sesuai, Mohon lakukan pengisian dengan benar!!!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            this.txtNIK.Focus();
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
                MessageBox.Show(" Flag BackList telah berubah \n Reason BackList harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtReason.Select();
                return;
            }
            //}
            // Update Yuda 06092021 //
            if (!string.IsNullOrEmpty(txtNIK.Text.Trim()) && string.IsNullOrEmpty(cbPemilikNIK.SelectedValue.ToString().Trim()))
            {
                MessageBox.Show(" Pemilik NIK Mohon Dipilih !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cbPemilikNIK.Select();
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
                MessageBox.Show("Post Code Billing Tidak boleh kosong!!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtPostCodeT3.Select();
                return;
            }
            if (string.IsNullOrEmpty(txtPostCodeT4.Text))
            {
                MessageBox.Show("Post Code Delivery Tidak boleh kosong!!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtPostCodeT4.Select();
                return;
            }
            if (string.IsNullOrEmpty(txtCustNameT3.Text))
            {
                MessageBox.Show("Billing Name Tidak boleh kosong!!! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtCustNameT3.Select();
                return;
            }
            /** end of Update 18082022 **/

            if (DVGSalesman.RowCount > 0)
            {
                DataGridViewRow dgr;
                try
                {
                    dgr = DVGSalesman.CurrentRow;
                    if (dgr != null)
                    {
                        bool rowDuplicate = false;
                        bool rowDuplicate2 = false;
                        bool rowDuplicate3 = false;
                        bool rowDuplicate4 = false;
                        foreach (DataGridViewRow gRow in DVGSalesman.Rows)
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
                            MessageBox.Show("Salesman belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate2 == true)
                        {
                            MessageBox.Show("Rute Belum diinput", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate3 == true)
                        {
                            MessageBox.Show("Hari Kunjungan belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate4 == true)
                        {
                            MessageBox.Show("Pola kunjungan belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                finally
                {
                    dgr = null;
                }
            }

            if (DGVHIR.RowCount > 0)
            {
                DataGridViewRow dgr;
                try
                {
                    dgr = DGVHIR.CurrentRow;

                    if (dgr != null)
                    {
                        bool rowDuplicate11 = false;
                        bool rowDuplicate2 = false;
                        foreach (DataGridViewRow gRow in DGVHIR.Rows)
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
                            MessageBox.Show("Produk Line belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (rowDuplicate2 == true)
                        {
                            MessageBox.Show("Type Outlet belum dipilih", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("Data Cluster kosong, silahkan isi terlebih dahulu sebelum disimpan data", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // checking bangunan
            if (CBBagunanT1.SelectedValue == null)
            {
                MessageBox.Show("Jenis bangunan belum diisi, silahkan isi terlebih dahulu sebelum disimpan data", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!isValidGroupPrice())
            {
                MessageBox.Show("Data Group Price STD (standart) harus ada, silahkan isi terlebih dahulu sebelum simpan data", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (txtShipArea.Text.Trim() == string.Empty)
            {
                MessageBox.Show(" Shipment Area harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtShipArea.Select();
                return;
            }

            if (__isFlagDC && skillidcodetb.Text.IsNullOrEmptyOrWhiteSpace())
            {
                MessageBox.Show(" Skill ID harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                skillidcodetb.Select();
                return;
            }

            if (__isFlagDC && txtTelephoneT4.Text.IsNullOrEmptyOrWhiteSpace())
            {
                MessageBox.Show(" Telephone Delivery Address harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtTelephoneT4.Select();
                return;
            }

            if (CBPembyaranFktT1.SelectedValue == null)
            {
                MessageBox.Show("Pembayaran Faktur belum diisi, silahkan isi terlebih dahulu sebelum disimpan data", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // PENGECEKAN NIK JIKA BERISI 1111 1111 1111 1111
            // MAKA PEMBAYARANG AUTO TUNAI
            // BASED ASK : 11060
            __isAutoTunai = false;
            string __znik = this.txtNIK.Text.Trim();
            string __ztypepemb = this.CBPembyaranFktT1.SelectedValue.ToString();

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
                            MessageBox.Show("Hirarky Pemerintahaan tidak sesuai, Pastikan Hirarky pemerintahaan benar.", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(string.Format("Error Pengecekan Hirarki: {0}", exc.Message), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            var stts = (TIRAItem)cbstatus.SelectedItem;
            if (stts == null)
            {
                MessageBox.Show(" Status harus diisi!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cbstatus.Select();
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
                MessageBox.Show("Kode Cust ini sudah terdapat dalam Data Base", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    string mulaiush = _clsGlobal.DateToDB(dateTimePicker1.Value.ToString("MM/dd/yyyy"));
                    string birthday = _clsGlobal.DateToDB(dateTimePicker2.Value.ToString("MM/dd/yyyy"));
                    string insfrom = _clsGlobal.DateToDB(dtTglFromIns.Value.ToString("MM/dd/yyyy"));
                    string insto = _clsGlobal.DateToDB(dtTglToIns.Value.ToString("MM/dd/yyyy"));
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
                        CBPembyaranFktT1.SelectedValue = "T";

                    string skillidcode = skillidcodetb.Text;

                    var stts = ((TIRAItem)cbstatus.SelectedItem) ?? new TIRAItem(__kode: "O", __nama: "OPEN NEW OUTLET");

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
                        "" + FmtStr(CBShipnBill.SelectedValue.ToString()) + ", " +
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
                        "" + FmtStr(CBPembyaranFktT1.SelectedValue.ToString()) + ", " +   // cm_payment_type                                                  
                        "" + FmtStr(cb3.Trim()) + ", " +
                        //"'" + _clsGlobal.DateToDB(dateTimePicker1.Text) + "', " + //cm_outlet_open
                        "'" + mulaiush + "', " + //cm_outlet_open
                        "" + FmtStr(CBBagunanT1.SelectedValue.ToString()) + ", " +
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
                        "" + ((comboBoxPajakT1.SelectedIndex.ToString() != "") ? "'" + comboBoxPajakT1.SelectedIndex.ToString() + "'" : "0") + ", " +
                        "" + FmtStr(cb7.Trim()) + ", " + // cm_top_by_cust
                        "" + FmtStr(txtMoidCode.Text.Trim()) + ", " +//100
                        "" + FmtStr(txtMerchanid.Text.Trim()) + ", " +//cm_merchant_id 

                        //"" + FmtStr("O") + ", " + // @SPcm_active_flag
                        "" + FmtStr(stts.Kode) + ", " + // @SPcm_active_flag

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
                        "" + CBDelivByDay.SelectedValue + ", " +
                        "" + FmtStr(skillidcode) + ", " +
                        "" + FmtStr(cbPemilikNIK.SelectedValue.ToString()) + ", " +
                        "" + FmtStr(txtBillBranch.Text.ToString()) + ", " +
                        "" + FmtStr(txtSPPKP.Text.ToString().Trim()) + ", " +
                        "" + FmtStr(cb11.Trim()) + ", " +
                        "'" + insfrom + "', " +
                        "'" + insto + "', " +
                        "'" + dtOpenHour.Value.ToString("HH:mm:ss") + "', " +
                        "'" + dtCloseHour.Value.ToString("HH:mm:ss") + "', " +
                        "'" + nmPriority.Value.ToString() + "', " +
                        "'" + nmUnloadingTime.Value.ToString() + "'";
                    if (__isAsk27918)
                    {
                        strSQL += "," + FmtStr(cbAddressChoice.SelectedValue.ToString()) + " ";
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
                    strSQL += "'" + getHottTaxCodes() +"', ";
                    if (cmbSource.SelectedValue.ToString() != null || cmbSource.SelectedValue.ToString() != "")
                    {
                        strSQL += "'" + cmbSource.SelectedValue.ToString() + "', ";  //cm_source
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
                    strSQL = strSQL + "'" + cbJenisIdentitas.SelectedValue.ToString() + "',"; //cm_jenis_identitas
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
                    strSQL = " select * from TBL_SD_CUSTPRDLINE WITH (NOLOCK) where cpl_cust_code1  = '" + txtCustCode.Text + "'AND cpl_cust_code2  = '" + txtCustCodeTo.Text + "' AND cpl_entity='" + ctrlEntityCustMaster2.txtCM.Text + "' and cpl_branch = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "'";
                    DataTable dtbl = _clsGlobal.ExecDT(strSQL);

                    if (dtbl.Rows.Count > 0)
                    {
                        //delete all
                        try
                        {
                            strSQL = "EXEC SP_AR_CUST_MASTER_CUSTPRDLINE '3', " +
                              "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text) + ", " +
                              "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text) + ", " +
                              "" + FmtStr(txtCustCode.Text) + ", " +
                              "" + FmtStr(txtCustCodeTo.Text) + ", " +
                              "  null, null, null " + " ";

                            _clsGlobal.BeginTrans();
                            _clsGlobal.ExecuteTrans(strSQL);
                            _clsGlobal.CommitTrans();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        int i;
                        _clsGlobal.BeginTrans();

                        for (i = 0; i < DGVHIR.Rows.Count; i++)
                        {

                            strSQL = "";
                            strSQL = "EXEC SP_AR_CUST_MASTER_CUSTPRDLINE '1', " +
                                "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text) + ", " +
                                "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text) + ", " +
                                "" + FmtStr(txtCustCode.Text.Trim()) + ", " +
                                "" + FmtStr(txtCustCodeTo.Text.Trim()) + ", " +
                                "" + FmtStr(DGVHIR.Rows[i].Cells["cpl_line_code"].Value.ToString()) + ", " +
                                "" + (i + 1).ToString() + ", " + //index
                                "" + FmtStr(DGVHIR.Rows[i].Cells["cpl_cust_type"].Value.ToString()) + " ";

                            _clsGlobal.ExecuteTrans(strSQL);

                        }
                        _clsGlobal.CommitTrans();
                    }
                }
                catch (Exception ex)
                {
                    _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //save salesman
                try
                {
                    strSQL = " select * from TBL_SD_CUSTCOVER where csc_cust_code1  = '" + txtCustCode.Text + "'AND csc_cust_code2  = '" + txtCustCodeTo.Text + "' AND csc_entity='" + ctrlEntityCustMaster2.txtCM.Text.Trim() + "' AND csc_branch='" + ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim() + "'";
                    DataTable dtbl = _clsGlobal.ExecDT(strSQL);

                    if (dtbl.Rows.Count > 0)
                    {
                        //hapus all
                        try
                        {
                            strSQL = "EXEC SP_AR_CUST_MASTER_CUSTCOVER '3', " +
                              "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + ", " +
                              "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + ", " +
                              "" + FmtStr(txtCustCode.Text) + ", " +
                              "" + FmtStr(txtCustCodeTo.Text) + ", " +
                              "  null, null , null, null, null, null, null, null , null" + " ";

                            _clsGlobal.BeginTrans();
                            _clsGlobal.ExecuteTrans(strSQL);
                            _clsGlobal.CommitTrans();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        int i, x = 0, z = 0;
                        _clsGlobal.BeginTrans();

                        for (i = 0; i < DVGSalesman.Rows.Count; i++)
                        {

                            z = i + 1;
                            if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Senin")
                            {
                                x = 1;
                            }
                            else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Selasa")
                            {
                                x = 2;
                            }
                            else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Rabu")
                            {
                                x = 3;
                            }

                            else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Kamis")
                            {
                                x = 4;
                            }

                            else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Jumat")
                            {
                                x = 5;
                            }

                            else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Sabtu")
                            {
                                x = 6;
                            }

                            else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Minggu")
                            {
                                x = 7;
                            }

                            strSQL = "EXEC SP_AR_CUST_MASTER_CUSTCOVER '1', " +
                                "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text) + ", " +
                                "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text) + ", " +
                                "" + FmtStr(txtCustCode.Text) + ", " +
                                "" + FmtStr(txtCustCodeTo.Text) + ", " +
                                "" + FmtStr(DVGSalesman.Rows[i].Cells["csc_salesman_id"].Value.ToString()) + ", " +
                                "" + x.ToString() + ", " +
                                "" + z + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week1"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week2"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week3"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week4"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + FmtStr(DVGSalesman.Rows[i].Cells["csc_route"].Value.ToString()) + ",'" + clsLogin.USERID + "' ";

                            _clsGlobal.ExecuteTrans(strSQL);

                        }
                        _clsGlobal.CommitTrans();
                    }
                }
                catch (Exception ex)
                {
                    _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                try
                {
                    if (isValidGroupPrice())
                    {
                        _clsGlobal.BeginTrans();

                        strSQL = "EXEC IP_INSERT_CUST_PRICE_GROUP " +
                        "1," +
                        "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + ", " +
                        "" + FmtStr(txtCustCode.Text) + ", " +
                        "" + FmtStr(txtCustCodeTo.Text) + ", " +
                        "NULL, " +
                        "NULL, " +
                        "NULL ";

                        _clsGlobal.ExecuteTrans(strSQL);

                        int i, x = 0;

                        for (i = 0; i < dgvGroupHarga.Rows.Count; i++)
                        {

                            x = i + 1;
                            strSQL = "EXEC IP_INSERT_CUST_PRICE_GROUP " +
                            "0," +
                            "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + ", " +
                            "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + ", " +
                            "" + FmtStr(txtCustCode.Text) + ", " +
                            "" + FmtStr(txtCustCodeTo.Text) + ", " +
                            "" + x + ", " +
                            "" + FmtStr(dgvGroupHarga.Rows[i].Cells["grp_code"].Value.ToString()) + ", " +
                            "'" + clsLogin.USERID + "' ";

                            _clsGlobal.ExecuteTrans(strSQL);
                        }

                        _clsGlobal.CommitTrans();
                    }
                    else
                    {

                        MessageBox.Show("Data Group Price STD (standart) harus ada, silahkan isi terlebih dahulu sebelum simpan data", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                }
                catch (Exception ex)
                {
                    _clsGlobal.RollbackTrans();
                    MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

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
                string mulaiush = _clsGlobal.DateToDB(dateTimePicker1.Value.ToString("MM/dd/yyyy"));
                string birthday = _clsGlobal.DateToDB(dateTimePicker2.Value.ToString("MM/dd/yyyy"));
                string insfrom = _clsGlobal.DateToDB(dtTglFromIns.Value.ToString("MM/dd/yyyy"));
                string insto = _clsGlobal.DateToDB(dtTglToIns.Value.ToString("MM/dd/yyyy"));
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
                    CBPembyaranFktT1.SelectedValue = "T";

                string skillidcode = skillidcodetb.Text;

                var stts = ((TIRAItem)cbstatus.SelectedItem) ?? new TIRAItem(__kode: "O", __nama: "OPEN NEW OUTLET");

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
                       "" + FmtStr(CBShipnBill.SelectedValue.ToString()) + ", " +
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
                       "" + FmtStr(CBPembyaranFktT1.SelectedValue.ToString()) + ", " +   // cm_payment_type                                                  
                       "" + FmtStr(cb3.Trim()) + ", " +
                       //"'" + _clsGlobal.DateToDB(dateTimepicker1.text) + "', " + //cm_outlet_open
                       "'" + mulaiush + "', " + //cm_outlet_open
                       "" + FmtStr(CBBagunanT1.SelectedValue.ToString()) + ", " +
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
                       "" + ((comboBoxPajakT1.SelectedIndex.ToString() != "") ? "'" + comboBoxPajakT1.SelectedIndex.ToString() + "'" : "0") + ", " +
                       "" + FmtStr(cb7.Trim()) + ", " + // cm_top_by_cust
                       "" + FmtStr(txtMoidCode.Text.Trim()) + ", " +//100
                       "" + FmtStr(txtMerchanid.Text.Trim()) + ", " +//cm_merchant_id 

                       //"" + FmtStr("O") + ", " + // @SPcm_active_flag
                       "" + FmtStr(stts.Kode) + ", " + // @SPcm_active_flag

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
                       "" + CBDelivByDay.SelectedValue + ", " +
                       "" + FmtStr(skillidcode) + ", " +
                       "" + FmtStr(cbPemilikNIK.SelectedValue.ToString()) + ", " +
                       "" + FmtStr(txtBillBranch.Text) + ", " +
                       "" + FmtStr(txtSPPKP.Text) + ", " +
                       "" + FmtStr(cb11.Trim()) + ", " +
                       "'" + insfrom + "', " +
                       "'" + insto + "', " +
                       "'" + dtOpenHour.Value.ToString("HH:mm:ss") + "', " +
                       "'" + dtCloseHour.Value.ToString("HH:mm:ss") + "', " +
                       "'" + nmPriority.Value.ToString() + "', " +
                       "'" + nmUnloadingTime.Value.ToString() + "' ";
                if (__isAsk27918)
                {
                    strSQL += "," + FmtStr(cbAddressChoice.SelectedValue.ToString()) + " ";
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
                strSQL = strSQL + "'" + cbJenisIdentitas.SelectedValue.ToString() + "',";
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
                int i;
                if (DGVHIR.RowCount > 0)
                {
                    //delete all
                    try
                    {
                        strSQL = "EXEC SP_AR_CUST_MASTER_CUSTPRDLINE '3', " +
                          "" + FmtStr(_prdEntityCode) + ", " +
                          "" + FmtStr(_prdBrandCode) + ", " +
                          "" + FmtStr(txtCustCode.Text) + ", " +
                          "" + FmtStr(txtCustCodeTo.Text) + ", " +
                          "  null, null, null " + " ";

                        _clsGlobal.BeginTrans();
                        _clsGlobal.ExecuteTrans(strSQL);
                        _clsGlobal.CommitTrans();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // insert data
                    _clsGlobal.BeginTrans();
                    for (i = 0; i < DGVHIR.Rows.Count; i++)
                    {
                        try
                        {
                            strSQL = "EXEC SP_AR_CUST_MASTER_CUSTPRDLINE '2', " +
                                "" + FmtStr(_prdEntityCode) + ", " +
                                "" + FmtStr(_prdBrandCode) + ", " +
                                "" + FmtStr(txtCustCode.Text.Trim()) + ", " +
                                "" + FmtStr(txtCustCodeTo.Text.Trim()) + ", " +
                                "" + FmtStr(DGVHIR.Rows[i].Cells["cpl_line_code"].Value.ToString()) + ", " +
                                "" + (i + 1).ToString() + ", " + //index
                                "" + FmtStr(DGVHIR.Rows[i].Cells["cpl_cust_type"].Value.ToString()) + " ";

                            _clsGlobal.ExecuteTrans(strSQL);
                        }
                        catch (Exception ex)
                        {
                            // MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    _clsGlobal.CommitTrans();
                }
                else
                {
                    try
                    {

                        strSQL = "EXEC SP_AR_CUST_MASTER_CUSTPRDLINE '3', " +
                          "" + FmtStr(_prdEntityCode) + ", " +
                          "" + FmtStr(_prdBrandCode) + ", " +
                          "" + FmtStr(txtCustCode.Text) + ", " +
                          "" + FmtStr(txtCustCodeTo.Text) + ", " +
                          "  null, null, null " + " ";

                        _clsGlobal.BeginTrans();
                        _clsGlobal.ExecuteTrans(strSQL);
                        _clsGlobal.CommitTrans();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //save salesman
            try
            {
                if (DVGSalesman.RowCount > 0)
                {
                    //hapus all
                    try
                    {
                        strSQL = "EXEC SP_AR_CUST_MASTER_CUSTCOVER '3', " +
                          "" + FmtStr(_prdEntityCode) + ", " +
                          "" + FmtStr(_prdBrandCode) + ", " +
                          "" + FmtStr(txtCustCode.Text) + ", " +
                          "" + FmtStr(txtCustCodeTo.Text) + ", " +
                          "  null, null , null, null, null, null, null, null, null " + " ";

                        _clsGlobal.BeginTrans();
                        _clsGlobal.ExecuteTrans(strSQL);
                        _clsGlobal.CommitTrans();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    //insert data
                    int i, x = 0, z = 0;
                    _clsGlobal.BeginTrans();
                    for (i = 0; i < DVGSalesman.Rows.Count; i++)
                    {
                        z = i + 1;

                        if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Senin")
                        {
                            x = 1;
                        }
                        else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Selasa")
                        {
                            x = 2;
                        }
                        else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Rabu")
                        {
                            x = 3;
                        }

                        else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Kamis")
                        {
                            x = 4;
                        }

                        else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Jumat")
                        {
                            x = 5;
                        }

                        else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Sabtu")
                        {
                            x = 6;
                        }

                        else if (DVGSalesman.Rows[i].Cells["csc_visit"].Value.ToString() == "Minggu")
                        {
                            x = 7;
                        }
                        try
                        {
                            strSQL = "EXEC SP_AR_CUST_MASTER_CUSTCOVER '2', " +
                                "" + FmtStr(_prdEntityCode) + ", " +
                                "" + FmtStr(_prdBrandCode) + ", " +
                                "" + FmtStr(txtCustCode.Text) + ", " +
                                "" + FmtStr(txtCustCodeTo.Text) + ", " +
                                "" + FmtStr(DVGSalesman.Rows[i].Cells["csc_salesman_id"].Value.ToString()) + ", " +
                                 "" + (x.ToString()) + ", " +
                                "" + z + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week1"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week2"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week3"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + ((DVGSalesman.Rows[i].Cells["csc_visit_week4"].Value.ToString() == "True") ? "Y" : "T") + ", " +
                                "" + FmtStr(DVGSalesman.Rows[i].Cells["csc_route"].Value.ToString()) + ",'" + clsLogin.USERID + "'  ";

                            _clsGlobal.ExecuteTrans(strSQL);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    _clsGlobal.CommitTrans();

                }
                else
                {

                    try
                    {
                        //delete all
                        strSQL = "EXEC SP_AR_CUST_MASTER_CUSTCOVER '3', " +
                            "" + FmtStr(_prdEntityCode) + ", " +
                            "" + FmtStr(_prdBrandCode) + ", " +
                            "" + FmtStr(txtCustCode.Text) + ", " +
                            "" + FmtStr(txtCustCodeTo.Text) + ", " +
                            "  null, null , null, null, null, null, null, null, null ";

                        _clsGlobal.BeginTrans();
                        _clsGlobal.ExecuteTrans(strSQL);
                        _clsGlobal.CommitTrans();
                    }
                    catch (Exception ex)
                    {
                        // MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                if (isValidGroupPrice())
                {
                    _clsGlobal.BeginTrans();

                    strSQL = "EXEC IP_INSERT_CUST_PRICE_GROUP " +
                    "1," +
                    "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + ", " +
                    "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + ", " +
                    "" + FmtStr(txtCustCode.Text) + ", " +
                    "" + FmtStr(txtCustCodeTo.Text) + ", " +
                    "NULL, " +
                    "NULL, " +
                    "NULL ";

                    _clsGlobal.ExecuteTrans(strSQL);

                    int i, x = 0;

                    for (i = 0; i < dgvGroupHarga.Rows.Count; i++)
                    {

                        x = i + 1;
                        strSQL = "EXEC IP_INSERT_CUST_PRICE_GROUP " +
                        "0," +
                        "" + FmtStr(ctrlEntityCustMaster2.txtCM.Text.Trim()) + ", " +
                        "" + FmtStr(ctrlEntityCustMaster2.txtBranchIdCM.Text.Trim()) + ", " +
                        "" + FmtStr(txtCustCode.Text) + ", " +
                        "" + FmtStr(txtCustCodeTo.Text) + ", " +
                        "" + x + ", " +
                        "" + FmtStr(dgvGroupHarga.Rows[i].Cells["grp_code"].Value.ToString()) + ", " +
                        "'" + clsLogin.USERID + "' ";

                        _clsGlobal.ExecuteTrans(strSQL);
                    }

                    _clsGlobal.CommitTrans();
                }
                else
                {
                    MessageBox.Show("Data Group Price STD (standart) harus ada, silahkan isi terlebih dahulu sebelum simpan data", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {

                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

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
            bool result = false;
            string stdtype = "";

            for (int i = 0; i < dgvGroupHarga.Rows.Count; i++)
            {
                stdtype = dgvGroupHarga.Rows[i].Cells["grp_code"].Value.ToString();
                if (stdtype == "STD")
                {
                    result = true;
                    break;
                }
            }

            return result;

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

        private void DVGSalesman_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column1_KeyPress);
            if (DVGSalesman.CurrentCell.ColumnIndex == 9) //Desired Column
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

        private void DVGSalesman_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                DataGridViewRow dgr, dgrB;
                string column0, column1;
                string maxB = "";
                string min = "";

                try
                {
                    dgr = DVGSalesman.Rows[DVGSalesman.CurrentRow.Index];

                    if (DVGSalesman.Rows[e.RowIndex].Cells["csc_salesman_id"].Value.ToString() != "")
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
                                if (DVGSalesman.RowCount > 1)
                                {
                                    dgrB = DVGSalesman.Rows[DVGSalesman.CurrentRow.Index - 1];
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

                                        MessageBox.Show("Salesman Sudah ada ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK);
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

            for (i = 0; i < DVGSalesman.Rows.Count; i++)
            {
                //cek nilai yg null atau kosong
                if ((DVGSalesman.Rows[i].Cells["csc_salesman_id"].Value == null) || (DVGSalesman.Rows[i].Cells["sgm_spgm_name"].Value == null) || (DVGSalesman.Rows[i].Cells["sgm_type_operasi"].Value == null))
                {
                    isDuplicate = true;
                    break;
                }

                string empEnt = DVGSalesman.Rows[i].Cells["csc_salesman_id"].Value.ToString().Trim();
                string empBrn = DVGSalesman.Rows[i].Cells["sgm_spgm_name"].Value.ToString().Trim();
                string empDiv = DVGSalesman.Rows[i].Cells["sgm_type_operasi"].Value.ToString().Trim();

                //cek duplicate data
                for (x = 0; x < DVGSalesman.Rows.Count; x++)
                {
                    if (i != x)//data yg dipilih tidak usah di compare
                    {
                        if (((empEnt == DVGSalesman.Rows[x].Cells["csc_salesman_id"].Value.ToString().Trim()) && (empBrn == DVGSalesman.Rows[x].Cells["sgm_spgm_name"].Value.ToString().Trim()) && (empDiv == DVGSalesman.Rows[x].Cells["sgm_type_operasi"].Value.ToString().Trim())) == true)
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

        private void getGroupHarga(TextBox objText)
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
            rtHeader.Height = this.DVGSalesman.ColumnHeadersHeight / 2;
            this.DVGSalesman.Invalidate(rtHeader);
        }

        private void DVGSalesman_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
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

        private void DVGSalesman_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            Rectangle rtHeader = this.DVGSalesman.DisplayRectangle;
            rtHeader.Height = this.DVGSalesman.ColumnHeadersHeight / 2;
            this.DVGSalesman.Invalidate(rtHeader);
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage.Name == tabPage1.Name)
            {
                txtReason.Visible = false;
                label18.Visible = false;
            }
            if (e.TabPage.Name == tabPage2.Name)
            {
                txtReason.Visible = false;
                label18.Visible = false;
            }
            if (e.TabPage.Name == tabPage3.Name)
            {
                txtReason.Visible = false;
                label18.Visible = false;
            }
            if (e.TabPage.Name == tabPage4.Name)
            {
                txtReason.Visible = false;
                label18.Visible = false;
            }
            if (e.TabPage.Name == tabPage5.Name)
            {
                txtReason.Visible = false;
                label18.Visible = false;
            }
            if (e.TabPage.Name == tabPage6.Name)
            {
                txtReason.Visible = false;
                label18.Visible = false;
            }
        }

        private void DGVBackList_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(DVGSalesman.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void DGVUnBackList_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(DVGSalesman.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void txtCustCode_TextChanged(object sender, EventArgs e)
        {
            IsiTextBill(CBShipnBill.SelectedIndex, txtCustCode.Text);
        }

        private void txtLeadtime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtTaxIdT2_TextChanged(object sender, EventArgs e)
        {
            getTaxID(txtTaxIdT2);
        }

        private void getTaxID(TextBox objText)
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

        void LoadTOP(TextBox objText)
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
            string __data = ((TextBox)sender).Text;
            ((TextBox)sender).Text = string.Format("{0:#.##}", removeFormat(__data));
        }
        private void tbNumeric_Leave(object sender, EventArgs e)
        {
            string __data = ((TextBox)sender).Text.Trim();
            try
            {
                decimal __value = 0;
                decimal.TryParse(__data, out __value);
                __data = string.Format("{0:#,##0.00}", __value);
            }
            catch (Exception) { }
            ((TextBox)sender).Text = __data;
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
                DataGridViewRow dgr;
                string column0, column1;
                try
                {
                    dgr = dgvGroupHarga.CurrentRow;

                    if (dgvGroupHarga.Rows.Count > 0)
                    {

                        column0 = dgr.Cells["grp_code"].Value.ToString();
                        column1 = dgr.Cells["grp_desc"].Value.ToString();

                        if (column0 != "")
                        {
                            //if (clsGlobal.MODE_TRX == 1)
                            if (__state == StateEntry.New)
                            {
                                dgvGroupHarga.Rows.Add("", "", "", "", "", "");
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
                            dgvGroupHarga.Rows.Add("", "", "", "", "", "");
                        }
                        else
                        {
                            if (dgvGroupHarga.RowCount > 0)
                            {
                                dgvGroupHarga.Rows.Add("", "", "", "", "", "");
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
                if (dgvGroupHarga.RowCount != 0)
                {
                    dgvGroupHarga.Rows.Remove(dgvGroupHarga.CurrentRow);
                }
            }
        }

        private void dgvGroupHarga_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var senderGrid = (DataGridView)sender;
            if (senderGrid.CurrentCell.ColumnIndex == dgvGroupHarga.Columns["btnGroup"].Index)
            {
                frmPopUp frm = new frmPopUp();
                frm.FrmText = "Search Group Price ";
                frm.Query = "select ppl_price_list_code [Group Price Code], ppl_price_list_description [Description] from OB_PRICE_LIST_HEADER  ";
                frm.ShowDialog();
                if (frm.ArrField != null)
                {
                    try
                    {
                        DataGridViewRow dgr, dgrB;
                        string maxB = "";
                        string min = "";
                        dgrB = null;

                        dgr = dgvGroupHarga.Rows[dgvGroupHarga.CurrentRow.Index];

                        if (dgvGroupHarga.RowCount > 1)
                        {
                            if (dgr.Index != 0)
                            {
                                dgrB = dgvGroupHarga.Rows[dgvGroupHarga.CurrentRow.Index - 1];
                                maxB = dgrB.Cells["grp_code"].Value.ToString();
                            }
                            else
                            {
                                min = frm.ArrField[0].Trim();
                                rowDuplicate = false;
                                exitData = false;

                                foreach (DataGridViewRow gRow in dgvGroupHarga.Rows)
                                {
                                    if ((gRow.Cells["grp_code"].Value.ToString().Trim() == min))
                                    {
                                        rowDuplicate = true;
                                        break;
                                    }
                                }
                                if (rowDuplicate == true)
                                {
                                    MessageBox.Show("Group Code " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Group Price yang berbeda ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                    dgvGroupHarga.Rows[e.RowIndex].Cells["grp_code"].Value = frm.ArrField[0].Trim();
                                    dgvGroupHarga.Rows[e.RowIndex].Cells["grp_desc"].Value = frm.ArrField[1].Trim();
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

                            foreach (DataGridViewRow gRow in dgvGroupHarga.Rows)
                            {
                                if ((gRow.Cells["grp_code"].Value.ToString().Trim() == min))
                                {
                                    rowDuplicate = true;
                                    break;
                                }
                            }

                            if (rowDuplicate == true)
                            {
                                MessageBox.Show("Group Price " + frm.ArrField[0].Trim() + "~" + frm.ArrField[1].Trim() + " sudah ada dalam baris sebelumnya.\nSilahkan pilih Group Prie yang berbeda ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                dgvGroupHarga.Rows[e.RowIndex].Cells["grp_code"].Value = frm.ArrField[0].Trim();
                                dgvGroupHarga.Rows[e.RowIndex].Cells["grp_desc"].Value = frm.ArrField[1].Trim();
                            }
                        }
                        else
                        {
                            if (exitData != true)
                            {
                                dgvGroupHarga.Rows[e.RowIndex].Cells["grp_code"].Value = frm.ArrField[0].Trim();
                                dgvGroupHarga.Rows[e.RowIndex].Cells["grp_desc"].Value = frm.ArrField[1].Trim();
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
                        MessageBox.Show("Data Shipment area tidak bisa di Sync mohon pilih mode manual.");
                        return;
                    }
                }

                strSQL = "";
                strSQL = "SELECT * FROM VW_SHIP_AREA_DETAIL_LOOKUP where msd_entity_id = '" + ctrlEntityCustMaster2.txtCM.Text + "' and msd_branch_id = '" + ctrlEntityCustMaster2.txtBranchIdCM.Text + "' and kl_prov_code = '" + popUpProvinsinCity1.txtPropinsiCMC.Text + "' and kl_city_code = '" + popUpProvinsinCity1.txtKabupatenCMC.Text + "' and kl_kecamatan_code = '" + popUpProvinsinCity1.txtKecamatan.Text + "' and kl_kelurahan_code = '" + popUpProvinsinCity1.txtKelurahan.Text + "' ";
                DataTable dt = _clsGlobal.ExecDT(strSQL);
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Data Shipment area tidak ditemukan untuk lokasi " + popUpProvinsinCity1.txtPropinsiCMC.Text + popUpProvinsinCity1.txtKabupatenCMC.Text + popUpProvinsinCity1.txtKecamatan.Text + popUpProvinsinCity1.txtKelurahan.Text + "");
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

        private void getShipArea(TextBox objText)
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
            frm.Query = " SELECT cm_entity[Entity], cm_branch[Branch], cm_cust_code1[Code1], cm_cust_code2 [Code2],cm_cust_name[Name], cm_delv_address1[Address]  From SO_CUST_MASTER ";
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
            DataGridViewTextBoxColumn col0 = new DataGridViewTextBoxColumn()
            {
                Name = "nocol",
                HeaderText = "No.",
                ValueType = typeof(int),
                ReadOnly = true,
                DataPropertyName = "no",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                CellTemplate = clsGlobalStatic.ICellInt,
                MinimumWidth = 40,
                Frozen = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn()
            {
                Name = "prdlinecol",
                HeaderText = "Product Line",
                ValueType = typeof(string),
                ReadOnly = false,
                DataPropertyName = "prdline_code",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                SortMode = DataGridViewColumnSortMode.NotSortable,

            };
            DataGridViewButtonColumn col1b = new DataGridViewButtonColumn()
            {
                Name = "prdlinebtncol",
                HeaderText = "",
                Text = "...",
                ReadOnly = false,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                Width = 40,
                DefaultCellStyle = new DataGridViewCellStyle()
                {
                    Padding = new Padding(0),
                },
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewTextBoxColumn col2 = new DataGridViewTextBoxColumn()
            {
                Name = "desccol",
                HeaderText = "Nama Product Line",
                ValueType = typeof(string),
                ReadOnly = true,
                DataPropertyName = "prdline_desc",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            //DataGridViewComboBoxColumn col3 = new DataGridViewComboBoxColumn()
            //{
            //    Name = "topcol",
            //    HeaderText = "TOP",
            //    ValueType = typeof(string),
            //    ReadOnly = false,
            //    DataPropertyName = "top",
            //    AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
            //    SortMode = DataGridViewColumnSortMode.NotSortable,
            //    DataSource = __top.DATA,
            //    DisplayMember = "Nama",
            //    ValueMember = "Kode",
            //    MinimumWidth = 70,
            //};
            DataGridViewTextBoxColumn col3 = new DataGridViewTextBoxColumn()
            {
                Name = "topcol",
                HeaderText = "TOP",
                ValueType = typeof(string),
                ReadOnly = false,
                DataPropertyName = "top_code",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                SortMode = DataGridViewColumnSortMode.NotSortable,

            };
            DataGridViewButtonColumn col3b = new DataGridViewButtonColumn()
            {
                Name = "topbtncol",
                HeaderText = "",
                Text = "...",
                ReadOnly = false,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                Width = 40,
                DefaultCellStyle = new DataGridViewCellStyle()
                {
                    Padding = new Padding(0),
                },
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewTextBoxColumn col4 = new DataGridViewTextBoxColumn()
            {
                Name = "topdesccol",
                HeaderText = "Nama TOP",
                ValueType = typeof(string),
                ReadOnly = true,
                DataPropertyName = "top_desc",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };

            topbyprdlinedgv.ReadOnly = false;
            topbyprdlinedgv.AutoGenerateColumns = false;
            topbyprdlinedgv.ColumnHeadersHeight = 30;
            topbyprdlinedgv.Columns.AddRange(col0, col1, col1b, col2, col3, col3b, col4);
            topbyprdlinedgv.AlternatingRowsDefaultCellStyle.BackColor = __COLOR;

            DataTable __source = new DataTable();
            foreach (DataGridViewColumn c in topbyprdlinedgv.Columns.Cast<DataGridViewColumn>().Where(x => x.ValueType != null))
                __source.Columns.Add(c.DataPropertyName, c.ValueType);
            __source.Rows.Add(__source.NewRow());
            topbyprdlinedgv.DataSource = __source;
            numberTopByPrdLine();
        }

        private void topbyprdlinedgv_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Insert)
                {
                    DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
                    if (__source == null) return;
                    __source.Rows.Add(__source.NewRow());
                    numberTopByPrdLine();
                }
                else if (e.KeyCode == Keys.Delete)
                {
                    DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
                    if (__source == null) return;
                    if (topbyprdlinedgv.CurrentCell == null) return;
                    int iRow = topbyprdlinedgv.CurrentCell.RowIndex;
                    var value = topbyprdlinedgv.Rows[iRow].Cells["nocol"].Value;
                    IEnumerable<DataRow> rows = new List<DataRow>(__source.Rows.Cast<DataRow>().Where(x => x["no"].Equals(value)));
                    foreach (DataRow r in rows)
                        __source.Rows.Remove(r);
                    numberTopByPrdLine();
                }
                else if (e.KeyCode == Keys.F4)
                {
                    if (topbyprdlinedgv.CurrentCell == null) return;
                    topbyprdlinedgv_CellContentClick(sender, new DataGridViewCellEventArgs(2, topbyprdlinedgv.CurrentCell.RowIndex));
                }
            }
            catch (Exception) { }
        }
        private void topbyprdlinedgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string name = topbyprdlinedgv.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinebtncol"))
                {
                    string key = topbyprdlinedgv["prdlinecol", e.RowIndex].Value.ToString();
                    __prdlines.ActionSelector(key, true, new IndexGrid(e));
                }
                else if (name.CompareC("topbtncol"))
                {
                    string key = topbyprdlinedgv["topcol", e.RowIndex].Value.ToString();
                    __top.ActionSelector(key, true, new IndexGrid(e));
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
        private void topbyprdlinedgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string name = topbyprdlinedgv.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinecol"))
                {
                    string key = topbyprdlinedgv[e.ColumnIndex, e.RowIndex].Value.ToString();
                    __prdlines.ActionSelector(key, false, new IndexGrid(e));
                }
                else if (name.CompareC("topcol"))
                {
                    string key = topbyprdlinedgv[e.ColumnIndex, e.RowIndex].Value.ToString();
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
                        this.topbyprdlinedgv["prdlinecol", e.Row].Value = string.Empty;
                        this.topbyprdlinedgv["desccol", e.Row].Value = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (__item == null) return;

                        DataTable __source = (DataTable)topbyprdlinedgv.DataSource;
                        if (__source == null) return;
                        if (__source.Rows.Cast<DataRow>().Count(x => checkExistCode(x["prdline_code"], __item.Kode)) > 1)
                        {
                            this.topbyprdlinedgv["prdlinecol", e.Row].Value = string.Empty;
                            this.topbyprdlinedgv["desccol", e.Row].Value = string.Empty;
                        }
                        else
                        {
                            this.topbyprdlinedgv["prdlinecol", e.Row].Value = __item.Kode;
                            this.topbyprdlinedgv["desccol", e.Row].Value = __item.Nama;
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
                        this.dgvPrdLineBlocking["prdlinecol", e.Row].Value = string.Empty;
                        this.dgvPrdLineBlocking["prdlinedesccol", e.Row].Value = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (__item == null) return;

                        DataTable __source = (DataTable)dgvPrdLineBlocking.DataSource;
                        if (__source == null) return;
                        if (__source.Rows.Cast<DataRow>().Count(x => checkExistCode(x["prdlinecode"], __item.Kode)) > 1)
                        {
                            this.dgvPrdLineBlocking["prdlinecol", e.Row].Value = string.Empty;
                            this.dgvPrdLineBlocking["prdlinedesccol", e.Row].Value = string.Empty;
                        }
                        else
                        {
                            this.dgvPrdLineBlocking["prdlinecol", e.Row].Value = __item.Kode;
                            this.dgvPrdLineBlocking["prdlinedesccol", e.Row].Value = __item.Nama;
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
                        this.topbyprdlinedgv["topcol", e.Row].Value = string.Empty;
                        this.topbyprdlinedgv["topdesccol", e.Row].Value = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (__item == null) return;
                        this.topbyprdlinedgv["topcol", e.Row].Value = __item.Kode;
                        this.topbyprdlinedgv["topdesccol", e.Row].Value = __item.Nama;
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

        private void topbyprdlinedgv_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (topbyprdlinedgv.CurrentCell == null) return;
            if (!topbyprdlinedgv.Columns[topbyprdlinedgv.CurrentCell.ColumnIndex].Name.CompareC("prdlinecol")) return;
            if (e.Control is TextBox)
                ((TextBox)(e.Control)).CharacterCasing = CharacterCasing.Upper;
        }

        private void topbyprdlinedgv_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            topbyprdlinedgv.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void txtDelivByDays_TextChanged(object sender, EventArgs e)
        {
            int i = 0;
            int.TryParse(txtDelivByDays.Text.Trim(), out i);
            if (i > 0) CBDelivByDay.SelectedValue = 0;
        }

        private void CBDelivByDay_SelectedValueChanged(object sender, EventArgs e)
        {
            int i = 0;
            int.TryParse(CBDelivByDay.SelectedValue.ToString().Trim(), out i);
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
            int[] tgl = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31 };

            DataGridViewComboBoxColumn col0 = new DataGridViewComboBoxColumn()
            {
                Name = "tglcol",
                HeaderText = "Tanggal",
                ValueType = typeof(int),
                ReadOnly = false,
                DataPropertyName = "tgl",
                DataSource = tgl,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                MinimumWidth = 80,
                Frozen = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn()
            {
                Name = "ketcol",
                HeaderText = "Keterangan",
                ValueType = typeof(string),
                ReadOnly = false,
                DataPropertyName = "ket",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };

            dgvschpayD.ReadOnly = false;
            dgvschpayD.AutoGenerateColumns = false;
            dgvschpayD.ColumnHeadersHeight = 20;
            dgvschpayD.Columns.AddRange(col0, col1);
            dgvschpayD.AlternatingRowsDefaultCellStyle.BackColor = __COLOR;
            StackedHeaderDecorator objREnderer = new StackedHeaderDecorator(dgvschpayD);

            DataTable __source = new DataTable();
            foreach (DataGridViewColumn c in dgvschpayD.Columns.Cast<DataGridViewColumn>().Where(x => x.ValueType != null))
                __source.Columns.Add(c.DataPropertyName, c.ValueType);
            __source.Columns["tgl"].DefaultValue = 1;
            //__source.Rows.Add(__source.NewRow());
            dgvschpayD.DataSource = __source;
        }

        private void SetGridViewSchPayW()
        {
            string[] hari = { "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu" };
            //List<Hari> hari = new List<Hari>();
            //int i = 0;

            //foreach (string nama in namahari)
            //    hari.Add(new Hari()
            //    {
            //        Code = i++,
            //        Name = nama,
            //    });

            DataGridViewComboBoxColumn col0 = new DataGridViewComboBoxColumn()
            {
                Name = "haricol",
                HeaderText = "Hari",
                ValueType = typeof(string),
                ReadOnly = false,
                DataPropertyName = "hari",
                //ValueMember = "Code",
                //DisplayMember = "Name",
                DisplayIndex = 1,
                DataSource = hari,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                MinimumWidth = 100,
                Frozen = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewCheckBoxColumn col1 = new DataGridViewCheckBoxColumn()
            {
                Name = "pola1col",
                HeaderText = "Pola.1",
                ValueType = typeof(bool),
                ReadOnly = false,
                DataPropertyName = "pola1",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                MinimumWidth = 40,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewCheckBoxColumn col2 = new DataGridViewCheckBoxColumn()
            {
                Name = "pola2col",
                HeaderText = "Pola.2",
                ValueType = typeof(bool),
                ReadOnly = false,
                DataPropertyName = "pola2",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                MinimumWidth = 40,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewCheckBoxColumn col3 = new DataGridViewCheckBoxColumn()
            {
                Name = "pola3col",
                HeaderText = "Pola.3",
                ValueType = typeof(bool),
                ReadOnly = false,
                DataPropertyName = "pola3",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                MinimumWidth = 40,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewCheckBoxColumn col4 = new DataGridViewCheckBoxColumn()
            {
                Name = "pola4col",
                HeaderText = "Pola.4",
                ValueType = typeof(bool),
                ReadOnly = false,
                DataPropertyName = "pola4",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                MinimumWidth = 40,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewTextBoxColumn col5 = new DataGridViewTextBoxColumn()
            {
                Name = "ketcol",
                HeaderText = "Keterangan",
                ValueType = typeof(string),
                ReadOnly = false,
                DataPropertyName = "ket",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };

            dgvschpayW.ReadOnly = false;
            dgvschpayW.AutoGenerateColumns = false;
            dgvschpayW.ColumnHeadersHeight = 20;
            dgvschpayW.Columns.AddRange(col0, col1, col2, col3, col4, col5);
            dgvschpayW.AlternatingRowsDefaultCellStyle.BackColor = __COLOR;
            StackedHeaderDecorator objREnderer = new StackedHeaderDecorator(dgvschpayW);

            DataTable __source = new DataTable();
            foreach (DataGridViewColumn c in dgvschpayW.Columns.Cast<DataGridViewColumn>().Where(x => x.ValueType != null))
                __source.Columns.Add(c.DataPropertyName, c.ValueType);
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
                    if (e.KeyCode == Keys.Insert)
                    {
                        DataTable __source = (DataTable)dgvschpayD.DataSource;
                        if (__source == null) return;
                        __source.Rows.Add(__source.NewRow());
                    }
                    else if (e.KeyCode == Keys.Delete)
                    {
                        DataTable __source = (DataTable)dgvschpayD.DataSource;
                        if (__source == null) return;
                        if (dgvschpayD.CurrentCell == null) return;
                        int iRow = dgvschpayD.CurrentCell.RowIndex;
                        DataRow row = __source.Rows[iRow];
                        __source.Rows.Remove(row);
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
                    if (e.KeyCode == Keys.Insert)
                    {
                        DataTable __source = (DataTable)dgvschpayW.DataSource;
                        if (__source == null) return;
                        __source.Rows.Add(__source.NewRow());
                    }
                    else if (e.KeyCode == Keys.Delete)
                    {
                        DataTable __source = (DataTable)dgvschpayW.DataSource;
                        if (__source == null) return;
                        if (dgvschpayW.CurrentCell == null) return;
                        int iRow = dgvschpayW.CurrentCell.RowIndex;
                        DataRow row = __source.Rows[iRow];
                        __source.Rows.Remove(row);
                    }
                }
            }
            catch (Exception) { }
        }

        private void dgvschpayW_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                if (dgvschpayW["pola1col", e.RowIndex].Value.ToString() == "True"
                    || dgvschpayW["pola2col", e.RowIndex].Value.ToString() == "True"
                    || dgvschpayW["pola3col", e.RowIndex].Value.ToString() == "True"
                    || dgvschpayW["pola4col", e.RowIndex].Value.ToString() == "True")
                {
                    dgvschpayW["pola1col", e.RowIndex].Value = false;
                    dgvschpayW["pola2col", e.RowIndex].Value = false;
                    dgvschpayW["pola3col", e.RowIndex].Value = false;
                    dgvschpayW["pola4col", e.RowIndex].Value = false;

                    dgvschpayW[e.ColumnIndex, e.RowIndex].Value = true;
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
            DataGridViewTextBoxColumn col0 = new DataGridViewTextBoxColumn()
            {
                Name = "nocol",
                HeaderText = "No.",
                ValueType = typeof(int),
                ReadOnly = true,
                DataPropertyName = "no",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader,
                CellTemplate = clsGlobalStatic.ICellInt,
                MinimumWidth = 40,
                Frozen = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn()
            {
                Name = "prdlinecol",
                HeaderText = "Product Line",
                ValueType = typeof(string),
                ReadOnly = false,
                DataPropertyName = "prdlinecode",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                SortMode = DataGridViewColumnSortMode.NotSortable,

            };
            DataGridViewButtonColumn col2 = new DataGridViewButtonColumn()
            {
                Name = "prdlinebtncol",
                HeaderText = "",
                Text = "...",
                ReadOnly = false,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                Width = 40,
                DefaultCellStyle = new DataGridViewCellStyle()
                {
                    Padding = new Padding(0),
                },
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewTextBoxColumn col3 = new DataGridViewTextBoxColumn()
            {
                Name = "prdlinedesccol",
                HeaderText = "Nama Product Line",
                ValueType = typeof(string),
                ReadOnly = true,
                DataPropertyName = "prdlinedesc",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewCheckBoxColumn col4 = new DataGridViewCheckBoxColumn()
            {
                Name = "chksalescol",
                HeaderText = "Sales",
                ValueType = typeof(bool),
                ReadOnly = false,
                DataPropertyName = "chksales",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            DataGridViewCheckBoxColumn col5 = new DataGridViewCheckBoxColumn()
            {
                Name = "chkreturcol",
                HeaderText = "Retur",
                ValueType = typeof(bool),
                ReadOnly = false,
                DataPropertyName = "chkretur",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };

            dgvPrdLineBlocking.ReadOnly = false;
            dgvPrdLineBlocking.AutoGenerateColumns = false;
            dgvPrdLineBlocking.ColumnHeadersHeight = 30;
            dgvPrdLineBlocking.Columns.AddRange(col0, col1, col2, col3, col4, col5);
            dgvPrdLineBlocking.AlternatingRowsDefaultCellStyle.BackColor = __COLOR;

            DataTable __source = new DataTable();
            foreach (DataGridViewColumn c in dgvPrdLineBlocking.Columns.Cast<DataGridViewColumn>().Where(x => x.ValueType != null))
                __source.Columns.Add(c.DataPropertyName, c.ValueType);
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

        private void dgvPrdLineBlocking_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string name = dgvPrdLineBlocking.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinecol"))
                {
                    string key = dgvPrdLineBlocking[e.ColumnIndex, e.RowIndex].Value.ToString();
                    __prdlineblocking.ActionSelector(key, false, new IndexGrid(e));
                }
            }
            catch (Exception) { }
        }

        private void dgvPrdLineBlocking_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgvPrdLineBlocking.CurrentCell == null) return;
            if (!dgvPrdLineBlocking.Columns[dgvPrdLineBlocking.CurrentCell.ColumnIndex].Name.CompareC("prdlinecol")) return;
            if (e.Control is TextBox)
                ((TextBox)(e.Control)).CharacterCasing = CharacterCasing.Upper;
        }

        private void dgvPrdLineBlocking_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //dgvPrdLineBlocking.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgvPrdLineBlocking_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Insert)
                {
                    DataTable __source = (DataTable)dgvPrdLineBlocking.DataSource;
                    if (__source == null) return;
                    __source.Rows.Add(__source.NewRow());
                    numberPrdLineBlocking();
                }
                else if (e.KeyCode == Keys.Delete)
                {
                    DataTable __source = (DataTable)dgvPrdLineBlocking.DataSource;
                    if (__source == null) return;
                    if (dgvPrdLineBlocking.CurrentCell == null) return;
                    int iRow = dgvPrdLineBlocking.CurrentCell.RowIndex;
                    var value = dgvPrdLineBlocking.Rows[iRow].Cells["nocol"].Value;
                    IEnumerable<DataRow> rows = new List<DataRow>(__source.Rows.Cast<DataRow>().Where(x => x["no"].Equals(value)));
                    foreach (DataRow r in rows)
                        __source.Rows.Remove(r);
                    numberPrdLineBlocking();
                }
                else if (e.KeyCode == Keys.F4)
                {
                    if (dgvPrdLineBlocking.CurrentCell == null) return;
                    dgvPrdLineBlocking_CellContentClick(sender, new DataGridViewCellEventArgs(2, dgvPrdLineBlocking.CurrentCell.RowIndex));
                }
            }
            catch (Exception) { }
        }

        private void dgvPrdLineBlocking_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string name = dgvPrdLineBlocking.Columns[e.ColumnIndex].Name;
                if (name.CompareC("prdlinebtncol"))
                {
                    string key = dgvPrdLineBlocking["prdlinecol", e.RowIndex].Value.ToString();
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

        private void gridPajak_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string pCode = "";
                var senderGrid = (DataGridView)sender;

                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                    e.RowIndex >= 0)
                {
                    int rowindex2 = gridPajak.CurrentRow.Index;
                    var row = this.gridPajak.Rows[rowindex2];
                    frmPopUp frm = new frmPopUp();
                    frm.FrmText = "Search Tipe Pajak";
                    frm.Query = "select stt_tax_code,stt_tax_desc from so_tax_type ";
                    frm.ShowDialog();
                    if (frm.ArrField != null)
                    {
                        row.Cells["code"].Value = frm.ArrField[0].Trim();
                        row.Cells["desc"].Value = frm.ArrField[1].Trim();
                        gridPajak_CellEndEdit(sender, e);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridPajak_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gridPajak.Columns[e.ColumnIndex].Name == "code" || gridPajak.Columns[e.ColumnIndex].Name == "btn")
                {
                    string condID = "";

                    if (e.RowIndex > -1)
                    {
                        if (gridPajak.Rows[e.RowIndex].Cells["code"].Value != null)
                        {
                            condID = gridPajak.Rows[e.RowIndex].Cells["code"].Value.ToString();
                        }
                        strSQL = "";
                        strSQL = "select stt_tax_code,stt_tax_desc from so_tax_type where stt_tax_code = '" + condID + "' ";

                        DataTable dtDesc = _clsGlobal.ExecDT(strSQL);

                        if (dtDesc.Rows.Count > 0)
                        {
                            if (pCheckValue(gridPajak.Rows[e.RowIndex].Cells["code"].Value.ToString().Trim()) == true)
                            {
                                gridPajak.Rows[e.RowIndex].Cells["code"].Value = dtDesc.Rows[0]["sct_cond_type_id"].ToString().Trim();
                                gridPajak.Rows[e.RowIndex].Cells["desc"].Value = dtDesc.Rows[0]["sct_cond_type_desc"].ToString().Trim();

                            }
                            else
                            {
                                MessageBox.Show("Data Sudah Ada!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                                gridPajak.Rows[e.RowIndex].Cells["code"].Value = "";
                                gridPajak.Rows[e.RowIndex].Cells["desc"].Value = "";

                            }
                        }
                        else
                        {
                            gridPajak.Rows[e.RowIndex].Cells["desc"].Value = "";
                        }

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        int max = 0;
        private void gridPajak_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(gridPajak.RowHeadersDefaultCellStyle.ForeColor))
            {
                SizeF sif = e.Graphics.MeasureString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font);
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
                max = Math.Max(max, (int)sif.Width + 20);
                ((DataGridView)sender).RowHeadersWidth = max;
            }
        }

        private void gridPajak_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Insert || e.KeyCode == Keys.Down)
                {
                    if (gridPajak.Rows.Count < LIMIT_COUNT_HOTT_TAX_CODES)
                    {
                        int rowIndex = this.gridPajak.Rows.Add();
                        var row = this.gridPajak.Rows[rowIndex];
                        row.Selected = true;
                        row.Cells[0].Selected = true;
                    }
                    else
                    {
                        MessageBox.Show("Maksmial "+ LIMIT_COUNT_HOTT_TAX_CODES +" Tipe Pajak !!", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                if (e.KeyCode == Keys.Delete)
                {
                    int rowIndex = ((DataGridView)sender).CurrentRow.Index;
                    var row = this.gridPajak.Rows[rowIndex];
                    gridPajak.Rows.Remove(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string getHottTaxCodes()
        {
            List<string> taxCodes = new List<string>();
            for(int i = 0; i < gridPajak.Rows.Count; i++)
            {
                object codeValue = gridPajak.Rows[i].Cells["code"].Value;
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
                gridPajak.Rows.Clear();
                if (dt.Rows.Count > 0)
                {
                    //kalo ada data detail
                    if (dt.Rows[0]["stt_tax_code"].ToString().Trim() != "")
                    {

                        foreach (DataRow rw in dt.Rows)
                        {
                            gridPajak.Rows.Insert(gridPajak.Rows.Count);
                            DataGridViewRow newRow = gridPajak.Rows[gridPajak.Rows.Count - 1];
                            newRow.Cells[0].Value = rw["stt_tax_code"];
                            newRow.Cells[2].Value = rw["stt_tax_desc"];
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool pCheckValue(string sValue)
        {
            bool hasil = false;
            int lCounter;

            for (lCounter = 0; lCounter < gridPajak.Rows.Count; lCounter++)
            {
                if (lCounter != gridPajak.CurrentRow.Index)
                {
                    if (gridPajak.Rows[lCounter].Cells[0].Value == null)
                    {
                        gridPajak.Rows[lCounter].Cells[0].Value = "";
                    }
                    if (gridPajak.Rows[lCounter].Cells[0].Value.ToString().Trim() == sValue.Trim().ToUpper())
                    {
                        hasil = false;
                        return hasil;
                    }
                }

            }
            hasil = true;
            return hasil;
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
                MessageBox.Show("Please Select Branch !", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            cmbSource.DataSource = _clsGlobal.ExecDT(strSQL);
            cmbSource.ValueMember = "gh_function_code";
            cmbSource.DisplayMember = "descr";

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
        public IndexGrid(DataGridViewCellEventArgs e)
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