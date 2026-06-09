using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;

namespace TIRASnDNet.INV.INVProductMaster
{
    public partial class frmINVProductMasterList : BaseListForm
    {
        // ============================================================
        // FIELDS
        // ============================================================
        LineHighlight __cursor = null;
        const int WidthList = 750;
        const int HeightList = 460;

        private clsGlobal _clsGlobal = new clsGlobal();

        string ProdID, Grade, Size;

        private string strSQL, paramMenuId;
        private bool _settingLookupValues = false;
        private const string LookupCodeField = "Code";
        private const string LookupDescriptionField = "Description";
        private const string LookupCategoryGroupField = "CategoryGroup";

        private const string BrandLookupQuery =
            "SELECT pl_prd_line_code AS [Code], pl_prd_line_desc AS [Description] " +
            "FROM IM_PRD_LINE WITH(NOLOCK) ORDER BY pl_prd_line_code";
        private const string ModelLookupQuery =
            "SELECT DISTINCT pm_prd_model_code AS [Code], pm_prd_model_desc AS [Description] " +
            "FROM IM_PRD_MODEL WITH(NOLOCK) ORDER BY pm_prd_model_code";
        private const string RawMaterialLookupQuery =
            "SELECT rmu_raw_mat_used_code AS [Code], rmu_raw_mat_used_desc AS [Description] " +
            "FROM IM_RM_USED WITH(NOLOCK) ORDER BY rmu_raw_mat_used_code";
        private const string ColorLookupQuery =
            "SELECT col_washing_collor_code AS [Code], col_washing_collor_desc AS [Description] " +
            "FROM IM_W_COLLOR WITH(NOLOCK) ORDER BY col_washing_collor_code";
        private const string CategoryLookupQuery =
            "SELECT pch_year AS [Code], pch_description AS [Description], pch_kelompok AS [CategoryGroup] " +
            "FROM IM_PRD_CATEGORY_HEADER WITH(NOLOCK) " +
            "WHERE pch_year LIKE '%' AND pch_kelompok LIKE '%' ORDER BY pch_year, pch_kelompok";

        private class ComboListItem
        {
            public string Id { get; private set; }
            public string Name { get; private set; }

            public ComboListItem(string id, string name)
            {
                Id = id ?? string.Empty;
                Name = name ?? string.Empty;
            }

            public override string ToString()
            {
                return Name;
            }
        }

        // ============================================================
        // CONSTRUCTOR
        // ============================================================
        public frmINVProductMasterList()
        {
            InitializeComponent();
        }

        // ============================================================
        // FORM LOAD
        // - DataSource untuk CBWarehose, CbNonStock, CbTech, CbAktif,
        //   CbOrder, CbSex sudah di-set di Designer (DataSource statis).
        // - CBProType di-load dari DB karena data dinamis.
        // ============================================================
        private void frmINVProductMasterList_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;

            // Inisialisasi nilai tanggal default - hari ini
            dateTimePicker1.DateTime = DateTime.Now;
            dateTimePicker2.DateTime = DateTime.Now;

            // Cursor highlight - tetap dipertahankan (legacy behavior)
            __cursor = new LineHighlight(this.layoutControl1);
            foreach (Control __child in this.layoutControl1.Controls)
            {
                if (__child.TabStop) __cursor.Add(__child);
            }
            txtBrandId.Focus();
            txtBrandId.Select();

            this.Text = clsLogin.MENUID + " - " + this.Text;
            paramMenuId = clsLogin.MENUID;
            AccessButton();

            // Load ComboBoxEdit static + Product Type dari database
            try
            {
                LoadDevExpressComboBoxes();
                LoadStaticSearchLookUpSources();
                ConfigureSearchLookUpPopupEvents();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }


        private void LoadDevExpressComboBoxes()
        {
            BindSimpleCombo(CBWarehose,
                new ComboListItem("All", "<ALL>"),
                new ComboListItem("Y", "Y-Yes"),
                new ComboListItem("N", "N-No"));

            BindSimpleCombo(CbNonStock,
                new ComboListItem("All", "<ALL>"),
                new ComboListItem("Y", "Y-Yes"),
                new ComboListItem("N", "N-No"));

            BindSimpleCombo(CbTech,
                new ComboListItem("All", "<ALL>"),
                new ComboListItem("Y", "Y-Yes"),
                new ComboListItem("N", "N-No"));

            BindSimpleCombo(CbAktif,
                new ComboListItem("All", "<ALL>"),
                new ComboListItem("Y", "Y-Yes"),
                new ComboListItem("N", "N-No"));

            BindSimpleCombo(CbOrder,
                new ComboListItem("0", "0-Standard"),
                new ComboListItem("1", "1-Group|Market Target"),
                new ComboListItem("2", "2-Market Target|Group"),
                new ComboListItem("3", "3-Group"),
                new ComboListItem("4", "4-Market Target"),
                new ComboListItem("5", "5-Raw Material|Market Target"),
                new ComboListItem("6", "6-Season"),
                new ComboListItem("7", "7-Collection"));

            BindSimpleCombo(CbSex,
                new ComboListItem("All", "<ALL>"),
                new ComboListItem("M", "M-Male"),
                new ComboListItem("F", "F-Female"),
                new ComboListItem("U", "U-Unisex"));

            LoadProductTypeCombo();
        }

        private void LoadProductTypeCombo()
        {
            strSQL = "SELECT * FROM (" +
                    "(select 0 AS gh_sequence_no, '' as gh_function_code, '<ALL>' as gh_function_desc) " +
                    "UNION " +
                    "(select gh_sequence_no, gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc " +
                    " from GS_GEN_HARDCODED " +
                    " where gh_sys = 'H' and gh_function_name = 'PRDTYPE'))VW order by gh_sequence_no";

            DataTable dt = _clsGlobal.ExecDT(strSQL);
            CBProType.Properties.Items.Clear();

            foreach (DataRow row in dt.Rows)
            {
                string id = row["gh_function_code"] == DBNull.Value ? string.Empty : row["gh_function_code"].ToString().Trim();
                string name = row["gh_function_desc"] == DBNull.Value ? string.Empty : row["gh_function_desc"].ToString().Trim();
                CBProType.Properties.Items.Add(new ComboListItem(id, name));
            }

            if (CBProType.Properties.Items.Count > 0)
                CBProType.SelectedIndex = 0;
        }

        private void BindSimpleCombo(ComboBoxEdit combo, params ComboListItem[] items)
        {
            if (combo == null) return;

            combo.Properties.Items.Clear();
            combo.Properties.Items.AddRange(items);
            combo.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;

            if (combo.Properties.Items.Count > 0)
                combo.SelectedIndex = 0;
        }

        // ============================================================
        // FILL GRID (Business Logic - tidak diubah)
        // ============================================================
        private void FillGridSP()
        {
            int CEK = 0;
            if (CHKTgl.Checked)
            {
                CEK = 1;
            }

            try
            {
                strSQL = "EXEC SP_INV_PRODUK_MASTER_LIST " +
                          "0, " +
                          SqlParam(txtProdMaster.Text) + ", " +
                          SqlParam(txtBrandId.Text) + ", " +
                          SqlParam(txtGroupId.Text) + ", " +
                          SqlParam(txtSubGId.Text) + ", " +
                          SqlParam(txtModelId.Text) + ", " +
                          SqlParam(txtRawMCodeId.Text) + ", " +
                          SqlParam(txtbarcode.Text) + ", " +
                          SqlParam(txtdesc.Text) + ", " +
                          SqlParam(txtColorId.Text) + ", " +
                          SqlParam(EditValueText(CBProType)) + ", " +
                          SqlParam(EditValueText(CBWarehose)) + ", " +
                          SqlParam(EditValueText(CbNonStock)) + ", " +
                          SqlParam(EditValueText(CbTech)) + ", " +
                          SqlParam(EditValueText(CbAktif)) + ", " +
                          SqlParam(txtTrend.Text) + ", " +
                          SqlParam(txtmarket.Text) + ", " +
                          SqlParam(txtseason.Text) + ", " +
                          SqlParam(txtKategori.Text) + ", " +
                          CEK.ToString() + ", " +
                          SqlParam(Convert.ToDateTime(dateTimePicker1.DateTime).ToString("yyyyMMdd")) + ", " +
                          SqlParam(Convert.ToDateTime(dateTimePicker2.DateTime).ToString("yyyyMMdd")) + ", " +
                          SqlParam(EditValueText(CbSex));

                gridControl1.DataSource = _clsGlobal.ExecDT(strSQL);
                strSQL = "";

                if (gridView1.RowCount > 0)
                {
                    gridView1.FocusedRowHandle = 0;
                    int rowHandle = gridView1.FocusedRowHandle;
                    ProdID = gridView1.GetRowCellValue(rowHandle, "prm_prd_master_code").ToString().Trim();
                    Grade = gridView1.GetRowCellValue(rowHandle, "prm_grade").ToString().Trim();
                    Size = gridView1.GetRowCellValue(rowHandle, "prm_prd_size").ToString().Trim();
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // GRID EVENTS
        // ============================================================
        void gridView1_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (e.RowHandle > -1)
            {
                ProdID = gridView1.GetRowCellValue(e.RowHandle, "prm_prd_master_code").ToString().Trim();
                Grade = gridView1.GetRowCellValue(e.RowHandle, "prm_grade").ToString().Trim();
                Size = gridView1.GetRowCellValue(e.RowHandle, "prm_prd_size").ToString().Trim();
            }
        }

        private void gridView1_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        private void gridView1_DoubleClick(object sender, EventArgs e)
        {
            if (gridView1.RowCount > 0)
            {
                int rowindex = gridView1.FocusedRowHandle;
                if (rowindex < 0) return;

                clsGlobal.MODE_TRX = 2;
                frmINVProductMasterEditor frm = new frmINVProductMasterEditor();
                frm.PrdBrandCode = gridView1.GetRowCellValue(rowindex, "prm_prd_master_code").ToString().Trim();
                frm.PrdGrade = gridView1.GetRowCellValue(rowindex, "prm_grade").ToString().Trim();
                frm.ShowDialog();
            }
        }

        // ============================================================
        // CHECKBOX EVENTS
        // ============================================================
        private void CHKTgl_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = CHKTgl.Checked;
            dateTimePicker2.Enabled = CHKTgl.Checked;
        }

        private void CBKLong_CheckedChanged(object sender, EventArgs e)
        {
            CBKLong.Text = CBKLong.Checked ? "Short Type" : "Long Type";
        }

        // ============================================================
        // BUTTON EXEC
        // ============================================================
        private void btnExec_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            FillGridSP();
            Cursor.Current = Cursors.Default;
        }

        // ============================================================
        // BASE FORM OVERRIDES
        // ============================================================
        protected override void OnEdit()
        {
            if (gridView1.RowCount > 0)
            {
                int rowindex = gridView1.FocusedRowHandle;

                clsGlobal.MODE_TRX = 2;
                frmINVProductMasterEditor frm = new frmINVProductMasterEditor();
                frm.PrdBrandCode = gridView1.GetRowCellValue(rowindex, "prm_prd_master_code").ToString().Trim();
                frm.PrdGrade = gridView1.GetRowCellValue(rowindex, "prm_grade").ToString().Trim();
                frm.ShowDialog();
            }
        }

        protected override void OnNew()
        {
            clsGlobal.MODE_TRX = 1;
            frmINVProductMasterEditor frm = new frmINVProductMasterEditor();
            frm.ShowDialog();
        }

        protected override void OnDelete()
        {
            try
            {
                if (gridView1.RowCount > 0)
                {
                    DialogResult dr;
                    dr = XtraMessageBox.Show(_clsGlobal.ApplMessage(10001), clsGlobal.APP_MSG_CAPTION_DELETE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.OK)
                    {
                        strSQL = "EXEC SP_INV_PRODUK_MASTER '3', " +
                            "'" + ProdID + "', " +
                            "'" + Grade + "', " +
                            "'" + Size + "', " +
                            "'', '', '', '', '','','','','','', " +
                            "'', '', '', '', '','','','','',''," +
                            "'', '', '', '', '','','','','','', " +
                            "'" + clsLogin.USERID + "','', '', '', '', '','', " +
                            "'', '', '', '', '','','','','','', " +
                            "'', '', '', '', '', '', '', '', '' ";

                        _clsGlobal.BeginTrans();
                        _clsGlobal.ExecuteTrans(strSQL);
                        _clsGlobal.CommitTrans();

                        FillGridSP();
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void OnPrint()
        {
            int CEK = 0;
            if (CHKTgl.Checked)
            {
                CEK = 1;
            }
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                DataTable dt = _clsGlobal.GetEntityBranch();
                string paramkode;

                frmReport frm = new frmReport();
                frm.ReportName("INVProductMasterList.rpt");

                strSQL = "SELECT * FROM VW_INV_PRODUK_MASTER where 1=1 ";
                if (txtBrandId.Text != "")
                    strSQL += " AND prm_prd_line_code='" + txtBrandId.Text + "'";
                if (txtGroupId.Text != "")
                    strSQL += " AND prm_prd_group_code ='" + txtGroupId.Text + "'";
                if (txtSubGId.Text != "")
                    strSQL += " AND prm_prd_sgroup_code ='" + txtSubGId.Text + "'";
                if (txtModelId.Text != "")
                    strSQL += " AND prm_prd_model_code='" + txtModelId.Text + "'";
                if (txtRawMCodeId.Text != "")
                    strSQL += " AND prm_raw_mat_used_code='" + txtRawMCodeId.Text + "'";
                if (txtTrend.Text != "")
                    strSQL += " AND prm_trend like '%" + txtTrend.Text + "%'";
                if (txtTrenddesc.Text != "")
                    strSQL += " AND desc_trend like '%" + txtTrenddesc.Text + "%'";
                if (txtmarket.Text != "")
                    strSQL += " AND prm_target like '%" + txtmarket.Text + "%'";
                if (txtmarketDesc.Text != "")
                    strSQL += " AND desc_target like '%" + txtmarketDesc.Text + "%'";
                if (txtProdMaster.Text != "")
                    strSQL += " AND prm_prd_master_code like '%" + txtProdMaster.Text + "%'";
                if (txtseason.Text != "")
                    strSQL += " AND prm_season like '%" + txtseason.Text + "%'";
                if (txtseasonDesc.Text != "")
                    strSQL += " AND desc_season like '%" + txtseasonDesc.Text + "%'";
                if (txtKategori.Text != "")
                    strSQL += " AND prm_mfg_periode_mm_yy like '%" + txtKategori.Text + "%'";
                if (txtkategoriTo.Text != "")
                    strSQL += " AND so_kelompok like '%" + txtkategoriTo.Text + "%'";
                if (txtKategoriDesc.Text != "")
                    strSQL += " AND pch_description like '%" + txtKategoriDesc.Text + "%'";
                if (txtdesc.Text != "")
                    strSQL += " AND prm_prd_desc like '%" + txtdesc.Text + "%'";
                if (txtColorId.Text != "")
                    strSQL += " AND prm_washing_collor_code ='" + txtColorId.Text + "'";
                if (EditValueText(CBProType) != string.Empty)
                    strSQL += "and prm_prd_type like '%" + EditValueText(CBProType) + "%' ";
                if (CBWarehose.SelectedIndex > 0)
                    strSQL += "and prm_wh_control_flag = '" + EditValueText(CBWarehose) + "' ";
                if (CbNonStock.SelectedIndex > 0)
                    strSQL += "and prm_non_stock_flag = '" + EditValueText(CbNonStock) + "' ";
                if (CbTech.SelectedIndex > 0)
                    strSQL += "and prm_tech_constrain_flag = '" + EditValueText(CbTech) + "' ";
                if (CbAktif.SelectedIndex > 0)
                    strSQL += "and prm_active_flag = '" + EditValueText(CbAktif) + "' ";
                if (CbSex.SelectedIndex > 0)
                    strSQL += "and prm_sex = '" + EditValueText(CbSex) + "' ";
                if (CHKTgl.Checked)
                    strSQL += "AND CONVERT(VARCHAR(8),prm_creation_date,112) between '" + _clsGlobal.DateToDB(dateTimePicker1.Text) + "' AND '" + _clsGlobal.DateToDB(dateTimePicker2.Text) + "' ";
                if (CBKLong.Checked)
                    strSQL += "AND prm_active_flag ='Y'";

                frm.QueryString(strSQL);
                frm.Parameter("ProgId", paramMenuId);
                frm.Parameter("Group", "" + (txtGroupId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtGroupId.Text.Trim()));
                frm.Parameter("SGroup", "" + (txtSubGId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtSubGId.Text.Trim()));
                frm.Parameter("Category", "" + (txtKategori.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtKategori.Text.Trim()));
                frm.Parameter("Style", "" + (txtTrend.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtTrend.Text.Trim()));
                frm.Parameter("LevelOrder", "" + EditValueText(CbOrder));
                frm.Parameter("Sex", "" + EditValueText(CbSex));
                frm.Parameter("Market", "" + (txtmarket.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtmarket.Text.Trim()));
                frm.Parameter("Color", "" + (txtColorId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtColorId.Text.Trim()));
                frm.Parameter("RawMat", "" + (txtRawMCodeId.Text.Trim().IsNullOrEmptyOrWhiteSpace() ? "All" : txtRawMCodeId.Text.Trim()));
                frm.Parameter("LevelDetail", "" + EditValueText(CbAktif));
                frm.Show();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // SEARCH LOOKUP HANDLERS
        // ============================================================
        private void LoadStaticSearchLookUpSources()
        {
            BindSearchLookUp(btnPopUpBrandIdE, _clsGlobal.ExecDT(BrandLookupQuery));
            BindSearchLookUp(btnPopUpModel, _clsGlobal.ExecDT(ModelLookupQuery));
            BindSearchLookUp(btnPopUpRawM, _clsGlobal.ExecDT(RawMaterialLookupQuery));
            BindSearchLookUp(btnPopUpColor, _clsGlobal.ExecDT(ColorLookupQuery));
            BindSearchLookUp(btnTrend, _clsGlobal.ExecDT(GetProductParameterLookupQuery("TREND")));
            BindSearchLookUp(btnMarket, _clsGlobal.ExecDT(GetProductParameterLookupQuery("TARGET")));
            BindSearchLookUp(btnSeason, _clsGlobal.ExecDT(GetProductParameterLookupQuery("SEASON")));
            BindSearchLookUp(btnKategori, _clsGlobal.ExecDT(CategoryLookupQuery));
        }

        private void BindSearchLookUp(SearchLookUpEdit lookup, DataTable dataSource)
        {
            if (lookup == null) return;

            if (lookup.Properties.View == null)
                lookup.Properties.View = new GridView();

            lookup.Properties.DataSource = dataSource;
            lookup.Properties.ValueMember = LookupCodeField;
            lookup.Properties.DisplayMember = LookupCodeField;
            lookup.Properties.NullText = string.Empty;
            lookup.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            lookup.Properties.View.OptionsBehavior.Editable = false;
            lookup.Properties.View.OptionsView.ColumnAutoWidth = false;
            lookup.Properties.View.OptionsView.ShowGroupPanel = false;
            lookup.Properties.PopulateViewColumns();

            ConfigureLookupColumn(lookup, LookupCodeField, "Code", 0, 120, true);
            ConfigureLookupColumn(lookup, LookupDescriptionField, "Description", 1, 360, true);
            ConfigureLookupColumn(lookup, LookupCategoryGroupField, "Category Group", 2, 120, false);
        }

        private void ConfigureLookupColumn(SearchLookUpEdit lookup, string fieldName, string caption, int visibleIndex, int width, bool visible)
        {
            if (lookup == null || lookup.Properties.View == null) return;

            DevExpress.XtraGrid.Columns.GridColumn column = lookup.Properties.View.Columns[fieldName];
            if (column == null) return;

            column.Caption = caption;
            column.Visible = visible;
            if (visible)
                column.VisibleIndex = visibleIndex;
            column.Width = width;
        }

        private string GetProductParameterLookupQuery(string functionName)
        {
            return "SELECT pp_function_code AS [Code], pp_function_desc AS [Description] " +
                   "FROM IM_PRD_PARAMETER WITH(NOLOCK) " +
                   "WHERE pp_function_name = '" + SqlText(functionName) + "' " +
                   "ORDER BY pp_sequence_no";
        }

        private string GetGroupLookupQuery()
        {
            return "SELECT DISTINCT pg_prd_group_code AS [Code], pg_prd_group_desc AS [Description] " +
                   "FROM IM_PRD_GROUP WITH(NOLOCK) " +
                   "WHERE pg_prd_line_code LIKE '" + SqlText(txtBrandId.Text.Trim()) + "%' " +
                   "ORDER BY pg_prd_group_code";
        }

        private string GetSubGroupLookupQuery()
        {
            return "SELECT DISTINCT psg_prd_sgroup_code AS [Code], psg_prd_sgroup_desc AS [Description] " +
                   "FROM IM_PRD_SGROUP WITH(NOLOCK) " +
                   "WHERE psg_prd_line LIKE '" + SqlText(txtBrandId.Text.Trim()) + "%' " +
                   "AND psg_prd_group_code LIKE '" + SqlText(txtGroupId.Text.Trim()) + "%' " +
                   "ORDER BY psg_prd_sgroup_code";
        }

        private string SqlText(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }

        private string EditValueText(BaseEdit editor)
        {
            if (editor == null || editor.EditValue == null || editor.EditValue == DBNull.Value)
                return string.Empty;

            ComboListItem comboItem = editor.EditValue as ComboListItem;
            if (comboItem != null)
                return comboItem.Id.Trim();

            return editor.EditValue.ToString().Trim();
        }

        private DataRow GetSelectedLookupRow(SearchLookUpEdit lookup)
        {
            if (lookup == null || lookup.EditValue == null || lookup.EditValue == DBNull.Value) return null;

            object rowObject = lookup.Properties.GetRowByKeyValue(lookup.EditValue);
            DataRowView rowView = rowObject as DataRowView;
            if (rowView != null) return rowView.Row;

            return rowObject as DataRow;
        }

        private string LookupRowText(DataRow row, string fieldName)
        {
            if (row == null || !row.Table.Columns.Contains(fieldName) || row[fieldName] == DBNull.Value)
                return string.Empty;

            return row[fieldName].ToString().Trim();
        }

        private bool SetSelectedLookupText(SearchLookUpEdit lookup, TextEdit codeEdit, TextEdit descriptionEdit)
        {
            if (_settingLookupValues) return false;

            DataRow row = GetSelectedLookupRow(lookup);
            if (row == null) return false;

            _settingLookupValues = true;
            try
            {
                codeEdit.Text = LookupRowText(row, LookupCodeField);
                descriptionEdit.Text = LookupRowText(row, LookupDescriptionField);
            }
            finally
            {
                _settingLookupValues = false;
            }

            return true;
        }
        private void SetTextEditPair(TextEdit codeEdit, TextEdit descriptionEdit, string code, string description)
        {
            _settingLookupValues = true;
            try
            {
                codeEdit.Text = code;
                descriptionEdit.Text = description;
            }
            finally
            {
                _settingLookupValues = false;
            }
        }
        private void SearchLookUpEdit_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            // Jangan panggil ShowPopup() manual.
            // Popup dibuka oleh DevExpress melalui ShowDropDown.SingleClick.
            // Data lookup di-bind pada event QueryPopUp.
        }

        private void ConfigureSearchLookUpPopupEvents()
        {
            SearchLookUpEdit[] lookups =
            {
                btnPopUpBrandIdE,
                btnPopUpGroup,
                btnPopUpSubGr,
                btnPopUpModel,
                btnPopUpRawM,
                btnPopUpColor,
                btnTrend,
                btnMarket,
                btnSeason,
                btnKategori
            };

            foreach (SearchLookUpEdit lookup in lookups)
            {
                if (lookup == null) continue;

                // Hindari popup kebuka 2x dari ButtonClick + ShowPopup manual.
                lookup.ButtonClick -= SearchLookUpEdit_ButtonClick;

                // Hindari datasource di-reset saat event Popup sedang opening.
                lookup.Popup -= btnPopUpBrandIdE_Popup;
                lookup.Popup -= btnPopUpGroup_Popup;
                lookup.Popup -= btnPopUpSubGr_Popup;
                lookup.Popup -= btnPopUpModel_Popup;
                lookup.Popup -= btnPopUpRawM_Popup;
                lookup.Popup -= btnPopUpColor_Popup;
                lookup.Popup -= btnTrend_Popup;
                lookup.Popup -= btnMarket_Popup;
                lookup.Popup -= btnSeason_Popup;
                lookup.Popup -= btnKategori_Popup;

                lookup.QueryPopUp -= SearchLookUpEdit_QueryPopUp;
                lookup.QueryPopUp += SearchLookUpEdit_QueryPopUp;

                lookup.Properties.ShowDropDown = DevExpress.XtraEditors.Controls.ShowDropDown.SingleClick;
            }
        }

        private void SearchLookUpEdit_QueryPopUp(object sender, CancelEventArgs e)
        {
            SearchLookUpEdit editor = sender as SearchLookUpEdit;
            if (editor == null) return;

            BindLookupBeforePopup(editor);
        }

        private void BindLookupBeforePopup(SearchLookUpEdit editor)
        {
            if (editor == btnPopUpBrandIdE)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(BrandLookupQuery));
            else if (editor == btnPopUpGroup)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(GetGroupLookupQuery()));
            else if (editor == btnPopUpSubGr)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(GetSubGroupLookupQuery()));
            else if (editor == btnPopUpModel)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(ModelLookupQuery));
            else if (editor == btnPopUpRawM)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(RawMaterialLookupQuery));
            else if (editor == btnPopUpColor)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(ColorLookupQuery));
            else if (editor == btnTrend)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(GetProductParameterLookupQuery("TREND")));
            else if (editor == btnMarket)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(GetProductParameterLookupQuery("TARGET")));
            else if (editor == btnSeason)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(GetProductParameterLookupQuery("SEASON")));
            else if (editor == btnKategori)
                BindSearchLookUp(editor, _clsGlobal.ExecDT(CategoryLookupQuery));
        }



        private void btnPopUpBrandIdE_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private string SqlParam(string value)
        {
            value = value == null ? string.Empty : value.Trim();

            if (value == string.Empty)
                return "NULL";

            return "'" + value.Replace("'", "''") + "'";
        }

        private void btnPopUpGroup_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnPopUpSubGr_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnPopUpModel_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnPopUpRawM_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnPopUpColor_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnTrend_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnMarket_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnSeason_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnKategori_Popup(object sender, EventArgs e)
        {
            // No-op.
            // Binding lookup dipindahkan ke QueryPopUp supaya popup tidak langsung tertutup.
        }

        private void btnPopUpBrandIdE_EditValueChanged(object sender, EventArgs e)
        {
            if (SetSelectedLookupText(btnPopUpBrandIdE, txtBrandId, txtBrandDesc))
            {
                SetTextEditPair(txtGroupId, txtGroupDesc, string.Empty, string.Empty);
                SetTextEditPair(txtSubGId, txtSubGDesc, string.Empty, string.Empty);
            }
        }

        private void btnPopUpGroup_EditValueChanged(object sender, EventArgs e)
        {
            if (SetSelectedLookupText(btnPopUpGroup, txtGroupId, txtGroupDesc))
                SetTextEditPair(txtSubGId, txtSubGDesc, string.Empty, string.Empty);
        }

        private void btnPopUpSubGr_EditValueChanged(object sender, EventArgs e)
        {
            SetSelectedLookupText(btnPopUpSubGr, txtSubGId, txtSubGDesc);
        }

        private void btnPopUpModel_EditValueChanged(object sender, EventArgs e)
        {
            SetSelectedLookupText(btnPopUpModel, txtModelId, txtModelDesc);
        }

        private void btnPopUpRawM_EditValueChanged(object sender, EventArgs e)
        {
            SetSelectedLookupText(btnPopUpRawM, txtRawMCodeId, txtRawMCodeDesc);
        }

        private void btnPopUpColor_EditValueChanged(object sender, EventArgs e)
        {
            SetSelectedLookupText(btnPopUpColor, txtColorId, txtColorDesc);
        }

        private void btnTrend_EditValueChanged(object sender, EventArgs e)
        {
            SetSelectedLookupText(btnTrend, txtTrend, txtTrenddesc);
        }

        private void btnMarket_EditValueChanged(object sender, EventArgs e)
        {
            SetSelectedLookupText(btnMarket, txtmarket, txtmarketDesc);
        }

        private void btnSeason_EditValueChanged(object sender, EventArgs e)
        {
            SetSelectedLookupText(btnSeason, txtseason, txtseasonDesc);
        }

        private void btnKategori_EditValueChanged(object sender, EventArgs e)
        {
            DataRow row = GetSelectedLookupRow(btnKategori);
            if (row == null) return;

            _settingLookupValues = true;
            try
            {
                txtKategori.Text = LookupRowText(row, LookupCodeField);
                txtKategoriDesc.Text = LookupRowText(row, LookupDescriptionField);
                txtkategoriTo.Text = LookupRowText(row, LookupCategoryGroupField);
            }
            finally
            {
                _settingLookupValues = false;
            }
        }

        private string GetFirstLookupValue(string query, string fieldName)
        {
            DataTable dt = _clsGlobal.ExecDT(query);
            if (dt.Rows.Count == 0 || !dt.Columns.Contains(fieldName) || dt.Rows[0][fieldName] == DBNull.Value)
                return string.Empty;

            return dt.Rows[0][fieldName].ToString().Trim();
        }

        private void txtBrandId_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                if (txtBrandId.Text.Trim() == string.Empty)
                {
                    txtBrandDesc.Text = string.Empty;
                    return;
                }

                strSQL = "SELECT pl_prd_line_desc AS [Description] " +
                         "FROM IM_PRD_LINE WITH(NOLOCK) " +
                         "WHERE pl_prd_line_code = '" + SqlText(txtBrandId.Text.Trim()) + "'";
                txtBrandDesc.Text = GetFirstLookupValue(strSQL, LookupDescriptionField);
            }
            catch (Exception ex) { }
        }

        private void txtGroupId_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                if (txtGroupId.Text.Trim() == string.Empty)
                {
                    txtGroupDesc.Text = string.Empty;
                    return;
                }

                strSQL = "SELECT DISTINCT pg_prd_group_desc AS [Description] " +
                         "FROM IM_PRD_GROUP WITH(NOLOCK) " +
                         "WHERE pg_prd_line_code LIKE '" + SqlText(txtBrandId.Text.Trim()) + "%' " +
                         "AND pg_prd_group_code = '" + SqlText(txtGroupId.Text.Trim()) + "'";
                txtGroupDesc.Text = GetFirstLookupValue(strSQL, LookupDescriptionField);
            }
            catch (Exception ex) { }
        }

        private void txtSubGId_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                if (txtSubGId.Text.Trim() == string.Empty)
                {
                    txtSubGDesc.Text = string.Empty;
                    return;
                }

                strSQL = "SELECT DISTINCT psg_prd_sgroup_desc AS [Description] " +
                         "FROM IM_PRD_SGROUP WITH(NOLOCK) " +
                         "WHERE psg_prd_line LIKE '" + SqlText(txtBrandId.Text.Trim()) + "%' " +
                         "AND psg_prd_group_code LIKE '" + SqlText(txtGroupId.Text.Trim()) + "%' " +
                         "AND psg_prd_sgroup_code = '" + SqlText(txtSubGId.Text.Trim()) + "'";
                txtSubGDesc.Text = GetFirstLookupValue(strSQL, LookupDescriptionField);
            }
            catch (Exception ex) { }
        }

        private void txtModelId_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                if (txtModelId.Text == "")
                {
                    txtModelDesc.Text = "";
                }
                else
                {
                    strSQL = "SELECT pm_prd_model_desc AS [Description] " +
                             "FROM IM_PRD_MODEL WITH(NOLOCK) " +
                             "WHERE pm_prd_model_code = '" + SqlText(txtModelId.Text.Trim()) + "'";
                    txtModelDesc.Text = GetFirstLookupValue(strSQL, LookupDescriptionField);
                }
            }
            catch (Exception ex) { }
        }

        private void txtRawMCodeId_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                if (txtRawMCodeId.Text.Trim() == string.Empty)
                {
                    txtRawMCodeDesc.Text = string.Empty;
                    return;
                }

                strSQL = "SELECT rmu_raw_mat_used_desc AS [Description] " +
                         "FROM IM_RM_USED WITH(NOLOCK) " +
                         "WHERE rmu_raw_mat_used_code = '" + SqlText(txtRawMCodeId.Text.Trim()) + "'";
                txtRawMCodeDesc.Text = GetFirstLookupValue(strSQL, LookupDescriptionField);
            }
            catch (Exception ex) { }
        }

        // ============================================================
        // TEXT CHANGED HANDLERS - KOLOM 2
        // ============================================================
        private void txtColorId_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                if (txtColorId.Text.Trim() == string.Empty)
                {
                    txtColorDesc.Text = string.Empty;
                    return;
                }

                strSQL = "SELECT col_washing_collor_desc AS [Description] " +
                         "FROM IM_W_COLLOR WITH(NOLOCK) " +
                         "WHERE col_washing_collor_code = '" + SqlText(txtColorId.Text.Trim()) + "'";
                txtColorDesc.Text = GetFirstLookupValue(strSQL, LookupDescriptionField);
            }
            catch (Exception ex) { }
        }

        // ============================================================
        // TEXT CHANGED HANDLERS - KOLOM 3 (Trend/Market/Kategori)
        // ============================================================
        private void txtTrend_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                string description = string.Empty;
                if (txtTrend.Text.Trim() != string.Empty)
                {
                    strSQL = "SELECT pp_function_desc AS [Description] " +
                             "FROM IM_PRD_PARAMETER WITH(NOLOCK) " +
                             "WHERE pp_function_name = 'TREND' " +
                             "AND pp_function_code = '" + SqlText(txtTrend.Text.Trim()) + "'";
                    description = GetFirstLookupValue(strSQL, LookupDescriptionField);
                }

                SetTextEditPair(txtTrend, txtTrenddesc, txtTrend.Text, description);
            }
            catch (Exception ex) { }
        }

        private void txtTrenddesc_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (txtTrenddesc.Text.ToString() == "")
                txtTrend.Text = string.Empty;
        }

        private void txtmarket_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                string description = string.Empty;
                if (txtmarket.Text.Trim() != string.Empty)
                {
                    strSQL = "SELECT pp_function_desc AS [Description] " +
                             "FROM IM_PRD_PARAMETER WITH(NOLOCK) " +
                             "WHERE pp_function_name = 'TARGET' " +
                             "AND pp_function_code = '" + SqlText(txtmarket.Text.Trim()) + "'";
                    description = GetFirstLookupValue(strSQL, LookupDescriptionField);
                }

                SetTextEditPair(txtmarket, txtmarketDesc, txtmarket.Text, description);
            }
            catch (Exception ex) { }
        }

        private void txtmarketDesc_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (txtmarketDesc.Text.ToString() == "")
                txtmarket.Text = string.Empty;
        }

        private void txtseason_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                string description = string.Empty;
                if (txtseason.Text.Trim() != string.Empty)
                {
                    strSQL = "SELECT pp_function_desc AS [Description] " +
                             "FROM IM_PRD_PARAMETER WITH(NOLOCK) " +
                             "WHERE pp_function_name = 'SEASON' " +
                             "AND pp_function_code = '" + SqlText(txtseason.Text.Trim()) + "'";
                    description = GetFirstLookupValue(strSQL, LookupDescriptionField);
                }

                SetTextEditPair(txtseason, txtseasonDesc, txtseason.Text, description);
            }
            catch (Exception ex) { }
        }

        private void txtKategori_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            try
            {
                if (txtKategori.Text.Trim() == string.Empty)
                {
                    _settingLookupValues = true;
                    try
                    {
                        txtKategoriDesc.Text = string.Empty;
                        txtkategoriTo.Text = string.Empty;
                    }
                    finally
                    {
                        _settingLookupValues = false;
                    }
                    return;
                }

                strSQL = "SELECT TOP 1 pch_description AS [Description], pch_kelompok AS [CategoryGroup] " +
                         "FROM IM_PRD_CATEGORY_HEADER WITH(NOLOCK) " +
                         "WHERE pch_year = '" + SqlText(txtKategori.Text.Trim()) + "' " +
                         "ORDER BY pch_year, pch_kelompok";
                DataTable dt = _clsGlobal.ExecDT(strSQL);

                _settingLookupValues = true;
                try
                {
                    if (dt.Rows.Count > 0)
                    {
                        txtKategoriDesc.Text = dt.Rows[0][LookupDescriptionField].ToString().Trim();
                        txtkategoriTo.Text = dt.Rows[0][LookupCategoryGroupField].ToString().Trim();
                    }
                    else
                    {
                        txtKategoriDesc.Text = string.Empty;
                        txtkategoriTo.Text = string.Empty;
                    }
                }
                finally
                {
                    _settingLookupValues = false;
                }
            }
            catch (Exception ex) { }
        }

        private void txtKategoriDesc_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (txtKategoriDesc.Text.ToString() == "")
            {
                txtkategoriTo.Text = string.Empty;
                txtKategori.Text = string.Empty;
            }
        }

        private void txtkategoriTo_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (txtkategoriTo.Text.ToString() == "")
            {
                txtKategoriDesc.Text = string.Empty;
                txtKategori.Text = string.Empty;
            }
        }
    }
}
