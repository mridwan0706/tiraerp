using FluentFTP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Net;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.DXErrorProvider;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraTab;


namespace TIRASnDNet.INV.INVProductMaster
{
    // DEVEXPRESS_NATIVE_LAYOUT_CONVERSION: Panel/Group/Tab containers converted to DevExpress native controls.
    public partial class frmINVProductMasterEditor : XtraForm
    {

        private FtpClient FTP1 = null;
        frmPopUp frm;
        clsGlobal _clsGlobal = new clsGlobal();
        BackgroundWorker __browseWorker = null;
        LoadingCircle __circle = null;

        bool __apparelSys = false;
        private string strSQL, paramMenuId, _prdBrandCode, _Grade, Sizegroupcode;

        private bool _settingLookupValues = false;
        private DXValidationProvider dxValidator = null;
        private bool _isRefreshingValidationState = false;
        private readonly List<Control> _registeredValidationControls = new List<Control>();
        private readonly Dictionary<Control, XtraTabPage> _validationTabMap = new Dictionary<Control, XtraTabPage>();
        private readonly Dictionary<Control, ValidationRule> _validationRuleMap = new Dictionary<Control, ValidationRule>();
        private readonly Dictionary<Control, Color> _validationOriginalBackColorMap = new Dictionary<Control, Color>();
        private readonly Dictionary<Control, Color> _validationOriginalBaseEditBackColorMap = new Dictionary<Control, Color>();
        private readonly HashSet<string> _invalidSizeGridCells = new HashSet<string>();
        private readonly HashSet<string> _invalidAlternateUomGridCells = new HashSet<string>();
        private const string LookupCodeField = "Code";
        private const string LookupDescriptionField = "Description";
        private const string LookupVendorCode2Field = "Code2";

        private const string BrandLookupQuery = "SELECT pl_prd_line_code AS [Code], pl_prd_line_desc AS [Description] FROM IM_PRD_LINE WITH(NOLOCK) ORDER BY pl_prd_line_code";
        private const string ModelLookupQuery = "SELECT DISTINCT pm_prd_model_code AS [Code], pm_prd_model_desc AS [Description] FROM IM_PRD_MODEL WITH(NOLOCK) ORDER BY pm_prd_model_code";
        private const string RawMaterialLookupQuery = "SELECT rmu_raw_mat_used_code AS [Code], rmu_raw_mat_used_desc AS [Description] FROM IM_RM_USED WITH(NOLOCK) ORDER BY rmu_raw_mat_used_code";
        private const string ColorLookupQuery = "SELECT col_washing_collor_code AS [Code], col_washing_collor_desc AS [Description] FROM IM_W_COLLOR WITH(NOLOCK) ORDER BY col_washing_collor_code";
        private const string ProcessCodeLookupQuery = "SELECT gh_function_code AS [Code], gh_function_desc AS [Description] FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_sys = 'H' AND gh_function_name = 'PROCESSCODE' ORDER BY gh_sequence_no";
        private const string VendorLookupQuery = "SELECT pvm_vendor_code1 AS [Code], pvm_vendor_code2 AS [Code2], pvm_vendor_namekey AS [Description] FROM PO_VENDOR_MASTER WITH(NOLOCK) ORDER BY pvm_vendor_code1, pvm_vendor_code2";
        private const string PrincipalLookupQuery = "SELECT pri_principal AS [Code], pri_principal_desc AS [Description] FROM PO_PRINCIPAL WITH(NOLOCK) ORDER BY pri_principal";
        private const string TaxCodeLookupQuery = "SELECT t_tax_id AS [Code], t_tax AS [Description] FROM IM_TAX WITH(NOLOCK) ORDER BY t_tax_id";
        private const string ValuationClassLookupQuery = "SELECT mvc_code AS [Code], mvc_description AS [Description] FROM PO_VALUATION_CLASS WITH(NOLOCK) ORDER BY mvc_code";
        private const string MaterialTypeLookupQuery = "SELECT mt_type_id AS [Code], mt_type_desc AS [Description] FROM IM_PRD_TYPE WITH(NOLOCK) ORDER BY mt_type_id";

        #region -- query and object
        const string __BRAND = "SELECT pl_prd_line_code AS [Kode], pl_prd_line_desc AS [Nama] FROM IM_PRD_LINE WITH(NOLOCK) ORDER BY pl_prd_line_code";

        // pg_prd_line_code = '" + txtLineBrndCd1.Text + "' 
        // ORDER BY pg_prd_group_code";
        const string __GROUP_F = "SELECT DISTINCT pg_prd_group_code AS [Kode], pg_prd_group_desc AS [Nama] FROM IM_PRD_GROUP WITH(NOLOCK) WHERE 1=1 {0}";

        // psg_prd_line = '" + txtLineBrndCd1.Text + "' 
        // AND psg_prd_group_code = '" + txtGroupCd1.Text + "' 
        // ORDER BY psg_prd_sgroup_code";
        const string __SUBGROUP_F = "SELECT DISTINCT psg_prd_sgroup_code AS [Kode], psg_prd_sgroup_desc AS [Nama] FROM IM_PRD_SGROUP WITH(NOLOCK) WHERE 1=1 {0}";
        const string __MODEL = "SELECT DISTINCT pm_prd_model_code AS [Kode],pm_prd_model_desc AS [Nama] FROM IM_PRD_MODEL WITH(NOLOCK) ORDER BY pm_prd_model_code";
        const string __RMCODE = "SELECT rmu_raw_mat_used_code AS [Kode], rmu_raw_mat_used_desc AS [Nama] FROM IM_RM_USED WITH(NOLOCK) ORDER BY rmu_raw_mat_used_code";
        const string __COLOR = "SELECT col_washing_collor_code AS [Kode], col_washing_collor_desc AS [Nama] FROM IM_W_COLLOR WITH(NOLOCK) ORDER BY col_washing_collor_code";
        const string __PROCESSCODE = "SELECT gh_function_code AS [Kode],gh_function_desc AS [Nama] FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_sys ='H' and gh_function_name ='PROCESSCODE'";
        const string __VENDOR = "SELECT pvm_vendor_code1 AS ID1, pvm_vendor_code2 AS ID2, pvm_vendor_namekey AS [Nama] FROM PO_VENDOR_MASTER WITH(NOLOCK)";
        const string __PRINCIPAL = "SELECT pri_principal AS [Kode], pri_principal_desc AS [Nama] FROM PO_PRINCIPAL WITH(NOLOCK)";
        const string __TAXCODE = "SELECT t_tax_id AS [Kode], t_tax AS [Nama] FROM IM_TAX WITH(NOLOCK)";
        const string __UOM_LOOKUP = "SELECT umc_uom_code AS [Kode], umc_uom_description AS [Nama] FROM IM_UNIT_MEASURE_CODES WITH(NOLOCK) ORDER BY umc_index";
        const string __PRD_UOM_CONVERSION_TABLE = @"
IF OBJECT_ID('dbo.IM_PRD_UOM_CONVERSION', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.IM_PRD_UOM_CONVERSION
    (
        prdu_prd_code        VARCHAR(15)     NOT NULL,
        prdu_alt_uom         VARCHAR(10)     NOT NULL,
        prdu_base_uom        VARCHAR(10)     NOT NULL,
        prdu_numerator       DECIMAL(18, 4)  NOT NULL CONSTRAINT DF_IM_PRD_UOM_CONVERSION_NUMERATOR DEFAULT (1),
        prdu_denominator     DECIMAL(18, 4)  NOT NULL CONSTRAINT DF_IM_PRD_UOM_CONVERSION_DENOMINATOR DEFAULT (1),
        prdu_is_base         CHAR(1)         NOT NULL CONSTRAINT DF_IM_PRD_UOM_CONVERSION_IS_BASE DEFAULT ('N'),
        prdu_is_active       CHAR(1)         NOT NULL CONSTRAINT DF_IM_PRD_UOM_CONVERSION_IS_ACTIVE DEFAULT ('Y'),
        prdu_created_by      VARCHAR(50)     NOT NULL,
        prdu_created_at      DATETIME        NOT NULL CONSTRAINT DF_IM_PRD_UOM_CONVERSION_CREATED_AT DEFAULT (GETDATE()),
        prdu_updated_by      VARCHAR(50)     NULL,
        prdu_updated_at      DATETIME        NULL,
        CONSTRAINT PK_IM_PRD_UOM_CONVERSION
            PRIMARY KEY (prdu_prd_code, prdu_alt_uom),
        CONSTRAINT CK_IM_PRD_UOM_CONVERSION_IS_BASE
            CHECK (prdu_is_base IN ('Y', 'N')),
        CONSTRAINT CK_IM_PRD_UOM_CONVERSION_IS_ACTIVE
            CHECK (prdu_is_active IN ('Y', 'N')),
        CONSTRAINT CK_IM_PRD_UOM_CONVERSION_DENOMINATOR
            CHECK (prdu_denominator <> 0)
    );
END";

        TIRAObject<TIRAItem> __brand = null;
        TIRAObject<TIRAItem> __group = null;
        TIRAObject<TIRAItem> __subgroup = null;
        TIRAObject<TIRAItem> __model = null;
        TIRAObject<TIRAItem> __rmcode = null;
        TIRAObject<TIRAItem> __color = null;
        TIRAObject<TIRAItem> __processcode = null;
        TIRAObject<VendorItem> __vendor = null;
        TIRAObject<TIRAItem> __principal = null;
        TIRAObject<TIRAItem> __taxcode = null;
        DataTable __alternateUomTable = null;
        bool __suppressUomSearchButtonEvent = false;
        string __legacySmallUom = "";
        string __legacyMiddleUom = "";
        string __legacyBigUom = "";
        string __legacySalesConversion = "1";
        string __legacyPurchaseConversion = "1";
        #endregion

        LineHighlight[] __cursor = null;

        public frmINVProductMasterEditor()
        {
            InitializeComponent();
            __circle = new LoadingCircle(lblCircle, 5, new Point(10, 1), new Size(10, 10));

            __browseWorker = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true,
            };
            __browseWorker.DoWork += __browseDoWork;
            __browseWorker.RunWorkerCompleted += __browseRunWorkerCompleted;


        }

        void __browseDoWork(object sender, DoWorkEventArgs e)
        {

            TIRAFTP __args = (TIRAFTP)e.Argument;
            string strFilename = linebrandcodetb0.Text.Trim() + "-" + groupcodetb0.Text.Trim() + "-" + txtProdMstrCd.Text.Trim() + __args.file_extension;

            try
            {
                FTP1 = new FtpClient(__args.host, __args.user, __args.pass);
                //FTP1.EncryptionMode = FtpEncryptionMode.Implicit; 
                FTP1.ConnectTimeout = 900000;
                FTP1.ReadTimeout = 900000;
                FTP1.DataConnectionType = FtpDataConnectionType.PASV;

                FTP1.Connect();
                FTP1.UploadFile(__args.filename, __args.directory + "/" + strFilename, FtpExists.Overwrite, true, FtpVerify.None);

            }
            catch (Exception __exc)
            {
                //string t = err.ToString();
                __args.IsSuccess = false;
                __args.State = (Exception)__exc;
            }
            finally
            {
                FTP1.Disconnect();
                __args.remote_path = __args.directory + "/" + strFilename;
                __args.IsSuccess = true;
                e.Result = __args;
            }
        }

        void __browseRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                TIRAFTP __im10 = (TIRAFTP)e.Result;
                if (__im10.IsSuccess)
                {

                    if (!__im10.Message.IsNullOrEmptyOrWhiteSpace())
                        throw new Exception(__im10.Message);

                    txtPhoto.Text = __im10.remote_path;

                    //if (__im30.IsAnyError)
                    //{
                    //    this.prosestbn0.Enabled = false;
                    //}
                    //else
                    //{
                    //    this.prosestbn0.Enabled = true;
                    //}



                }
                else throw (Exception)__im10.State;
            }
            catch (Exception __exc)
            {
                MessageBox.Show(__exc.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                __circle.Stop();
            }
        }





        public string PrdBrandCode
        {
            get { return _prdBrandCode; }
            set { _prdBrandCode = value; }
        }
        public string PrdGrade
        {
            get { return _Grade; }
            set { _Grade = value; }
        }
        private void frmINVProductMasterEditor_Load(object sender, EventArgs e)
        {
            //this.Width += 50;
            //this.Height += 70;

            loaddropdown();
            ConfigureEditorSearchLookUps();
            InitializeDxValidation();
            if (clsGlobal.MODE_TRX == 2)
            {
                FillData();
                CBGroupCode.Enabled = false;
                textGrade.ReadOnly = true;
            }
            FillGridSizeGrp();
            LoadAlternateUomConversions();

            __cursor = new LineHighlight[] { new LineHighlight(tabPage1), new LineHighlight(tabPage2), new LineHighlight(tabPageAlternateUom) };
            foreach (LineHighlight h in __cursor)
                foreach (Control c in h.Parent.Controls.OfType<CheckBox>())
                    if (c.TabStop) h.Add(c);

            __apparelSys = getAppSys();
            //__apparelSys = true;
        }
        bool getAppSys()
        {
            try
            {
                string __query = string.Format("SELECT TOP 1 gp_apparel_sys FROM IM_GENERAL_PARAMETER WITH(NOLOCK) WHERE gp_entity_id='{0}' AND gp_branch_id='{1}'", clsGlobal.ENTITYID, clsGlobal.BRANCHID);
                return _clsGlobal.GetFieldValue(__query).CompareC("Y");
            }
            catch (Exception) { return false; }
        }


        private void ConfigureEditorSearchLookUps()
        {
            SearchLookUpEdit[] lookups =
            {
                linebrandcodetb0n, groupcodetb0n, subgroupcodetb0n, modelcodetb0n,
                rmcodetb0n, colorcodetb0n, proccodetb0n, vendortb0n,
                principaltb0n, taxcodetb0n, srcValClass, srcMatType
            };

            foreach (SearchLookUpEdit lookup in lookups)
            {
                if (lookup == null) continue;

                lookup.Properties.Buttons.Clear();
                lookup.Properties.Buttons.Add(new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Search));
                lookup.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
                lookup.Properties.NullText = string.Empty;
                lookup.Properties.PopupFormSize = new Size(650, 350);
                lookup.Properties.ShowDropDown = DevExpress.XtraEditors.Controls.ShowDropDown.Never;

                lookup.ButtonClick -= EditorSearchLookUp_ButtonClick;
                lookup.ButtonClick += EditorSearchLookUp_ButtonClick;

                lookup.QueryPopUp -= EditorSearchLookUp_QueryPopUp;
                lookup.QueryPopUp += EditorSearchLookUp_QueryPopUp;
            }

            linebrandcodetb0n.EditValueChanged -= linebrandcodetb0n_EditValueChanged;
            linebrandcodetb0n.EditValueChanged += linebrandcodetb0n_EditValueChanged;
            groupcodetb0n.EditValueChanged -= groupcodetb0n_EditValueChanged;
            groupcodetb0n.EditValueChanged += groupcodetb0n_EditValueChanged;
            subgroupcodetb0n.EditValueChanged -= subgroupcodetb0n_EditValueChanged;
            subgroupcodetb0n.EditValueChanged += subgroupcodetb0n_EditValueChanged;
            modelcodetb0n.EditValueChanged -= modelcodetb0n_EditValueChanged;
            modelcodetb0n.EditValueChanged += modelcodetb0n_EditValueChanged;
            rmcodetb0n.EditValueChanged -= rmcodetb0n_EditValueChanged;
            rmcodetb0n.EditValueChanged += rmcodetb0n_EditValueChanged;
            colorcodetb0n.EditValueChanged -= colorcodetb0n_EditValueChanged;
            colorcodetb0n.EditValueChanged += colorcodetb0n_EditValueChanged;
            proccodetb0n.EditValueChanged -= proccodetb0n_EditValueChanged;
            proccodetb0n.EditValueChanged += proccodetb0n_EditValueChanged;
            vendortb0n.EditValueChanged -= vendortb0n_EditValueChanged;
            vendortb0n.EditValueChanged += vendortb0n_EditValueChanged;
            principaltb0n.EditValueChanged -= principaltb0n_EditValueChanged;
            principaltb0n.EditValueChanged += principaltb0n_EditValueChanged;
            taxcodetb0n.EditValueChanged -= taxcodetb0n_EditValueChanged;
            taxcodetb0n.EditValueChanged += taxcodetb0n_EditValueChanged;
            srcValClass.EditValueChanged -= srcValClass_EditValueChanged;
            srcValClass.EditValueChanged += srcValClass_EditValueChanged;
            srcMatType.EditValueChanged -= srcMatType_EditValueChanged;
            srcMatType.EditValueChanged += srcMatType_EditValueChanged;

            linebrandcodetb0.TextChanged -= linebrandcodetb0_TextChanged;
            linebrandcodetb0.TextChanged += linebrandcodetb0_TextChanged;
            groupcodetb0.TextChanged -= groupcodetb0_TextChanged;
            groupcodetb0.TextChanged += groupcodetb0_TextChanged;
            subgroupcodetb0.TextChanged -= subgroupcodetb0_TextChanged;
            subgroupcodetb0.TextChanged += subgroupcodetb0_TextChanged;
            modelcodetb0.TextChanged -= modelcodetb0_TextChanged;
            modelcodetb0.TextChanged += modelcodetb0_TextChanged;
            rmcodetb0.TextChanged -= rmcodetb0_TextChanged;
            rmcodetb0.TextChanged += rmcodetb0_TextChanged;
            colorcodetb0.TextChanged -= colorcodetb0_TextChanged;
            colorcodetb0.TextChanged += colorcodetb0_TextChanged;
            proccodetb0.TextChanged -= proccodetb0_TextChanged;
            proccodetb0.TextChanged += proccodetb0_TextChanged;
            txtValClassCode.TextChanged -= txtValClassCode_TextChanged;
            txtValClassCode.TextChanged += txtValClassCode_TextChanged;
            txtMatTypeCode.TextChanged -= txtMatTypeCode_TextChanged;
            txtMatTypeCode.TextChanged += txtMatTypeCode_TextChanged;

            BindLookupBeforePopup(linebrandcodetb0n);
            BindLookupBeforePopup(groupcodetb0n);
            BindLookupBeforePopup(subgroupcodetb0n);
            BindLookupBeforePopup(modelcodetb0n);
            BindLookupBeforePopup(rmcodetb0n);
            BindLookupBeforePopup(colorcodetb0n);
            BindLookupBeforePopup(proccodetb0n);
            BindLookupBeforePopup(vendortb0n);
            BindLookupBeforePopup(principaltb0n);
            BindLookupBeforePopup(taxcodetb0n);
            BindLookupBeforePopup(srcValClass);
            BindLookupBeforePopup(srcMatType);
        }

        private void EditorSearchLookUp_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            SearchLookUpEdit editor = sender as SearchLookUpEdit;
            if (editor == null) return;

            BindLookupBeforePopup(editor);

            BeginInvoke(new Action(delegate
            {
                if (!editor.IsDisposed && !editor.IsPopupOpen)
                    editor.ShowPopup();
            }));
        }

        private void EditorSearchLookUp_QueryPopUp(object sender, CancelEventArgs e)
        {
            SearchLookUpEdit editor = sender as SearchLookUpEdit;
            if (editor == null) return;

            BindLookupBeforePopup(editor);
        }

        private void BindLookupBeforePopup(SearchLookUpEdit editor)
        {
            if (editor == linebrandcodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(BrandLookupQuery));
            else if (editor == groupcodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(GetEditorGroupLookupQuery()));
            else if (editor == subgroupcodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(GetEditorSubGroupLookupQuery()));
            else if (editor == modelcodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(ModelLookupQuery));
            else if (editor == rmcodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(RawMaterialLookupQuery));
            else if (editor == colorcodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(ColorLookupQuery));
            else if (editor == proccodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(ProcessCodeLookupQuery));
            else if (editor == vendortb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(VendorLookupQuery), true);
            else if (editor == principaltb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(PrincipalLookupQuery));
            else if (editor == taxcodetb0n)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(TaxCodeLookupQuery));
            else if (editor == srcValClass)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(ValuationClassLookupQuery));
            else if (editor == srcMatType)
                BindEditorSearchLookUp(editor, _clsGlobal.ExecDT(MaterialTypeLookupQuery));
        }

        private void BindEditorSearchLookUp(SearchLookUpEdit lookup, DataTable dataSource)
        {
            BindEditorSearchLookUp(lookup, dataSource, false);
        }

        private void BindEditorSearchLookUp(SearchLookUpEdit lookup, DataTable dataSource, bool showCode2)
        {
            if (lookup == null) return;

            if (lookup.Properties.View == null)
                lookup.Properties.View = new GridView();

            lookup.Properties.DataSource = dataSource;
            lookup.Properties.ValueMember = LookupCodeField;
            lookup.Properties.DisplayMember = LookupCodeField;
            lookup.Properties.NullText = string.Empty;
            lookup.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            lookup.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
            lookup.Properties.View.OptionsBehavior.Editable = false;
            lookup.Properties.View.OptionsView.ColumnAutoWidth = false;
            lookup.Properties.View.OptionsView.ShowGroupPanel = false;
            lookup.Properties.PopulateViewColumns();

            ConfigureEditorLookupColumn(lookup, LookupCodeField, "Code", 0, 120, true);
            ConfigureEditorLookupColumn(lookup, LookupVendorCode2Field, "Code 2", 1, 80, showCode2);
            ConfigureEditorLookupColumn(lookup, LookupDescriptionField, "Description", showCode2 ? 2 : 1, 360, true);
        }

        private void ConfigureEditorLookupColumn(SearchLookUpEdit lookup, string fieldName, string caption, int visibleIndex, int width, bool visible)
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

        private string GetEditorGroupLookupQuery()
        {
            return "SELECT DISTINCT pg_prd_group_code AS [Code], pg_prd_group_desc AS [Description] " +
                   "FROM IM_PRD_GROUP WITH(NOLOCK) " +
                   "WHERE pg_prd_line_code LIKE '" + SqlText(linebrandcodetb0.Text.Trim()) + "%' " +
                   "ORDER BY pg_prd_group_code";
        }

        private string GetEditorSubGroupLookupQuery()
        {
            return "SELECT DISTINCT psg_prd_sgroup_code AS [Code], psg_prd_sgroup_desc AS [Description] " +
                   "FROM IM_PRD_SGROUP WITH(NOLOCK) " +
                   "WHERE psg_prd_line LIKE '" + SqlText(linebrandcodetb0.Text.Trim()) + "%' " +
                   "AND psg_prd_group_code LIKE '" + SqlText(groupcodetb0.Text.Trim()) + "%' " +
                   "ORDER BY psg_prd_sgroup_code";
        }

        private string SqlText(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }

        private DataRow GetEditorSelectedLookupRow(SearchLookUpEdit lookup)
        {
            if (lookup == null || lookup.EditValue == null || lookup.EditValue == DBNull.Value) return null;

            object rowObject = lookup.Properties.GetRowByKeyValue(lookup.EditValue);
            DataRowView rowView = rowObject as DataRowView;
            if (rowView != null) return rowView.Row;

            return rowObject as DataRow;
        }

        private string EditorLookupRowText(DataRow row, string fieldName)
        {
            if (row == null || !row.Table.Columns.Contains(fieldName) || row[fieldName] == DBNull.Value)
                return string.Empty;

            return row[fieldName].ToString().Trim();
        }

        private bool SetEditorLookupText(SearchLookUpEdit lookup, TextEdit codeEdit, TextEdit descriptionEdit)
        {
            if (_settingLookupValues) return false;

            DataRow row = GetEditorSelectedLookupRow(lookup);
            if (row == null) return false;

            _settingLookupValues = true;
            try
            {
                codeEdit.Text = EditorLookupRowText(row, LookupCodeField);
                if (descriptionEdit != null)
                    descriptionEdit.Text = EditorLookupRowText(row, LookupDescriptionField);
            }
            finally
            {
                _settingLookupValues = false;
            }

            return true;
        }


        private void linebrandcodetb0_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(linebrandcodetb0.Text))
            {
                ClearTextEdit(linebrandcodetb1);
                ClearTextEdit(groupcodetb0);
                ClearTextEdit(groupcodetb1);
                ClearTextEdit(subgroupcodetb0);
                ClearTextEdit(subgroupcodetb1);
            }
            else
            {
                linebrandcodetb1.Text = FirstFieldValue(
                    "SELECT pl_prd_line_desc AS [Description] FROM IM_PRD_LINE WITH(NOLOCK) WHERE pl_prd_line_code = '" + SqlText(linebrandcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void groupcodetb0_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(groupcodetb0.Text))
            {
                ClearTextEdit(groupcodetb1);
                ClearTextEdit(subgroupcodetb0);
                ClearTextEdit(subgroupcodetb1);
            }
            else
            {
                groupcodetb1.Text = FirstFieldValue(
                    "SELECT DISTINCT pg_prd_group_desc AS [Description] FROM IM_PRD_GROUP WITH(NOLOCK) WHERE pg_prd_line_code LIKE '" + SqlText(linebrandcodetb0.Text.Trim()) + "%' AND pg_prd_group_code = '" + SqlText(groupcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void subgroupcodetb0_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(subgroupcodetb0.Text))
            {
                ClearTextEdit(subgroupcodetb1);
            }
            else
            {
                subgroupcodetb1.Text = FirstFieldValue(
                    "SELECT DISTINCT psg_prd_sgroup_desc AS [Description] FROM IM_PRD_SGROUP WITH(NOLOCK) WHERE psg_prd_line LIKE '" + SqlText(linebrandcodetb0.Text.Trim()) + "%' AND psg_prd_group_code LIKE '" + SqlText(groupcodetb0.Text.Trim()) + "%' AND psg_prd_sgroup_code = '" + SqlText(subgroupcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void modelcodetb0_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(modelcodetb0.Text))
            {
                ClearTextEdit(modelcodetb1);
            }
            else
            {
                modelcodetb1.Text = FirstFieldValue(
                    "SELECT pm_prd_model_desc AS [Description] FROM IM_PRD_MODEL WITH(NOLOCK) WHERE pm_prd_model_code = '" + SqlText(modelcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void rmcodetb0_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(rmcodetb0.Text))
            {
                ClearTextEdit(rmcodetb1);
            }
            else
            {
                rmcodetb1.Text = FirstFieldValue(
                    "SELECT rmu_raw_mat_used_desc AS [Description] FROM IM_RM_USED WITH(NOLOCK) WHERE rmu_raw_mat_used_code = '" + SqlText(rmcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void colorcodetb0_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(colorcodetb0.Text))
            {
                ClearTextEdit(colorcodetb1);
            }
            else
            {
                colorcodetb1.Text = FirstFieldValue(
                    "SELECT col_washing_collor_desc AS [Description] FROM IM_W_COLLOR WITH(NOLOCK) WHERE col_washing_collor_code = '" + SqlText(colorcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void proccodetb0_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(proccodetb0.Text))
            {
                ClearTextEdit(proccodetb1);
            }
            else
            {
                proccodetb1.Text = FirstFieldValue(
                    "SELECT gh_function_desc AS [Description] FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_sys = 'H' AND gh_function_name = 'PROCESSCODE' AND gh_function_code = '" + SqlText(proccodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void txtValClassCode_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(txtValClassCode.Text))
            {
                ClearTextEdit(txtValClassDesc);
            }
            else
            {
                txtValClassDesc.Text = FirstFieldValue(
                    "SELECT mvc_description AS [Description] FROM PO_VALUATION_CLASS WITH(NOLOCK) WHERE mvc_code = '" + SqlText(txtValClassCode.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }

        private void txtMatTypeCode_TextChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            if (string.IsNullOrWhiteSpace(txtMatTypeCode.Text))
            {
                ClearTextEdit(txtMatTypeDesc);
            }
            else
            {
                txtMatTypeDesc.Text = FirstFieldValue(
                    "SELECT mt_type_desc AS [Description] FROM IM_PRD_TYPE WITH(NOLOCK) WHERE mt_type_id = '" + SqlText(txtMatTypeCode.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
        }


        private void linebrandcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            if (SetEditorLookupText(linebrandcodetb0n, linebrandcodetb0, linebrandcodetb1))
            {
                _settingLookupValues = true;
                try
                {
                    groupcodetb0.Text = string.Empty;
                    groupcodetb1.Text = string.Empty;
                    subgroupcodetb0.Text = string.Empty;
                    subgroupcodetb1.Text = string.Empty;
                }
                finally
                {
                    _settingLookupValues = false;
                }
            }
        }

        private void groupcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            if (SetEditorLookupText(groupcodetb0n, groupcodetb0, groupcodetb1))
            {
                _settingLookupValues = true;
                try
                {
                    subgroupcodetb0.Text = string.Empty;
                    subgroupcodetb1.Text = string.Empty;
                }
                finally
                {
                    _settingLookupValues = false;
                }
            }
        }

        private void subgroupcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(subgroupcodetb0n, subgroupcodetb0, subgroupcodetb1);
        }

        private void modelcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(modelcodetb0n, modelcodetb0, modelcodetb1);
        }

        private void rmcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(rmcodetb0n, rmcodetb0, rmcodetb1);
        }

        private void colorcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(colorcodetb0n, colorcodetb0, colorcodetb1);
        }

        private void proccodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(proccodetb0n, proccodetb0, proccodetb1);
        }

        private void vendortb0n_EditValueChanged(object sender, EventArgs e)
        {
            if (_settingLookupValues) return;

            DataRow row = GetEditorSelectedLookupRow(vendortb0n);
            if (row == null) return;

            _settingLookupValues = true;
            try
            {
                vendortb0.Text = EditorLookupRowText(row, LookupCodeField);
                vendortb1.Text = EditorLookupRowText(row, LookupVendorCode2Field);
            }
            finally
            {
                _settingLookupValues = false;
            }
        }

        private void principaltb0n_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(principaltb0n, principaltb0, null);
        }

        private void taxcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(taxcodetb0n, taxcodetb0, null);
        }

        private void srcValClass_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(srcValClass, txtValClassCode, txtValClassDesc);
        }

        private void srcMatType_EditValueChanged(object sender, EventArgs e)
        {
            SetEditorLookupText(srcMatType, txtMatTypeCode, txtMatTypeDesc);
        }


        void _initilize()
        {
            // Native DevExpress SearchLookUpEdit digunakan.
            // Legacy TIRAObject browse tidak dipakai agar popup sama seperti frmINVProductMasterList.
        }

        #region -- end action
        void _endBrand(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.linebrandcodetb1.Text = "";
                        groupcodetb0.Text = string.Empty;
                        subgroupcodetb0.Text = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (!this.linebrandcodetb0.Text.CompareC(__item.Kode))
                        {
                            groupcodetb0.Text = string.Empty;
                            subgroupcodetb0.Text = string.Empty;
                        }

                        //TIRAItem __item = (TIRAItem)__result.Item;
                        this.linebrandcodetb0.Text = __item.Kode;
                        this.linebrandcodetb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endGroup(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.groupcodetb1.Text = "";
                        subgroupcodetb0.Text = string.Empty;
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        if (!this.groupcodetb0.Text.CompareC(__item.Kode))
                        {
                            subgroupcodetb0.Text = string.Empty;
                        }

                        this.groupcodetb0.Text = __item.Kode;
                        this.groupcodetb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endSubGroup(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.subgroupcodetb1.Text = "";
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        this.subgroupcodetb0.Text = __item.Kode;
                        this.subgroupcodetb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endModel(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.modelcodetb1.Text = "";
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        this.modelcodetb0.Text = __item.Kode;
                        this.modelcodetb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endRMCode(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.rmcodetb1.Text = "";
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        this.rmcodetb0.Text = __item.Kode;
                        this.rmcodetb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endColor(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.colorcodetb1.Text = "";
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        this.colorcodetb0.Text = __item.Kode;
                        this.colorcodetb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endProcessCode(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.proccodetb1.Text = "";
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        this.proccodetb0.Text = __item.Kode;
                        this.proccodetb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endVendor(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        this.vendortb1.Text = "";
                    }
                    else
                    {
                        VendorItem __item = (VendorItem)__result.Item;
                        this.vendortb0.Text = __item.ID1;
                        this.vendortb1.Text = __item.ID2;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endPrincipal(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        //this.principaltb0.Text = "";
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        this.principaltb0.Text = __item.Kode;
                        //this.branchtb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        void _endTaxCode(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate ()
                {
                    if (__result.IsError)
                    {
                        //this.branchtb1.Text = "";
                    }
                    else
                    {
                        TIRAItem __item = (TIRAItem)__result.Item;
                        this.taxcodetb0.Text = __item.Kode;
                        //this.branchtb1.Text = __item.Nama;
                    }
                }));
            }
            catch (Exception) { }
        }
        #endregion

        void loaddropdown()
        {
            try
            {
                //Size Group Code

                strSQL = "SELECT szg_size_group_code From IM_PRD_SIZE_GRP GROUP BY szg_size_group_code ORDER BY szg_size_group_code";
                DataTable dt = new DataTable();
                dt = _clsGlobal.ExecDT(strSQL);
                CBGroupCode.Properties.DataSource = dt;
                CBGroupCode.Properties.ValueMember = "szg_size_group_code";
                CBGroupCode.Properties.DisplayMember = "szg_size_group_code";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



            try
            {
                //Status
                strSQL = "select gh_function_code, gh_function_code +' - '+gh_function_desc as gh_function_desc from GS_GEN_HARDCODED where gh_sys = 'H' and gh_function_name = 'STATUS_PRD_MASTER' order by gh_sequence_no ";
                CBStatus.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBStatus.Properties.ValueMember = "gh_function_code";
                CBStatus.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //trend
                strSQL = "select  pp_function_code, pp_function_code+' - '+pp_function_desc pp_function_desc from IM_PRD_PARAMETER  where  pp_function_name = 'TREND' order by pp_sequence_no  ";
                CBPacksize.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPacksize.Properties.ValueMember = "pp_function_code";
                CBPacksize.Properties.DisplayMember = "pp_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //target
                strSQL = "select  pp_function_code, pp_function_code+' - '+pp_function_desc pp_function_desc from IM_PRD_PARAMETER  where  pp_function_name = 'TARGET' order by pp_sequence_no  ";
                CBPackType.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPackType.Properties.ValueMember = "pp_function_code";
                CBPackType.Properties.DisplayMember = "pp_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //season
                strSQL = "select  pp_function_code, pp_function_code+' - '+pp_function_desc pp_function_desc from IM_PRD_PARAMETER  where  pp_function_name = 'SEASON' order by pp_sequence_no  ";
                CBSeason.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBSeason.Properties.ValueMember = "pp_function_code";
                CBSeason.Properties.DisplayMember = "pp_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //period indicator
                strSQL = "select gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc  from GS_GEN_HARDCODED  where gh_sys = 'H' and gh_function_name = 'PERIOD_INDC'  order by gh_sequence_no";
                CBPeriodIndicator.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPeriodIndicator.Properties.ValueMember = "gh_function_code";
                CBPeriodIndicator.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //produk type
                strSQL = "select gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc  from GS_GEN_HARDCODED  where gh_sys = 'H' and gh_function_name = 'PRDTYPE'  order by gh_sequence_no";
                CBType.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBType.Properties.ValueMember = "gh_function_code";
                CBType.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //Class
                strSQL = "select gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc  from GS_GEN_HARDCODED  where gh_sys = 'H' and gh_function_name = 'SEX_DESIGN'  order by gh_sequence_no";
                CBGrade.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBGrade.Properties.ValueMember = "gh_function_code";
                CBGrade.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //Rumusan OB
                strSQL = "SELECT gh_function_code, gh_function_code + ' ~ ' + gh_function_desc as gh_function_desc  From GS_GEN_HARDCODED WHERE gh_sys = 'H'   AND gh_function_name = 'RUMUSAN_OB'";
                CBRumusOB.Properties.DataSource = _clsGlobal.ExecDT(strSQL);
                CBRumusOB.Properties.ValueMember = "gh_function_code";
                CBRumusOB.Properties.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            BindAlternateUomLookup();

        }

        private void BindAlternateUomLookup()
        {
            try
            {
                DataTable dtUom = _clsGlobal.ExecDT(__UOM_LOOKUP);
                repositoryItemUomLookup.DataSource = dtUom;
                repositoryItemBaseUomLookup.DataSource = dtUom.Copy();

                // FIX: the visible columns (Alternate / Measurement / Base Unit) use
                // repositoryItemUomSearchButton as their editor. Its DataSource was never
                // assigned, which is why the search popup showed no data. Bind it here.
                repositoryItemUomSearchButton.DataSource = dtUom.Copy();

                EnsureUomLookupValue(__legacySmallUom);
                EnsureUomLookupValue(__legacyMiddleUom);
                EnsureUomLookupValue(__legacyBigUom);

                // Make sure the picked value is written back to the matching unit column,
                // and keep the (unbound) search-button cells clean.
                gridAlternateUomView.CustomUnboundColumnData -= gridAlternateUomView_CustomUnboundColumnData;
                gridAlternateUomView.CustomUnboundColumnData += gridAlternateUomView_CustomUnboundColumnData;

                repositoryItemUomSearchButton.EditValueChanged -= repositoryItemUomSearchButton_EditValueChanged;
                repositoryItemUomSearchButton.EditValueChanged += repositoryItemUomSearchButton_EditValueChanged;

                colMeasurementUom.OptionsColumn.AllowEdit = false;
                colMeasurementUom.OptionsColumn.ReadOnly = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EnsureUomLookupValue(string uomCode)
        {
            uomCode = (uomCode ?? string.Empty).Trim();
            if (uomCode == string.Empty) return;

            EnsureUomLookupValue(repositoryItemUomLookup.DataSource as DataTable, uomCode);
            EnsureUomLookupValue(repositoryItemBaseUomLookup.DataSource as DataTable, uomCode);
            EnsureUomLookupValue(repositoryItemUomSearchButton.DataSource as DataTable, uomCode);
        }

        private void EnsureUomLookupValue(DataTable dtUom, string uomCode)
        {
            if (dtUom == null) return;

            foreach (DataRow row in dtUom.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                if (row["Kode"].ToString().Trim().Equals(uomCode, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            DataRow newRow = dtUom.NewRow();
            newRow["Kode"] = uomCode;
            newRow["Nama"] = uomCode;
            dtUom.Rows.Add(newRow);
        }

        private string GetUomDescription(string uomCode)
        {
            uomCode = (uomCode ?? string.Empty).Trim();
            if (uomCode == string.Empty) return string.Empty;

            DataTable dtUom = repositoryItemUomLookup.DataSource as DataTable;
            if (dtUom == null || dtUom.Rows.Count == 0)
            {
                dtUom = _clsGlobal.ExecDT(__UOM_LOOKUP);
                repositoryItemUomLookup.DataSource = dtUom;
            }

            foreach (DataRow row in dtUom.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;

                string code = RowText(row, "Kode");
                if (code.Equals(uomCode, StringComparison.OrdinalIgnoreCase))
                    return RowText(row, "Nama");
            }

            return uomCode;
        }

        #region -- button event click
        //private void btnUpBrandCode_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Brand ";
        //    frm.Query = "SELECT pl_prd_line_code as 'Prod. Line', pl_prd_line_desc as 'Description' " +
        //                "  From IM_PRD_LINE " +
        //                " ORDER BY pl_prd_line_code";

        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtLineBrndCd1.Text = frm.ArrField[0].Trim();
        //        txtLineBrndCd2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnUPGrupCode_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Group ";
        //    frm.Query = "SELECT distinct pg_prd_group_code as 'Prod. Group', pg_prd_group_desc as 'Description' " +
        //                "  From IM_PRD_GROUP " +
        //                "  WHERE pg_prd_line_code = '" + txtLineBrndCd1.Text + "'" +
        //                " ORDER BY pg_prd_group_code";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtGroupCd1.Text = frm.ArrField[0].Trim();
        //        txtGroupCd2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnUpSubGroup_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Sub Group ";
        //    frm.Query = "SELECT distinct psg_prd_sgroup_code as 'Prod. Subgroup', psg_prd_sgroup_desc as 'Description ' " +
        //                "  From IM_PRD_SGROUP " +
        //                " WHERE psg_prd_line = '" + txtLineBrndCd1.Text + "' " +
        //                "   AND psg_prd_group_code = '" + txtGroupCd1.Text + "' " +
        //                " ORDER BY psg_prd_sgroup_code";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {
        //        txtSubGrpCd1.Text = frm.ArrField[0].Trim();
        //        txtSubGrpCd2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnUPModeCode_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Model ";
        //    frm.Query = "SELECT distinct pm_prd_model_code as 'Model Code',pm_prd_model_desc as 'Description' " +
        //                "  From IM_PRD_MODEL " +
        //                " ORDER BY pm_prd_model_code";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtMdlCd1.Text = frm.ArrField[0].Trim();
        //        txtMdlCd2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnUpRMCode_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Group ";
        //    frm.Query = "SELECT rmu_raw_mat_used_code as 'Raw Material', rmu_raw_mat_used_desc as 'Description' " +
        //                "  From IM_RM_USED " +
        //                " ORDER BY rmu_raw_mat_used_code";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtRMSrceCd1.Text = frm.ArrField[0].Trim();
        //        txtRMSrceCd2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnPopUpColor_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Color ";
        //    frm.Query = "SELECT col_washing_collor_code as 'Color', col_washing_collor_desc as 'Description' " +
        //                "  From IM_W_COLLOR " +
        //                " ORDER BY col_washing_collor_code";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtClrCd1.Text = frm.ArrField[0].Trim();
        //        txtClrcd2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnPopUpProcessCode_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Process Code ";
        //    frm.Query = " SELECT gh_function_code As Process ,gh_function_desc as Proc_Desc from   GS_GEN_HARDCODED WHERE gh_sys ='H' and gh_function_name ='PROCESSCODE' ";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtprcsCd1.Text = frm.ArrField[0].Trim();
        //        txtprcsCd2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnPopUpVendor_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Process Code ";
        //    frm.Query = "  select pvm_vendor_code1 as ID1, pvm_vendor_code2 as ID2, pvm_vendor_namekey as [Desc] from PO_VENDOR_MASTER ";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtVndrSupplier1.Text = frm.ArrField[0].Trim();
        //        txtVndrSupplier2.Text = frm.ArrField[1].Trim();

        //    }
        //}
        //private void btnUPPrincipal_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Principal ";
        //    frm.Query = " select pri_principal as ID, pri_principal_desc as [Desc] from PO_PRINCIPAL  ";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txtPrncpl.Text = frm.ArrField[0].Trim();

        //    }
        //}
        //private void btnUPTax_Click(object sender, EventArgs e)
        //{
        //    frm = new frmPopUp();
        //    frm.FrmText = "Search Tax COde ";
        //    frm.Query = " SELECT t_tax_id, t_tax FROM IM_TAX   ";
        //    frm.ShowDialog();

        //    if (frm.ArrField != null)
        //    {

        //        txttaxcode.Text = frm.ArrField[0].Trim();

        //    }
        //}
        #endregion

        #region -- combo selection index change
        private void ComboBixStts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBStatus.ItemIndex > -1)
            { }
        }
        private void ComboSzGroupCd_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBGroupCode.ItemIndex > -1)
            { }
        }
        private void checkBoxComboPckSzTrnd_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBPacksize.ItemIndex > -1)
            { }
        }
        private void checkBoxComboPckTypeTrgt_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBPackType.ItemIndex > -1)
            { }
        }
        private void checkBoxComboSeason_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBSeason.ItemIndex > 0)
            { }
        }
        #endregion

        private void FillGridSizeGrp()
        {
            try
            {
                strSQL = " select distinct szg_index, szg_prd_size, opl_het as Price , cs_standard_unit_cost as Cost from IM_COST_STD  inner  join IM_PRD_SIZE_GRP on szg_prd_size = cs_size " +
                " inner join VW_INV_PRODUK_MASTER on  prm_size_group_code = szg_size_group_code AND prm_prd_master_code = cs_prd_master_code " +
                " where cs_prd_master_code ='" + txtProdMstrCd.Text + "' and cs_grade='" + textGrade.Text + "'  ORDER BY szg_index";


                DataTable dtSize = _clsGlobal.ExecDT(strSQL);
                BindSizeGrid(dtSize);

                if (dtSize.Rows.Count > 0)
                {
                    SetFirstSizeGridRowSelected(true);

                    //DataGridViewCellEventArgs DataGridViewCellEventArgs = new DataGridViewCellEventArgs(dGV1.Columns["szg_index"].Index, dGV1.CurrentRow.Index);
                    //this.dGV1_CellClick(dGV1.CurrentRow.Index, DataGridViewCellEventArgs);

                }
                else
                {
                    strSQL = "SELECT DISTINCT szg_index, szg_prd_size, '' as Price, '' as Cost FROM IM_PRD_SIZE_GRP WHERE szg_size_group_code= '" + CBGroupCode.EditValue + "' ORDER BY szg_index";
                    dtSize = _clsGlobal.ExecDT(strSQL);
                    BindSizeGrid(dtSize);
                    if (dtSize.Rows.Count > 0)
                    {
                        if (clsGlobal.MODE_TRX == 1)
                        {
                            SetFirstSizeGridRowSelected(false);
                        }
                        else
                        {
                            SetFirstSizeGridRowSelected(true);
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindSizeGrid(DataTable dtSize)
        {
            if (!dtSize.Columns.Contains("szg_selected"))
                dtSize.Columns.Add("szg_selected", typeof(bool));

            foreach (DataRow row in dtSize.Rows)
                row["szg_selected"] = false;

            dGV1.DataSource = dtSize;
        }

        private void SetFirstSizeGridRowSelected(bool selected)
        {
            DataTable dtSize = dGV1.DataSource as DataTable;
            if (dtSize != null && dtSize.Rows.Count > 0)
                dtSize.Rows[0]["szg_selected"] = selected;
        }

        private void dGV1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CheckBox chkisdefault = new CheckBox();
                chkisdefault.Checked = false;
                chkisdefault.Enabled = false;

                if (chkisdefault.Checked)
                {
                    chkisdefault.Checked = false;
                }
            }
        }

        private void checkFlagOBRnd_CheckedChanged(object sender, EventArgs e)
        {
            if (CKObTidak.Checked)
            {
                CKObTidak.Text = "Ya";

            }
            else
            {
                CKObTidak.Text = "Tidak";
            }
        }

        private void checkFlagBatch_CheckedChanged(object sender, EventArgs e)
        {
            if (CKBatchTidak.Checked)
            {
                CKBatchTidak.Text = "Ya";
                lblPeriodvvv.Visible = true;
                lblPeriod.Visible = true;
                txtMimimum.Visible = true;
                CBPeriodIndicator.Visible = true;

            }
            else
            {
                CKBatchTidak.Text = "Tidak";
                lblPeriodvvv.Visible = false;
                lblPeriod.Visible = false;
                txtMimimum.Visible = false;
                CBPeriodIndicator.Visible = false;
            }
        }



        private void ClearTextEdit(TextEdit edit)
        {
            if (edit != null)
                edit.Text = string.Empty;
        }


        private string RowText(DataRow row, string fieldName)
        {
            if (row == null || row.Table == null || !row.Table.Columns.Contains(fieldName) || row[fieldName] == DBNull.Value)
                return string.Empty;

            return row[fieldName].ToString().Trim();
        }

        private string FirstFieldValue(string query, string fieldName)
        {
            DataTable dt = _clsGlobal.ExecDT(query);
            if (dt.Rows.Count == 0 || !dt.Columns.Contains(fieldName) || dt.Rows[0][fieldName] == DBNull.Value)
                return string.Empty;

            return dt.Rows[0][fieldName].ToString().Trim();
        }

        private void ReloadLookupDescriptions()
        {
            _settingLookupValues = true;
            try
            {
                linebrandcodetb1.Text = FirstFieldValue(
                    "SELECT pl_prd_line_desc AS [Description] FROM IM_PRD_LINE WITH(NOLOCK) WHERE pl_prd_line_code = '" + SqlText(linebrandcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);

                groupcodetb1.Text = FirstFieldValue(
                    "SELECT DISTINCT pg_prd_group_desc AS [Description] FROM IM_PRD_GROUP WITH(NOLOCK) WHERE pg_prd_line_code LIKE '" + SqlText(linebrandcodetb0.Text.Trim()) + "%' AND pg_prd_group_code = '" + SqlText(groupcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);

                subgroupcodetb1.Text = FirstFieldValue(
                    "SELECT DISTINCT psg_prd_sgroup_desc AS [Description] FROM IM_PRD_SGROUP WITH(NOLOCK) WHERE psg_prd_line LIKE '" + SqlText(linebrandcodetb0.Text.Trim()) + "%' AND psg_prd_group_code LIKE '" + SqlText(groupcodetb0.Text.Trim()) + "%' AND psg_prd_sgroup_code = '" + SqlText(subgroupcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);

                modelcodetb1.Text = FirstFieldValue(
                    "SELECT pm_prd_model_desc AS [Description] FROM IM_PRD_MODEL WITH(NOLOCK) WHERE pm_prd_model_code = '" + SqlText(modelcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);

                rmcodetb1.Text = FirstFieldValue(
                    "SELECT rmu_raw_mat_used_desc AS [Description] FROM IM_RM_USED WITH(NOLOCK) WHERE rmu_raw_mat_used_code = '" + SqlText(rmcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);

                colorcodetb1.Text = FirstFieldValue(
                    "SELECT col_washing_collor_desc AS [Description] FROM IM_W_COLLOR WITH(NOLOCK) WHERE col_washing_collor_code = '" + SqlText(colorcodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);

                proccodetb1.Text = FirstFieldValue(
                    "SELECT gh_function_desc AS [Description] FROM GS_GEN_HARDCODED WITH(NOLOCK) WHERE gh_sys = 'H' AND gh_function_name = 'PROCESSCODE' AND gh_function_code = '" + SqlText(proccodetb0.Text.Trim()) + "'",
                    LookupDescriptionField);

                txtValClassDesc.Text = FirstFieldValue(
                    "SELECT mvc_description AS [Description] FROM PO_VALUATION_CLASS WITH(NOLOCK) WHERE mvc_code = '" + SqlText(txtValClassCode.Text.Trim()) + "'",
                    LookupDescriptionField);

                txtMatTypeDesc.Text = FirstFieldValue(
                    "SELECT mt_type_desc AS [Description] FROM IM_PRD_TYPE WITH(NOLOCK) WHERE mt_type_id = '" + SqlText(txtMatTypeCode.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
            finally
            {
                _settingLookupValues = false;
            }
        }

        private string FormatDecimalSql(string value)
        {
            string text = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
                return "0";

            decimal parsed;
            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed) ||
                decimal.TryParse(text, NumberStyles.Any, new CultureInfo("id-ID"), out parsed))
            {
                return parsed.ToString(CultureInfo.InvariantCulture);
            }

            return FmtStr(text);
        }

        private void LoadAdditionalProductMasterFields(DataRow row)
        {
            if (row == null) return;

            txtNetGross.Text = RowText(row, "prm_net_weight").Replace(".00000", "");
            txtSupplierMatNo.Text = RowText(row, "prm_supplier_mat_id");
            txtBarcodeOuter.Text = RowText(row, "prm_barcode_outer");
            txtValClassCode.Text = RowText(row, "prm_valuation_class");
            txtMatTypeCode.Text = RowText(row, "prm_type_id");

            if (string.IsNullOrWhiteSpace(txtNetGross.Text) &&
                string.IsNullOrWhiteSpace(txtSupplierMatNo.Text) &&
                string.IsNullOrWhiteSpace(txtBarcodeOuter.Text) &&
                string.IsNullOrWhiteSpace(txtValClassCode.Text) &&
                string.IsNullOrWhiteSpace(txtMatTypeCode.Text))
            {
                DataTable dtAdditional = _clsGlobal.ExecDT(
                    "SELECT prm_net_weight, prm_supplier_mat_id, prm_barcode_outer, prm_valuation_class, prm_type_id " +
                    "FROM IM_PRD_MASTER WITH(NOLOCK) " +
                    "WHERE prm_prd_master_code = " + FmtStr(txtProdMstrCd.Text.Trim()) + " AND prm_grade = " + FmtStr(textGrade.Text.Trim()));

                if (dtAdditional.Rows.Count > 0)
                {
                    txtNetGross.Text = RowText(dtAdditional.Rows[0], "prm_net_weight").Replace(".00000", "");
                    txtSupplierMatNo.Text = RowText(dtAdditional.Rows[0], "prm_supplier_mat_id");
                    txtBarcodeOuter.Text = RowText(dtAdditional.Rows[0], "prm_barcode_outer");
                    txtValClassCode.Text = RowText(dtAdditional.Rows[0], "prm_valuation_class");
                    txtMatTypeCode.Text = RowText(dtAdditional.Rows[0], "prm_type_id");
                }
            }

            ReloadAdditionalProductMasterDescriptions();
        }

        private void ReloadAdditionalProductMasterDescriptions()
        {
            _settingLookupValues = true;
            try
            {
                txtValClassDesc.Text = FirstFieldValue(
                    "SELECT mvc_description AS [Description] FROM PO_VALUATION_CLASS WITH(NOLOCK) WHERE mvc_code = '" + SqlText(txtValClassCode.Text.Trim()) + "'",
                    LookupDescriptionField);

                txtMatTypeDesc.Text = FirstFieldValue(
                    "SELECT mt_type_desc AS [Description] FROM IM_PRD_TYPE WITH(NOLOCK) WHERE mt_type_id = '" + SqlText(txtMatTypeCode.Text.Trim()) + "'",
                    LookupDescriptionField);
            }
            finally
            {
                _settingLookupValues = false;
            }
        }

        private void SaveAdditionalProductMasterFields()
        {
            string updateAdditionalSql =
                "UPDATE IM_PRD_MASTER SET " +
                "prm_net_weight = " + FormatDecimalSql(txtNetGross.Text.Trim()) + ", " +
                "prm_supplier_mat_id = " + FmtStr(txtSupplierMatNo.Text.Trim()) + ", " +
                "prm_barcode_outer = " + FmtStr(txtBarcodeOuter.Text.Trim()) + ", " +
                "prm_valuation_class = " + FmtStr(txtValClassCode.Text.Trim()) + ", " +
                "prm_type_id = " + FmtStr(txtMatTypeCode.Text.Trim()) + " " +
                "WHERE prm_prd_master_code = " + FmtStr(txtProdMstrCd.Text.Trim()) + " AND prm_grade = " + FmtStr(textGrade.Text.Trim());

            _clsGlobal.ExecuteTrans(updateAdditionalSql);
        }


        private void FillData()
        {
            try
            {
                strSQL = "";
                strSQL = " SELECT * FROM VW_INV_PRODUK_MASTER WITH(NOLOCK) " +
                         " WHERE prm_prd_master_code = " + FmtStr(_prdBrandCode) + " AND prm_grade = " + FmtStr(_Grade) + " ";

                DataTable dt = new DataTable();
                dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    linebrandcodetb0.Text = dt.Rows[0]["prm_prd_line_code"].ToString().Trim();
                    groupcodetb0.Text = dt.Rows[0]["prm_prd_group_code"].ToString().Trim();
                    subgroupcodetb0.Text = dt.Rows[0]["prm_prd_sgroup_code"].ToString().Trim();
                    modelcodetb0.Text = dt.Rows[0]["prm_prd_model_code"].ToString().Trim();
                    rmcodetb0.Text = dt.Rows[0]["prm_raw_mat_used_code"].ToString().Trim();
                    colorcodetb0.Text = dt.Rows[0]["prm_washing_collor_code"].ToString().Trim();
                    proccodetb0.Text = dt.Rows[0]["prm_process_code"].ToString().Trim();
                    vendortb0.Text = dt.Rows[0]["prm_vendor_id1"].ToString().Trim();
                    vendortb1.Text = dt.Rows[0]["prm_vendor_id2"].ToString().Trim();
                    principaltb0.Text = dt.Rows[0]["prm_principal"].ToString().Trim();

                    // Setelah SearchLookUpEdit diubah native DevExpress, description tidak lagi diisi oleh legacy TIRAObject.
                    // Isi ulang description dari VW_INV_PRODUK_MASTER dulu, lalu fallback ke master jika view tidak punya kolom desc.
                    _settingLookupValues = true;
                    try
                    {
                        linebrandcodetb1.Text = RowText(dt.Rows[0], "pl_prd_line_desc");
                        groupcodetb1.Text = RowText(dt.Rows[0], "pg_prd_group_desc");
                        subgroupcodetb1.Text = RowText(dt.Rows[0], "psg_prd_sgroup_desc");
                        modelcodetb1.Text = RowText(dt.Rows[0], "pm_prd_model_desc");
                        rmcodetb1.Text = RowText(dt.Rows[0], "rmu_raw_mat_used_desc");
                        colorcodetb1.Text = RowText(dt.Rows[0], "col_washing_collor_desc");
                        proccodetb1.Text = RowText(dt.Rows[0], "gh_function_desc");
                    }
                    finally
                    {
                        _settingLookupValues = false;
                    }

                    if (string.IsNullOrWhiteSpace(linebrandcodetb1.Text) ||
                        string.IsNullOrWhiteSpace(groupcodetb1.Text) ||
                        string.IsNullOrWhiteSpace(subgroupcodetb1.Text) ||
                        string.IsNullOrWhiteSpace(modelcodetb1.Text) ||
                        string.IsNullOrWhiteSpace(rmcodetb1.Text) ||
                        string.IsNullOrWhiteSpace(colorcodetb1.Text) ||
                        string.IsNullOrWhiteSpace(proccodetb1.Text))
                    {
                        ReloadLookupDescriptions();
                    }

                    txtProdMstrCd.Text = dt.Rows[0]["prm_prd_master_code"].ToString().Trim();
                    txtdescrptn.Text = dt.Rows[0]["prm_prd_desc"].ToString().Trim();
                    txtshrtdescrp.Text = dt.Rows[0]["prm_prd_short"].ToString().Trim();
                    //txtmnfctrCd.Text = ((dt.Rows[0]["prm_mfg_code"].ToString().Trim() != "") ? "'" + dt.Rows[0]["prm_mfg_code"].ToString().Trim() + "'" : "1");
                    txtmnfctrCd.Text = dt.Rows[0]["prm_mfg_code"].ToString().Trim();
                    txtNoOfItem.Text = dt.Rows[0]["prm_no_of_item"].ToString().Trim();
                    taxcodetb0.Text = dt.Rows[0]["prm_tax_code"].ToString().Trim();
                    txtgrpitm.Text = dt.Rows[0]["prm_item_group"].ToString().Trim();
                    if (dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString() != "")
                    {
                        if (dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString().Length >= 2)
                        {
                            txtperiode1.Text = dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString().Trim().Substring(0, 2);
                        }
                        else
                        {
                            txtperiode1.Text = dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString().Trim();
                        }
                    }
                    if (dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString() != "")
                    {
                        if (dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString().Length >= 4)
                        {
                            txtperiode2.Text = dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString().Trim().Substring(2);
                        }
                        else
                        {
                            txtperiode2.Text = dt.Rows[0]["prm_mfg_periode_mm_yy"].ToString().Trim();
                        }
                    }
                    textGrade.Text = dt.Rows[0]["prm_grade"].ToString().Trim();
                    LoadAdditionalProductMasterFields(dt.Rows[0]);
                    txtDcmlPrcsn.Text = dt.Rows[0]["prm_decimal_point"].ToString().Trim();
                    txtMnmnItm.Text = dt.Rows[0]["prm_pline"].ToString().Trim();
                    txtPknAwl.Text = dt.Rows[0]["prm_weekno"].ToString().Trim();
                    txtRprtName1.Text = dt.Rows[0]["prm_report_name1"].ToString().Trim();
                    txtRprtName2.Text = dt.Rows[0]["prm_report_name2"].ToString().Trim();
                    txtRprtName3.Text = dt.Rows[0]["prm_report_name3"].ToString().Trim();
                    txtWeight.Text = dt.Rows[0]["prm_unit_weight_gr"].ToString().Trim().Replace(".00000", "");
                    txtvlme1.Text = dt.Rows[0]["prm_unit_volume_cm3"].ToString().Trim().Replace(".00000", "");
                    txtvlme2.Text = dt.Rows[0]["prm_unit_volume_lt"].ToString().Trim().Replace(".00000", "");
                    //txtWeight.Text = dt.Rows[0]["prm_unit_weight_gr"].ToString().Trim().Substring(0, 1);
                    //txtvlme1.Text = dt.Rows[0]["prm_unit_volume_cm3"].ToString().Trim().Substring(0, 1);
                    //txtvlme2.Text = dt.Rows[0]["prm_unit_volume_lt"].ToString().Trim().Substring(0, 1);
                    CBType.EditValue = dt.Rows[0]["prm_prd_type"].ToString().Trim();
                    CBGroupCode.EditValue = dt.Rows[0]["prm_size_group_code"].ToString().Trim();
                    Sizegroupcode = dt.Rows[0]["prm_size_group_code"].ToString().Trim();
                    CBGrade.EditValue = dt.Rows[0]["prm_sex"].ToString().Trim();
                    __legacySmallUom = dt.Rows[0]["prm_um"].ToString().Trim();
                    __legacyMiddleUom = dt.Rows[0]["prm_um_sales"].ToString().Trim();
                    __legacySalesConversion = DefaultOne(dt.Rows[0]["prm_conversion_sales"].ToString().Trim());
                    __legacyBigUom = dt.Rows[0]["prm_um_purc"].ToString().Trim();
                    __legacyPurchaseConversion = DefaultOne(dt.Rows[0]["prm_conversion_purc"].ToString().Trim());
                    CBStatus.EditValue = dt.Rows[0]["prm_status"].ToString().Trim();
                    CBPackType.EditValue = dt.Rows[0]["prm_target"].ToString().Trim();
                    CBPacksize.EditValue = dt.Rows[0]["prm_trend"].ToString().Trim();
                    CBSeason.EditValue = dt.Rows[0]["prm_season"].ToString().Trim();

                    if (dt.Rows[0]["prm_wh_control_flag"].ToString().Trim() == "Y")
                    {
                        checkWrhsCntrl.Checked = true;
                    }
                    else
                    {
                        checkWrhsCntrl.Checked = false;
                    }

                    if (dt.Rows[0]["prm_non_stock_flag"].ToString().Trim() == "Y")
                    {
                        checkNnStckItm.Checked = true;
                    }
                    else
                    {
                        checkNnStckItm.Checked = false;
                    }
                    if (dt.Rows[0]["prm_tech_constrain_flag"].ToString().Trim() == "Y")
                    {
                        CBKhusus.Checked = true;
                    }
                    else
                    {
                        CBKhusus.Checked = false;
                    }
                    if (dt.Rows[0]["prm_prd_cust_flag"].ToString().Trim() == "Y")
                    {
                        CBReguler.Checked = true;
                    }
                    else
                    {
                        CBReguler.Checked = false;
                    }
                    //txtbrcd.Text = ((dt.Rows[0]["prm_barcode"].ToString().Trim() != "") ? "'" + dt.Rows[0]["prm_barcode"].ToString().Trim() + "'" : "11111");
                    txtbrcd.Text = dt.Rows[0]["prm_barcode"].ToString().Trim();
                    txtSftyBox.Text = dt.Rows[0]["prm_sftstock"].ToString().Trim();
                    txtStockMax.Text = dt.Rows[0]["prm_stock_max"].ToString().Trim();
                    txtStockMin.Text = dt.Rows[0]["prm_min_stock"].ToString().Trim();
                    CBRumusOB.EditValue = dt.Rows[0]["prm_rumus_ob"].ToString().Trim();

                    if (dt.Rows[0]["prm_flaq_ob_round"].ToString().Trim() == "Y")
                    {
                        CKObTidak.Checked = true;
                    }
                    else
                    {
                        CKObTidak.Checked = false;
                    }
                    if (dt.Rows[0]["prm_batch_flaq"].ToString().Trim() == "Y")
                    {
                        CKBatchTidak.Checked = true;
                    }
                    else
                    {
                        CKBatchTidak.Checked = false;
                    }
                    txtMimimum.Text = dt.Rows[0]["prm_min_rem_sled"].ToString().Trim();
                    CBPeriodIndicator.EditValue = dt.Rows[0]["prm_period_ind"].ToString().Trim();
                    txtPhoto.Text = dt.Rows[0]["prm_prd_photo"].ToString().Trim();
                    if (dt.Rows[0]["prm_ext_pajak_flag"].ToString() == "Y")
                    {
                        check_ext_pajak_flag.Checked = true;
                    }
                    else
                    {
                        check_ext_pajak_flag.Checked = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        string _checkedYN(bool __ischecked)
        {
            if (__ischecked)
                return "Y";
            return "N";
        }

        private DataTable CreateAlternateUomTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("prdu_prd_code", typeof(string));
            dt.Columns.Add("prdu_alt_uom", typeof(string));
            dt.Columns.Add("prdu_measurement_uom", typeof(string));
            dt.Columns.Add("equivalent", typeof(string));
            dt.Columns.Add("prdu_base_uom", typeof(string));
            dt.Columns.Add("prdu_numerator", typeof(decimal));
            dt.Columns.Add("prdu_denominator", typeof(decimal));
            dt.Columns.Add("prdu_is_base_bool", typeof(bool));
            dt.Columns.Add("prdu_is_active_bool", typeof(bool));
            return dt;
        }

        private void LoadAlternateUomConversions()
        {
            __alternateUomTable = CreateAlternateUomTable();
            gridAlternateUom.DataSource = __alternateUomTable;

            EnsureUomLookupValue(__legacySmallUom);
            EnsureUomLookupValue(__legacyMiddleUom);
            EnsureUomLookupValue(__legacyBigUom);

            if (txtProdMstrCd.Text.Trim() == "")
            {
                AddDefaultAlternateUomRowsFromLegacy();
                return;
            }

            try
            {
                _clsGlobal.ExecuteTrans(__PRD_UOM_CONVERSION_TABLE);

                string query = "SELECT prdu_prd_code, prdu_alt_uom, ISNULL(umc.umc_uom_description, prdu_alt_uom) AS prdu_measurement_uom, '<=>' AS equivalent, " +
                               "prdu_base_uom, prdu_numerator, prdu_denominator, " +
                               "CASE WHEN prdu_is_base = 'Y' THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS prdu_is_base_bool, " +
                               "CASE WHEN prdu_is_active = 'Y' THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS prdu_is_active_bool " +
                               "FROM IM_PRD_UOM_CONVERSION WITH(NOLOCK) " +
                               "LEFT JOIN IM_UNIT_MEASURE_CODES umc WITH(NOLOCK) ON umc.umc_uom_code = prdu_alt_uom " +
                               "WHERE prdu_prd_code = " + FmtStr(txtProdMstrCd.Text.Trim()) + " " +
                               "ORDER BY prdu_is_base DESC, prdu_alt_uom";

                DataTable dt = _clsGlobal.ExecDT(query);
                foreach (DataRow source in dt.Rows)
                {
                    string altUom = source["prdu_alt_uom"].ToString().Trim();
                    string measurementUom = source["prdu_measurement_uom"].ToString().Trim();
                    string baseUom = source["prdu_base_uom"].ToString().Trim();

                    EnsureUomLookupValue(altUom);
                    EnsureUomLookupValue(baseUom);

                    DataRow row = __alternateUomTable.NewRow();
                    row["prdu_prd_code"] = source["prdu_prd_code"].ToString().Trim();
                    row["prdu_alt_uom"] = altUom;
                    row["prdu_measurement_uom"] = measurementUom;
                    row["equivalent"] = "<=>";
                    row["prdu_base_uom"] = baseUom;
                    row["prdu_numerator"] = ToDecimal(source["prdu_numerator"], 1m);
                    row["prdu_denominator"] = ToDecimal(source["prdu_denominator"], 1m);
                    row["prdu_is_base_bool"] = ToBool(source["prdu_is_base_bool"]);
                    row["prdu_is_active_bool"] = ToBool(source["prdu_is_active_bool"]);
                    __alternateUomTable.Rows.Add(row);
                }

                if (__alternateUomTable.Rows.Count == 0)
                    AddDefaultAlternateUomRowsFromLegacy();
            }
            catch
            {
                __alternateUomTable.Clear();
                AddDefaultAlternateUomRowsFromLegacy();
            }
        }

        private void AddDefaultAlternateUomRowsFromLegacy()
        {
            AddAlternateUomRowFromLegacy(__legacySmallUom, __legacySmallUom, 1m, 1m, true);

            if (!string.IsNullOrWhiteSpace(__legacyMiddleUom) &&
                !__legacyMiddleUom.Trim().Equals(__legacySmallUom.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                AddAlternateUomRowFromLegacy(__legacyMiddleUom, __legacySmallUom, ToDecimal(__legacySalesConversion, 1m), 1m, false);
            }

            if (!string.IsNullOrWhiteSpace(__legacyBigUom) &&
                !__legacyBigUom.Trim().Equals(__legacySmallUom.Trim(), StringComparison.OrdinalIgnoreCase) &&
                !__legacyBigUom.Trim().Equals(__legacyMiddleUom.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                AddAlternateUomRowFromLegacy(__legacyBigUom, __legacySmallUom, ToDecimal(__legacyPurchaseConversion, 1m), 1m, false);
            }
        }

        private void AddAlternateUomRowFromLegacy(string altUom, string baseUom, decimal numerator, decimal denominator, bool isBase)
        {
            altUom = (altUom ?? string.Empty).Trim();
            baseUom = (baseUom ?? string.Empty).Trim();

            if (altUom == string.Empty) return;
            if (baseUom == string.Empty) baseUom = altUom;

            EnsureUomLookupValue(altUom);
            EnsureUomLookupValue(baseUom);

            DataRow row = __alternateUomTable.NewRow();
            row["prdu_prd_code"] = txtProdMstrCd.Text.Trim();
            row["prdu_alt_uom"] = altUom;
            row["prdu_measurement_uom"] = GetUomDescription(altUom);
            row["equivalent"] = "<=>";
            row["prdu_base_uom"] = baseUom;
            row["prdu_numerator"] = numerator == 0m ? 1m : numerator;
            row["prdu_denominator"] = denominator == 0m ? 1m : denominator;
            row["prdu_is_base_bool"] = isBase;
            row["prdu_is_active_bool"] = true;
            __alternateUomTable.Rows.Add(row);
        }

        private void SaveAlternateUomConversions()
        {
            gridAlternateUomView.CloseEditor();
            gridAlternateUomView.UpdateCurrentRow();

            _clsGlobal.ExecuteTrans(__PRD_UOM_CONVERSION_TABLE);
            _clsGlobal.ExecuteTrans("DELETE FROM IM_PRD_UOM_CONVERSION WHERE prdu_prd_code = " + FmtStr(txtProdMstrCd.Text.Trim()));

            DataTable dt = __alternateUomTable ?? gridAlternateUom.DataSource as DataTable;
            if (dt == null)
                return;

            foreach (DataRow row in dt.Rows)
            {
                if (row.RowState == DataRowState.Deleted)
                    continue;

                string altUom = GetRowString(row, "prdu_alt_uom");
                string baseUom = GetRowString(row, "prdu_base_uom");
                if (altUom == "" && baseUom == "")
                    continue;

                decimal numerator = ToDecimal(row["prdu_numerator"], 1m);
                decimal denominator = ToDecimal(row["prdu_denominator"], 1m);
                string isBase = ToBool(row["prdu_is_base_bool"]) ? "Y" : "N";
                string isActive = ToBool(row["prdu_is_active_bool"]) ? "Y" : "N";

                string insertSql = "INSERT INTO IM_PRD_UOM_CONVERSION " +
                                   "(prdu_prd_code, prdu_alt_uom, prdu_base_uom, prdu_numerator, prdu_denominator, prdu_is_base, prdu_is_active, prdu_created_by, prdu_created_at, prdu_updated_by, prdu_updated_at) VALUES (" +
                                   FmtStr(txtProdMstrCd.Text.Trim()) + ", " +
                                   FmtStr(altUom) + ", " +
                                   FmtStr(baseUom) + ", " +
                                   SqlDecimal(numerator) + ", " +
                                   SqlDecimal(denominator) + ", " +
                                   FmtStr(isBase) + ", " +
                                   FmtStr(isActive) + ", " +
                                   FmtStr(clsLogin.USERID ?? "") + ", GETDATE(), " +
                                   FmtStr(clsLogin.USERID ?? "") + ", GETDATE())";

                _clsGlobal.ExecuteTrans(insertSql);
            }
        }

        private string GetRowString(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName) || row[columnName] == DBNull.Value)
                return "";

            return row[columnName].ToString().Trim();
        }

        private decimal ToDecimal(object value, decimal defaultValue)
        {
            if (value == null || value == DBNull.Value || value.ToString().Trim() == "")
                return defaultValue;

            decimal result;
            if (decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out result))
                return result;

            if (decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out result))
                return result;

            return defaultValue;
        }

        private bool ToBool(object value)
        {
            if (value == null || value == DBNull.Value)
                return false;

            if (value is bool)
                return (bool)value;

            return value.ToString().Trim().CompareC("Y") || value.ToString().Trim().CompareC("1") || value.ToString().Trim().CompareC("True");
        }

        private string SqlDecimal(decimal value)
        {
            return value.ToString("0.####", CultureInfo.InvariantCulture);
        }

        private string CurrentBaseUom()
        {
            DataTable dt = __alternateUomTable ?? gridAlternateUom.DataSource as DataTable;
            if (dt != null)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row.RowState == DataRowState.Deleted)
                        continue;

                    if (ToBool(row["prdu_is_base_bool"]))
                    {
                        string baseUom = GetRowString(row, "prdu_base_uom");
                        if (baseUom != "")
                            return baseUom;

                        baseUom = GetRowString(row, "prdu_alt_uom");
                        if (baseUom != "")
                            return baseUom;
                    }
                }

                foreach (DataRow row in dt.Rows)
                {
                    if (row.RowState == DataRowState.Deleted)
                        continue;

                    string baseUom = GetRowString(row, "prdu_base_uom");
                    if (baseUom != "")
                        return baseUom;
                }
            }

            return __legacySmallUom;
        }

        private string LegacyMiddleUom()
        {
            string baseUom = CurrentBaseUom();
            if (__legacyMiddleUom != "")
                return __legacyMiddleUom;

            return baseUom;
        }

        private string LegacyBigUom()
        {
            string baseUom = CurrentBaseUom();
            if (__legacyBigUom != "")
                return __legacyBigUom;

            return baseUom;
        }

        private string DefaultOne(string value)
        {
            if (value == null || value.Trim() == "" || value.Trim() == "0")
                return "1";

            return value.Trim();
        }

        private void SaveNew()
        {
            try
            {
                string priode = txtperiode1.Text.Trim() + "" + txtperiode2.Text.Trim();

                //string cb1, cb2, cb3, cb4, cb5, cb6;
                //if (CKObTidak.Checked)
                //{
                //    cb1 = "Y";
                //}
                //else
                //{
                //    cb1 = "N";
                //}
                //if (CKBatchTidak.Checked)
                //{
                //    cb2 = "Y";
                //}
                //else
                //{
                //    cb2 = "N";
                //}
                //if (checkWrhsCntrl.Checked)
                //{
                //    cb3 = "Y";
                //}
                //else
                //{
                //    cb3 = "N";
                //}
                //if (CBKhusus.Checked)
                //{
                //    cb4 = "Y";
                //}
                //else
                //{
                //    cb4 = "N";
                //}
                //if (checkNnStckItm.Checked)
                //{
                //    cb5 = "Y";
                //}
                //else
                //{
                //    cb5 = "N";
                //}
                //if (CBReguler.Checked)
                //{
                //    cb6 = "Y";
                //}
                //else
                //{
                //    cb6 = "N";
                //}
                string cb1 = _checkedYN(CKObTidak.Checked);
                string cb2 = _checkedYN(CKBatchTidak.Checked);
                string cb3 = _checkedYN(checkWrhsCntrl.Checked);
                string cb4 = _checkedYN(CBKhusus.Checked);
                string cb5 = _checkedYN(checkNnStckItm.Checked);
                string cb6 = _checkedYN(CBReguler.Checked);
                string cbflagextract = _checkedYN(check_ext_pajak_flag.Checked);
                if (!taxcodetb0.Text.ToString().ToUpper().Equals("PPN0"))
                {
                    cbflagextract = "N";
                }

                strSQL = "";
                strSQL = "EXEC [SP_INV_PRODUK_MASTER] '1', ";
                strSQL += "" + FmtStr(txtProdMstrCd.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(textGrade.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(textGrade.Text.Trim()) + ", ";
                //strSQL+=   // "" + ((CBPacksize.EditValue != "") ? "" + CBPacksize.EditValue + "" : "NULL") + "," +
                //strSQL+=   //"" + ((CBPackType.EditValue != "") ? "" + CBPackType.EditValue + "" : "NULL") + "," +
                //strSQL+=   //"" + ((CBSeason.EditValue != "") ? "" + CBSeason.EditValue + "" : "NULL") + "," +

                if (CBPacksize.EditValue != null)
                    strSQL += "" + FmtStr(CBPacksize.EditValue.ToString()) + ", ";
                else
                    strSQL += "'',";

                if (CBPackType.EditValue != null)
                    strSQL += "" + FmtStr(CBPackType.EditValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                if (CBSeason.EditValue != null)
                    strSQL += "" + FmtStr(CBSeason.EditValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                strSQL += "" + FmtStr(linebrandcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(groupcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(subgroupcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(modelcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(rmcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(colorcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtdescrptn.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtshrtdescrp.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtmnfctrCd.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtgrpitm.Text.Trim()) + ", ";
                strSQL += "" + ((txtNoOfItem.Text != "") ? "'" + txtNoOfItem.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(taxcodetb0.Text.Trim()) + ",'', ";
                //strSQL +=   //prm_default_disc_code

                strSQL += "" + FmtStr(CurrentBaseUom()) + ", ";
                strSQL += "" + FmtStr(LegacyMiddleUom()) + ", ";
                strSQL += "" + FmtStr(LegacyBigUom()) + ", ";

                strSQL += "" + FmtStr(DefaultOne(__legacySalesConversion)) + ", ";
                strSQL += "" + FmtStr(DefaultOne(__legacyPurchaseConversion)) + ", ";
                strSQL += "" + ((txtWeight.Text != "") ? "'" + txtWeight.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtvlme1.Text != "") ? "'" + txtvlme1.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtvlme2.Text != "") ? "'" + txtvlme2.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(cb3.Trim()) + ", ";
                strSQL += "" + FmtStr(cb4.Trim()) + ", ";
                strSQL += "" + FmtStr(cb5.Trim()) + ",'Y', ";
                //strSQL+=   //"" + FmtStr(cb6.Trim()) + ", " +
                strSQL += "" + FmtStr(CBType.EditValue.ToString()) + ",'', ";
                //strSQL += "" + clsLogin.USERID + ", ";
                strSQL += string.Format("'{0}', ", clsLogin.USERID);
                strSQL += "" + FmtStr(CBGrade.EditValue.ToString()) + ", ";
                strSQL += "" + ((txtDcmlPrcsn.Text != "") ? "'" + txtDcmlPrcsn.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(priode.Trim()) + ", ";
                strSQL += "" + FmtStr(CBGroupCode.EditValue.ToString()) + ", ";
                strSQL += "" + FmtStr(proccodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(vendortb0.Text.Trim()) + ",";
                strSQL += "" + FmtStr(vendortb1.Text.Trim()) + ",";
                strSQL += "" + FmtStr(principaltb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(CBStatus.EditValue.ToString()) + ", ";
                //strSQL+=    //"" + ((CBStatus.EditValue != "") ? "" + CBStatus.EditValue + "" : "NULL") + "," +
                strSQL += "" + FmtStr(txtRprtName1.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName2.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName3.Text.Trim()) + ", ";
                strSQL += "" + ((txtDcmlPrcsn.Text != "") ? "'" + txtMnmnItm.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtPknAwl.Text != "") ? "'" + txtPknAwl.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(txtbrcd.Text.Trim()) + ", ";
                strSQL += "" + ((txtStockMax.Text != "") ? "'" + txtStockMax.Text + "'" : "0") + ", ";
                //strSQL+=     //"" + ((CBRumusOB.EditValue != "") ? "" + CBRumusOB.EditValue + "" : "NULL") + "," +
                strSQL += "" + FmtStr(CBRumusOB.EditValue.ToString()) + ", ";
                strSQL += "" + FmtStr(cb1.Trim()) + ", ";
                strSQL += "" + FmtStr(cb2.Trim()) + ", ";
                strSQL += "" + ((txtSftyBox.Text != "") ? "'" + txtSftyBox.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtStockMin.Text != "") ? "'" + txtStockMin.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtMimimum.Text != "") ? "'" + txtMimimum.Text + "'" : "0") + ", ";
                //strSQL+=    //"" + ((CBPeriodIndicator.EditValue != "") ? "" + CBPeriodIndicator.EditValue + "" : "NULL") + "," +
                strSQL += "" + FmtStr(CBPeriodIndicator.EditValue.ToString()) + ",";
                strSQL += "" + FmtStr(cb6.Trim()) + ",";
                strSQL += "'N' ,";
                strSQL += "" + FmtStr(txtPhoto.Text.Trim()) + ",";
                strSQL += "" + FmtStr(cbflagextract) + "";
                //prm_prd_cust_flag
                //prm_prd_promo_flag 58
                _clsGlobal.BeginTrans();
                _clsGlobal.ExecuteTrans(strSQL);
                SaveAdditionalProductMasterFields();
                SaveAlternateUomConversions();
                _clsGlobal.CommitTrans();

                MessageBox.Show(_clsGlobal.ApplMessage(40048), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                _clsGlobal.RollbackTrans();
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveEdit()
        {
            bool begTrans = false;
            try
            {
                string priode = txtperiode1.Text.Trim() + "" + txtperiode2.Text.Trim();

                //string cb1, cb2, cb3, cb4, cb5, cb6;
                //if (CKObTidak.Checked)
                //{
                //    cb1 = "Y";
                //}
                //else
                //{
                //    cb1 = "N";
                //}
                //if (CKBatchTidak.Checked)
                //{
                //    cb2 = "Y";
                //}
                //else
                //{
                //    cb2 = "N";
                //}
                //if (checkWrhsCntrl.Checked)
                //{
                //    cb3 = "Y";
                //}
                //else
                //{
                //    cb3 = "N";
                //}
                //if (CBKhusus.Checked)
                //{ 
                //    cb4 = "Y"; 
                //}
                //else
                //{
                //    cb4 = "N";
                //}
                //if (checkNnStckItm.Checked)
                //{
                //    cb5 = "Y"; 
                //}
                //else
                //{
                //    cb5 = "N";
                //}
                //if (CBReguler.Checked)
                //{
                //    cb6 = "Y"; 
                //}
                //else 
                //{
                //    cb6 = "N";
                //}

                string cb1 = _checkedYN(CKObTidak.Checked);
                string cb2 = _checkedYN(CKBatchTidak.Checked);
                string cb3 = _checkedYN(checkWrhsCntrl.Checked);
                string cb4 = _checkedYN(CBKhusus.Checked);
                string cb5 = _checkedYN(checkNnStckItm.Checked);
                string cb6 = _checkedYN(CBReguler.Checked);
                string cbflagextract = _checkedYN(check_ext_pajak_flag.Checked);
                if (!taxcodetb0.Text.ToString().ToUpper().Equals("PPN0"))
                {
                    cbflagextract = "N";
                }

                strSQL = "";
                strSQL = "EXEC [SP_INV_PRODUK_MASTER] '2', ";
                strSQL += "" + FmtStr(txtProdMstrCd.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(textGrade.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(textGrade.Text.Trim()) + ", ";
                strSQL += "" + ((CBPacksize.EditValue != null) ? "'" + CBPacksize.EditValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + ((CBPackType.EditValue != null) ? "'" + CBPackType.EditValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + ((CBSeason.EditValue != null) ? "'" + CBSeason.EditValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + FmtStr(linebrandcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(groupcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(subgroupcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(modelcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(rmcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(colorcodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtdescrptn.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtshrtdescrp.Text.Trim()) + ",";
                strSQL += "" + FmtStr(txtmnfctrCd.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtgrpitm.Text.Trim()) + ", ";
                strSQL += "" + ((txtNoOfItem.Text != "") ? "'" + txtNoOfItem.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(taxcodetb0.Text.Trim()) + ", ";
                strSQL += " NULL ,";

                strSQL += "" + FmtStr(CurrentBaseUom()) + ", ";
                strSQL += "" + FmtStr(LegacyMiddleUom()) + ", ";
                strSQL += "" + FmtStr(LegacyBigUom()) + ", ";

                strSQL += "" + FmtStr(DefaultOne(__legacySalesConversion)) + ", ";
                strSQL += "" + FmtStr(DefaultOne(__legacyPurchaseConversion)) + ", ";
                strSQL += "" + FmtStr(txtWeight.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtvlme1.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtvlme2.Text.Trim()) + ", ";
                //strSQL+=     //"" + ((txtWeight.Text != "") ? "'" + txtWeight.Text + "'" : "0") + ", " +
                //strSQL+=     //"" + ((txtvlme1.Text != "") ? "'" + txtvlme1.Text + "'" : "0") + ", " +
                //strSQL+=     //"" + ((txtvlme2.Text != "") ? "'" + txtvlme2.Text + "'" : "0") + ", " +
                strSQL += "" + FmtStr(cb3.Trim()) + ", ";
                strSQL += "" + FmtStr(cb4.Trim()) + ", ";
                strSQL += "" + FmtStr(cb5.Trim()) + ", ";
                strSQL += " NULL ,";
                strSQL += "" + ((CBType.EditValue != null) ? "'" + CBType.EditValue.ToString() + "'" : "NULL") + ",";
                strSQL += " NULL ,";
                // strSQL += "" + clsLogin.USERID + ", ";
                strSQL += string.Format("'{0}', ", clsLogin.USERID);

                if (CBGrade.EditValue != null)
                    strSQL += "" + FmtStr(CBGrade.EditValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                strSQL += "" + ((txtDcmlPrcsn.Text != "") ? "'" + txtDcmlPrcsn.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(priode.Trim()) + ", ";

                if (CBGroupCode.EditValue != null)
                    strSQL += "" + FmtStr(CBGroupCode.EditValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                strSQL += "" + FmtStr(proccodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(vendortb0.Text.Trim()) + ",";
                strSQL += "" + FmtStr(vendortb1.Text.Trim()) + ",";
                strSQL += "" + FmtStr(principaltb0.Text.Trim()) + ", ";
                strSQL += "" + ((CBStatus.EditValue != null) ? "'" + CBStatus.EditValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + FmtStr(txtRprtName1.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName2.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName3.Text.Trim()) + ", ";
                strSQL += "" + ((txtMnmnItm.Text != "") ? "'" + txtMnmnItm.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtPknAwl.Text != "") ? "'" + txtPknAwl.Text + "'" : "0") + ",";
                strSQL += "" + ((txtbrcd.Text != "") ? "'" + txtbrcd.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtStockMax.Text != "") ? "'" + txtStockMax.Text + "'" : "0") + ", ";
                strSQL += "" + ((CBRumusOB.EditValue != null) ? "'" + CBRumusOB.EditValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + FmtStr(cb1.Trim()) + ", ";
                strSQL += "" + FmtStr(cb2.Trim()) + ", ";
                strSQL += "" + ((txtSftyBox.Text != "") ? "'" + txtSftyBox.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtStockMin.Text != "") ? "'" + txtStockMin.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtMimimum.Text != "") ? "'" + txtMimimum.Text + "'" : "0") + ", ";
                strSQL += "" + ((CBPeriodIndicator.EditValue != null) ? "'" + CBPeriodIndicator.EditValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + FmtStr(cb6.Trim()) + ",";
                strSQL += "'N' ,";
                strSQL += "" + FmtStr(txtPhoto.Text.Trim()) + ",";
                strSQL += "" + FmtStr(cbflagextract) + " ";
                begTrans = true;
                _clsGlobal.BeginTrans();
                _clsGlobal.ExecuteTrans(strSQL);
                SaveAdditionalProductMasterFields();
                SaveAlternateUomConversions();
                _clsGlobal.CommitTrans();

                MessageBox.Show(_clsGlobal.ApplMessage(40048), clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                if (begTrans)
                {
                    _clsGlobal.RollbackTrans();
                }
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string FmtStr(string value_str)
        {
            string formatStr;
            formatStr = "'" + value_str.Trim().Replace("'", "''") + "'";

            return formatStr;
        }

        #region filter textbox events
        //private void txtLineBrndCd1_Leave(object sender, EventArgs e)
        //{
        //    if (txtLineBrndCd2.Text == "")
        //    {
        //        txtLineBrndCd1.Text = "";
        //    }
        //}
        //private void txtLineBrndCd1_TextChanged(object sender, EventArgs e)
        //{
        //    getBranch(txtLineBrndCd1);

        //}
        //private void getBranch(TextBox objText)
        //{
        //    try
        //    {
        //        strSQL = "";
        //        strSQL = "SELECT distinct pl_prd_line_code as 'Prod. Line', pl_prd_line_desc as 'Description' " +
        //                    "  From IM_PRD_LINE " +
        //                    "  WHERE pl_prd_line_code = '" + objText.Text.Trim() + "'";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtLineBrndCd2.Text = dt.Rows[0]["Description"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtLineBrndCd2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void txtGroupCd1_TextChanged(object sender, EventArgs e)
        //{
        //    getGrup(txtGroupCd1);
        //}
        //private void txtGroupCd1_Leave(object sender, EventArgs e)
        //{
        //    if (txtGroupCd2.Text == "")
        //    {
        //        txtGroupCd1.Text = "";
        //    }
        //}
        //private void getGrup(TextBox objText2)
        //{
        //    try
        //    {
        //        strSQL = "SELECT  pg_prd_group_code, pg_prd_group_desc" +
        //                "  From IM_PRD_GROUP " +
        //                "  WHERE pg_prd_line_code = '" + txtLineBrndCd1.Text.Trim() + "'" +
        //                "  AND pg_prd_group_code = '" + objText2.Text.Trim() + "'";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtGroupCd2.Text = dt.Rows[0]["pg_prd_group_desc"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtGroupCd2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void txtSubGrpCd1_TextChanged(object sender, EventArgs e)
        //{
        //    getSubGrup(txtSubGrpCd1);
        //}
        //private void txtSubGrpCd1_Leave(object sender, EventArgs e)
        //{

        //    if (txtSubGrpCd2.Text == "")
        //    {
        //        txtSubGrpCd1.Text = "";
        //    }
        //}
        //private void getSubGrup(TextBox objText1)
        //{
        //    try
        //    {
        //        strSQL = "";
        //        strSQL = "SELECT  psg_prd_sgroup_code, psg_prd_sgroup_desc" +
        //                "  From IM_PRD_SGROUP " +
        //                " WHERE psg_prd_line = '" + txtLineBrndCd1.Text.Trim() + "' " +
        //                " AND psg_prd_group_code = '" + txtGroupCd1.Text.Trim() + "' " +
        //                " AND psg_prd_sgroup_code = '" + objText1.Text.Trim() + "' ";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtSubGrpCd2.Text = dt.Rows[0]["psg_prd_sgroup_desc"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtSubGrpCd2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void txtMdlCd1_TextChanged(object sender, EventArgs e)
        //{
        //    getModel(txtMdlCd1);
        //}
        //private void txtMdlCd1_Leave(object sender, EventArgs e)
        //{
        //    if (txtMdlCd2.Text == "")
        //    {
        //        txtMdlCd1.Text = "";
        //    }
        //}
        //private void getModel(TextBox objText1)
        //{
        //    try
        //    {
        //        strSQL = "SELECT  pm_prd_model_code,pm_prd_model_desc" +
        //                 " From IM_PRD_MODEL WHERE pm_prd_model_code = '" + objText1.Text.Trim() + "' " +
        //                " ORDER BY pm_prd_model_code";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtMdlCd2.Text = dt.Rows[0]["pm_prd_model_desc"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtMdlCd2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void txtRMSrceCd1_TextChanged(object sender, EventArgs e)
        //{
        //    getRMcODE(txtRMSrceCd1);
        //}
        //private void txtRMSrceCd1_Leave(object sender, EventArgs e)
        //{
        //    if (txtRMSrceCd2.Text == "")
        //    {
        //        txtRMSrceCd1.Text = "";
        //    }
        //}
        //private void getRMcODE(TextBox objText1)
        //{
        //    try
        //    {
        //        strSQL = "";
        //        strSQL = "SELECT  rmu_raw_mat_used_code, rmu_raw_mat_used_desc" +
        //                "  From IM_RM_USED where rmu_raw_mat_used_code= '" + objText1.Text.Trim() + "' " +
        //                " ORDER BY rmu_raw_mat_used_code";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtRMSrceCd2.Text = dt.Rows[0]["rmu_raw_mat_used_desc"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtRMSrceCd2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void txtClrCd1_Leave(object sender, EventArgs e)
        //{
        //    if (txtClrcd2.Text == "")
        //    {
        //        txtClrCd1.Text = "";
        //    }
        //}
        //private void txtClrCd1_TextChanged(object sender, EventArgs e)
        //{
        //    getColor(txtClrCd1);
        //}
        //private void getColor(TextBox objText1)
        //{
        //    try
        //    {
        //        strSQL = "";
        //        strSQL = "SELECT col_washing_collor_code, col_washing_collor_desc" +
        //                "  From IM_W_COLLOR where col_washing_collor_code ='" + objText1.Text.Trim() + "'" +
        //                " ORDER BY col_washing_collor_code";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtClrcd2.Text = dt.Rows[0]["col_washing_collor_desc"].ToString().Trim();

        //        }
        //        else
        //        {

        //            txtClrcd2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void txtprcsCd1_TextChanged(object sender, EventArgs e)
        //{
        //    getProsescode();
        //}
        //private void txtprcsCd1_Leave(object sender, EventArgs e)
        //{
        //    if (txtprcsCd2.Text == "")
        //    {
        //        txtprcsCd1.Text = "";
        //    }
        //}
        //private void getProsescode()
        //{
        //    try
        //    {
        //        strSQL = "";
        //        strSQL = "SELECT gh_function_code,gh_function_desc from   GS_GEN_HARDCODED WHERE gh_sys ='H' and gh_function_name ='PROCESSCODE' ";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtprcsCd2.Text = dt.Rows[0]["gh_function_desc"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtprcsCd2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void txtVndrSupplier1_TextChanged(object sender, EventArgs e)
        //{
        //    getVendor();
        //}
        //private void txtVndrSupplier1_Leave(object sender, EventArgs e)
        //{
        //    if (txtVndrSupplier2.Text == "")
        //    {
        //        txtVndrSupplier1.Text = "";
        //    }
        //}
        //private void getVendor()
        //{
        //    try
        //    {
        //        strSQL = "";
        //        strSQL = "select pvm_vendor_code1, pvm_vendor_code2, pvm_vendor_namekey from PO_VENDOR_MASTER ";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtVndrSupplier2.Text = dt.Rows[0]["pvm_vendor_code2"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtVndrSupplier2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void getPrincipal()
        //{
        //    try
        //    {
        //        strSQL = "";
        //        strSQL = "select pri_principal, pri_principal_desc from PO_PRINCIPAL ";

        //        DataTable dt = _clsGlobal.ExecDT(strSQL);

        //        if (dt.Rows.Count > 0)
        //        {
        //            txtVndrSupplier2.Text = dt.Rows[0]["pri_principal_desc"].ToString().Trim();
        //        }
        //        else
        //        {

        //            txtVndrSupplier2.Text = string.Empty;

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        #endregion

        private void InitializeDxValidation()
        {
            if (dxValidator != null) return;

            dxValidator = new DXValidationProvider(this.components);
            dxValidator.ValidationMode = ValidationMode.Manual;
            RegisterProductValidationRules();

            dGV1View.RowCellStyle -= ProductGridView_RowCellStyle;
            dGV1View.RowCellStyle += ProductGridView_RowCellStyle;
            dGV1View.CellValueChanged -= ProductGridView_CellValueChanged;
            dGV1View.CellValueChanged += ProductGridView_CellValueChanged;

            gridAlternateUomView.RowCellStyle -= AlternateUomGridView_RowCellStyle;
            gridAlternateUomView.RowCellStyle += AlternateUomGridView_RowCellStyle;

            CKBatchTidak.CheckedChanged -= CKBatchTidak_CheckedChanged;
            CKBatchTidak.CheckedChanged += CKBatchTidak_CheckedChanged;
        }

        private void RegisterProductValidationRules()
        {
            _registeredValidationControls.Clear();
            _validationTabMap.Clear();
            _validationRuleMap.Clear();

            AddRequiredRule(linebrandcodetb0, null, "Line Brand Code wajib diisi.");
            AddRequiredRule(groupcodetb0, null, "Group wajib diisi.");
            AddRequiredRule(subgroupcodetb0, null, "SubGroup wajib diisi.");
            AddRequiredRule(modelcodetb0, null, "Model wajib diisi.");
            AddRequiredRule(rmcodetb0, null, "RM/Source Code wajib diisi.");
            AddRequiredRule(colorcodetb0, null, "Color Code wajib diisi.");
            AddRequiredRule(vendortb0, null, "Vendor / Supplier wajib diisi.");
            AddRequiredRule(principaltb0, null, "Principal wajib diisi.");

            AddRequiredRule(txtProdMstrCd, tabPage1, "Prod Master Code wajib diisi.");
            AddRequiredRule(txtdescrptn, tabPage1, "Description wajib diisi.");
            AddRequiredRule(txtmnfctrCd, tabPage1, "Manufacture Code wajib diisi.");
            AddRequiredRule(txtMnmnItm, tabPage1, "Minimum Item wajib diisi.");
            AddRequiredRule(txtPknAwl, tabPage1, "Pekan Awal wajib diisi.");
            AddRequiredRule(txtRprtName1, tabPage1, "Report Name 1 wajib diisi.");
            AddRequiredRule(txtRprtName2, tabPage1, "Report Name 2 wajib diisi.");
            AddRequiredRule(txtRprtName3, tabPage1, "Report Name 3 wajib diisi.");
            AddRequiredRule(taxcodetb0, tabPage1, "Tax Code wajib diisi.");

            AddRequiredRule(txtbrcd, tabPage2, "Barcode wajib diisi.");
            AddRule(txtMimimum, tabPage2, new MinimumShelfLifeRule(this), "Min. Remaining Shelf Life tidak boleh 0 saat Batch dicentang.");
        }

        private void AddRequiredRule(Control control, XtraTabPage tabPage, string errorText)
        {
            AddRule(control, tabPage, new RequiredControlRule(), errorText);
        }

        private void AddRule(Control control, XtraTabPage tabPage, ValidationRule rule, string errorText)
        {
            if (control == null || dxValidator == null || rule == null) return;

            rule.ErrorText = errorText;
            rule.ErrorType = ErrorType.Critical;
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

        private bool ValidateProductMasterForm()
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

            if (!ValidateSizeGrid())
                isValid = false;

            if (!ValidateAlternateUomGrid())
                isValid = false;

            RefreshValidationTabColors(false);

            XtraTabPage firstInvalidTab = GetFirstInvalidTab();
            if (firstInvalidTab != null)
                tabProduct.SelectedTabPage = firstInvalidTab;

            FocusFirstInvalidControlOrGrid();
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

            if (control == txtMimimum && !CKBatchTidak.Checked)
                return false;

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
                return;
            }

            if (!_validationOriginalBackColorMap.ContainsKey(control))
                _validationOriginalBackColorMap[control] = control.BackColor;

            control.BackColor = Color.MistyRose;
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
                return;
            }

            if (_validationOriginalBackColorMap.ContainsKey(control))
                control.BackColor = _validationOriginalBackColorMap[control];
        }

        private void ResetValidationTabColors()
        {
            ResetTabHeader(tabPage1);
            ResetTabHeader(tabPage2);
            ResetTabHeader(tabPageAlternateUom);
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
                    if (control == null || !IsControlValidationActive(control)) continue;

                    bool controlValid = validateControls ? dxValidator.Validate(control) : EvaluateValidationRuleOnly(control);
                    if (!controlValid)
                    {
                        XtraTabPage tabPage;
                        if (_validationTabMap.TryGetValue(control, out tabPage))
                            MarkTabInvalid(tabPage);
                    }
                }

                if (_invalidSizeGridCells.Count > 0)
                    MarkTabInvalid(tabPage1);

                if (_invalidAlternateUomGridCells.Count > 0)
                    MarkTabInvalid(tabPageAlternateUom);
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
            tabPage.Appearance.Header.BackColor = Color.MistyRose;
            tabPage.Appearance.Header.Font = new Font(tabProduct.Font, FontStyle.Bold);
            tabPage.Appearance.Header.Options.UseForeColor = true;
            tabPage.Appearance.Header.Options.UseBackColor = true;
            tabPage.Appearance.Header.Options.UseFont = true;
        }

        private void ResetTabHeader(XtraTabPage tabPage)
        {
            if (tabPage == null) return;

            tabPage.Appearance.Header.Options.UseForeColor = false;
            tabPage.Appearance.Header.Options.UseBackColor = false;
            tabPage.Appearance.Header.Options.UseFont = false;
        }

        private XtraTabPage GetFirstInvalidTab()
        {
            foreach (Control control in _registeredValidationControls)
            {
                if (control == null || !IsControlValidationActive(control)) continue;

                if (!EvaluateValidationRuleOnly(control))
                {
                    XtraTabPage tabPage;
                    if (_validationTabMap.TryGetValue(control, out tabPage))
                        return tabPage;
                }
            }

            if (_invalidSizeGridCells.Count > 0)
                return tabPage1;

            if (_invalidAlternateUomGridCells.Count > 0)
                return tabPageAlternateUom;

            return null;
        }

        private void FocusFirstInvalidControlOrGrid()
        {
            foreach (Control control in _registeredValidationControls)
            {
                if (control == null || !IsControlValidationActive(control)) continue;
                if (EvaluateValidationRuleOnly(control)) continue;

                XtraTabPage tabPage;
                if (_validationTabMap.TryGetValue(control, out tabPage))
                    tabProduct.SelectedTabPage = tabPage;

                if (control.CanFocus)
                    control.Focus();
                return;
            }

            if (_invalidSizeGridCells.Count > 0)
            {
                tabProduct.SelectedTabPage = tabPage1;
                dGV1.Focus();
                return;
            }

            if (_invalidAlternateUomGridCells.Count > 0)
            {
                tabProduct.SelectedTabPage = tabPageAlternateUom;
                gridAlternateUom.Focus();
            }
        }

        private bool ValidateSizeGrid()
        {
            _invalidSizeGridCells.Clear();

            if (dGV1View.RowCount > 0)
            {
                for (int rowHandle = 0; rowHandle < dGV1View.RowCount; rowHandle++)
                {
                    if (dGV1View.IsNewItemRow(rowHandle))
                        continue;

                    object costValue = dGV1View.GetRowCellValue(rowHandle, "Cost");
                    object priceValue = dGV1View.GetRowCellValue(rowHandle, "Price");

                    if (costValue == null || costValue == DBNull.Value || costValue.ToString().Trim() == "")
                        _invalidSizeGridCells.Add(GridCellKey(rowHandle, "Cost"));

                    if (priceValue == null || priceValue == DBNull.Value || priceValue.ToString().Trim() == "")
                        _invalidSizeGridCells.Add(GridCellKey(rowHandle, "Price"));
                }
            }

            dGV1View.RefreshData();
            if (_invalidSizeGridCells.Count > 0)
                MarkTabInvalid(tabPage1);

            return _invalidSizeGridCells.Count == 0;
        }

        private bool ValidateAlternateUomGrid()
        {
            _invalidAlternateUomGridCells.Clear();
            gridAlternateUomView.CloseEditor();
            gridAlternateUomView.UpdateCurrentRow();

            Dictionary<string, int> altUomRows = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int rowHandle = 0; rowHandle < gridAlternateUomView.RowCount; rowHandle++)
            {
                if (gridAlternateUomView.IsNewItemRow(rowHandle))
                    continue;

                string altUom = GridCellText(gridAlternateUomView.GetRowCellValue(rowHandle, colAltUom));
                string baseUom = GridCellText(gridAlternateUomView.GetRowCellValue(rowHandle, colBaseUom));
                if (altUom == "" && baseUom == "")
                    continue;

                if (altUom == "")
                    _invalidAlternateUomGridCells.Add(GridCellKey(rowHandle, colAltUom.FieldName));

                if (baseUom == "")
                    _invalidAlternateUomGridCells.Add(GridCellKey(rowHandle, colBaseUom.FieldName));

                if (ToDecimal(gridAlternateUomView.GetRowCellValue(rowHandle, colAltDenominator), 1m) == 0m)
                    _invalidAlternateUomGridCells.Add(GridCellKey(rowHandle, colAltDenominator.FieldName));

                if (altUom != "")
                {
                    int firstRowHandle;
                    if (altUomRows.TryGetValue(altUom, out firstRowHandle))
                    {
                        _invalidAlternateUomGridCells.Add(GridCellKey(firstRowHandle, colAltUom.FieldName));
                        _invalidAlternateUomGridCells.Add(GridCellKey(rowHandle, colAltUom.FieldName));
                    }
                    else
                    {
                        altUomRows[altUom] = rowHandle;
                    }
                }
            }

            gridAlternateUomView.RefreshData();
            if (_invalidAlternateUomGridCells.Count > 0)
                MarkTabInvalid(tabPageAlternateUom);

            return _invalidAlternateUomGridCells.Count == 0;
        }

        private string GridCellText(object value)
        {
            if (value == null || value == DBNull.Value) return string.Empty;
            return value.ToString().Trim();
        }

        private string GridCellKey(int rowHandle, string fieldName)
        {
            return rowHandle.ToString(CultureInfo.InvariantCulture) + "|" + (fieldName ?? string.Empty);
        }

        private void ProductGridView_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (_invalidSizeGridCells.Contains(GridCellKey(e.RowHandle, e.Column.FieldName)))
                e.Appearance.BackColor = Color.MistyRose;
        }

        private void AlternateUomGridView_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (_invalidAlternateUomGridCells.Contains(GridCellKey(e.RowHandle, e.Column.FieldName)))
                e.Appearance.BackColor = Color.MistyRose;
        }

        private void ProductGridView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            ValidateSizeGrid();
            RefreshValidationTabColors(false);
        }

        private void CKBatchTidak_CheckedChanged(object sender, EventArgs e)
        {
            ValidateSingleControlAndUpdateUi(txtMimimum);
            RefreshValidationTabColors(false);
        }

        private class RequiredControlRule : ValidationRule
        {
            public override bool Validate(Control control, object value)
            {
                return frmINVProductMasterEditor.HasValue(control, value);
            }
        }

        private class MinimumShelfLifeRule : ValidationRule
        {
            private readonly frmINVProductMasterEditor _owner;

            public MinimumShelfLifeRule(frmINVProductMasterEditor owner)
            {
                _owner = owner;
            }

            public override bool Validate(Control control, object value)
            {
                if (_owner == null || !_owner.CKBatchTidak.Checked) return true;
                if (!frmINVProductMasterEditor.HasValue(control, value)) return false;

                decimal parsed;
                if (!decimal.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.CurrentCulture, out parsed) &&
                    !decimal.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed))
                    return false;

                return parsed != 0m;
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

        #region events save change
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!ValidateProductMasterForm())
                return;

            if (clsGlobal.MODE_TRX == 1)//new
            {
                SaveNew();
            }
            else if (clsGlobal.MODE_TRX == 2)//edit
            {
                SaveEdit();
            }
        }
        #endregion

        #region textbox keypress or down
        private void txtProdMstrCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && __apparelSys)
            {
                strSQL = "";
                strSQL = " select (max(vw.id) + 1) as id from ( SELECT  cast(prm_prd_master_code as integer) as id from IM_PRD_MASTER where SUBSTRING(prm_prd_master_code,1,2)  ='11') vw ";

                DataTable dt = _clsGlobal.ExecDT(strSQL);

                if (dt.Rows.Count > 0)
                {
                    txtProdMstrCd.Text = dt.Rows[0]["id"].ToString().Trim();
                }


                SetFirstSizeGridRowSelected(true);
            }
        }
        private void CBGroupCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGridSizeGrp();
        }

        private void gridAlternateUomView_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
        {
            gridAlternateUomView.SetRowCellValue(e.RowHandle, colAltDenominator, 1m);
            gridAlternateUomView.SetRowCellValue(e.RowHandle, colAltNumerator, 1m);
            gridAlternateUomView.SetRowCellValue(e.RowHandle, colAltEquivalent, "<=>");
            gridAlternateUomView.SetRowCellValue(e.RowHandle, colBaseUom, CurrentBaseUom());
            gridAlternateUomView.SetRowCellValue(e.RowHandle, colAltIsBase, false);
            gridAlternateUomView.SetRowCellValue(e.RowHandle, colAltIsActive, true);
        }

        private void gridAlternateUomView_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column == colAltUom)
            {
                string altUom = (e.Value == null || e.Value == DBNull.Value) ? string.Empty : e.Value.ToString().Trim();
                gridAlternateUomView.SetRowCellValue(e.RowHandle, colMeasurementUom, GetUomDescription(altUom));
            }

            object equivalent = gridAlternateUomView.GetRowCellValue(e.RowHandle, colAltEquivalent);
            if (equivalent == null || equivalent == DBNull.Value || equivalent.ToString().Trim() == "")
                gridAlternateUomView.SetRowCellValue(e.RowHandle, colAltEquivalent, "<=>");

            BeginInvoke(new Action(delegate
            {
                ValidateAlternateUomGrid();
                RefreshValidationTabColors(false);
            }));
        }

        private void gridAlternateUomView_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
            // The search-button columns are unbound; they hold no persistent value.
            if (e.Column == colAltUomSearchButton ||
                e.Column == colBaseUomSearchButton)
            {
                if (e.IsGetData)
                    e.Value = string.Empty;
                // IsSetData: ignore – value is transferred to the unit column instead.
            }
        }

        private void repositoryItemUomSearchButton_EditValueChanged(object sender, EventArgs e)
        {
            if (__suppressUomSearchButtonEvent) return;

            SearchLookUpEdit editor = sender as SearchLookUpEdit;
            if (editor == null) return;

            string picked = (editor.EditValue == null || editor.EditValue == DBNull.Value)
                                ? string.Empty
                                : editor.EditValue.ToString().Trim();
            if (picked == string.Empty) return;

            int rowHandle = gridAlternateUomView.FocusedRowHandle;
            DevExpress.XtraGrid.Columns.GridColumn focused = gridAlternateUomView.FocusedColumn;

            DevExpress.XtraGrid.Columns.GridColumn target =
                (focused == colAltUomSearchButton) ? colAltUom :
                (focused == colBaseUomSearchButton) ? colBaseUom : null;

            if (target == null) return;

            __suppressUomSearchButtonEvent = true;
            try
            {
                gridAlternateUomView.SetRowCellValue(rowHandle, target, picked);

                if (target == colAltUom)
                {
                    gridAlternateUomView.SetRowCellValue(rowHandle, colMeasurementUom, GetUomDescription(picked));
                }
            }
            finally
            {
                __suppressUomSearchButtonEvent = false;
            }

            // Clear the helper editor (after this event) so the same code can be picked again
            // and the unbound cell stays blank.
            BeginInvoke(new Action(delegate
            {
                if (editor.IsDisposed) return;
                __suppressUomSearchButtonEvent = true;
                try
                {
                    if (editor.IsPopupOpen) editor.ClosePopup();
                    editor.EditValue = null;
                    gridAlternateUomView.CloseEditor();
                }
                catch { }
                finally { __suppressUomSearchButtonEvent = false; }
            }));
        }

        private void gridAlternateUomView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && gridAlternateUomView.FocusedRowHandle >= 0)
            {
                gridAlternateUomView.DeleteSelectedRows();
                e.Handled = true;
            }
        }

        private void textGrade_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtNoOfItem_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txttaxcode_Leave(object sender, EventArgs e)
        {

            // taxcodetb0.Text = "";

        }
        private void txtWeight_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtvlme1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtvlme2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtMnmnItm_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtPknAwl_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtSftyBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtStockMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtStockMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtbrcd_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        #endregion

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void photoBtn_Click(object sender, EventArgs e)
        {

            if (__browseWorker.IsBusy) return;
            String SQLStr = "";
            try
            {
                TIRAFTP __im10 = new TIRAFTP();
                SQLStr = "SELECT gh_function_code FROM GS_GEN_HARDCODED with(nolock) WHERE gh_sys = 'H' and gh_function_name = 'FTP_PRODUCT_PHOTO' and gh_sequence_no = 1";
                DataTable dtf = _clsGlobal.ExecDT(SQLStr);
                if (dtf.Rows.Count > 0)
                {
                    __im10.host = dtf.Rows[0]["gh_function_code"].ToString().Trim();
                }
                else
                {

                    MessageBox.Show("Setting FTP belum ada", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                SQLStr = "SELECT gh_function_code FROM GS_GEN_HARDCODED with(nolock) WHERE gh_sys = 'H' and gh_function_name = 'FTP_PRODUCT_PHOTO' and gh_sequence_no = 2";
                DataTable dtg = _clsGlobal.ExecDT(SQLStr);
                if (dtg.Rows.Count > 0)
                {
                    __im10.user = dtg.Rows[0]["gh_function_code"].ToString().Trim();
                }

                SQLStr = "SELECT gh_function_code FROM GS_GEN_HARDCODED with(nolock) WHERE gh_sys = 'H' and gh_function_name = 'FTP_PRODUCT_PHOTO' and gh_sequence_no = 3";
                DataTable dth = _clsGlobal.ExecDT(SQLStr);
                if (dth.Rows.Count > 0)
                {
                    __im10.pass = dth.Rows[0]["gh_function_code"].ToString().Trim();
                }

                SQLStr = "SELECT gh_function_code FROM GS_GEN_HARDCODED with(nolock) WHERE gh_sys = 'H' and gh_function_name = 'FTP_PRODUCT_PHOTO' and gh_sequence_no = 4";
                DataTable dti = _clsGlobal.ExecDT(SQLStr);
                if (dti.Rows.Count > 0)
                {
                    __im10.directory = dti.Rows[0]["gh_function_code"].ToString().Trim();
                }
                else
                {

                    MessageBox.Show("Belum ada setting folder upload pada ftp", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }




                using (OpenFileDialog __ofdlog = new OpenFileDialog()
                {
                    Filter = "Image Files (*.JPG)|*.JPG",
                    FilterIndex = 1,
                    Title = "Select file",
                })
                    if (__ofdlog.ShowDialog() == DialogResult.OK)
                    {
                        //StringBuilder sb = new StringBuilder();
                        //sb.Append(__ofdlog.FileName);
                        //sb.Append(".");
                        //sb.Append(Path.GetExtension(__ofdlog.FileName));
                        //__im10.filename = sb.ToString();
                        __im10.filename = __ofdlog.FileName;
                        __im10.file_extension = Path.GetExtension(__ofdlog.FileName);
                        this.txtPhoto.Text = __im10.filename;


                        if (System.IO.Path.GetFileNameWithoutExtension(__ofdlog.FileName) != txtProdMstrCd.Text.Trim())
                        {
                            MessageBox.Show("Nama file photo product berbeda dengan kode product", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        __circle.Start();
                        __browseWorker.RunWorkerAsync(__im10);
                    }
            }
            catch (Exception __exc)
            {
                MessageBox.Show(__exc.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

    }

    public class TIRAItem
    {
        public string Kode { get; set; }
        public string Nama { get; set; }
    }

    public class TIRAFTP
    {
        public string host { get; set; }
        public string directory { get; set; }
        public string user { get; set; }
        public string pass { get; set; }
        public string filename { get; set; }
        public string file_extension { get; set; }
        public bool IsSuccess { get; set; }
        public object State { get; set; }
        public string remote_path { get; set; }
        public string Message { get; set; }
    }

    public class VendorItem
    {
        public string ID1 { get; set; }
        public string ID2 { get; set; }
        public string Nama { get; set; }
    }

    internal static class TIRAObjectDevExpressButtonAdapter
    {
        public static void AddRangeButton<T>(this TIRAObject<T> source, SearchLookUpEdit searchButton) where T : class, new()
        {
            if (source == null || searchButton == null)
                return;

            Button hiddenButton = new Button();
            hiddenButton.Visible = false;
            hiddenButton.Width = 0;
            hiddenButton.Height = 0;

            source.AddRangeButton(hiddenButton);

            searchButton.ButtonClick += delegate
            {
                hiddenButton.PerformClick();
            };

            searchButton.Click += delegate
            {
                hiddenButton.PerformClick();
            };
        }
    }

}

