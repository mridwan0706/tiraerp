using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace TIRASnDNet.GS.GSProductSubstitution
{
    public partial class frmProductSubsList : TIRASnDNet.BaseListForm
    {
        private clsGlobal _clsGlobal = new clsGlobal();
        private string SQLStr;
        private string paramMenuId;
        private string captionForm;
        private DataTable dt = new DataTable();

        const string copyCell = "Copy Cell";
        const string copyRow = "Copy Row";
        const string copyAll = "Copy All Grid to Excel/Notepad";
        const string deleteAll = "Delete All";
        const int maxCopy = 10000;

        public static string Entity;
        public static string Branch;

        public frmProductSubsList()
        {
            InitializeComponent();
        }

        private string GetLookupValue(SearchLookUpEdit lookup)
        {
            return Convert.ToString(lookup.EditValue ?? string.Empty).Trim();
        }

        private void SetLookupValue(SearchLookUpEdit lookup, object value)
        {
            lookup.EditValue = value == null ? null : Convert.ToString(value).Trim();
        }

        private void SetLookupDescription(SearchLookUpEdit lookup, TextEdit descriptionText, params string[] descriptionColumns)
        {
            descriptionText.Text = string.Empty;

            DataRowView row = lookup.Properties.View.GetFocusedRow() as DataRowView;
            if (row == null)
                return;

            foreach (string columnName in descriptionColumns)
            {
                if (row.Row.Table.Columns.Contains(columnName))
                {
                    descriptionText.Text = Convert.ToString(row[columnName]).Trim();
                    return;
                }
            }
        }

        private void SetGridView()
        {
            gvList.Columns.Clear();
            gvList.OptionsBehavior.Editable = false;
            gvList.OptionsView.ShowGroupPanel = false;
            gvList.OptionsView.ColumnAutoWidth = false;
            gvList.OptionsSelection.EnableAppearanceFocusedCell = false;
            gvList.OptionsSelection.MultiSelect = false;

            gvList.Columns.AddVisible("tiranumb", "No").Width = 40;
            gvList.Columns.AddVisible("prds_line", "Product Line").Width = 100;
            gvList.Columns.AddVisible("prds_index", "Index").Width = 70;
            gvList.Columns.AddVisible("prds_parent", "Product Parent").Width = 120;
            gvList.Columns.AddVisible("prds_parent_desc", "Parent Description").Width = 200;            
            gvList.Columns.AddVisible("prds_child", "Product Child").Width = 150;
            gvList.Columns.AddVisible("prds_child_desc", "Child Description").Width = 200;

            if (gvList.Columns["tiranumb"] != null)
                gvList.Columns["tiranumb"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            if (gvList.Columns["prds_line"] != null)
                gvList.Columns["prds_line"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            if (gvList.Columns["prds_index"] != null)
                gvList.Columns["prds_index"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            if (gvList.Columns["prds_parent"] != null)
                gvList.Columns["prds_parent"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            
        }

        private void frmProductSubsList_Load(object sender, EventArgs e)
        {
            LoadEntityLookup();
            getPilihanEntity();
            LoadBranchLookup();
            getDefBranch();
            LoadProductLineLookup();
            LoadProductParentLookup();
            LoadProductChildLookup();

            this.Text = clsLogin.MENUID + " - " + clsLogin.MENUNAME + " " + this.Text;
            captionForm = clsLogin.MENUID + " - " + clsLogin.MENUNAME;
            paramMenuId = clsLogin.MENUID;
            SetGridView();
            AccessButton();
        }

        protected override void OnNew()
        {
            tsb_new_Click(this, EventArgs.Empty);
        }

        protected override void OnDelete()
        {
            tsb_delete_Click(this, EventArgs.Empty);
        }

        protected override void OnCloseForm()
        {
            this.Close();
        }

        private void ConfigureLookup(SearchLookUpEdit lookup, DataTable source, string valueMember, string displayMember)
        {
            lookup.Properties.DataSource = source;
            lookup.Properties.ValueMember = valueMember;
            lookup.Properties.DisplayMember = displayMember;
            lookup.Properties.NullText = string.Empty;
            lookup.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.True;
            lookup.Properties.ImmediatePopup = true;
            lookup.Properties.PopupFilterMode = DevExpress.XtraEditors.PopupFilterMode.Contains;
            lookup.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;

            GridView view = lookup.Properties.View;
            view.Columns.Clear();
            foreach (DataColumn col in source.Columns)
            {
                view.Columns.AddVisible(col.ColumnName, col.ColumnName);
            }
            view.OptionsView.ShowAutoFilterRow = true;
            view.OptionsView.ShowGroupPanel = false;
            view.BestFitColumns();
        }

        private void LoadEntityLookup()
        {
            SQLStr = "select ge_entity_id [Entity], ge_entity [Desc] from GS_ENTITY WITH(NOLOCK) inner join GL_USER_ENTITY WITH(NOLOCK) on ge_entity_id = ue_entity_id where ue_user_id = '" + clsLogin.USERID + "' order by ge_entity_id";
            DataTable source = _clsGlobal.ExecDT(SQLStr);
            ConfigureLookup(txtEntityCode, source, "Entity", "Entity");
        }

        private void LoadBranchLookup()
        {
            SQLStr = "select gu_branch [Branch], br_branch_desc [Branch Desc] FROM VW_USERS_SECURITY WITH(NOLOCK) where gu_entity = '" + GetLookupValue(txtEntityCode) + "' and gu_user_id = '" + clsLogin.USERID + "'";
            DataTable source = _clsGlobal.ExecDT(SQLStr);
            ConfigureLookup(txtBranchCode, source, "Branch", "Branch");
        }

        private void LoadProductLineLookup()
        {
            SQLStr = $@"SELECT distinct prds_line [Kode Product Line], pl_prd_line_desc [Description] FROM TBL_PRD_SUBS WITH(NOLOCK)
                        INNER JOIN IM_PRD_LINE WITH (NOLOCK) ON pl_prd_line_code=prds_line
                    WHERE prds_entity_id = '" + GetLookupValue(txtEntityCode) + "' AND prds_branch_id = '" + GetLookupValue(txtBranchCode) + "' order by prds_Line";
            DataTable source = _clsGlobal.ExecDT(SQLStr);
            ConfigureLookup(txtProdline, source, "Kode Product Line", "Kode Product Line");
        }

        private void LoadProductParentLookup()
        {
            SQLStr = $@"SELECT distinct prds_parent [Kode Product Parent], prm_prd_desc [Description] FROM TBL_PRD_SUBS WITH (NOLOCK)
             LEFT JOIN IM_PRD_MASTER WITH (NOLOCK) ON prds_parent=prm_prd_master_code and prds_grade=prm_grade and prds_size=prm_prd_size
             WHERE prds_entity_id = '" + GetLookupValue(txtEntityCode) + "' AND prds_branch_id = '" + GetLookupValue(txtBranchCode) + "' AND prds_Line = '" + GetLookupValue(txtProdline) + "' order by prds_parent";
            DataTable source = _clsGlobal.ExecDT(SQLStr);
            ConfigureLookup(txtProdParent, source, "Kode Product Parent", "Kode Product Parent");
        }

        private void LoadProductChildLookup()
        {
            SQLStr = $@"SELECT distinct prds_child [Kode Product Child], prds_child [Description] FROM TBL_PRD_SUBS WITH(NOLOCK)
                    LEFT JOIN IM_PRD_MASTER WITH (NOLOCK) ON prds_child=prm_prd_master_code and prds_grade=prm_grade and prds_size=prm_prd_size
                    WHERE prds_entity_id = '" + GetLookupValue(txtEntityCode) + "' AND prds_branch_id = '" + GetLookupValue(txtBranchCode) + "' AND prds_Line = '" + GetLookupValue(txtProdline) + "' AND prds_parent = '" + GetLookupValue(txtProdParent) + "' order by prds_child";
            DataTable source = _clsGlobal.ExecDT(SQLStr);
            ConfigureLookup(txtProdChild, source, "Kode Product Child", "Kode Product Child");
        }

        public void getDefBranch()
        {
            try
            {
                SQLStr = "select gu_branch [Branch], br_branch_desc [Branch Desc] FROM VW_USERS_SECURITY with(nolock) where gu_entity = '" + GetLookupValue(txtEntityCode) + "' and gu_user_id = '" + clsLogin.USERID + "' ";

                DataTable dt = _clsGlobal.ExecDT(SQLStr);

                if (dt.Rows.Count > 0)
                {
                    SetLookupValue(txtBranchCode, dt.Rows[0]["Branch"]);
                    txtBranchDesc.Text = dt.Rows[0]["Branch Desc"].ToString().Trim();
                }
                else
                {
                    SetLookupValue(txtBranchCode, null);
                    txtBranchDesc.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void getPilihanEntity()
        {
            try
            {
                SQLStr = " select gu_entity, CONCAT(gu_entity, ' - ',ge_entity) AS ge_entity from VW_USERS_SECURITY WITH(NOLOCK) where gu_user_id = '" + clsLogin.USERID + "'";

                DataTable dt = _clsGlobal.ExecDT(SQLStr);

                if (dt.Rows.Count > 0)
                {
                    SetLookupValue(txtEntityCode, dt.Rows[0]["gu_entity"]);
                    txtEntityDesc.Text = dt.Rows[0]["ge_entity"].ToString().Trim();
                }
                else
                {
                    SetLookupValue(txtEntityCode, null);
                    txtEntityDesc.Text = string.Empty;
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnlistData_Click(object sender, EventArgs e)
        {
            if (GetLookupValue(txtEntityCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Entity terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtEntityCode.Select();
                return;
            }

            if (GetLookupValue(txtBranchCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Branch terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtBranchCode.Select();
                return;
            }

            FillGrid();
        }

        private void FillGrid()
        {
            StringBuilder sql = new StringBuilder();
            sql.Append("SELECT prds_entity_id, prds_branch_id, prds_line, prds_parent, parent.prm_prd_desc [prds_parent_desc],  prds_child,child.prm_prd_desc [prds_child_desc], prds_index ");
            sql.Append(",prds_valid_from,prds_valid_to,prds_mrp_ind ");
            sql.Append("FROM TBL_PRD_SUBS WITH(NOLOCK) ");
            sql.Append("LEFT JOIN IM_PRD_LINE WITH(NOLOCK) ON prds_line=pl_prd_line_code ");
            sql.Append("LEFT JOIN IM_PRD_MASTER  parent WITH (NOLOCK) ON  prds_parent=parent.prm_prd_master_code ");
            sql.Append("LEFT JOIN IM_PRD_MASTER child WITH (NOLOCK) ON  prds_child=child.prm_prd_master_code ");
            sql.Append("WHERE prds_entity_id = '").Append(GetLookupValue(txtEntityCode)).Append("' AND prds_branch_id = '").Append(GetLookupValue(txtBranchCode)).Append("' ");

            if (GetLookupValue(txtProdline).Length > 0)
                sql.Append("and prds_line = '").Append(GetLookupValue(txtProdline)).Append("' ");

            if (GetLookupValue(txtProdParent).Length > 0)
                sql.Append("and prds_parent = '").Append(GetLookupValue(txtProdParent)).Append("' ");

            if (GetLookupValue(txtProdChild).Length > 0)
                sql.Append("and prds_child = '").Append(GetLookupValue(txtProdChild)).Append("' ");

            if (dateValidFrom.EditValue != null)
                sql.Append("and CONVERT(date, prds_valid_from) >= '").Append(dateValidFrom.DateTime.ToString("yyyy-MM-dd")).Append("' ");

            if (dateValidTo.EditValue != null)
                sql.Append("and CONVERT(date, prds_valid_to) <= '").Append(dateValidTo.DateTime.ToString("yyyy-MM-dd")).Append("' ");

            sql.Append("order by prds_entity_id, prds_branch_id, prds_line, prds_parent, prds_index, prds_child ");
            SQLStr = sql.ToString();

            gcList.DataSource = _clsGlobal.ExecDTWithAutoIncrement(SQLStr);
            gvList.BestFitColumns();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (GetLookupValue(txtEntityCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Entity terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtEntityCode.Select();
                return;
            }

            if (GetLookupValue(txtBranchCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Branch terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtBranchCode.Select();
                return;
            }
            Entity = GetLookupValue(txtEntityCode);
            Branch = GetLookupValue(txtBranchCode);

            frmProductSubsEditor frm = new frmProductSubsEditor();
            frm.Text = captionForm + " " + frm.Text;
            frm.ShowDialog();
        }

        private void gvList_PopupMenuShowing(object sender, DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs e)
        {
            if (e.MenuType != DevExpress.XtraGrid.Views.Grid.GridMenuType.Row) return;

            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add(copyCell, null, ToolStripMenu_Click);
            menu.Items.Add(copyRow, null, ToolStripMenu_Click);
            menu.Items.Add(copyAll, null, ToolStripMenu_Click);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(deleteAll, null, ToolStripMenu_Click);
            menu.Show(gcList, e.Point);
        }

        private void ToolStripMenu_Click(object sender, EventArgs e)
        {
            string action = sender.ToString();
            Clipboard.Clear();

            if (action == copyCell)
            {
                object val = gvList.GetFocusedValue();
                Clipboard.SetText(Convert.ToString(val));
            }
            else if (action == copyRow)
            {
                DataRow row = gvList.GetFocusedDataRow();
                if (row == null) return;

                StringBuilder selectedRow = new StringBuilder();
                foreach (DataColumn col in row.Table.Columns)
                    selectedRow.Append(Convert.ToString(row[col])).Append("\t");
                Clipboard.SetText(selectedRow.ToString());
            }
            else if (action == copyAll)
            {
                DataTable source = gcList.DataSource as DataTable;
                if (source == null) return;

                StringBuilder selectedAll = new StringBuilder();
                foreach (DataColumn col in source.Columns)
                    selectedAll.Append(col.ColumnName).Append("\t");
                selectedAll.Append(Environment.NewLine);

                int rowCount = Math.Min(source.Rows.Count, maxCopy);
                for (int i = 0; i < rowCount; i++)
                {
                    foreach (DataColumn col in source.Columns)
                        selectedAll.Append(Convert.ToString(source.Rows[i][col])).Append("\t");
                    selectedAll.Append(Environment.NewLine);
                }

                Clipboard.SetText(selectedAll.ToString());

                if (source.Rows.Count < maxCopy)
                    XtraMessageBox.Show("Sukses Copy to Clipboard", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    XtraMessageBox.Show("Maaf hanya 10.000 baris yang bisa di copy", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (action == deleteAll)
            {
                if (XtraMessageBox.Show("Apakah Anda Yakin ingin menghapus Semua Data Product Substitusi?", "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel) return;
                SQLStr = "DELETE FROM TBL_PRD_SUBS ";
                SQLStr += "WHERE prds_entity_id = '" + GetLookupValue(txtEntityCode) + "' ";
                SQLStr += "AND prds_branch_id = '" + GetLookupValue(txtBranchCode) + "' ";
                _clsGlobal.Execute(SQLStr);
                FillGrid();
            }
        }

        private void tsb_delete_Click(object sender, EventArgs e)
        {
            if (GetLookupValue(txtEntityCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Entity terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtEntityCode.Select();
                return;
            }

            if (GetLookupValue(txtBranchCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Branch terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtBranchCode.Select();
                return;
            }

            DataRow row = gvList.GetFocusedDataRow();
            if (row == null)
            {
                XtraMessageBox.Show("Silahkan pilih data terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SQLStr = "DELETE FROM TBL_PRD_SUBS WHERE prds_entity_id = '" + GetLookupValue(txtEntityCode) + "' ";
            SQLStr += " AND prds_branch_id = '" + GetLookupValue(txtBranchCode) + "' ";
            SQLStr += " AND prds_entity_id ='" + GetLookupValue(txtEntityCode) + "'";
            SQLStr += " AND prds_line ='" + Convert.ToString(row["prds_line"]) + "'";
            SQLStr += " AND prds_parent ='" + Convert.ToString(row["prds_parent"]) + "'";
            SQLStr += " AND prds_index ='" + Convert.ToString(row["prds_index"]) + "'";
            SQLStr += " AND prds_child ='" + Convert.ToString(row["prds_child"]) + "'";
            _clsGlobal.Execute(SQLStr);

            XtraMessageBox.Show("Delete Product Success", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
            FillGrid();
        }

        private void tsb_new_Click(object sender, EventArgs e)
        {
            if (GetLookupValue(txtEntityCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Entity terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtEntityCode.Select();
                return;
            }

            if (GetLookupValue(txtBranchCode).Length < 1)
            {
                XtraMessageBox.Show("Silahkan pilih Branch terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtBranchCode.Select();
                return;
            }
            clsGlobal.MODE_TRX = 1;
            frmProductSubsNewEditor frm = new frmProductSubsNewEditor();
            frm.ShowDialog();
        }

        private void txtEntityCode_EditValueChanged(object sender, EventArgs e)
        {
            DataRowView row = txtEntityCode.Properties.View.GetFocusedRow() as DataRowView;
            txtEntityDesc.Text = row == null ? string.Empty : Convert.ToString(row["Desc"]);
            LoadBranchLookup();
            SetLookupValue(txtBranchCode, null);
            txtBranchDesc.Text = string.Empty;
        }

        private void txtBranchCode_EditValueChanged(object sender, EventArgs e)
        {
            if (txtBranchCode.EditValue == null || string.IsNullOrWhiteSpace(txtBranchCode.EditValue.ToString()))
            {
                txtBranchDesc.EditValue = null;

                SetLookupValue(txtProdline, null);
                txtProdlineDesc.EditValue = null;

                SetLookupValue(txtProdParent, null);
                txtProdParentDesc.EditValue = null;

                SetLookupValue(txtProdChild, null);
                txtProdChildDesc.EditValue = null;

                return;
            }

            DataRowView row = txtBranchCode.Properties.View.GetFocusedRow() as DataRowView;
            txtBranchDesc.EditValue = row == null ? null : Convert.ToString(row["Branch Desc"]);

            LoadProductLineLookup();
            SetLookupValue(txtProdline, null);
            txtProdlineDesc.EditValue = null;

            SetLookupValue(txtProdParent, null);
            txtProdParentDesc.EditValue = null;

            SetLookupValue(txtProdChild, null);
            txtProdChildDesc.EditValue = null;
        }

        private void txtProdline_EditValueChanged(object sender, EventArgs e)
        {
            if (txtProdline.EditValue == null || string.IsNullOrWhiteSpace(txtProdline.EditValue.ToString()))
            {
                // clear semua desc
                txtProdlineDesc.EditValue = null;

                SetLookupValue(txtProdParent, null);
                txtProdParentDesc.EditValue = null;

                SetLookupValue(txtProdChild, null);
                txtProdChildDesc.EditValue = null;

                return;
            }


            SetLookupDescription(txtProdline, txtProdlineDesc, "Description", "Kode Product Line");
            LoadProductParentLookup();
            SetLookupValue(txtProdParent, null);
            txtProdParentDesc.Text = string.Empty;
            SetLookupValue(txtProdChild, null);
            txtProdChildDesc.Text = string.Empty;
        }

        private void txtProdParent_EditValueChanged(object sender, EventArgs e)
        {

            if (txtProdParent.EditValue == null || string.IsNullOrWhiteSpace(txtProdParent.EditValue.ToString()))
            {
                txtProdParentDesc.Text = string.Empty;
                SetLookupValue(txtProdChild, null);
                txtProdChildDesc.Text = string.Empty;

                return;
            }


            SetLookupDescription(txtProdParent, txtProdParentDesc, "Description", "Kode Product Parent");
            LoadProductChildLookup();
            SetLookupValue(txtProdChild, null);
            txtProdChildDesc.Text = string.Empty;
        }

        private void txtProdChild_EditValueChanged(object sender, EventArgs e)
        {
            SetLookupDescription(txtProdChild, txtProdChildDesc, "Description", "Kode Product Child");
        }
    }
}
