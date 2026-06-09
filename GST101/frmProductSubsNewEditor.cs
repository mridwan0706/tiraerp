using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TIRASnDNet.GS.GSProductSubstitution
{
    public partial class frmProductSubsNewEditor : Form
    {
        private clsGlobal _clsGlobal = new clsGlobal();
        private string SQLStr;
        string strprdline;
        string strprdparent;
        string strprdindex;
        string strprdchild;
        string strKeterangan;
        private DataTable dt = new DataTable();
        private int intKet = 0;
        public frmProductSubsNewEditor()
        {
            InitializeComponent();
        }

        private void btnPopUpEntityCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Entity";
            frm.Query = "select ge_entity_id [Entity], ge_entity [Desc]  from GS_ENTITY WITH(NOLOCK) inner join GL_USER_ENTITY WITH(NOLOCK) on ge_entity_id = ue_entity_id where ue_user_id = '" + clsLogin.USERID + "' order by ge_entity_id ";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtEntityCode.Text = frm.ArrField[0].Trim();
                txtEntityDesc.Text = frm.ArrField[1].Trim();
            }
        }

        private void btnPopUpBranchCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Branch";
            frm.Query = "select gu_branch [Branch], br_branch_desc [Branch Desc] FROM VW_USERS_SECURITY where gu_entity = '" + txtEntityCode.Text.Trim() + "' and gu_user_id = '" + clsLogin.USERID + "'";
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtBranchCode.Text = frm.ArrField[0].Trim();
                txtBranchDesc.Text = frm.ArrField[1].Trim();
            }
        }
        
        private void btnPopUpPrdLineCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Product Line";
            SQLStr = " SELECT pl_prd_line_code [Kode Product Line], ";
            SQLStr += " pl_prd_line_desc [Description Line] ";
            SQLStr += " FROM IM_PRD_LINE ";
            //SQLStr += "WHERE prds_entity_id = '" + txtEntityCode.Text.Trim() + "' ";
            //SQLStr += "AND prds_branch_id = '" + txtBranchCode.Text.Trim() + "' ";
            SQLStr += " order by pl_prd_line_code";
            frm.Query = SQLStr;
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtProdline.Text = frm.ArrField[0].Trim();
                txtProdlineDesc.Text = frm.ArrField[1].Trim();
            }

        }

        private void btnPopUpPrdParentCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Product Parent";
            SQLStr = " SELECT distinct prm_prd_master_code [Kode Product Parent], ";
            SQLStr += " prm_prd_desc [Description Product Parent] ";
            SQLStr += " FROM IM_PRD_MASTER ";
            //SQLStr += "WHERE prds_entity_id = '" + txtEntityCode.Text.Trim() + "' ";
            //SQLStr += "AND prds_branch_id = '" + txtBranchCode.Text.Trim() + "' ";
            SQLStr += "WHERE prm_prd_line_code = '" + txtProdline.Text.Trim() + "' ";
            SQLStr += " order by prm_prd_master_code";
            frm.Query = SQLStr;
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtProdParent.Text = frm.ArrField[0].Trim();
                txtProdParentDesc.Text = frm.ArrField[1].Trim();
            }
        }

        private void btnPopUpPrdChildCode_Click(object sender, EventArgs e)
        {
            frmPopUp frm = new frmPopUp();
            frm.FrmText = "Search Product Child";
            SQLStr = " SELECT distinct prm_prd_master_code [Kode Product Child], ";
            SQLStr += " prm_prd_desc [Description Product Child] ";
            SQLStr += " FROM IM_PRD_MASTER ";
            //SQLStr += "WHERE prds_entity_id = '" + txtEntityCode.Text.Trim() + "' ";
            //SQLStr += "AND prds_branch_id = '" + txtBranchCode.Text.Trim() + "' ";
            SQLStr += "WHERE prm_prd_line_code = '" + txtProdline.Text.Trim() + "' ";
            //SQLStr += "AND prds_parent = '" + txtProdParent.Text.Trim() + "' ";
            SQLStr += " order by prm_prd_master_code";
            frm.Query = SQLStr;
            frm.ShowDialog();
            if (frm.ArrField != null)
            {
                txtProdChild.Text = frm.ArrField[0].Trim();
                txtProdChildDesc.Text = frm.ArrField[1].Trim();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtEntityCode.Text.Length < 1)
                {
                    MessageBox.Show("Silahkan pilih Entity terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtEntityCode.Select();
                    return;
                }

                if (txtBranchCode.Text.Length < 1)
                {
                    MessageBox.Show("Silahkan pilih Branch terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtBranchCode.Select();
                    return;
                }

                if (txtProdline.Text.Length < 1)
                {
                    MessageBox.Show("Silahkan pilih Line terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtProdline.Select();
                    return;
                }

                if (txtProdParent.Text.Length < 1)
                {
                    MessageBox.Show("Silahkan pilih Product Parent terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtProdParent.Select();
                    return;
                }

                if (txtProdChild.Text.Length < 1)
                {
                    MessageBox.Show("Silahkan pilih Product Child terlebih dahulu", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtProdChild.Select();
                    return;
                }

                strprdline = txtProdline.Text.Trim();
                strprdparent = txtProdParent.Text.Trim();
                strprdchild = txtProdChild.Text.Trim();

                SQLStr = "select pl_prd_line_code from IM_PRD_LINE where pl_prd_line_code = '" + strprdline + "' ";
                dt = _clsGlobal.ExecDT(SQLStr);

                if (dt.Rows.Count < 1)
                {
                    strKeterangan = "Kode Produk Line tidak terdaftar";
                    intKet = intKet + 1;
                }

                SQLStr = "select prm_prd_master_code from IM_PRD_MASTER where prm_prd_master_code = '" + strprdparent + "' ";
                dt = _clsGlobal.ExecDT(SQLStr);

                if (dt.Rows.Count < 1)
                {
                    strKeterangan = "Kode Produk Parent tidak terdaftar";
                    intKet = intKet + 1;
                }

                SQLStr = "select prm_prd_master_code from IM_PRD_MASTER where prm_prd_master_code = '" + strprdchild + "' ";
                dt = _clsGlobal.ExecDT(SQLStr);

                if (dt.Rows.Count < 1)
                {
                    strKeterangan = "Kode Produk Child tidak terdaftar";
                    intKet = intKet + 1;
                }

                if(intKet > 1)
                {
                    MessageBox.Show(strKeterangan, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                SQLStr = "Select MAX(prds_index) as Jumlah from TBL_PRD_SUBS where prds_line= '" + txtProdline.Text.Trim() + "' ";
                SQLStr += "AND prds_parent = '" + txtProdParent.Text.Trim() + "' and prds_entity_id = '" + txtEntityCode.Text.Trim() + "' and prds_branch_id = '" + txtBranchCode.Text.Trim() + "' ";
                dt = _clsGlobal.ExecDT(SQLStr);

                if (dt.Rows[0][0].ToString() != "")
                {
                    int index = Convert.ToInt32(dt.Rows[0][0].ToString());
                    index += 1;

                    SQLStr = "insert into TBL_PRD_SUBS ( prds_entity_id, prds_branch_id, prds_line, prds_parent, ";
                    SQLStr += "prds_index, prds_child, prds_grade, prds_size, ";
                    SQLStr += "prds_created_date, prds_created_user)";
                    SQLStr += "values( ";
                    SQLStr += "'" + txtEntityCode.Text.Trim() + "', ";
                    SQLStr += "'" + txtBranchCode.Text.Trim() + "', ";
                    SQLStr += "'" + txtProdline.Text.Trim() + "', ";  //'prds_line
                    SQLStr += "'" + txtProdParent.Text.Trim() + "', ";  //'prds_parent
                    SQLStr += "'" + index.ToString("D2") + "', ";  //'prds_index
                    SQLStr += "'" + txtProdChild.Text.Trim() + "', ";  //'prds_child
                    SQLStr += "'00',";  //'prds_grade
                    SQLStr += "'00',";  //'prds_size
                    SQLStr += "GETDATE(), ";                                           //'prds_created_date
                    SQLStr += "'" + clsLogin.USERID + "' ) ";                          //'prds_created_user

                    _clsGlobal.Execute(SQLStr);

                    MessageBox.Show("Add Product Substitution", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);

                    return;
                }

                else
                {
                    int __sIdx = 1;
                    SQLStr = "insert into TBL_PRD_SUBS ( prds_entity_id, prds_branch_id, prds_line, prds_parent, ";
                    SQLStr += "prds_index, prds_child, prds_grade, prds_size, ";
                    SQLStr += "prds_created_date, prds_created_user)";
                    SQLStr += "values( ";
                    SQLStr += "'" + txtEntityCode.Text.Trim() + "', ";
                    SQLStr += "'" + txtBranchCode.Text.Trim() + "', ";
                    SQLStr += "'" + txtProdline.Text.Trim() + "', ";  //'prds_line
                    SQLStr += "'" + txtProdParent.Text.Trim() + "', ";  //'prds_parent
                    SQLStr += "'" + ""+ __sIdx.ToString("D2") +"" + "', ";  //'prds_index
                    SQLStr += "'" + txtProdChild.Text.Trim() + "', ";  //'prds_child
                    SQLStr += "'00',";  //'prds_grade
                    SQLStr += "'00',";  //'prds_size
                    SQLStr += "GETDATE(), ";                                           //'prds_created_date
                    SQLStr += "'" + clsLogin.USERID + "' ) ";                          //'prds_created_user

                    _clsGlobal.Execute(SQLStr);

                    MessageBox.Show("Add Product Substitution", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}
