using FluentFTP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;


namespace TIRASnDNet.INV.INVProductMaster
{
    public partial class frmINVProductMasterEditor : Form
    {

        private FtpClient FTP1 = null;
        frmPopUp frm;
        clsGlobal _clsGlobal = new clsGlobal();
        BackgroundWorker __browseWorker = null;        
        LoadingCircle __circle = null;
        
        bool __apparelSys = false;
        private string strSQL, paramMenuId, _prdBrandCode, _Grade, Sizegroupcode;

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
        const string __VENDOR_LOOKUP = "SELECT pvm_vendor_code1 AS [Kode], pvm_vendor_code2 AS [ID2], pvm_vendor_namekey AS [Nama] FROM PO_VENDOR_MASTER WITH(NOLOCK) ORDER BY pvm_vendor_code1";
        const string __PRINCIPAL = "SELECT pri_principal AS [Kode], pri_principal_desc AS [Nama] FROM PO_PRINCIPAL WITH(NOLOCK)";
        const string __TAXCODE = "SELECT t_tax_id AS [Kode], t_tax AS [Nama] FROM IM_TAX WITH(NOLOCK)";

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
            _initilize();
            if (clsGlobal.MODE_TRX == 2)
            {
                FillData();
                CBGroupCode.Enabled = false;
                textGrade.ReadOnly = true;
            }
            FillGridSizeGrp();

            __cursor = new LineHighlight[]{new LineHighlight(tabPage1),new  LineHighlight(tabPage2)};
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

        void _initilize()
        {
            try
            {
                __brand = new TIRAObject<TIRAItem>(__BRAND, 0)
                {
                    Title = "Search Brand",
                    EndAction = _endBrand,
                };
                __brand.AddRangeTextBox(linebrandcodetb0);

                __group = new TIRAObject<TIRAItem>(string.Format(__GROUP_F, ""), 0)
                {
                    Title = "Search Group",
                    EndAction = _endGroup,
                    IsAlwaysExecute = true,
                };
                __group.AddRangeTextBox(groupcodetb0);
                __group.AddParameter(new TIRAParameter("AND pg_prd_line_code =", linebrandcodetb0, true));
                __group.AddParameter(new TIRAParameter("ORDER BY pg_prd_group_code", null, false));

                __subgroup = new TIRAObject<TIRAItem>(string.Format(__SUBGROUP_F, ""), 0)
                {
                    Title = "Search Sub Group",
                    EndAction = _endSubGroup,
                    IsAlwaysExecute = true,
                };
                __subgroup.AddRangeTextBox(subgroupcodetb0);
                __subgroup.AddParameter(new TIRAParameter("AND psg_prd_line =", linebrandcodetb0, true));
                __subgroup.AddParameter(new TIRAParameter("AND psg_prd_group_code =", groupcodetb0, true));
                __subgroup.AddParameter(new TIRAParameter("ORDER BY psg_prd_sgroup_code", null, false));

                __model = new TIRAObject<TIRAItem>(__MODEL, 0)
                {
                    Title = "Search Model",
                    EndAction = _endModel,
                };
                __model.AddRangeTextBox(modelcodetb0);

                __rmcode = new TIRAObject<TIRAItem>(__RMCODE, 0)
                {
                    Title = "Search RM Code",
                    EndAction = _endRMCode,
                };
                __rmcode.AddRangeTextBox(rmcodetb0);

                __color = new TIRAObject<TIRAItem>(__COLOR, 0)
                {
                    Title = "Search Color",
                    EndAction = _endColor,
                };
                __color.AddRangeTextBox(colorcodetb0);

                __processcode = new TIRAObject<TIRAItem>(__PROCESSCODE, 0)
                {
                    Title = "Search Process Code",
                    EndAction = _endProcessCode,
                };
                __processcode.AddRangeTextBox(proccodetb0);

                __vendor = new TIRAObject<VendorItem>(__VENDOR, 0)
                {
                    Title = "Search Vendor",
                    EndAction = _endVendor,
                };
                __vendor.AddRangeTextBox(vendortb0);

                __principal = new TIRAObject<TIRAItem>(__PRINCIPAL, 0)
                {
                    Title = "Search Principal",
                    EndAction = _endPrincipal,
                };
                __principal.AddRangeTextBox(principaltb0);

                __taxcode = new TIRAObject<TIRAItem>(__TAXCODE, 0)
                {
                    Title = "Search Tax Code",
                    EndAction = _endTaxCode,
                };
                __taxcode.AddRangeTextBox(taxcodetb0);


                __brand.Initialize();
                __group.Initialize();
                __subgroup.Initialize();
                __model.Initialize();
                __rmcode.Initialize();
                __color.Initialize();
                __processcode.Initialize();
                __vendor.Initialize();
                __principal.Initialize();
                __taxcode.Initialize();
            }
            catch (Exception) { }
        }

        #region -- native search lookup
        private void linebrandcodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            linebrandcodetb0n.EditValue = null;
            linebrandcodetb0n.Properties.DataSource = _clsGlobal.ExecDT(__BRAND);
        }

        private void groupcodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            groupcodetb0n.EditValue = null;
            string filter = "";
            if (!string.IsNullOrWhiteSpace(linebrandcodetb0.Text))
                filter += " AND pg_prd_line_code = " + FmtStr(linebrandcodetb0.Text.Trim());

            groupcodetb0n.Properties.DataSource = _clsGlobal.ExecDT(string.Format(__GROUP_F, filter + " ORDER BY pg_prd_group_code"));
        }

        private void subgroupcodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            subgroupcodetb0n.EditValue = null;
            string filter = "";
            if (!string.IsNullOrWhiteSpace(linebrandcodetb0.Text))
                filter += " AND psg_prd_line = " + FmtStr(linebrandcodetb0.Text.Trim());
            if (!string.IsNullOrWhiteSpace(groupcodetb0.Text))
                filter += " AND psg_prd_group_code = " + FmtStr(groupcodetb0.Text.Trim());

            subgroupcodetb0n.Properties.DataSource = _clsGlobal.ExecDT(string.Format(__SUBGROUP_F, filter + " ORDER BY psg_prd_sgroup_code"));
        }

        private void modelcodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            modelcodetb0n.EditValue = null;
            modelcodetb0n.Properties.DataSource = _clsGlobal.ExecDT(__MODEL);
        }

        private void rmcodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            rmcodetb0n.EditValue = null;
            rmcodetb0n.Properties.DataSource = _clsGlobal.ExecDT(__RMCODE);
        }

        private void colorcodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            colorcodetb0n.EditValue = null;
            colorcodetb0n.Properties.DataSource = _clsGlobal.ExecDT(__COLOR);
        }

        private void proccodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            proccodetb0n.EditValue = null;
            proccodetb0n.Properties.DataSource = _clsGlobal.ExecDT(__PROCESSCODE);
        }

        private void vendortb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            vendortb0n.EditValue = null;
            vendortb0n.Properties.DataSource = _clsGlobal.ExecDT(__VENDOR_LOOKUP);
        }

        private void principaltb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            principaltb0n.EditValue = null;
            principaltb0n.Properties.DataSource = _clsGlobal.ExecDT(__PRINCIPAL);
        }

        private void taxcodetb0n_QueryPopUp(object sender, CancelEventArgs e)
        {
            taxcodetb0n.EditValue = null;
            taxcodetb0n.Properties.DataSource = _clsGlobal.ExecDT(__TAXCODE);
        }

        private void linebrandcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            DataRow row = GetSelectedLookupRow(linebrandcodetb0n);
            if (row == null) return;

            string code = LookupRowText(row, "Kode");
            if (!linebrandcodetb0.Text.CompareC(code))
            {
                groupcodetb0.Text = string.Empty;
                groupcodetb1.Text = string.Empty;
                subgroupcodetb0.Text = string.Empty;
                subgroupcodetb1.Text = string.Empty;
            }

            linebrandcodetb0.Text = code;
            linebrandcodetb1.Text = LookupRowText(row, "Nama");
        }

        private void groupcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            DataRow row = GetSelectedLookupRow(groupcodetb0n);
            if (row == null) return;

            string code = LookupRowText(row, "Kode");
            if (!groupcodetb0.Text.CompareC(code))
            {
                subgroupcodetb0.Text = string.Empty;
                subgroupcodetb1.Text = string.Empty;
            }

            groupcodetb0.Text = code;
            groupcodetb1.Text = LookupRowText(row, "Nama");
        }

        private void subgroupcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            ApplySelectedLookup(subgroupcodetb0n, subgroupcodetb0, subgroupcodetb1);
        }

        private void modelcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            ApplySelectedLookup(modelcodetb0n, modelcodetb0, modelcodetb1);
        }

        private void rmcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            ApplySelectedLookup(rmcodetb0n, rmcodetb0, rmcodetb1);
        }

        private void colorcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            ApplySelectedLookup(colorcodetb0n, colorcodetb0, colorcodetb1);
        }

        private void proccodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            ApplySelectedLookup(proccodetb0n, proccodetb0, proccodetb1);
        }

        private void vendortb0n_EditValueChanged(object sender, EventArgs e)
        {
            DataRow row = GetSelectedLookupRow(vendortb0n);
            if (row == null) return;

            vendortb0.Text = LookupRowText(row, "Kode");
            vendortb1.Text = LookupRowText(row, "ID2");
        }

        private void principaltb0n_EditValueChanged(object sender, EventArgs e)
        {
            ApplySelectedLookup(principaltb0n, principaltb0, null);
        }

        private void taxcodetb0n_EditValueChanged(object sender, EventArgs e)
        {
            ApplySelectedLookup(taxcodetb0n, taxcodetb0, null);
        }

        private bool ApplySelectedLookup(DevExpress.XtraEditors.SearchLookUpEdit lookup, TextBox codeTextBox, TextBox descriptionTextBox)
        {
            DataRow row = GetSelectedLookupRow(lookup);
            if (row == null) return false;

            codeTextBox.Text = LookupRowText(row, "Kode");
            if (descriptionTextBox != null)
                descriptionTextBox.Text = LookupRowText(row, "Nama");

            return true;
        }

        private DataRow GetSelectedLookupRow(DevExpress.XtraEditors.SearchLookUpEdit lookup)
        {
            if (lookup == null || lookup.EditValue == null || lookup.EditValue == DBNull.Value)
                return null;

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
        #endregion

        #region -- end action
        void _endBrand(IAsyncResult __iasync)
        {
            try
            {
                TIRAResult __result = (TIRAResult)__iasync.AsyncState;
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                this.Invoke(new MethodInvoker(delegate()
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
                CBGroupCode.DataSource = dt;
                CBGroupCode.ValueMember = "szg_size_group_code";
                CBGroupCode.DisplayMember = "szg_size_group_code";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



            try
            {
                //Status
                strSQL = "select gh_function_code, gh_function_code +' - '+gh_function_desc as gh_function_desc from GS_GEN_HARDCODED where gh_sys = 'H' and gh_function_name = 'STATUS_PRD_MASTER' order by gh_sequence_no ";
                CBStatus.DataSource = _clsGlobal.ExecDT(strSQL);
                CBStatus.ValueMember = "gh_function_code";
                CBStatus.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //trend
                strSQL = "select  pp_function_code, pp_function_code+' - '+pp_function_desc pp_function_desc from IM_PRD_PARAMETER  where  pp_function_name = 'TREND' order by pp_sequence_no  ";
                CBPacksize.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPacksize.ValueMember = "pp_function_code";
                CBPacksize.DisplayMember = "pp_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //target
                strSQL = "select  pp_function_code, pp_function_code+' - '+pp_function_desc pp_function_desc from IM_PRD_PARAMETER  where  pp_function_name = 'TARGET' order by pp_sequence_no  ";
                CBPackType.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPackType.ValueMember = "pp_function_code";
                CBPackType.DisplayMember = "pp_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //season
                strSQL = "select  pp_function_code, pp_function_code+' - '+pp_function_desc pp_function_desc from IM_PRD_PARAMETER  where  pp_function_name = 'SEASON' order by pp_sequence_no  ";
                CBSeason.DataSource = _clsGlobal.ExecDT(strSQL);
                CBSeason.ValueMember = "pp_function_code";
                CBSeason.DisplayMember = "pp_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //period indicator
                strSQL = "select gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc  from GS_GEN_HARDCODED  where gh_sys = 'H' and gh_function_name = 'PERIOD_INDC'  order by gh_sequence_no";
                CBPeriodIndicator.DataSource = _clsGlobal.ExecDT(strSQL);
                CBPeriodIndicator.ValueMember = "gh_function_code";
                CBPeriodIndicator.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //ukuran kecil
                strSQL = "select umc_uom_code,umc_uom_code+' - '+umc_uom_description umc_uom_description from  IM_UNIT_MEASURE_CODES ORDER BY umc_index ";
                CbKecil.DataSource = _clsGlobal.ExecDT(strSQL);
                CbKecil.ValueMember = "umc_uom_code";
                CbKecil.DisplayMember = "umc_uom_description";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //ukuran tengah
                strSQL = "select umc_uom_code,umc_uom_code+' - '+umc_uom_description umc_uom_description from  IM_UNIT_MEASURE_CODES ORDER BY umc_index ";
                CBTengah.DataSource = _clsGlobal.ExecDT(strSQL);
                CBTengah.ValueMember = "umc_uom_code";
                CBTengah.DisplayMember = "umc_uom_description";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //ukuran besar
                strSQL = "select umc_uom_code,umc_uom_code+' - '+umc_uom_description umc_uom_description from  IM_UNIT_MEASURE_CODES ORDER BY umc_index ";
                CBBesar.DataSource = _clsGlobal.ExecDT(strSQL);
                CBBesar.ValueMember = "umc_uom_code";
                CBBesar.DisplayMember = "umc_uom_description";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //produk type
                strSQL = "select gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc  from GS_GEN_HARDCODED  where gh_sys = 'H' and gh_function_name = 'PRDTYPE'  order by gh_sequence_no";
                CBType.DataSource = _clsGlobal.ExecDT(strSQL);
                CBType.ValueMember = "gh_function_code";
                CBType.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //Class
                strSQL = "select gh_function_code, gh_function_code+' - '+gh_function_desc gh_function_desc  from GS_GEN_HARDCODED  where gh_sys = 'H' and gh_function_name = 'SEX_DESIGN'  order by gh_sequence_no";
                CBGrade.DataSource = _clsGlobal.ExecDT(strSQL);
                CBGrade.ValueMember = "gh_function_code";
                CBGrade.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                //Rumusan OB
                strSQL = "SELECT gh_function_code, gh_function_code + ' ~ ' + gh_function_desc as gh_function_desc  From GS_GEN_HARDCODED WHERE gh_sys = 'H'   AND gh_function_name = 'RUMUSAN_OB'";
                CBRumusOB.DataSource = _clsGlobal.ExecDT(strSQL);
                CBRumusOB.ValueMember = "gh_function_code";
                CBRumusOB.DisplayMember = "gh_function_desc";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

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
            if (CBStatus.SelectedIndex > -1)
            { }
        }
        private void ComboSzGroupCd_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBGroupCode.SelectedIndex > -1)
            { }
        }
        private void checkBoxComboPckSzTrnd_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBPacksize.SelectedIndex > -1)
            { }
        }
        private void checkBoxComboPckTypeTrgt_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBPackType.SelectedIndex > -1)
            { }
        }
        private void checkBoxComboSeason_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBSeason.SelectedIndex > 0)
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


                dGV1.DataSource = _clsGlobal.ExecDT(strSQL);

                if (dGV1.Rows.Count > 0)
                {
                    this.dGV1.Rows[0].Cells["szg_index"].Value = true;

                    //DataGridViewCellEventArgs DataGridViewCellEventArgs = new DataGridViewCellEventArgs(dGV1.Columns["szg_index"].Index, dGV1.CurrentRow.Index);
                    //this.dGV1_CellClick(dGV1.CurrentRow.Index, DataGridViewCellEventArgs);

                }
                else
                {
                    strSQL = "SELECT DISTINCT szg_index, szg_prd_size, '' as Price, '' as Cost FROM IM_PRD_SIZE_GRP WHERE szg_size_group_code= '" + CBGroupCode.SelectedValue + "' ORDER BY szg_index";
                    dGV1.DataSource = _clsGlobal.ExecDT(strSQL);
                    if (dGV1.RowCount > 0)
                    {
                        if (clsGlobal.MODE_TRX == 1)
                        {
                            this.dGV1.Rows[0].Cells["szg_index"].Value = false;
                        }
                        else
                        {
                            this.dGV1.Rows[0].Cells["szg_index"].Value = true;
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                    CBType.SelectedValue = dt.Rows[0]["prm_prd_type"].ToString().Trim();
                    CBGroupCode.SelectedValue = dt.Rows[0]["prm_size_group_code"].ToString().Trim();
                    Sizegroupcode = dt.Rows[0]["prm_size_group_code"].ToString().Trim();
                    CBGrade.SelectedValue = dt.Rows[0]["prm_sex"].ToString().Trim();
                    CbKecil.SelectedValue = dt.Rows[0]["prm_um"].ToString().Trim();
                    CBTengah.SelectedValue = dt.Rows[0]["prm_um_sales"].ToString().Trim();
                    txttengah.Text = dt.Rows[0]["prm_conversion_sales"].ToString().Trim();
                    CBBesar.SelectedValue = dt.Rows[0]["prm_um_purc"].ToString().Trim();
                    txtbesar.Text = dt.Rows[0]["prm_conversion_purc"].ToString().Trim();
                    CBStatus.SelectedValue = dt.Rows[0]["prm_status"].ToString().Trim();
                    CBPackType.SelectedValue = dt.Rows[0]["prm_target"].ToString().Trim();
                    CBPacksize.SelectedValue = dt.Rows[0]["prm_trend"].ToString().Trim();
                    CBSeason.SelectedValue = dt.Rows[0]["prm_season"].ToString().Trim();

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
                    CBRumusOB.SelectedValue = dt.Rows[0]["prm_rumus_ob"].ToString().Trim();

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
                    CBPeriodIndicator.SelectedValue = dt.Rows[0]["prm_period_ind"].ToString().Trim();
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
                if(!taxcodetb0.Text.ToString().ToUpper().Equals("PPN0"))
                {
                    cbflagextract = "N";
                }

                strSQL = "";
                strSQL = "EXEC [SP_INV_PRODUK_MASTER] '1', ";
                strSQL += "" + FmtStr(txtProdMstrCd.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(textGrade.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(textGrade.Text.Trim()) + ", ";
                //strSQL+=   // "" + ((CBPacksize.SelectedValue != "") ? "" + CBPacksize.SelectedValue + "" : "NULL") + "," +
                //strSQL+=   //"" + ((CBPackType.SelectedValue != "") ? "" + CBPackType.SelectedValue + "" : "NULL") + "," +
                //strSQL+=   //"" + ((CBSeason.SelectedValue != "") ? "" + CBSeason.SelectedValue + "" : "NULL") + "," +

                if (CBPacksize.SelectedValue != null)
                    strSQL += "" + FmtStr(CBPacksize.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'',";

                if (CBPackType.SelectedValue != null)
                    strSQL += "" + FmtStr(CBPackType.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                if (CBSeason.SelectedValue != null)
                    strSQL += "" + FmtStr(CBSeason.SelectedValue.ToString()) + ", ";
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

                if (CbKecil.SelectedValue != null)
                    strSQL += "" + FmtStr(CbKecil.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                if (CBTengah.SelectedValue != null)
                    strSQL += "" + FmtStr(CBTengah.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                if (CBBesar.SelectedValue != null)
                    strSQL += "" + FmtStr(CBBesar.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                strSQL += "" + ((txttengah.Text != "") ? "'" + txttengah.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtbesar.Text != "") ? "'" + txtbesar.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtWeight.Text != "") ? "'" + txtWeight.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtvlme1.Text != "") ? "'" + txtvlme1.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtvlme2.Text != "") ? "'" + txtvlme2.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(cb3.Trim()) + ", ";
                strSQL += "" + FmtStr(cb4.Trim()) + ", ";
                strSQL += "" + FmtStr(cb5.Trim()) + ",'Y', ";
                //strSQL+=   //"" + FmtStr(cb6.Trim()) + ", " +
                strSQL += "" + FmtStr(CBType.SelectedValue.ToString()) + ",'', ";
                //strSQL += "" + clsLogin.USERID + ", ";
                strSQL += string.Format("'{0}', ", clsLogin.USERID);
                strSQL += "" + FmtStr(CBGrade.SelectedValue.ToString()) + ", ";
                strSQL += "" + ((txtDcmlPrcsn.Text != "") ? "'" + txtDcmlPrcsn.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(priode.Trim()) + ", ";
                strSQL += "" + FmtStr(CBGroupCode.SelectedValue.ToString()) + ", ";
                strSQL += "" + FmtStr(proccodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(vendortb0.Text.Trim()) + ",";
                strSQL += "" + FmtStr(vendortb1.Text.Trim()) + ",";
                strSQL += "" + FmtStr(principaltb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(CBStatus.SelectedValue.ToString()) + ", ";
                //strSQL+=    //"" + ((CBStatus.SelectedValue != "") ? "" + CBStatus.SelectedValue + "" : "NULL") + "," +
                strSQL += "" + FmtStr(txtRprtName1.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName2.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName3.Text.Trim()) + ", ";
                strSQL += "" + ((txtDcmlPrcsn.Text != "") ? "'" + txtMnmnItm.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtPknAwl.Text != "") ? "'" + txtPknAwl.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(txtbrcd.Text.Trim()) + ", ";
                strSQL += "" + ((txtStockMax.Text != "") ? "'" + txtStockMax.Text + "'" : "0") + ", ";
                //strSQL+=     //"" + ((CBRumusOB.SelectedValue != "") ? "" + CBRumusOB.SelectedValue + "" : "NULL") + "," +
                strSQL += "" + FmtStr(CBRumusOB.SelectedValue.ToString()) + ", ";
                strSQL += "" + FmtStr(cb1.Trim()) + ", ";
                strSQL += "" + FmtStr(cb2.Trim()) + ", ";
                strSQL += "" + ((txtSftyBox.Text != "") ? "'" + txtSftyBox.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtStockMin.Text != "") ? "'" + txtStockMin.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtMimimum.Text != "") ? "'" + txtMimimum.Text + "'" : "0") + ", ";
                //strSQL+=    //"" + ((CBPeriodIndicator.SelectedValue != "") ? "" + CBPeriodIndicator.SelectedValue + "" : "NULL") + "," +
                strSQL += "" + FmtStr(CBPeriodIndicator.SelectedValue.ToString()) + ",";
                strSQL += "" + FmtStr(cb6.Trim()) + ",";
                strSQL += "'N' ,";
                strSQL += "" + FmtStr(txtPhoto.Text.Trim()) + ",";
                strSQL += "" + FmtStr(cbflagextract) + "";
                //prm_prd_cust_flag
                //prm_prd_promo_flag 58
                _clsGlobal.BeginTrans();
                _clsGlobal.ExecuteTrans(strSQL);
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
                strSQL += "" + ((CBPacksize.SelectedValue != null) ? "'" + CBPacksize.SelectedValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + ((CBPackType.SelectedValue != null) ? "'" + CBPackType.SelectedValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + ((CBSeason.SelectedValue != null) ? "'" + CBSeason.SelectedValue.ToString() + "'" : "NULL") + ",";
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

                if (CbKecil.SelectedValue != null)
                    strSQL += "" + FmtStr(CbKecil.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                if (CBTengah.SelectedValue != null)
                    strSQL += "" + FmtStr(CBTengah.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                if (CBBesar.SelectedValue != null)
                    strSQL += "" + FmtStr(CBBesar.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                strSQL += "" + ((txttengah.Text != "") ? "'" + txttengah.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtbesar.Text != "") ? "'" + txtbesar.Text + "'" : "0") + ", ";
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
                strSQL += "" + ((CBType.SelectedValue != null) ? "'" + CBType.SelectedValue.ToString() + "'" : "NULL") + ",";
                strSQL += " NULL ,";
                // strSQL += "" + clsLogin.USERID + ", ";
                strSQL += string.Format("'{0}', ", clsLogin.USERID);

                if (CBGrade.SelectedValue != null)
                    strSQL += "" + FmtStr(CBGrade.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                strSQL += "" + ((txtDcmlPrcsn.Text != "") ? "'" + txtDcmlPrcsn.Text + "'" : "0") + ", ";
                strSQL += "" + FmtStr(priode.Trim()) + ", ";

                if (CBGroupCode.SelectedValue != null)
                    strSQL += "" + FmtStr(CBGroupCode.SelectedValue.ToString()) + ", ";
                else
                    strSQL += "'', ";

                strSQL += "" + FmtStr(proccodetb0.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(vendortb0.Text.Trim()) + ",";
                strSQL += "" + FmtStr(vendortb1.Text.Trim()) + ",";
                strSQL += "" + FmtStr(principaltb0.Text.Trim()) + ", ";
                strSQL += "" + ((CBStatus.SelectedValue != null) ? "'" + CBStatus.SelectedValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + FmtStr(txtRprtName1.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName2.Text.Trim()) + ", ";
                strSQL += "" + FmtStr(txtRprtName3.Text.Trim()) + ", ";
                strSQL += "" + ((txtMnmnItm.Text != "") ? "'" + txtMnmnItm.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtPknAwl.Text != "") ? "'" + txtPknAwl.Text + "'" : "0") + ",";
                strSQL += "" + ((txtbrcd.Text != "") ? "'" + txtbrcd.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtStockMax.Text != "") ? "'" + txtStockMax.Text + "'" : "0") + ", ";
                strSQL += "" + ((CBRumusOB.SelectedValue != null) ? "'" + CBRumusOB.SelectedValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + FmtStr(cb1.Trim()) + ", ";
                strSQL += "" + FmtStr(cb2.Trim()) + ", ";
                strSQL += "" + ((txtSftyBox.Text != "") ? "'" + txtSftyBox.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtStockMin.Text != "") ? "'" + txtStockMin.Text + "'" : "0") + ", ";
                strSQL += "" + ((txtMimimum.Text != "") ? "'" + txtMimimum.Text + "'" : "0") + ", ";
                strSQL += "" + ((CBPeriodIndicator.SelectedValue != null) ? "'" + CBPeriodIndicator.SelectedValue.ToString() + "'" : "NULL") + ",";
                strSQL += "" + FmtStr(cb6.Trim()) + ",";
                strSQL += "'N' ,";
                strSQL += "" + FmtStr(txtPhoto.Text.Trim()) + ",";
                strSQL += "" + FmtStr(cbflagextract) + " ";
                begTrans = true;
                _clsGlobal.BeginTrans();
                _clsGlobal.ExecuteTrans(strSQL);
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

        #region events save change
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btnOk_Click(object sender, EventArgs e)
        {
            if (linebrandcodetb0.Text.Trim() == "")
            {
                MessageBox.Show(linebrandcodetb0.Text + " Line Brand Code Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                linebrandcodetb0.Select();
                return;
            }

            if (groupcodetb0.Text.Trim() == "")
            {
                MessageBox.Show(groupcodetb0.Text + " Group Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                groupcodetb0.Select();
                return;
            }
            if (subgroupcodetb0.Text.Trim() == "")
            {
                MessageBox.Show(subgroupcodetb0.Text + " SubGroup Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                subgroupcodetb0.Select();
                return;
            }
            if (modelcodetb0.Text.Trim() == "")
            {
                MessageBox.Show(modelcodetb0.Text + " Model Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                modelcodetb0.Select();
                return;
            }

            if (rmcodetb0.Text.Trim() == "")
            {
                MessageBox.Show(rmcodetb0.Text + " RM/Source Code Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                rmcodetb0.Select();
                return;
            }
            if (colorcodetb0.Text.Trim() == "")
            {
                MessageBox.Show(colorcodetb0.Text + " Color Code Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                colorcodetb0.Select();
                return;
            }
            if (vendortb0.Text.Trim() == "")
            {
                MessageBox.Show(vendortb0.Text + " Vendor / Supplier Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                vendortb0.Select();
                return;
            }
            if (principaltb0.Text.Trim() == "")
            {
                MessageBox.Show(principaltb0.Text + " Principal Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                principaltb0.Select();
                return;
            }
            if (txtProdMstrCd.Text.Trim() == "")
            {
                MessageBox.Show(txtProdMstrCd.Text + " Prod Master Code Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtProdMstrCd.Select();
                return;
            }
            if (txtdescrptn.Text.Trim() == "")
            {
                MessageBox.Show(txtdescrptn.Text + " Description Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtdescrptn.Select();
                return;
            }
            if (txtmnfctrCd.Text.Trim() == "")
            {
                MessageBox.Show(txtmnfctrCd.Text + " Manufacture Code Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtmnfctrCd.Select();
                return;
            }
            if (txtMnmnItm.Text.Trim() == "")
            {
                MessageBox.Show(txtMnmnItm.Text + " Minimum Item Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtMnmnItm.Select();
                return;
            }
            if (txtPknAwl.Text.Trim() == "")
            {
                MessageBox.Show(txtPknAwl.Text + " Pekan Awal Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtPknAwl.Select();
                return;
            }
            if (txtRprtName1.Text.Trim() == "")
            {
                MessageBox.Show(txtRprtName1.Text + " Report Name 1 Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtRprtName1.Select();
                return;
            }
            if (txtRprtName2.Text.Trim() == "")
            {
                MessageBox.Show(txtRprtName2.Text + " Report Name 2 Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtRprtName2.Select();
                return;
            }
            if (txtRprtName3.Text.Trim() == "")
            {
                MessageBox.Show(txtRprtName3.Text + " Report Name 3 Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtRprtName3.Select();
                return;
            }
            if (txttengah.Text.Trim() == "")
            {
                MessageBox.Show(txttengah.Text + " Tengah Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txttengah.Select();
                return;
            }
            if (txtbesar.Text.Trim() == "")
            {
                MessageBox.Show(txtbesar.Text + " Besar Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                txtbesar.Select();
                return;
            }
            if (CBTengah.Text == "")
            {
                MessageBox.Show(CBTengah.Text + " Combo Tengah Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                CBTengah.Select();
                return;
            }
            if (CBBesar.Text == "")
            {
                MessageBox.Show(CBBesar.Text + " Combo Besar Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                CBBesar.Select();
                return;
            }
            if (dGV1.RowCount > 0)
            {
                DataGridViewRow dgr;
                try
                {
                    dgr = dGV1.CurrentRow;

                    if (dgr != null)
                    {
                        bool rowDuplicate11 = false;
                        bool rowDuplicate2 = false;
                        foreach (DataGridViewRow gRow in dGV1.Rows)
                        {
                            if ((gRow.Cells["Cost"].Value.ToString().Trim() == ""))
                            {
                                rowDuplicate11 = true;
                                break;
                            }
                            if (gRow.Cells["Price"].Value.ToString() == "")
                            {
                                rowDuplicate2 = true;
                                break;
                            }
                        }
                        if (rowDuplicate11 == true)
                        {
                            MessageBox.Show("Cost Belum diisi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            dGV1.Select();
                            return;
                        }
                        if (rowDuplicate2 == true)
                        {
                            MessageBox.Show("Price Belum diisi", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            dGV1.Select();
                            return;
                        }
                    }
                }
                finally
                {
                    dgr = null;
                }
            }
            if (CKBatchTidak.Checked)
            {
                if (txtMimimum.Text.Trim() == "0" || txtMimimum.Text.Trim() == "")
                {
                    MessageBox.Show(txtMimimum.Text + " Min.Remaining Shelf Life Must not 0  ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    tabProduct.SelectedTab = tabPage1;
                    txtMimimum.Select();
                    return;
                }
            }
            if (txtbrcd.Text.Trim() == "")
            {
                MessageBox.Show(txtbrcd.Text + " Barcode Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage2;
                txtbrcd.Focus();
                return;
            }
            if (taxcodetb0.Text.Trim() == "")
            {
                MessageBox.Show(taxcodetb0.Text + " Tax Code Does not Allow Empty ! ", clsGlobal.APP_MSG_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabProduct.SelectedTab = tabPage1;
                taxcodetb0.Focus();
                return;
            }

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


                if (dGV1.RowCount > 0)
                {
                    this.dGV1.Rows[0].Cells["szg_index"].Value = true;
                }
            }
        }
        private void CBGroupCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGridSizeGrp();
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
        private void txtkecil_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txttengah_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtbesar_KeyPress(object sender, KeyPressEventArgs e)
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
                TIRAFTP __im10= new TIRAFTP();
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
}

