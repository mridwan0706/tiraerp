namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    partial class frmSOManualEntryList
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSOManualEntryList));
            this.panelMain = new DevExpress.XtraEditors.PanelControl();
            this.lcFilter = new DevExpress.XtraLayout.LayoutControl();
            this.txtEntityCode = new DevExpress.XtraEditors.TextEdit();
            this.btnPopUpEntityCode = new DevExpress.XtraEditors.SimpleButton();
            this.txtEntityDesc = new DevExpress.XtraEditors.TextEdit();
            this.txtBranchCode = new DevExpress.XtraEditors.TextEdit();
            this.btnPopUpBranchCode = new DevExpress.XtraEditors.SimpleButton();
            this.txtBranchDesc = new DevExpress.XtraEditors.TextEdit();
            this.txtSalesCode = new DevExpress.XtraEditors.TextEdit();
            this.btnPopUpSalesCode = new DevExpress.XtraEditors.SimpleButton();
            this.txtSalesDesc = new DevExpress.XtraEditors.TextEdit();
            this.txtOutletCode = new DevExpress.XtraEditors.TextEdit();
            this.btnPopUpOutlet = new DevExpress.XtraEditors.SimpleButton();
            this.txtOutletDesc = new DevExpress.XtraEditors.TextEdit();
            this.txtOutlet2 = new DevExpress.XtraEditors.TextEdit();
            this.txtNoOrder = new DevExpress.XtraEditors.TextEdit();
            this.btnPopUpNoOrder = new DevExpress.XtraEditors.SimpleButton();
            this.txtNoPO = new DevExpress.XtraEditors.TextEdit();
            this.btnPopUpNoPO = new DevExpress.XtraEditors.SimpleButton();
            this.cbSOType = new DevExpress.XtraEditors.LookUpEdit();
            this.cbSOStatus = new DevExpress.XtraEditors.LookUpEdit();
            this.cbdateselection = new DevExpress.XtraEditors.LookUpEdit();
            this.dtTglOrder = new DevExpress.XtraEditors.DateEdit();
            this.dtTglOrderTo = new DevExpress.XtraEditors.DateEdit();
            this.cbSource = new DevExpress.XtraEditors.LookUpEdit();
            this.cbFlagSloc = new DevExpress.XtraEditors.LookUpEdit();
            this.txtShippingPlant = new DevExpress.XtraEditors.TextEdit();
            this.btnShipping = new DevExpress.XtraEditors.SimpleButton();
            this.btnExec = new DevExpress.XtraEditors.SimpleButton();
            this.lblD = new DevExpress.XtraEditors.LabelControl();
            this.lblK = new DevExpress.XtraEditors.LabelControl();
            this.lblR = new DevExpress.XtraEditors.LabelControl();
            this.lblA = new DevExpress.XtraEditors.LabelControl();
            this.lcgFilter = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lcgColA = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lciEntity = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBtnEntity = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciEntityDesc = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBranch = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBtnBranch = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBranchDesc = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciSalesman = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBtnSales = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciSalesDesc = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciOutlet = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBtnOutlet = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciOutletDesc = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciOutlet2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciNoOrder = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBtnNoOrder = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciNoPO = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBtnNoPO = new DevExpress.XtraLayout.LayoutControlItem();
            this.lcgColB = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lciSOType = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciSOStatus = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciDateSel = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciTglOrder = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciTglOrderTo = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.lcgColC = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lciSource = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciFlagSloc = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciShipPlant = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciBtnShip = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptyExec = new DevExpress.XtraLayout.EmptySpaceItem();
            this.lcgColD = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lciLblD = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciLblK = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciLblR = new DevExpress.XtraLayout.LayoutControlItem();
            this.lciLblA = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.lciExec = new DevExpress.XtraLayout.LayoutControlItem();
            this.txtShippingPlantId = new DevExpress.XtraEditors.TextEdit();
            this.gcSalesHeader = new DevExpress.XtraGrid.GridControl();
            this.gvSalesHeader = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gcSalesDetail = new DevExpress.XtraGrid.GridControl();
            this.gvSalesDetail = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gcPromosiRp = new DevExpress.XtraGrid.GridControl();
            this.gvPromosiRp = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gcPromosiQty = new DevExpress.XtraGrid.GridControl();
            this.gvPromosiQty = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gcDisc = new DevExpress.XtraGrid.GridControl();
            this.gvDisc = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.groupTotal = new DevExpress.XtraEditors.GroupControl();
            this.lblSubTotal = new DevExpress.XtraEditors.LabelControl();
            this.txtSubTotal = new DevExpress.XtraEditors.TextEdit();
            this.lblCashDisc = new DevExpress.XtraEditors.LabelControl();
            this.txtCashDiscP = new DevExpress.XtraEditors.TextEdit();
            this.lblPercent = new DevExpress.XtraEditors.LabelControl();
            this.lblPromosiU = new DevExpress.XtraEditors.LabelControl();
            this.txtPromosiU = new DevExpress.XtraEditors.TextEdit();
            this.lblRpCashDisc = new DevExpress.XtraEditors.LabelControl();
            this.txtCashDiscR = new DevExpress.XtraEditors.TextEdit();
            this.lblDisc1 = new DevExpress.XtraEditors.LabelControl();
            this.txtDisc1 = new DevExpress.XtraEditors.TextEdit();
            this.lblPPN = new DevExpress.XtraEditors.LabelControl();
            this.txtPPN = new DevExpress.XtraEditors.TextEdit();
            this.lblTotalInvoice = new DevExpress.XtraEditors.LabelControl();
            this.txtTtlInvoice = new DevExpress.XtraEditors.TextEdit();
            this.lblBebanCashDisc = new DevExpress.XtraEditors.LabelControl();
            this.txtBebanCashDisc = new DevExpress.XtraEditors.TextEdit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMarqueeProgressBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemProgressBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.contentPanel)).BeginInit();
            this.contentPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelMain)).BeginInit();
            this.panelMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.lcFilter)).BeginInit();
            this.lcFilter.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtEntityCode.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtEntityDesc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBranchCode.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBranchDesc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesCode.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesDesc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtOutletCode.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtOutletDesc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtOutlet2.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtNoOrder.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtNoPO.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbSOType.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbSOStatus.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbdateselection.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrder.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrder.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrderTo.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrderTo.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbSource.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbFlagSloc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtShippingPlant.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgFilter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColA)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciEntity)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnEntity)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciEntityDesc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBranch)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnBranch)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBranchDesc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSalesman)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnSales)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSalesDesc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciOutlet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnOutlet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciOutletDesc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciOutlet2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciNoOrder)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnNoOrder)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciNoPO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnNoPO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColB)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSOType)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSOStatus)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciDateSel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciTglOrder)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciTglOrderTo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColC)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciFlagSloc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciShipPlant)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnShip)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptyExec)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColD)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblD)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblK)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblR)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblA)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciExec)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtShippingPlantId.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcSalesHeader)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvSalesHeader)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcSalesDetail)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvSalesDetail)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcPromosiRp)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPromosiRp)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcPromosiQty)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPromosiQty)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcDisc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvDisc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupTotal)).BeginInit();
            this.groupTotal.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtSubTotal.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtCashDiscP.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPromosiU.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtCashDiscR.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtDisc1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPPN.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTtlInvoice.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBebanCashDisc.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // contentPanel
            // 
            this.contentPanel.Controls.Add(this.panelMain);
            this.contentPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.contentPanel.Location = new System.Drawing.Point(0, 24);
            this.contentPanel.Size = new System.Drawing.Size(1120, 622);
            // 
            // panelMain
            // 
            this.panelMain.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.panelMain.Controls.Add(this.lcFilter);
            this.panelMain.Controls.Add(this.txtShippingPlantId);
            this.panelMain.Controls.Add(this.gcSalesHeader);
            this.panelMain.Controls.Add(this.gcSalesDetail);
            this.panelMain.Controls.Add(this.gcPromosiRp);
            this.panelMain.Controls.Add(this.gcPromosiQty);
            this.panelMain.Controls.Add(this.gcDisc);
            this.panelMain.Controls.Add(this.groupTotal);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.Location = new System.Drawing.Point(0, 0);
            this.panelMain.Name = "panelMain";
            this.panelMain.Padding = new System.Windows.Forms.Padding(6);
            this.panelMain.Size = new System.Drawing.Size(1120, 622);
            this.panelMain.TabIndex = 9;
            // 
            // lcFilter
            // 
            this.lcFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lcFilter.Controls.Add(this.txtEntityCode);
            this.lcFilter.Controls.Add(this.btnPopUpEntityCode);
            this.lcFilter.Controls.Add(this.txtEntityDesc);
            this.lcFilter.Controls.Add(this.txtBranchCode);
            this.lcFilter.Controls.Add(this.btnPopUpBranchCode);
            this.lcFilter.Controls.Add(this.txtBranchDesc);
            this.lcFilter.Controls.Add(this.txtSalesCode);
            this.lcFilter.Controls.Add(this.btnPopUpSalesCode);
            this.lcFilter.Controls.Add(this.txtSalesDesc);
            this.lcFilter.Controls.Add(this.txtOutletCode);
            this.lcFilter.Controls.Add(this.btnPopUpOutlet);
            this.lcFilter.Controls.Add(this.txtOutletDesc);
            this.lcFilter.Controls.Add(this.txtOutlet2);
            this.lcFilter.Controls.Add(this.txtNoOrder);
            this.lcFilter.Controls.Add(this.btnPopUpNoOrder);
            this.lcFilter.Controls.Add(this.txtNoPO);
            this.lcFilter.Controls.Add(this.btnPopUpNoPO);
            this.lcFilter.Controls.Add(this.cbSOType);
            this.lcFilter.Controls.Add(this.cbSOStatus);
            this.lcFilter.Controls.Add(this.cbdateselection);
            this.lcFilter.Controls.Add(this.dtTglOrder);
            this.lcFilter.Controls.Add(this.dtTglOrderTo);
            this.lcFilter.Controls.Add(this.cbSource);
            this.lcFilter.Controls.Add(this.cbFlagSloc);
            this.lcFilter.Controls.Add(this.txtShippingPlant);
            this.lcFilter.Controls.Add(this.btnShipping);
            this.lcFilter.Controls.Add(this.btnExec);
            this.lcFilter.Controls.Add(this.lblD);
            this.lcFilter.Controls.Add(this.lblK);
            this.lcFilter.Controls.Add(this.lblR);
            this.lcFilter.Controls.Add(this.lblA);
            this.lcFilter.Location = new System.Drawing.Point(6, 6);
            this.lcFilter.Name = "lcFilter";
            this.lcFilter.Root = this.lcgFilter;
            this.lcFilter.Size = new System.Drawing.Size(1108, 175);
            this.lcFilter.TabIndex = 0;
            this.lcFilter.Text = "lcFilter";
            // 
            // txtEntityCode
            // 
            this.txtEntityCode.Location = new System.Drawing.Point(105, 24);
            this.txtEntityCode.Name = "txtEntityCode";
            this.txtEntityCode.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtEntityCode.Properties.MaxLength = 5;
            this.txtEntityCode.Size = new System.Drawing.Size(74, 20);
            this.txtEntityCode.StyleController = this.lcFilter;
            this.txtEntityCode.TabIndex = 1;
            this.txtEntityCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEntityCode_KeyDown);
            this.txtEntityCode.Leave += new System.EventHandler(this.txtEntityCode_Leave);
            // 
            // btnPopUpEntityCode
            // 
            this.btnPopUpEntityCode.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.Binoculars2;
            this.btnPopUpEntityCode.Location = new System.Drawing.Point(183, 24);
            this.btnPopUpEntityCode.Name = "btnPopUpEntityCode";
            this.btnPopUpEntityCode.Size = new System.Drawing.Size(40, 22);
            this.btnPopUpEntityCode.StyleController = this.lcFilter;
            this.btnPopUpEntityCode.TabIndex = 2;
            this.btnPopUpEntityCode.Click += new System.EventHandler(this.btnPopUpEntityCode_Click);
            // 
            // txtEntityDesc
            // 
            this.txtEntityDesc.Location = new System.Drawing.Point(227, 24);
            this.txtEntityDesc.Name = "txtEntityDesc";
            this.txtEntityDesc.Properties.Appearance.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.txtEntityDesc.Properties.Appearance.ForeColor = System.Drawing.SystemColors.Window;
            this.txtEntityDesc.Properties.Appearance.Options.UseBackColor = true;
            this.txtEntityDesc.Properties.Appearance.Options.UseForeColor = true;
            this.txtEntityDesc.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtEntityDesc.Properties.ReadOnly = true;
            this.txtEntityDesc.Size = new System.Drawing.Size(227, 20);
            this.txtEntityDesc.StyleController = this.lcFilter;
            this.txtEntityDesc.TabIndex = 3;
            // 
            // txtBranchCode
            // 
            this.txtBranchCode.Location = new System.Drawing.Point(105, 50);
            this.txtBranchCode.Name = "txtBranchCode";
            this.txtBranchCode.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtBranchCode.Properties.MaxLength = 5;
            this.txtBranchCode.Size = new System.Drawing.Size(74, 20);
            this.txtBranchCode.StyleController = this.lcFilter;
            this.txtBranchCode.TabIndex = 5;
            this.txtBranchCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBranchCode_KeyDown);
            this.txtBranchCode.Leave += new System.EventHandler(this.txtBranchCode_Leave);
            // 
            // btnPopUpBranchCode
            // 
            this.btnPopUpBranchCode.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.Binoculars2;
            this.btnPopUpBranchCode.Location = new System.Drawing.Point(183, 50);
            this.btnPopUpBranchCode.Name = "btnPopUpBranchCode";
            this.btnPopUpBranchCode.Size = new System.Drawing.Size(40, 22);
            this.btnPopUpBranchCode.StyleController = this.lcFilter;
            this.btnPopUpBranchCode.TabIndex = 6;
            this.btnPopUpBranchCode.Click += new System.EventHandler(this.btnPopUpBranchCode_Click);
            // 
            // txtBranchDesc
            // 
            this.txtBranchDesc.Location = new System.Drawing.Point(227, 50);
            this.txtBranchDesc.Name = "txtBranchDesc";
            this.txtBranchDesc.Properties.Appearance.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.txtBranchDesc.Properties.Appearance.ForeColor = System.Drawing.SystemColors.Window;
            this.txtBranchDesc.Properties.Appearance.Options.UseBackColor = true;
            this.txtBranchDesc.Properties.Appearance.Options.UseForeColor = true;
            this.txtBranchDesc.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtBranchDesc.Properties.ReadOnly = true;
            this.txtBranchDesc.Size = new System.Drawing.Size(227, 20);
            this.txtBranchDesc.StyleController = this.lcFilter;
            this.txtBranchDesc.TabIndex = 7;
            // 
            // txtSalesCode
            // 
            this.txtSalesCode.Location = new System.Drawing.Point(105, 76);
            this.txtSalesCode.Name = "txtSalesCode";
            this.txtSalesCode.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtSalesCode.Properties.MaxLength = 10;
            this.txtSalesCode.Size = new System.Drawing.Size(74, 20);
            this.txtSalesCode.StyleController = this.lcFilter;
            this.txtSalesCode.TabIndex = 9;
            this.txtSalesCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSalesCode_KeyDown);
            this.txtSalesCode.Leave += new System.EventHandler(this.txtSalesCode_Leave);
            // 
            // btnPopUpSalesCode
            // 
            this.btnPopUpSalesCode.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.Binoculars2;
            this.btnPopUpSalesCode.Location = new System.Drawing.Point(183, 76);
            this.btnPopUpSalesCode.Name = "btnPopUpSalesCode";
            this.btnPopUpSalesCode.Size = new System.Drawing.Size(40, 22);
            this.btnPopUpSalesCode.StyleController = this.lcFilter;
            this.btnPopUpSalesCode.TabIndex = 10;
            this.btnPopUpSalesCode.Click += new System.EventHandler(this.btnPopUpSalesCode_Click);
            // 
            // txtSalesDesc
            // 
            this.txtSalesDesc.Location = new System.Drawing.Point(227, 76);
            this.txtSalesDesc.Name = "txtSalesDesc";
            this.txtSalesDesc.Properties.Appearance.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.txtSalesDesc.Properties.Appearance.ForeColor = System.Drawing.SystemColors.Window;
            this.txtSalesDesc.Properties.Appearance.Options.UseBackColor = true;
            this.txtSalesDesc.Properties.Appearance.Options.UseForeColor = true;
            this.txtSalesDesc.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtSalesDesc.Properties.ReadOnly = true;
            this.txtSalesDesc.Size = new System.Drawing.Size(227, 20);
            this.txtSalesDesc.StyleController = this.lcFilter;
            this.txtSalesDesc.TabIndex = 11;
            // 
            // txtOutletCode
            // 
            this.txtOutletCode.Location = new System.Drawing.Point(105, 102);
            this.txtOutletCode.Name = "txtOutletCode";
            this.txtOutletCode.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtOutletCode.Properties.MaxLength = 6;
            this.txtOutletCode.Size = new System.Drawing.Size(74, 20);
            this.txtOutletCode.StyleController = this.lcFilter;
            this.txtOutletCode.TabIndex = 13;
            this.txtOutletCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOutletCode_KeyDown);
            this.txtOutletCode.Leave += new System.EventHandler(this.txtOutletCode_Leave);
            // 
            // btnPopUpOutlet
            // 
            this.btnPopUpOutlet.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.Binoculars2;
            this.btnPopUpOutlet.Location = new System.Drawing.Point(183, 102);
            this.btnPopUpOutlet.Name = "btnPopUpOutlet";
            this.btnPopUpOutlet.Size = new System.Drawing.Size(40, 22);
            this.btnPopUpOutlet.StyleController = this.lcFilter;
            this.btnPopUpOutlet.TabIndex = 14;
            this.btnPopUpOutlet.Click += new System.EventHandler(this.btnPopUpOutlet_Click);
            // 
            // txtOutletDesc
            // 
            this.txtOutletDesc.Location = new System.Drawing.Point(227, 102);
            this.txtOutletDesc.Name = "txtOutletDesc";
            this.txtOutletDesc.Properties.Appearance.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.txtOutletDesc.Properties.Appearance.ForeColor = System.Drawing.SystemColors.Window;
            this.txtOutletDesc.Properties.Appearance.Options.UseBackColor = true;
            this.txtOutletDesc.Properties.Appearance.Options.UseForeColor = true;
            this.txtOutletDesc.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtOutletDesc.Properties.ReadOnly = true;
            this.txtOutletDesc.Size = new System.Drawing.Size(172, 20);
            this.txtOutletDesc.StyleController = this.lcFilter;
            this.txtOutletDesc.TabIndex = 15;
            // 
            // txtOutlet2
            // 
            this.txtOutlet2.Location = new System.Drawing.Point(403, 102);
            this.txtOutlet2.Name = "txtOutlet2";
            this.txtOutlet2.Properties.Appearance.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.txtOutlet2.Properties.Appearance.Options.UseBackColor = true;
            this.txtOutlet2.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtOutlet2.Size = new System.Drawing.Size(51, 20);
            this.txtOutlet2.StyleController = this.lcFilter;
            this.txtOutlet2.TabIndex = 16;
            this.txtOutlet2.Visible = false;
            // 
            // txtNoOrder
            // 
            this.txtNoOrder.Location = new System.Drawing.Point(105, 128);
            this.txtNoOrder.Name = "txtNoOrder";
            this.txtNoOrder.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtNoOrder.Properties.MaxLength = 10;
            this.txtNoOrder.Size = new System.Drawing.Size(74, 20);
            this.txtNoOrder.StyleController = this.lcFilter;
            this.txtNoOrder.TabIndex = 18;
            this.txtNoOrder.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNoOrder_KeyDown);
            this.txtNoOrder.Leave += new System.EventHandler(this.txtNoOrder_Leave);
            // 
            // btnPopUpNoOrder
            // 
            this.btnPopUpNoOrder.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.Binoculars2;
            this.btnPopUpNoOrder.Location = new System.Drawing.Point(183, 128);
            this.btnPopUpNoOrder.Name = "btnPopUpNoOrder";
            this.btnPopUpNoOrder.Size = new System.Drawing.Size(40, 22);
            this.btnPopUpNoOrder.StyleController = this.lcFilter;
            this.btnPopUpNoOrder.TabIndex = 19;
            this.btnPopUpNoOrder.Click += new System.EventHandler(this.btnPopUpNoOrder_Click);
            // 
            // txtNoPO
            // 
            this.txtNoPO.Location = new System.Drawing.Point(308, 128);
            this.txtNoPO.Name = "txtNoPO";
            this.txtNoPO.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtNoPO.Properties.MaxLength = 10;
            this.txtNoPO.Size = new System.Drawing.Size(102, 20);
            this.txtNoPO.StyleController = this.lcFilter;
            this.txtNoPO.TabIndex = 21;
            this.txtNoPO.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNoPO_KeyDown);
            this.txtNoPO.Leave += new System.EventHandler(this.txtNoPO_Leave);
            // 
            // btnPopUpNoPO
            // 
            this.btnPopUpNoPO.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.Binoculars2;
            this.btnPopUpNoPO.Location = new System.Drawing.Point(414, 128);
            this.btnPopUpNoPO.Name = "btnPopUpNoPO";
            this.btnPopUpNoPO.Size = new System.Drawing.Size(40, 22);
            this.btnPopUpNoPO.StyleController = this.lcFilter;
            this.btnPopUpNoPO.TabIndex = 22;
            this.btnPopUpNoPO.Click += new System.EventHandler(this.btnPopUpNoPO_Click);
            // 
            // cbSOType
            // 
            this.cbSOType.Location = new System.Drawing.Point(563, 24);
            this.cbSOType.Name = "cbSOType";
            this.cbSOType.Properties.NullText = "";
            this.cbSOType.Properties.ShowFooter = false;
            this.cbSOType.Properties.ShowHeader = false;
            this.cbSOType.Size = new System.Drawing.Size(153, 20);
            this.cbSOType.StyleController = this.lcFilter;
            this.cbSOType.TabIndex = 24;
            // 
            // cbSOStatus
            // 
            this.cbSOStatus.Location = new System.Drawing.Point(563, 48);
            this.cbSOStatus.Name = "cbSOStatus";
            this.cbSOStatus.Properties.NullText = "";
            this.cbSOStatus.Properties.ShowFooter = false;
            this.cbSOStatus.Properties.ShowHeader = false;
            this.cbSOStatus.Size = new System.Drawing.Size(153, 20);
            this.cbSOStatus.StyleController = this.lcFilter;
            this.cbSOStatus.TabIndex = 26;
            // 
            // cbdateselection
            // 
            this.cbdateselection.Location = new System.Drawing.Point(563, 72);
            this.cbdateselection.Name = "cbdateselection";
            this.cbdateselection.Properties.NullText = "";
            this.cbdateselection.Properties.ShowFooter = false;
            this.cbdateselection.Properties.ShowHeader = false;
            this.cbdateselection.Size = new System.Drawing.Size(153, 20);
            this.cbdateselection.StyleController = this.lcFilter;
            this.cbdateselection.TabIndex = 28;
            // 
            // dtTglOrder
            // 
            this.dtTglOrder.EditValue = new System.DateTime(2026, 4, 24, 0, 0, 0, 0);
            this.dtTglOrder.Location = new System.Drawing.Point(482, 96);
            this.dtTglOrder.Name = "dtTglOrder";
            this.dtTglOrder.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dtTglOrder.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dtTglOrder.Properties.DisplayFormat.FormatString = "dd MMM yyyy";
            this.dtTglOrder.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dtTglOrder.Properties.EditFormat.FormatString = "dd MMM yyyy";
            this.dtTglOrder.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dtTglOrder.Properties.Mask.UseMaskAsDisplayFormat = true;
            this.dtTglOrder.Properties.MaskSettings.Set("mask", "dd MMM yyyy");
            this.dtTglOrder.Size = new System.Drawing.Size(103, 20);
            this.dtTglOrder.StyleController = this.lcFilter;
            this.dtTglOrder.TabIndex = 29;
            // 
            // dtTglOrderTo
            // 
            this.dtTglOrderTo.EditValue = new System.DateTime(2026, 4, 24, 0, 0, 0, 0);
            this.dtTglOrderTo.Location = new System.Drawing.Point(614, 96);
            this.dtTglOrderTo.Name = "dtTglOrderTo";
            this.dtTglOrderTo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dtTglOrderTo.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dtTglOrderTo.Properties.DisplayFormat.FormatString = "dd MMM yyyy";
            this.dtTglOrderTo.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dtTglOrderTo.Properties.EditFormat.FormatString = "dd MMM yyyy";
            this.dtTglOrderTo.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dtTglOrderTo.Properties.Mask.UseMaskAsDisplayFormat = true;
            this.dtTglOrderTo.Properties.MaskSettings.Set("mask", "dd MMM yyyy");
            this.dtTglOrderTo.Size = new System.Drawing.Size(102, 20);
            this.dtTglOrderTo.StyleController = this.lcFilter;
            this.dtTglOrderTo.TabIndex = 31;
            // 
            // cbSource
            // 
            this.cbSource.Location = new System.Drawing.Point(825, 24);
            this.cbSource.Name = "cbSource";
            this.cbSource.Properties.NullText = "";
            this.cbSource.Properties.ShowFooter = false;
            this.cbSource.Properties.ShowHeader = false;
            this.cbSource.Size = new System.Drawing.Size(102, 20);
            this.cbSource.StyleController = this.lcFilter;
            this.cbSource.TabIndex = 33;
            // 
            // cbFlagSloc
            // 
            this.cbFlagSloc.Location = new System.Drawing.Point(825, 48);
            this.cbFlagSloc.Name = "cbFlagSloc";
            this.cbFlagSloc.Properties.NullText = "";
            this.cbFlagSloc.Properties.ShowFooter = false;
            this.cbFlagSloc.Properties.ShowHeader = false;
            this.cbFlagSloc.Size = new System.Drawing.Size(102, 20);
            this.cbFlagSloc.StyleController = this.lcFilter;
            this.cbFlagSloc.TabIndex = 35;
            // 
            // txtShippingPlant
            // 
            this.txtShippingPlant.Location = new System.Drawing.Point(825, 72);
            this.txtShippingPlant.Name = "txtShippingPlant";
            this.txtShippingPlant.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtShippingPlant.Properties.ReadOnly = true;
            this.txtShippingPlant.Size = new System.Drawing.Size(75, 20);
            this.txtShippingPlant.StyleController = this.lcFilter;
            this.txtShippingPlant.TabIndex = 37;
            // 
            // btnShipping
            // 
            this.btnShipping.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.Binoculars2;
            this.btnShipping.Location = new System.Drawing.Point(904, 72);
            this.btnShipping.Name = "btnShipping";
            this.btnShipping.Size = new System.Drawing.Size(23, 22);
            this.btnShipping.StyleController = this.lcFilter;
            this.btnShipping.TabIndex = 39;
            this.btnShipping.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnExec
            // 
            this.btnExec.ImageOptions.Image = global::TIRASnDNet.Properties.Resources.database_lightning_ok;
            this.btnExec.Location = new System.Drawing.Point(955, 129);
            this.btnExec.Name = "btnExec";
            this.btnExec.Size = new System.Drawing.Size(129, 22);
            this.btnExec.StyleController = this.lcFilter;
            this.btnExec.TabIndex = 40;
            this.btnExec.Text = "E&xecute";
            this.btnExec.Click += new System.EventHandler(this.btnExec_Click);
            // 
            // lblD
            // 
            this.lblD.Location = new System.Drawing.Point(955, 24);
            this.lblD.Name = "lblD";
            this.lblD.Size = new System.Drawing.Size(63, 13);
            this.lblD.StyleController = this.lcFilter;
            this.lblD.TabIndex = 41;
            this.lblD.Text = "D = Overdue";
            // 
            // lblK
            // 
            this.lblK.Location = new System.Drawing.Point(955, 41);
            this.lblK.Name = "lblK";
            this.lblK.Size = new System.Drawing.Size(72, 13);
            this.lblK.StyleController = this.lcFilter;
            this.lblK.TabIndex = 42;
            this.lblK.Text = "K = Kredit Limit";
            // 
            // lblR
            // 
            this.lblR.Location = new System.Drawing.Point(955, 58);
            this.lblR.Name = "lblR";
            this.lblR.Size = new System.Drawing.Size(59, 13);
            this.lblR.StyleController = this.lcFilter;
            this.lblR.TabIndex = 43;
            this.lblR.Text = "R = Release";
            // 
            // lblA
            // 
            this.lblA.Location = new System.Drawing.Point(955, 75);
            this.lblA.Name = "lblA";
            this.lblA.Size = new System.Drawing.Size(129, 13);
            this.lblA.StyleController = this.lcFilter;
            this.lblA.TabIndex = 44;
            this.lblA.Text = "A = Overdue + Kredit Limit";
            // 
            // lcgFilter
            // 
            this.lcgFilter.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.lcgFilter.GroupBordersVisible = false;
            this.lcgFilter.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lcgColA,
            this.lcgColB,
            this.lcgColC,
            this.lcgColD});
            this.lcgFilter.Name = "Root";
            this.lcgFilter.Size = new System.Drawing.Size(1108, 175);
            this.lcgFilter.TextVisible = false;
            // 
            // lcgColA
            // 
            this.lcgColA.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.lcgColA.GroupBordersVisible = false;
            this.lcgColA.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lciEntity,
            this.lciBtnEntity,
            this.lciEntityDesc,
            this.lciBranch,
            this.lciBtnBranch,
            this.lciBranchDesc,
            this.lciSalesman,
            this.lciBtnSales,
            this.lciSalesDesc,
            this.lciOutlet,
            this.lciBtnOutlet,
            this.lciOutletDesc,
            this.lciOutlet2,
            this.lciNoOrder,
            this.lciBtnNoOrder,
            this.lciNoPO,
            this.lciBtnNoPO});
            this.lcgColA.Location = new System.Drawing.Point(0, 0);
            this.lcgColA.Name = "lcgColA";
            this.lcgColA.Size = new System.Drawing.Size(458, 155);
            this.lcgColA.TextVisible = false;
            // 
            // lciEntity
            // 
            this.lciEntity.Control = this.txtEntityCode;
            this.lciEntity.Location = new System.Drawing.Point(0, 0);
            this.lciEntity.Name = "lciEntity";
            this.lciEntity.Size = new System.Drawing.Size(159, 26);
            this.lciEntity.Text = "Entity";
            this.lciEntity.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciBtnEntity
            // 
            this.lciBtnEntity.Control = this.btnPopUpEntityCode;
            this.lciBtnEntity.Location = new System.Drawing.Point(159, 0);
            this.lciBtnEntity.Name = "lciBtnEntity";
            this.lciBtnEntity.Size = new System.Drawing.Size(44, 26);
            this.lciBtnEntity.TextSize = new System.Drawing.Size(0, 0);
            this.lciBtnEntity.TextVisible = false;
            // 
            // lciEntityDesc
            // 
            this.lciEntityDesc.Control = this.txtEntityDesc;
            this.lciEntityDesc.Location = new System.Drawing.Point(203, 0);
            this.lciEntityDesc.Name = "lciEntityDesc";
            this.lciEntityDesc.Size = new System.Drawing.Size(231, 26);
            this.lciEntityDesc.TextSize = new System.Drawing.Size(0, 0);
            this.lciEntityDesc.TextVisible = false;
            // 
            // lciBranch
            // 
            this.lciBranch.Control = this.txtBranchCode;
            this.lciBranch.Location = new System.Drawing.Point(0, 26);
            this.lciBranch.Name = "lciBranch";
            this.lciBranch.Size = new System.Drawing.Size(159, 26);
            this.lciBranch.Text = "Branch";
            this.lciBranch.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciBtnBranch
            // 
            this.lciBtnBranch.Control = this.btnPopUpBranchCode;
            this.lciBtnBranch.Location = new System.Drawing.Point(159, 26);
            this.lciBtnBranch.Name = "lciBtnBranch";
            this.lciBtnBranch.Size = new System.Drawing.Size(44, 26);
            this.lciBtnBranch.TextSize = new System.Drawing.Size(0, 0);
            this.lciBtnBranch.TextVisible = false;
            // 
            // lciBranchDesc
            // 
            this.lciBranchDesc.Control = this.txtBranchDesc;
            this.lciBranchDesc.Location = new System.Drawing.Point(203, 26);
            this.lciBranchDesc.Name = "lciBranchDesc";
            this.lciBranchDesc.Size = new System.Drawing.Size(231, 26);
            this.lciBranchDesc.TextSize = new System.Drawing.Size(0, 0);
            this.lciBranchDesc.TextVisible = false;
            // 
            // lciSalesman
            // 
            this.lciSalesman.Control = this.txtSalesCode;
            this.lciSalesman.Location = new System.Drawing.Point(0, 52);
            this.lciSalesman.Name = "lciSalesman";
            this.lciSalesman.Size = new System.Drawing.Size(159, 26);
            this.lciSalesman.Text = "Salesman";
            this.lciSalesman.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciBtnSales
            // 
            this.lciBtnSales.Control = this.btnPopUpSalesCode;
            this.lciBtnSales.Location = new System.Drawing.Point(159, 52);
            this.lciBtnSales.Name = "lciBtnSales";
            this.lciBtnSales.Size = new System.Drawing.Size(44, 26);
            this.lciBtnSales.TextSize = new System.Drawing.Size(0, 0);
            this.lciBtnSales.TextVisible = false;
            // 
            // lciSalesDesc
            // 
            this.lciSalesDesc.Control = this.txtSalesDesc;
            this.lciSalesDesc.Location = new System.Drawing.Point(203, 52);
            this.lciSalesDesc.Name = "lciSalesDesc";
            this.lciSalesDesc.Size = new System.Drawing.Size(231, 26);
            this.lciSalesDesc.TextSize = new System.Drawing.Size(0, 0);
            this.lciSalesDesc.TextVisible = false;
            // 
            // lciOutlet
            // 
            this.lciOutlet.Control = this.txtOutletCode;
            this.lciOutlet.Location = new System.Drawing.Point(0, 78);
            this.lciOutlet.Name = "lciOutlet";
            this.lciOutlet.Size = new System.Drawing.Size(159, 26);
            this.lciOutlet.Text = "Outlet";
            this.lciOutlet.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciBtnOutlet
            // 
            this.lciBtnOutlet.Control = this.btnPopUpOutlet;
            this.lciBtnOutlet.Location = new System.Drawing.Point(159, 78);
            this.lciBtnOutlet.Name = "lciBtnOutlet";
            this.lciBtnOutlet.Size = new System.Drawing.Size(44, 26);
            this.lciBtnOutlet.TextSize = new System.Drawing.Size(0, 0);
            this.lciBtnOutlet.TextVisible = false;
            // 
            // lciOutletDesc
            // 
            this.lciOutletDesc.Control = this.txtOutletDesc;
            this.lciOutletDesc.Location = new System.Drawing.Point(203, 78);
            this.lciOutletDesc.Name = "lciOutletDesc";
            this.lciOutletDesc.Size = new System.Drawing.Size(176, 26);
            this.lciOutletDesc.TextSize = new System.Drawing.Size(0, 0);
            this.lciOutletDesc.TextVisible = false;
            // 
            // lciOutlet2
            // 
            this.lciOutlet2.Control = this.txtOutlet2;
            this.lciOutlet2.Location = new System.Drawing.Point(379, 78);
            this.lciOutlet2.Name = "lciOutlet2";
            this.lciOutlet2.Size = new System.Drawing.Size(55, 26);
            this.lciOutlet2.TextSize = new System.Drawing.Size(0, 0);
            this.lciOutlet2.TextVisible = false;
            // 
            // lciNoOrder
            // 
            this.lciNoOrder.Control = this.txtNoOrder;
            this.lciNoOrder.Location = new System.Drawing.Point(0, 104);
            this.lciNoOrder.Name = "lciNoOrder";
            this.lciNoOrder.Size = new System.Drawing.Size(159, 27);
            this.lciNoOrder.Text = "No Order";
            this.lciNoOrder.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciBtnNoOrder
            // 
            this.lciBtnNoOrder.Control = this.btnPopUpNoOrder;
            this.lciBtnNoOrder.Location = new System.Drawing.Point(159, 104);
            this.lciBtnNoOrder.Name = "lciBtnNoOrder";
            this.lciBtnNoOrder.Size = new System.Drawing.Size(44, 27);
            this.lciBtnNoOrder.TextSize = new System.Drawing.Size(0, 0);
            this.lciBtnNoOrder.TextVisible = false;
            // 
            // lciNoPO
            // 
            this.lciNoPO.Control = this.txtNoPO;
            this.lciNoPO.Location = new System.Drawing.Point(203, 104);
            this.lciNoPO.Name = "lciNoPO";
            this.lciNoPO.Size = new System.Drawing.Size(187, 27);
            this.lciNoPO.Text = "No PO";
            this.lciNoPO.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciBtnNoPO
            // 
            this.lciBtnNoPO.Control = this.btnPopUpNoPO;
            this.lciBtnNoPO.Location = new System.Drawing.Point(390, 104);
            this.lciBtnNoPO.Name = "lciBtnNoPO";
            this.lciBtnNoPO.Size = new System.Drawing.Size(44, 27);
            this.lciBtnNoPO.TextSize = new System.Drawing.Size(0, 0);
            this.lciBtnNoPO.TextVisible = false;
            // 
            // lcgColB
            // 
            this.lcgColB.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.lcgColB.GroupBordersVisible = false;
            this.lcgColB.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lciSOType,
            this.lciSOStatus,
            this.lciDateSel,
            this.lciTglOrder,
            this.lciTglOrderTo,
            this.emptySpaceItem2});
            this.lcgColB.Location = new System.Drawing.Point(458, 0);
            this.lcgColB.Name = "lcgColB";
            this.lcgColB.Size = new System.Drawing.Size(262, 155);
            this.lcgColB.TextVisible = false;
            // 
            // lciSOType
            // 
            this.lciSOType.Control = this.cbSOType;
            this.lciSOType.Location = new System.Drawing.Point(0, 0);
            this.lciSOType.Name = "lciSOType";
            this.lciSOType.Size = new System.Drawing.Size(238, 24);
            this.lciSOType.Text = "SO Type";
            this.lciSOType.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciSOStatus
            // 
            this.lciSOStatus.Control = this.cbSOStatus;
            this.lciSOStatus.Location = new System.Drawing.Point(0, 24);
            this.lciSOStatus.Name = "lciSOStatus";
            this.lciSOStatus.Size = new System.Drawing.Size(238, 24);
            this.lciSOStatus.Text = "Status SO";
            this.lciSOStatus.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciDateSel
            // 
            this.lciDateSel.Control = this.cbdateselection;
            this.lciDateSel.Location = new System.Drawing.Point(0, 48);
            this.lciDateSel.Name = "lciDateSel";
            this.lciDateSel.Size = new System.Drawing.Size(238, 24);
            this.lciDateSel.Text = "Date Selection";
            this.lciDateSel.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciTglOrder
            // 
            this.lciTglOrder.Control = this.dtTglOrder;
            this.lciTglOrder.Location = new System.Drawing.Point(0, 72);
            this.lciTglOrder.Name = "lciTglOrder";
            this.lciTglOrder.Size = new System.Drawing.Size(107, 59);
            this.lciTglOrder.TextSize = new System.Drawing.Size(0, 0);
            this.lciTglOrder.TextVisible = false;
            // 
            // lciTglOrderTo
            // 
            this.lciTglOrderTo.Control = this.dtTglOrderTo;
            this.lciTglOrderTo.Location = new System.Drawing.Point(107, 72);
            this.lciTglOrderTo.Name = "lciTglOrderTo";
            this.lciTglOrderTo.Size = new System.Drawing.Size(131, 24);
            this.lciTglOrderTo.Text = "To";
            this.lciTglOrderTo.TextAlignMode = DevExpress.XtraLayout.TextAlignModeItem.CustomSize;
            this.lciTglOrderTo.TextSize = new System.Drawing.Size(20, 20);
            this.lciTglOrderTo.TextToControlDistance = 5;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(107, 96);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(131, 35);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // lcgColC
            // 
            this.lcgColC.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.lcgColC.GroupBordersVisible = false;
            this.lcgColC.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lciSource,
            this.lciFlagSloc,
            this.lciShipPlant,
            this.lciBtnShip,
            this.emptyExec});
            this.lcgColC.Location = new System.Drawing.Point(720, 0);
            this.lcgColC.Name = "lcgColC";
            this.lcgColC.Size = new System.Drawing.Size(211, 155);
            this.lcgColC.TextVisible = false;
            // 
            // lciSource
            // 
            this.lciSource.Control = this.cbSource;
            this.lciSource.Location = new System.Drawing.Point(0, 0);
            this.lciSource.Name = "lciSource";
            this.lciSource.Size = new System.Drawing.Size(187, 24);
            this.lciSource.Text = "Source";
            this.lciSource.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciFlagSloc
            // 
            this.lciFlagSloc.Control = this.cbFlagSloc;
            this.lciFlagSloc.Location = new System.Drawing.Point(0, 24);
            this.lciFlagSloc.Name = "lciFlagSloc";
            this.lciFlagSloc.Size = new System.Drawing.Size(187, 24);
            this.lciFlagSloc.Text = "Flag Sloc";
            this.lciFlagSloc.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciShipPlant
            // 
            this.lciShipPlant.Control = this.txtShippingPlant;
            this.lciShipPlant.Location = new System.Drawing.Point(0, 48);
            this.lciShipPlant.Name = "lciShipPlant";
            this.lciShipPlant.Size = new System.Drawing.Size(160, 26);
            this.lciShipPlant.Text = "Ship. Plant";
            this.lciShipPlant.TextSize = new System.Drawing.Size(69, 13);
            // 
            // lciBtnShip
            // 
            this.lciBtnShip.Control = this.btnShipping;
            this.lciBtnShip.Location = new System.Drawing.Point(160, 48);
            this.lciBtnShip.Name = "lciBtnShip";
            this.lciBtnShip.Size = new System.Drawing.Size(27, 26);
            this.lciBtnShip.TextSize = new System.Drawing.Size(0, 0);
            this.lciBtnShip.TextVisible = false;
            // 
            // emptyExec
            // 
            this.emptyExec.AllowHotTrack = false;
            this.emptyExec.Location = new System.Drawing.Point(0, 74);
            this.emptyExec.Name = "emptyExec";
            this.emptyExec.Size = new System.Drawing.Size(187, 57);
            this.emptyExec.TextSize = new System.Drawing.Size(0, 0);
            // 
            // lcgColD
            // 
            this.lcgColD.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.lcgColD.GroupBordersVisible = false;
            this.lcgColD.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lciLblD,
            this.lciLblK,
            this.lciLblR,
            this.lciLblA,
            this.emptySpaceItem1,
            this.lciExec});
            this.lcgColD.Location = new System.Drawing.Point(931, 0);
            this.lcgColD.Name = "lcgColD";
            this.lcgColD.Size = new System.Drawing.Size(157, 155);
            this.lcgColD.TextVisible = false;
            // 
            // lciLblD
            // 
            this.lciLblD.Control = this.lblD;
            this.lciLblD.Location = new System.Drawing.Point(0, 0);
            this.lciLblD.Name = "lciLblD";
            this.lciLblD.Size = new System.Drawing.Size(133, 17);
            this.lciLblD.TextSize = new System.Drawing.Size(0, 0);
            this.lciLblD.TextVisible = false;
            // 
            // lciLblK
            // 
            this.lciLblK.Control = this.lblK;
            this.lciLblK.Location = new System.Drawing.Point(0, 17);
            this.lciLblK.Name = "lciLblK";
            this.lciLblK.Size = new System.Drawing.Size(133, 17);
            this.lciLblK.TextSize = new System.Drawing.Size(0, 0);
            this.lciLblK.TextVisible = false;
            // 
            // lciLblR
            // 
            this.lciLblR.Control = this.lblR;
            this.lciLblR.Location = new System.Drawing.Point(0, 34);
            this.lciLblR.Name = "lciLblR";
            this.lciLblR.Size = new System.Drawing.Size(133, 17);
            this.lciLblR.TextSize = new System.Drawing.Size(0, 0);
            this.lciLblR.TextVisible = false;
            // 
            // lciLblA
            // 
            this.lciLblA.Control = this.lblA;
            this.lciLblA.Location = new System.Drawing.Point(0, 51);
            this.lciLblA.Name = "lciLblA";
            this.lciLblA.Size = new System.Drawing.Size(133, 17);
            this.lciLblA.TextSize = new System.Drawing.Size(0, 0);
            this.lciLblA.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 68);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(133, 37);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // lciExec
            // 
            this.lciExec.Control = this.btnExec;
            this.lciExec.Location = new System.Drawing.Point(0, 105);
            this.lciExec.Name = "lciExec";
            this.lciExec.Size = new System.Drawing.Size(133, 26);
            this.lciExec.TextSize = new System.Drawing.Size(0, 0);
            this.lciExec.TextVisible = false;
            // 
            // txtShippingPlantId
            // 
            this.txtShippingPlantId.Location = new System.Drawing.Point(784, 63);
            this.txtShippingPlantId.Name = "txtShippingPlantId";
            this.txtShippingPlantId.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtShippingPlantId.Properties.ReadOnly = true;
            this.txtShippingPlantId.Size = new System.Drawing.Size(114, 20);
            this.txtShippingPlantId.TabIndex = 38;
            this.txtShippingPlantId.Visible = false;
            // 
            // gcSalesHeader
            // 
            this.gcSalesHeader.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gcSalesHeader.Location = new System.Drawing.Point(6, 173);
            this.gcSalesHeader.MainView = this.gvSalesHeader;
            this.gcSalesHeader.Name = "gcSalesHeader";
            this.gcSalesHeader.Size = new System.Drawing.Size(1114, 123);
            this.gcSalesHeader.TabIndex = 1;
            this.gcSalesHeader.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvSalesHeader});
            // 
            // gvSalesHeader
            // 
            this.gvSalesHeader.GridControl = this.gcSalesHeader;
            this.gvSalesHeader.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            this.gvSalesHeader.Name = "gvSalesHeader";
            this.gvSalesHeader.OptionsView.ColumnAutoWidth = false;
            this.gvSalesHeader.OptionsView.ShowGroupPanel = false;
            // 
            // gcSalesDetail
            // 
            this.gcSalesDetail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gcSalesDetail.Location = new System.Drawing.Point(6, 299);
            this.gcSalesDetail.MainView = this.gvSalesDetail;
            this.gcSalesDetail.Name = "gcSalesDetail";
            this.gcSalesDetail.Size = new System.Drawing.Size(1114, 95);
            this.gcSalesDetail.TabIndex = 2;
            this.gcSalesDetail.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvSalesDetail});
            // 
            // gvSalesDetail
            // 
            this.gvSalesDetail.GridControl = this.gcSalesDetail;
            this.gvSalesDetail.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            this.gvSalesDetail.Name = "gvSalesDetail";
            this.gvSalesDetail.OptionsView.ColumnAutoWidth = false;
            this.gvSalesDetail.OptionsView.ShowGroupPanel = false;
            // 
            // gcPromosiRp
            // 
            this.gcPromosiRp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.gcPromosiRp.Location = new System.Drawing.Point(6, 397);
            this.gcPromosiRp.MainView = this.gvPromosiRp;
            this.gcPromosiRp.Name = "gcPromosiRp";
            this.gcPromosiRp.Size = new System.Drawing.Size(449, 60);
            this.gcPromosiRp.TabIndex = 3;
            this.gcPromosiRp.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvPromosiRp});
            // 
            // gvPromosiRp
            // 
            this.gvPromosiRp.GridControl = this.gcPromosiRp;
            this.gvPromosiRp.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            this.gvPromosiRp.Name = "gvPromosiRp";
            this.gvPromosiRp.OptionsView.ColumnAutoWidth = false;
            this.gvPromosiRp.OptionsView.ShowGroupPanel = false;
            // 
            // gcPromosiQty
            // 
            this.gcPromosiQty.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gcPromosiQty.Location = new System.Drawing.Point(457, 397);
            this.gcPromosiQty.MainView = this.gvPromosiQty;
            this.gcPromosiQty.Name = "gcPromosiQty";
            this.gcPromosiQty.Size = new System.Drawing.Size(654, 60);
            this.gcPromosiQty.TabIndex = 4;
            this.gcPromosiQty.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvPromosiQty});
            // 
            // gvPromosiQty
            // 
            this.gvPromosiQty.GridControl = this.gcPromosiQty;
            this.gvPromosiQty.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            this.gvPromosiQty.Name = "gvPromosiQty";
            this.gvPromosiQty.OptionsView.ColumnAutoWidth = false;
            this.gvPromosiQty.OptionsView.ShowGroupPanel = false;
            // 
            // gcDisc
            // 
            this.gcDisc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gcDisc.Location = new System.Drawing.Point(6, 460);
            this.gcDisc.MainView = this.gvDisc;
            this.gcDisc.Name = "gcDisc";
            this.gcDisc.Size = new System.Drawing.Size(1105, 70);
            this.gcDisc.TabIndex = 5;
            this.gcDisc.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gvDisc});
            // 
            // gvDisc
            // 
            this.gvDisc.GridControl = this.gcDisc;
            this.gvDisc.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always;
            this.gvDisc.Name = "gvDisc";
            this.gvDisc.OptionsView.ColumnAutoWidth = false;
            this.gvDisc.OptionsView.ShowGroupPanel = false;
            // 
            // groupTotal
            // 
            this.groupTotal.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupTotal.Controls.Add(this.lblSubTotal);
            this.groupTotal.Controls.Add(this.txtSubTotal);
            this.groupTotal.Controls.Add(this.lblCashDisc);
            this.groupTotal.Controls.Add(this.txtCashDiscP);
            this.groupTotal.Controls.Add(this.lblPercent);
            this.groupTotal.Controls.Add(this.lblPromosiU);
            this.groupTotal.Controls.Add(this.txtPromosiU);
            this.groupTotal.Controls.Add(this.lblRpCashDisc);
            this.groupTotal.Controls.Add(this.txtCashDiscR);
            this.groupTotal.Controls.Add(this.lblDisc1);
            this.groupTotal.Controls.Add(this.txtDisc1);
            this.groupTotal.Controls.Add(this.lblPPN);
            this.groupTotal.Controls.Add(this.txtPPN);
            this.groupTotal.Controls.Add(this.lblTotalInvoice);
            this.groupTotal.Controls.Add(this.txtTtlInvoice);
            this.groupTotal.Controls.Add(this.lblBebanCashDisc);
            this.groupTotal.Controls.Add(this.txtBebanCashDisc);
            this.groupTotal.Location = new System.Drawing.Point(6, 536);
            this.groupTotal.Name = "groupTotal";
            this.groupTotal.ShowCaption = false;
            this.groupTotal.Size = new System.Drawing.Size(1105, 70);
            this.groupTotal.TabIndex = 6;
            // 
            // lblSubTotal
            // 
            this.lblSubTotal.Location = new System.Drawing.Point(9, 14);
            this.lblSubTotal.Name = "lblSubTotal";
            this.lblSubTotal.Size = new System.Drawing.Size(45, 13);
            this.lblSubTotal.TabIndex = 0;
            this.lblSubTotal.Text = "Sub Total";
            // 
            // txtSubTotal
            // 
            this.txtSubTotal.Location = new System.Drawing.Point(72, 12);
            this.txtSubTotal.Name = "txtSubTotal";
            this.txtSubTotal.Properties.Appearance.Options.UseTextOptions = true;
            this.txtSubTotal.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtSubTotal.Properties.ReadOnly = true;
            this.txtSubTotal.Size = new System.Drawing.Size(141, 20);
            this.txtSubTotal.TabIndex = 1;
            // 
            // lblCashDisc
            // 
            this.lblCashDisc.Location = new System.Drawing.Point(9, 37);
            this.lblCashDisc.Name = "lblCashDisc";
            this.lblCashDisc.Size = new System.Drawing.Size(46, 13);
            this.lblCashDisc.TabIndex = 2;
            this.lblCashDisc.Text = "Cash Disc";
            // 
            // txtCashDiscP
            // 
            this.txtCashDiscP.Location = new System.Drawing.Point(71, 34);
            this.txtCashDiscP.Name = "txtCashDiscP";
            this.txtCashDiscP.Properties.Appearance.Options.UseTextOptions = true;
            this.txtCashDiscP.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtCashDiscP.Properties.ReadOnly = true;
            this.txtCashDiscP.Size = new System.Drawing.Size(81, 20);
            this.txtCashDiscP.TabIndex = 3;
            // 
            // lblPercent
            // 
            this.lblPercent.Location = new System.Drawing.Point(158, 37);
            this.lblPercent.Name = "lblPercent";
            this.lblPercent.Size = new System.Drawing.Size(11, 13);
            this.lblPercent.TabIndex = 4;
            this.lblPercent.Text = "%";
            // 
            // lblPromosiU
            // 
            this.lblPromosiU.Location = new System.Drawing.Point(234, 14);
            this.lblPromosiU.Name = "lblPromosiU";
            this.lblPromosiU.Size = new System.Drawing.Size(65, 13);
            this.lblPromosiU.TabIndex = 5;
            this.lblPromosiU.Text = "Promosi Uang";
            // 
            // txtPromosiU
            // 
            this.txtPromosiU.Location = new System.Drawing.Point(309, 12);
            this.txtPromosiU.Name = "txtPromosiU";
            this.txtPromosiU.Properties.Appearance.Options.UseTextOptions = true;
            this.txtPromosiU.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtPromosiU.Properties.ReadOnly = true;
            this.txtPromosiU.Size = new System.Drawing.Size(141, 20);
            this.txtPromosiU.TabIndex = 6;
            // 
            // lblRpCashDisc
            // 
            this.lblRpCashDisc.Location = new System.Drawing.Point(234, 37);
            this.lblRpCashDisc.Name = "lblRpCashDisc";
            this.lblRpCashDisc.Size = new System.Drawing.Size(62, 13);
            this.lblRpCashDisc.TabIndex = 7;
            this.lblRpCashDisc.Text = "Rp Cash Disc";
            // 
            // txtCashDiscR
            // 
            this.txtCashDiscR.Location = new System.Drawing.Point(309, 34);
            this.txtCashDiscR.Name = "txtCashDiscR";
            this.txtCashDiscR.Properties.Appearance.Options.UseTextOptions = true;
            this.txtCashDiscR.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtCashDiscR.Properties.ReadOnly = true;
            this.txtCashDiscR.Size = new System.Drawing.Size(141, 20);
            this.txtCashDiscR.TabIndex = 8;
            // 
            // lblDisc1
            // 
            this.lblDisc1.Location = new System.Drawing.Point(475, 14);
            this.lblDisc1.Name = "lblDisc1";
            this.lblDisc1.Size = new System.Drawing.Size(57, 13);
            this.lblDisc1.TabIndex = 9;
            this.lblDisc1.Text = "Disc 1 sd 10";
            // 
            // txtDisc1
            // 
            this.txtDisc1.Location = new System.Drawing.Point(545, 13);
            this.txtDisc1.Name = "txtDisc1";
            this.txtDisc1.Properties.Appearance.Options.UseTextOptions = true;
            this.txtDisc1.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtDisc1.Properties.ReadOnly = true;
            this.txtDisc1.Size = new System.Drawing.Size(141, 20);
            this.txtDisc1.TabIndex = 10;
            // 
            // lblPPN
            // 
            this.lblPPN.Location = new System.Drawing.Point(513, 37);
            this.lblPPN.Name = "lblPPN";
            this.lblPPN.Size = new System.Drawing.Size(19, 13);
            this.lblPPN.TabIndex = 11;
            this.lblPPN.Text = "PPN";
            // 
            // txtPPN
            // 
            this.txtPPN.Location = new System.Drawing.Point(545, 35);
            this.txtPPN.Name = "txtPPN";
            this.txtPPN.Properties.Appearance.Options.UseTextOptions = true;
            this.txtPPN.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtPPN.Properties.ReadOnly = true;
            this.txtPPN.Size = new System.Drawing.Size(141, 20);
            this.txtPPN.TabIndex = 12;
            // 
            // lblTotalInvoice
            // 
            this.lblTotalInvoice.Location = new System.Drawing.Point(719, 15);
            this.lblTotalInvoice.Name = "lblTotalInvoice";
            this.lblTotalInvoice.Size = new System.Drawing.Size(62, 13);
            this.lblTotalInvoice.TabIndex = 13;
            this.lblTotalInvoice.Text = "Total Invoice";
            // 
            // txtTtlInvoice
            // 
            this.txtTtlInvoice.Location = new System.Drawing.Point(722, 34);
            this.txtTtlInvoice.Name = "txtTtlInvoice";
            this.txtTtlInvoice.Properties.Appearance.Options.UseTextOptions = true;
            this.txtTtlInvoice.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtTtlInvoice.Properties.ReadOnly = true;
            this.txtTtlInvoice.Size = new System.Drawing.Size(141, 20);
            this.txtTtlInvoice.TabIndex = 14;
            // 
            // lblBebanCashDisc
            // 
            this.lblBebanCashDisc.Location = new System.Drawing.Point(875, 14);
            this.lblBebanCashDisc.Name = "lblBebanCashDisc";
            this.lblBebanCashDisc.Size = new System.Drawing.Size(79, 13);
            this.lblBebanCashDisc.TabIndex = 15;
            this.lblBebanCashDisc.Text = "Beban Cash Disc";
            // 
            // txtBebanCashDisc
            // 
            this.txtBebanCashDisc.Location = new System.Drawing.Point(878, 33);
            this.txtBebanCashDisc.Name = "txtBebanCashDisc";
            this.txtBebanCashDisc.Properties.Appearance.Options.UseTextOptions = true;
            this.txtBebanCashDisc.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.txtBebanCashDisc.Properties.ReadOnly = true;
            this.txtBebanCashDisc.Size = new System.Drawing.Size(111, 20);
            this.txtBebanCashDisc.TabIndex = 16;
            // 
            // frmSOManualEntryList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1120, 646);
            this.IconOptions.Icon = ((System.Drawing.Icon)(resources.GetObject("frmSOManualEntryList.IconOptions.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmSOManualEntryList";
            this.Text = "SO Manual Entry List";
            this.Load += new System.EventHandler(this.frmSOManualEntryList_Load);
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMarqueeProgressBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemProgressBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.contentPanel)).EndInit();
            this.contentPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelMain)).EndInit();
            this.panelMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.lcFilter)).EndInit();
            this.lcFilter.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txtEntityCode.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtEntityDesc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBranchCode.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBranchDesc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesCode.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesDesc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtOutletCode.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtOutletDesc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtOutlet2.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtNoOrder.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtNoPO.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbSOType.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbSOStatus.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbdateselection.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrder.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrder.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrderTo.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglOrderTo.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbSource.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbFlagSloc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtShippingPlant.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgFilter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColA)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciEntity)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnEntity)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciEntityDesc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBranch)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnBranch)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBranchDesc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSalesman)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnSales)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSalesDesc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciOutlet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnOutlet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciOutletDesc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciOutlet2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciNoOrder)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnNoOrder)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciNoPO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnNoPO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColB)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSOType)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSOStatus)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciDateSel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciTglOrder)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciTglOrderTo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColC)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciFlagSloc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciShipPlant)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciBtnShip)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptyExec)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lcgColD)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblD)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblK)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblR)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciLblA)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lciExec)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtShippingPlantId.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcSalesHeader)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvSalesHeader)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcSalesDetail)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvSalesDetail)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcPromosiRp)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPromosiRp)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcPromosiQty)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPromosiQty)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gcDisc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvDisc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupTotal)).EndInit();
            this.groupTotal.ResumeLayout(false);
            this.groupTotal.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtSubTotal.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtCashDiscP.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPromosiU.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtCashDiscR.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtDisc1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPPN.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTtlInvoice.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtBebanCashDisc.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private DevExpress.XtraEditors.PanelControl panelMain;
        private DevExpress.XtraLayout.LayoutControl lcFilter;
        private DevExpress.XtraLayout.LayoutControlGroup lcgFilter;
        private DevExpress.XtraLayout.LayoutControlGroup lcgColA;
        private DevExpress.XtraLayout.LayoutControlGroup lcgColB;
        private DevExpress.XtraLayout.LayoutControlGroup lcgColC;
        private DevExpress.XtraLayout.LayoutControlGroup lcgColD;
        private DevExpress.XtraLayout.LayoutControlItem lciEntity;
        private DevExpress.XtraLayout.LayoutControlItem lciBtnEntity;
        private DevExpress.XtraLayout.LayoutControlItem lciEntityDesc;
        private DevExpress.XtraLayout.LayoutControlItem lciBranch;
        private DevExpress.XtraLayout.LayoutControlItem lciBtnBranch;
        private DevExpress.XtraLayout.LayoutControlItem lciBranchDesc;
        private DevExpress.XtraLayout.LayoutControlItem lciSalesman;
        private DevExpress.XtraLayout.LayoutControlItem lciBtnSales;
        private DevExpress.XtraLayout.LayoutControlItem lciSalesDesc;
        private DevExpress.XtraLayout.LayoutControlItem lciOutlet;
        private DevExpress.XtraLayout.LayoutControlItem lciBtnOutlet;
        private DevExpress.XtraLayout.LayoutControlItem lciOutletDesc;
        private DevExpress.XtraLayout.LayoutControlItem lciOutlet2;
        private DevExpress.XtraLayout.LayoutControlItem lciNoOrder;
        private DevExpress.XtraLayout.LayoutControlItem lciBtnNoOrder;
        private DevExpress.XtraLayout.LayoutControlItem lciNoPO;
        private DevExpress.XtraLayout.LayoutControlItem lciBtnNoPO;
        private DevExpress.XtraLayout.LayoutControlItem lciSOType;
        private DevExpress.XtraLayout.LayoutControlItem lciSOStatus;
        private DevExpress.XtraLayout.LayoutControlItem lciDateSel;
        private DevExpress.XtraLayout.LayoutControlItem lciTglOrder;
        private DevExpress.XtraLayout.LayoutControlItem lciTglOrderTo;
        private DevExpress.XtraLayout.LayoutControlItem lciSource;
        private DevExpress.XtraLayout.LayoutControlItem lciFlagSloc;
        private DevExpress.XtraLayout.LayoutControlItem lciShipPlant;
        private DevExpress.XtraLayout.LayoutControlItem lciBtnShip;
        private DevExpress.XtraLayout.LayoutControlItem lciExec;
        private DevExpress.XtraLayout.LayoutControlItem lciLblD;
        private DevExpress.XtraLayout.LayoutControlItem lciLblK;
        private DevExpress.XtraLayout.LayoutControlItem lciLblR;
        private DevExpress.XtraLayout.LayoutControlItem lciLblA;
        private DevExpress.XtraLayout.EmptySpaceItem emptyExec;
        private DevExpress.XtraEditors.GroupControl groupTotal;
        private DevExpress.XtraEditors.TextEdit txtEntityCode;
        private DevExpress.XtraEditors.TextEdit txtEntityDesc;
        private DevExpress.XtraEditors.TextEdit txtBranchCode;
        private DevExpress.XtraEditors.TextEdit txtBranchDesc;
        private DevExpress.XtraEditors.TextEdit txtSalesCode;
        private DevExpress.XtraEditors.TextEdit txtSalesDesc;
        private DevExpress.XtraEditors.TextEdit txtOutletCode;
        private DevExpress.XtraEditors.TextEdit txtOutletDesc;
        private DevExpress.XtraEditors.TextEdit txtOutlet2;
        private DevExpress.XtraEditors.TextEdit txtNoOrder;
        private DevExpress.XtraEditors.TextEdit txtNoPO;
        private DevExpress.XtraEditors.TextEdit txtShippingPlant;
        private DevExpress.XtraEditors.TextEdit txtShippingPlantId;
        private DevExpress.XtraEditors.TextEdit txtSubTotal;
        private DevExpress.XtraEditors.TextEdit txtCashDiscP;
        private DevExpress.XtraEditors.TextEdit txtPromosiU;
        private DevExpress.XtraEditors.TextEdit txtCashDiscR;
        private DevExpress.XtraEditors.TextEdit txtDisc1;
        private DevExpress.XtraEditors.TextEdit txtPPN;
        private DevExpress.XtraEditors.TextEdit txtTtlInvoice;
        private DevExpress.XtraEditors.TextEdit txtBebanCashDisc;
        private DevExpress.XtraEditors.SimpleButton btnPopUpEntityCode;
        private DevExpress.XtraEditors.SimpleButton btnPopUpBranchCode;
        private DevExpress.XtraEditors.SimpleButton btnPopUpSalesCode;
        private DevExpress.XtraEditors.SimpleButton btnPopUpOutlet;
        private DevExpress.XtraEditors.SimpleButton btnPopUpNoOrder;
        private DevExpress.XtraEditors.SimpleButton btnPopUpNoPO;
        private DevExpress.XtraEditors.SimpleButton btnShipping;
        private DevExpress.XtraEditors.SimpleButton btnExec;
        private DevExpress.XtraEditors.LookUpEdit cbSOType;
        private DevExpress.XtraEditors.LookUpEdit cbSOStatus;
        private DevExpress.XtraEditors.LookUpEdit cbdateselection;
        private DevExpress.XtraEditors.LookUpEdit cbSource;
        private DevExpress.XtraEditors.LookUpEdit cbFlagSloc;
        private DevExpress.XtraEditors.DateEdit dtTglOrder;
        private DevExpress.XtraEditors.DateEdit dtTglOrderTo;
        private DevExpress.XtraEditors.LabelControl lblD;
        private DevExpress.XtraEditors.LabelControl lblK;
        private DevExpress.XtraEditors.LabelControl lblR;
        private DevExpress.XtraEditors.LabelControl lblA;
        private DevExpress.XtraEditors.LabelControl lblSubTotal;
        private DevExpress.XtraEditors.LabelControl lblCashDisc;
        private DevExpress.XtraEditors.LabelControl lblPercent;
        private DevExpress.XtraEditors.LabelControl lblPromosiU;
        private DevExpress.XtraEditors.LabelControl lblRpCashDisc;
        private DevExpress.XtraEditors.LabelControl lblDisc1;
        private DevExpress.XtraEditors.LabelControl lblPPN;
        private DevExpress.XtraEditors.LabelControl lblTotalInvoice;
        private DevExpress.XtraEditors.LabelControl lblBebanCashDisc;
        private DevExpress.XtraGrid.GridControl gcSalesHeader;
        private DevExpress.XtraGrid.Views.Grid.GridView gvSalesHeader;
        private DevExpress.XtraGrid.GridControl gcSalesDetail;
        private DevExpress.XtraGrid.Views.Grid.GridView gvSalesDetail;
        private DevExpress.XtraGrid.GridControl gcPromosiRp;
        private DevExpress.XtraGrid.Views.Grid.GridView gvPromosiRp;
        private DevExpress.XtraGrid.GridControl gcPromosiQty;
        private DevExpress.XtraGrid.Views.Grid.GridView gvPromosiQty;
        private DevExpress.XtraGrid.GridControl gcDisc;
        private DevExpress.XtraGrid.Views.Grid.GridView gvDisc;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
    }
}
