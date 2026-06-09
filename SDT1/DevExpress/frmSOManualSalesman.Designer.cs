namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    partial class frmSOManualSalesman
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSOManualSalesman));
            this.btnCancel = new DevExpress.XtraEditors.SimpleButton();
            this.btnOk = new DevExpress.XtraEditors.SimpleButton();
            this.grbSalesman = new DevExpress.XtraEditors.GroupControl();
            this.txtSalesDesc = new DevExpress.XtraEditors.TextEdit();
            this.txtSalesIDView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.cbBranchIDView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.lblWHLoc2 = new DevExpress.XtraEditors.TextEdit();
            this.lblWHLoc1 = new DevExpress.XtraEditors.TextEdit();
            this.lblEmployee = new DevExpress.XtraEditors.TextEdit();
            this.lblOrderType = new DevExpress.XtraEditors.TextEdit();
            this.lblKdTipe = new DevExpress.XtraEditors.TextEdit();
            this.cbBranchID = new DevExpress.XtraEditors.SearchLookUpEdit();
            this.cbEntityID = new DevExpress.XtraEditors.LookUpEdit();
            this.label25 = new DevExpress.XtraEditors.LabelControl();
            this.label23 = new DevExpress.XtraEditors.LabelControl();
            this.txtGudang = new DevExpress.XtraEditors.TextEdit();
            this.txtTeam = new DevExpress.XtraEditors.TextEdit();
            this.dtTglTrans = new DevExpress.XtraEditors.DateEdit();
            this.txtTipeSales = new DevExpress.XtraEditors.TextEdit();
            this.txtSalesID = new DevExpress.XtraEditors.SearchLookUpEdit();
            this.label6 = new DevExpress.XtraEditors.LabelControl();
            this.label5 = new DevExpress.XtraEditors.LabelControl();
            this.label4 = new DevExpress.XtraEditors.LabelControl();
            this.label3 = new DevExpress.XtraEditors.LabelControl();
            this.label2 = new DevExpress.XtraEditors.LabelControl();
            this.label1 = new DevExpress.XtraEditors.LabelControl();
            this.rdAllAktif = new DevExpress.XtraEditors.CheckEdit();
            this.rdAktif = new DevExpress.XtraEditors.CheckEdit();
            ((System.ComponentModel.ISupportInitialize)(this.grbSalesman)).BeginInit();
            this.grbSalesman.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesDesc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesIDView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbBranchIDView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblWHLoc2.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblWHLoc1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblEmployee.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblOrderType.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblKdTipe.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbBranchID.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbEntityID.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtGudang.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTeam.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglTrans.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglTrans.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTipeSales.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesID.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rdAllAktif.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rdAktif.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.ImageOptions.Image = global::TIRASnDNet.Properties.Resources._1490217733_3;
            this.btnCancel.Location = new System.Drawing.Point(408, 260);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(85, 22);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOk
            // 
            this.btnOk.ImageOptions.Image = global::TIRASnDNet.Properties.Resources._1490217776_4;
            this.btnOk.Location = new System.Drawing.Point(314, 260);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(85, 22);
            this.btnOk.TabIndex = 7;
            this.btnOk.Text = "&OK";
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // grbSalesman
            // 
            this.grbSalesman.Controls.Add(this.txtSalesDesc);
            this.grbSalesman.Controls.Add(this.lblWHLoc2);
            this.grbSalesman.Controls.Add(this.lblWHLoc1);
            this.grbSalesman.Controls.Add(this.lblEmployee);
            this.grbSalesman.Controls.Add(this.lblOrderType);
            this.grbSalesman.Controls.Add(this.lblKdTipe);
            this.grbSalesman.Controls.Add(this.cbBranchID);
            this.grbSalesman.Controls.Add(this.cbEntityID);
            this.grbSalesman.Controls.Add(this.label25);
            this.grbSalesman.Controls.Add(this.label23);
            this.grbSalesman.Controls.Add(this.txtGudang);
            this.grbSalesman.Controls.Add(this.txtTeam);
            this.grbSalesman.Controls.Add(this.dtTglTrans);
            this.grbSalesman.Controls.Add(this.txtTipeSales);
            this.grbSalesman.Controls.Add(this.txtSalesID);
            this.grbSalesman.Controls.Add(this.label6);
            this.grbSalesman.Controls.Add(this.label5);
            this.grbSalesman.Controls.Add(this.label4);
            this.grbSalesman.Controls.Add(this.label3);
            this.grbSalesman.Controls.Add(this.label2);
            this.grbSalesman.Controls.Add(this.label1);
            this.grbSalesman.Controls.Add(this.rdAllAktif);
            this.grbSalesman.Controls.Add(this.rdAktif);
            this.grbSalesman.Location = new System.Drawing.Point(12, 8);
            this.grbSalesman.Name = "grbSalesman";
            this.grbSalesman.ShowCaption = false;
            this.grbSalesman.Size = new System.Drawing.Size(484, 245);
            this.grbSalesman.TabIndex = 13;
            // 
            // txtSalesDesc
            // 
            this.txtSalesDesc.EditValue = "";
            this.txtSalesDesc.Location = new System.Drawing.Point(267, 87);
            this.txtSalesDesc.Name = "txtSalesDesc";
            this.txtSalesDesc.Properties.Appearance.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.txtSalesDesc.Properties.Appearance.ForeColor = System.Drawing.Color.White;
            this.txtSalesDesc.Properties.Appearance.Options.UseBackColor = true;
            this.txtSalesDesc.Properties.Appearance.Options.UseForeColor = true;
            this.txtSalesDesc.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtSalesDesc.Properties.MaxLength = 5;
            this.txtSalesDesc.Properties.ReadOnly = true;
            this.txtSalesDesc.Size = new System.Drawing.Size(214, 20);
            this.txtSalesDesc.TabIndex = 200;
            this.txtSalesDesc.TabStop = false;
            // hidden controls
            this.lblWHLoc2.Location = new System.Drawing.Point(345, 210);
            this.lblWHLoc2.Name = "lblWHLoc2";
            this.lblWHLoc2.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.lblWHLoc2.Properties.MaxLength = 5;
            this.lblWHLoc2.Size = new System.Drawing.Size(25, 20);
            this.lblWHLoc2.TabIndex = 181;
            this.lblWHLoc2.Visible = false;
            this.lblWHLoc1.Location = new System.Drawing.Point(314, 210);
            this.lblWHLoc1.Name = "lblWHLoc1";
            this.lblWHLoc1.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.lblWHLoc1.Properties.MaxLength = 5;
            this.lblWHLoc1.Size = new System.Drawing.Size(25, 20);
            this.lblWHLoc1.TabIndex = 180;
            this.lblWHLoc1.Visible = false;
            this.lblEmployee.Location = new System.Drawing.Point(444, 116);
            this.lblEmployee.Name = "lblEmployee";
            this.lblEmployee.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.lblEmployee.Properties.MaxLength = 5;
            this.lblEmployee.Size = new System.Drawing.Size(25, 20);
            this.lblEmployee.TabIndex = 179;
            this.lblEmployee.Visible = false;
            this.lblOrderType.Location = new System.Drawing.Point(286, 116);
            this.lblOrderType.Name = "lblOrderType";
            this.lblOrderType.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.lblOrderType.Properties.MaxLength = 5;
            this.lblOrderType.Size = new System.Drawing.Size(25, 20);
            this.lblOrderType.TabIndex = 178;
            this.lblOrderType.Visible = false;
            this.lblKdTipe.Location = new System.Drawing.Point(255, 116);
            this.lblKdTipe.Name = "lblKdTipe";
            this.lblKdTipe.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.lblKdTipe.Properties.MaxLength = 5;
            this.lblKdTipe.Size = new System.Drawing.Size(25, 20);
            this.lblKdTipe.TabIndex = 177;
            this.lblKdTipe.Visible = false;
            // 
            // cbBranchID
            // 
            this.cbBranchID.Location = new System.Drawing.Point(116, 35);
            this.cbBranchID.Name = "cbBranchID";
            this.cbBranchID.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] { new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo) });
            this.cbBranchID.Properties.NullText = "";
            this.cbBranchID.Properties.PopupView = this.cbBranchIDView;
            this.cbBranchID.Properties.ShowClearButton = true;
            this.cbBranchID.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            this.cbBranchID.Size = new System.Drawing.Size(298, 20);
            this.cbBranchID.TabIndex = 1;
            this.cbBranchID.EditValueChanged += new System.EventHandler(this.cbBranchID_EditValueChanged);
            this.cbBranchID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cbBranchID_KeyPress);
            // 
            // cbBranchIDView
            // 
            this.cbBranchIDView.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
            this.cbBranchIDView.Name = "cbBranchIDView";
            this.cbBranchIDView.OptionsSelection.EnableAppearanceFocusedCell = false;
            this.cbBranchIDView.OptionsView.ShowGroupPanel = false;
            // 
            // cbEntityID
            // 
            this.cbEntityID.Location = new System.Drawing.Point(116, 11);
            this.cbEntityID.Name = "cbEntityID";
            this.cbEntityID.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] { new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo) });
            this.cbEntityID.Properties.NullText = "";
            this.cbEntityID.Properties.ShowFooter = false;
            this.cbEntityID.Properties.ShowHeader = false;
            this.cbEntityID.Size = new System.Drawing.Size(298, 20);
            this.cbEntityID.TabIndex = 0;
            this.cbEntityID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cbEntityID_KeyPress);
            // 
            // labels and editable fields
            // 
            this.label25.Location = new System.Drawing.Point(76, 14);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(25, 13);
            this.label25.Text = "Entity";
            this.label23.Location = new System.Drawing.Point(75, 38);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(31, 13);
            this.label23.Text = "Branch";
            this.txtGudang.Location = new System.Drawing.Point(117, 199);
            this.txtGudang.Name = "txtGudang";
            this.txtGudang.Size = new System.Drawing.Size(199, 20);
            this.txtGudang.TabIndex = 6;
            this.txtTeam.Location = new System.Drawing.Point(117, 174);
            this.txtTeam.Name = "txtTeam";
            this.txtTeam.Size = new System.Drawing.Size(199, 20);
            this.txtTeam.TabIndex = 5;
            this.dtTglTrans.EditValue = System.DateTime.Today;
            this.dtTglTrans.Location = new System.Drawing.Point(117, 148);
            this.dtTglTrans.Name = "dtTglTrans";
            this.dtTglTrans.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] { new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo) });
            this.dtTglTrans.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] { new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo) });
            this.dtTglTrans.Properties.DisplayFormat.FormatString = "dd MMM yyyy";
            this.dtTglTrans.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dtTglTrans.Properties.EditFormat.FormatString = "dd MMM yyyy";
            this.dtTglTrans.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.dtTglTrans.Properties.Mask.EditMask = "dd MMM yyyy";
            this.dtTglTrans.Size = new System.Drawing.Size(108, 20);
            this.dtTglTrans.TabIndex = 4;
            this.txtTipeSales.Location = new System.Drawing.Point(117, 122);
            this.txtTipeSales.Name = "txtTipeSales";
            this.txtTipeSales.Size = new System.Drawing.Size(139, 20);
            this.txtTipeSales.TabIndex = 3;
            this.txtSalesID.Location = new System.Drawing.Point(117, 87);
            this.txtSalesID.Name = "txtSalesID";
            this.txtSalesID.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] { new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo) });
            this.txtSalesID.Properties.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtSalesID.Properties.NullText = "";
            this.txtSalesID.Properties.PopupView = this.txtSalesIDView;
            this.txtSalesID.Properties.ShowClearButton = true;
            this.txtSalesID.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            this.txtSalesID.Size = new System.Drawing.Size(140, 20);
            this.txtSalesID.TabIndex = 2;
            this.txtSalesID.EditValueChanged += new System.EventHandler(this.txtSalesID_EditValueChanged);
            this.txtSalesID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSalesID_KeyDown);
            this.txtSalesID.Leave += new System.EventHandler(this.txtSalesID_Leave);
            // 
            // txtSalesIDView
            // 
            this.txtSalesIDView.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
            this.txtSalesIDView.Name = "txtSalesIDView";
            this.txtSalesIDView.OptionsSelection.EnableAppearanceFocusedCell = false;
            this.txtSalesIDView.OptionsView.ShowGroupPanel = false;
            this.label6.Location = new System.Drawing.Point(67, 202);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(39, 13);
            this.label6.Text = "Gudang";
            this.label5.Location = new System.Drawing.Point(80, 177);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(26, 13);
            this.label5.Text = "Team";
            this.label4.Location = new System.Drawing.Point(43, 151);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 13);
            this.label4.Text = "Tgl Transaksi";
            this.label3.Location = new System.Drawing.Point(28, 125);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 13);
            this.label3.Text = "Tipe Salesman";
            this.label2.Location = new System.Drawing.Point(24, 90);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 13);
            this.label2.Text = "Kode Salesman";
            this.label1.Location = new System.Drawing.Point(56, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 13);
            this.label1.Text = "Salesman";
            this.rdAllAktif.Location = new System.Drawing.Point(176, 62);
            this.rdAllAktif.Name = "rdAllAktif";
            this.rdAllAktif.Properties.Caption = "Semua (aktif dan Non)";
            this.rdAllAktif.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.Radio;
            this.rdAllAktif.Properties.RadioGroupIndex = 1;
            this.rdAllAktif.Size = new System.Drawing.Size(152, 20);
            this.rdAllAktif.TabIndex = 4;
            this.rdAllAktif.TabStop = false;
            this.rdAllAktif.CheckedChanged += new System.EventHandler(this.rdAllAktif_CheckedChanged);
            this.rdAktif.EditValue = true;
            this.rdAktif.Location = new System.Drawing.Point(117, 62);
            this.rdAktif.Name = "rdAktif";
            this.rdAktif.Properties.Caption = "Aktif";
            this.rdAktif.Properties.CheckStyle = DevExpress.XtraEditors.Controls.CheckStyles.Radio;
            this.rdAktif.Properties.RadioGroupIndex = 1;
            this.rdAktif.Size = new System.Drawing.Size(52, 20);
            this.rdAktif.TabIndex = 3;
            this.rdAktif.CheckedChanged += new System.EventHandler(this.rdAktif_CheckedChanged);
            // 
            // frmSOManualSalesman
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(508, 289);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.grbSalesman);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSOManualSalesman";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Load += new System.EventHandler(this.frmSOManualSalesman_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grbSalesman)).EndInit();
            this.grbSalesman.ResumeLayout(false);
            this.grbSalesman.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesDesc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesIDView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbBranchIDView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblWHLoc2.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblWHLoc1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblEmployee.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblOrderType.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblKdTipe.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbBranchID.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbEntityID.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtGudang.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTeam.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglTrans.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtTglTrans.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTipeSales.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtSalesID.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rdAllAktif.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rdAktif.Properties)).EndInit();
            this.ResumeLayout(false);
        }

        #endregion

        private DevExpress.XtraEditors.SimpleButton btnCancel;
        private DevExpress.XtraEditors.SimpleButton btnOk;
        private DevExpress.XtraEditors.GroupControl grbSalesman;
        public DevExpress.XtraEditors.TextEdit lblWHLoc2;
        public DevExpress.XtraEditors.TextEdit lblWHLoc1;
        public DevExpress.XtraEditors.TextEdit lblEmployee;
        public DevExpress.XtraEditors.TextEdit lblOrderType;
        public DevExpress.XtraEditors.TextEdit lblKdTipe;
        private DevExpress.XtraEditors.SearchLookUpEdit cbBranchID;
        private DevExpress.XtraEditors.LookUpEdit cbEntityID;
        private DevExpress.XtraEditors.LabelControl label25;
        private DevExpress.XtraEditors.LabelControl label23;
        private DevExpress.XtraEditors.TextEdit txtGudang;
        private DevExpress.XtraEditors.TextEdit txtTeam;
        private DevExpress.XtraEditors.DateEdit dtTglTrans;
        private DevExpress.XtraEditors.TextEdit txtTipeSales;
        private DevExpress.XtraEditors.SearchLookUpEdit txtSalesID;
        private DevExpress.XtraGrid.Views.Grid.GridView txtSalesIDView;
        private DevExpress.XtraGrid.Views.Grid.GridView cbBranchIDView;
        private DevExpress.XtraEditors.LabelControl label6;
        private DevExpress.XtraEditors.LabelControl label5;
        private DevExpress.XtraEditors.LabelControl label4;
        private DevExpress.XtraEditors.LabelControl label3;
        private DevExpress.XtraEditors.LabelControl label2;
        private DevExpress.XtraEditors.LabelControl label1;
        private DevExpress.XtraEditors.CheckEdit rdAllAktif;
        private DevExpress.XtraEditors.CheckEdit rdAktif;
        public DevExpress.XtraEditors.TextEdit txtSalesDesc;
    }
}
