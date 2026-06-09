namespace TIRASnDNet
{
    partial class BaseListForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();

            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BaseListForm));
            this.barManager1 = new DevExpress.XtraBars.BarManager(this.components);
            this.barTools = new DevExpress.XtraBars.Bar();
            this.barBtnNew = new DevExpress.XtraBars.BarButtonItem();
            this.barBtnEdit = new DevExpress.XtraBars.BarButtonItem();
            this.barBtnDelete = new DevExpress.XtraBars.BarButtonItem();
            this.barBtnPrint = new DevExpress.XtraBars.BarButtonItem();
            this.barBtnRefresh = new DevExpress.XtraBars.BarButtonItem();
            this.barBtnClose = new DevExpress.XtraBars.BarButtonItem();
            this.barEditLoading = new DevExpress.XtraBars.BarEditItem();
            this.repositoryItemMarqueeProgressBar1 = new DevExpress.XtraEditors.Repository.RepositoryItemMarqueeProgressBar();
            this.barStaticLoading = new DevExpress.XtraBars.BarStaticItem();
            this.barDockControlTop = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlBottom = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlLeft = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlRight = new DevExpress.XtraBars.BarDockControl();
            this.contentPanel = new DevExpress.XtraEditors.PanelControl();
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMarqueeProgressBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.contentPanel)).BeginInit();
            this.SuspendLayout();
            // 
            // barManager1
            // 
            this.barManager1.Bars.AddRange(new DevExpress.XtraBars.Bar[] {
            this.barTools});
            this.barManager1.DockControls.Add(this.barDockControlTop);
            this.barManager1.DockControls.Add(this.barDockControlBottom);
            this.barManager1.DockControls.Add(this.barDockControlLeft);
            this.barManager1.DockControls.Add(this.barDockControlRight);
            this.barManager1.Form = this;
            this.barManager1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.barBtnNew,
            this.barBtnEdit,
            this.barBtnDelete,
            this.barBtnPrint,
            this.barBtnRefresh,
            this.barBtnClose,
            this.barStaticLoading,
            this.barEditLoading});
            this.barManager1.MaxItemId = 8;
            this.barManager1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemMarqueeProgressBar1});
            // 
            // barTools
            // 
            this.barTools.BarName = "Tools";
            this.barTools.DockCol = 0;
            this.barTools.DockRow = 0;
            this.barTools.DockStyle = DevExpress.XtraBars.BarDockStyle.Top;
            this.barTools.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
            new DevExpress.XtraBars.LinkPersistInfo(this.barBtnNew),
            new DevExpress.XtraBars.LinkPersistInfo(this.barBtnEdit),
            new DevExpress.XtraBars.LinkPersistInfo(this.barBtnDelete),
            new DevExpress.XtraBars.LinkPersistInfo(this.barBtnPrint),
            new DevExpress.XtraBars.LinkPersistInfo(this.barBtnClose),
            new DevExpress.XtraBars.LinkPersistInfo(this.barEditLoading, true),
            new DevExpress.XtraBars.LinkPersistInfo(this.barStaticLoading)});
            this.barTools.OptionsBar.AllowQuickCustomization = false;
            this.barTools.OptionsBar.DrawDragBorder = false;
            this.barTools.OptionsBar.UseWholeRow = true;
            this.barTools.Text = "Tools";
            // 
            // barBtnNew
            // 
            this.barBtnNew.Caption = "&New";
            this.barBtnNew.Id = 0;
            this.barBtnNew.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barBtnNew.ImageOptions.Image")));
            this.barBtnNew.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barBtnNew.ImageOptions.LargeImage")));
            this.barBtnNew.Name = "barBtnNew";
            this.barBtnNew.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.barBtnNew.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barBtnNew_ItemClick);
            // 
            // barBtnEdit
            // 
            this.barBtnEdit.Caption = "&Edit";
            this.barBtnEdit.Id = 1;
            this.barBtnEdit.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barBtnEdit.ImageOptions.Image")));
            this.barBtnEdit.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barBtnEdit.ImageOptions.LargeImage")));
            this.barBtnEdit.Name = "barBtnEdit";
            this.barBtnEdit.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.barBtnEdit.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barBtnEdit_ItemClick);
            // 
            // barBtnDelete
            // 
            this.barBtnDelete.Caption = "&Delete";
            this.barBtnDelete.Id = 2;
            this.barBtnDelete.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barBtnDelete.ImageOptions.Image")));
            this.barBtnDelete.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barBtnDelete.ImageOptions.LargeImage")));
            this.barBtnDelete.Name = "barBtnDelete";
            this.barBtnDelete.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.barBtnDelete.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barBtnDelete_ItemClick);
            // 
            // barBtnPrint
            // 
            this.barBtnPrint.Caption = "&Print";
            this.barBtnPrint.Id = 3;
            this.barBtnPrint.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barBtnPrint.ImageOptions.Image")));
            this.barBtnPrint.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barBtnPrint.ImageOptions.LargeImage")));
            this.barBtnPrint.Name = "barBtnPrint";
            this.barBtnPrint.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.barBtnPrint.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barBtnPrint_ItemClick);
            // 
            // barBtnRefresh
            // 
            this.barBtnRefresh.Caption = "&Refresh";
            this.barBtnRefresh.Id = 4;
            this.barBtnRefresh.Name = "barBtnRefresh";
            this.barBtnRefresh.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.barBtnRefresh.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barBtnRefresh_ItemClick);
            // 
            // barBtnClose
            // 
            this.barBtnClose.Caption = "&Close";
            this.barBtnClose.Id = 5;
            this.barBtnClose.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barBtnClose.ImageOptions.Image")));
            this.barBtnClose.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barBtnClose.ImageOptions.LargeImage")));
            this.barBtnClose.Name = "barBtnClose";
            this.barBtnClose.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.barBtnClose.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barBtnClose_ItemClick);
            // 
            // barEditLoading
            // 
            this.barEditLoading.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right;
            this.barEditLoading.Caption = "LoadingBar";
            this.barEditLoading.Edit = this.repositoryItemMarqueeProgressBar1;
            this.barEditLoading.EditWidth = 160;
            this.barEditLoading.Id = 7;
            this.barEditLoading.Name = "barEditLoading";
            this.barEditLoading.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            // 
            // repositoryItemMarqueeProgressBar1
            // 
            this.repositoryItemMarqueeProgressBar1.Name = "repositoryItemMarqueeProgressBar1";
            // 
            // barStaticLoading
            // 
            this.barStaticLoading.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right;
            this.barStaticLoading.Caption = "Loading...";
            this.barStaticLoading.Id = 6;
            this.barStaticLoading.Name = "barStaticLoading";
            this.barStaticLoading.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            // 
            // barDockControlTop
            // 
            this.barDockControlTop.CausesValidation = false;
            this.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.barDockControlTop.Location = new System.Drawing.Point(0, 0);
            this.barDockControlTop.Manager = this.barManager1;
            this.barDockControlTop.Size = new System.Drawing.Size(1200, 24);
            // 
            // barDockControlBottom
            // 
            this.barDockControlBottom.CausesValidation = false;
            this.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.barDockControlBottom.Location = new System.Drawing.Point(0, 700);
            this.barDockControlBottom.Manager = this.barManager1;
            this.barDockControlBottom.Size = new System.Drawing.Size(1200, 0);
            // 
            // barDockControlLeft
            // 
            this.barDockControlLeft.CausesValidation = false;
            this.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.barDockControlLeft.Location = new System.Drawing.Point(0, 24);
            this.barDockControlLeft.Manager = this.barManager1;
            this.barDockControlLeft.Size = new System.Drawing.Size(0, 676);
            // 
            // barDockControlRight
            // 
            this.barDockControlRight.CausesValidation = false;
            this.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right;
            this.barDockControlRight.Location = new System.Drawing.Point(1200, 24);
            this.barDockControlRight.Manager = this.barManager1;
            this.barDockControlRight.Size = new System.Drawing.Size(0, 676);
            // 
            // contentPanel
            // 
            this.contentPanel.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.contentPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contentPanel.Location = new System.Drawing.Point(0, 24);
            this.contentPanel.Name = "contentPanel";
            this.contentPanel.Size = new System.Drawing.Size(1200, 676);
            this.contentPanel.TabIndex = 4;
            // 
            // BaseListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1200, 700);
            this.Controls.Add(this.contentPanel);
            this.Controls.Add(this.barDockControlLeft);
            this.Controls.Add(this.barDockControlRight);
            this.Controls.Add(this.barDockControlBottom);
            this.Controls.Add(this.barDockControlTop);
            this.Name = "BaseListForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Base List Form";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemMarqueeProgressBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.contentPanel)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        protected DevExpress.XtraBars.BarManager barManager1;
        protected DevExpress.XtraBars.Bar barTools;
        protected DevExpress.XtraBars.BarButtonItem barBtnNew;
        protected DevExpress.XtraBars.BarButtonItem barBtnEdit;
        protected DevExpress.XtraBars.BarButtonItem barBtnDelete;
        protected DevExpress.XtraBars.BarButtonItem barBtnPrint;
        protected DevExpress.XtraBars.BarButtonItem barBtnRefresh;
        protected DevExpress.XtraBars.BarButtonItem barBtnClose;
        protected DevExpress.XtraBars.BarStaticItem barStaticLoading;
        protected DevExpress.XtraBars.BarEditItem barEditLoading;
        protected DevExpress.XtraEditors.Repository.RepositoryItemMarqueeProgressBar repositoryItemMarqueeProgressBar1;
        protected DevExpress.XtraBars.BarDockControl barDockControlTop;
        protected DevExpress.XtraBars.BarDockControl barDockControlBottom;
        protected DevExpress.XtraBars.BarDockControl barDockControlLeft;
        protected DevExpress.XtraBars.BarDockControl barDockControlRight;
        protected DevExpress.XtraEditors.PanelControl contentPanel;
    }
}