using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using System;

namespace TIRASnDNet
{
    public partial class BaseListForm : XtraForm
    {
        public BaseListForm()
        {
            InitializeComponent();
            InitializeLoadingBar();
        }

        #region Virtual Toolbar Actions

        protected virtual void OnNew()
        {
        }

        protected virtual void OnEdit()
        {
        }

        protected virtual void OnDelete()
        {
        }

        protected virtual void OnPrint()
        {
        }

        protected virtual void OnRefreshData()
        {
        }

        protected virtual void OnCloseForm()
        {
            this.Close();
        }

        #endregion

        #region Access Button

        protected virtual void AccessButton()
        {
            barBtnNew.Enabled = clsLogin.BTNNEW;
            barBtnEdit.Enabled = clsLogin.BTNEDIT;
            barBtnDelete.Enabled = clsLogin.BTNDELETE;
            barBtnPrint.Enabled = clsLogin.BTNPRINT;
            barBtnRefresh.Enabled = true;
            barBtnClose.Enabled = true;
        }

        #endregion

        #region Toolbar Events

        private void barBtnNew_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (!barBtnNew.Enabled) return;
            OnNew();
        }

        private void barBtnEdit_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (!barBtnEdit.Enabled) return;
            OnEdit();
        }

        private void barBtnDelete_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (!barBtnDelete.Enabled) return;
            OnDelete();
        }

        private void barBtnPrint_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (!barBtnPrint.Enabled) return;
            OnPrint();
        }

        private void barBtnRefresh_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (!barBtnRefresh.Enabled) return;
            OnRefreshData();
        }

        private void barBtnClose_ItemClick(object sender, ItemClickEventArgs e)
        {
            OnCloseForm();
        }

        #endregion

        #region Loading Bar

        private void InitializeLoadingBar()
        {
            repositoryItemMarqueeProgressBar1.MarqueeAnimationSpeed = 30;
            repositoryItemMarqueeProgressBar1.ProgressViewStyle = DevExpress.XtraEditors.Controls.ProgressViewStyle.Solid;
            repositoryItemMarqueeProgressBar1.ShowTitle = false;

            barStaticLoading.Caption = "Ready";
            barStaticLoading.Visibility = BarItemVisibility.Never;

            barEditLoading.EditValue = 0;
            barEditLoading.Visibility = BarItemVisibility.Never;
            barEditLoading.Width = 160;
        }

        public void ShowLoading(string message = "Loading...")
        {
            if (this.IsDisposed) return;

            if (InvokeRequired)
            {
                this.Invoke(new Action(() => ShowLoading(message)));
                return;
            }

            barStaticLoading.Caption = message;
            barStaticLoading.Visibility = BarItemVisibility.Always;
            barEditLoading.Visibility = BarItemVisibility.Always;
            this.Refresh();
        }

        public void HideLoading()
        {
            if (this.IsDisposed) return;

            if (InvokeRequired)
            {
                this.Invoke(new Action(HideLoading));
                return;
            }

            barStaticLoading.Visibility = BarItemVisibility.Never;
            barEditLoading.Visibility = BarItemVisibility.Never;
            this.Refresh();
        }

        public void SetLoadingMessage(string message)
        {
            if (this.IsDisposed) return;

            if (InvokeRequired)
            {
                this.Invoke(new Action(() => SetLoadingMessage(message)));
                return;
            }

            barStaticLoading.Caption = message;
        }

        #endregion

        #region Helpers

        protected virtual void InitBaseForm()
        {
            AccessButton();
            HideLoading();
        }

        public PanelControl GetContentPanel()
        {
            return contentPanel;
        }

        #endregion

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            InitBaseForm();
        }
    }
}