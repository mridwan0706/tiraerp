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
            InitializeLoadingBars();
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

        private void barBtnClose_ItemClick(object sender, ItemClickEventArgs e)
        {
            OnCloseForm();
        }

        #endregion

        #region Loading

        private void InitializeLoadingBars()
        {
            repositoryItemMarqueeProgressBar1.MarqueeAnimationSpeed = 30;
            repositoryItemMarqueeProgressBar1.ProgressViewStyle = DevExpress.XtraEditors.Controls.ProgressViewStyle.Solid;
            repositoryItemMarqueeProgressBar1.ShowTitle = false;

            repositoryItemProgressBar1.Minimum = 0;
            repositoryItemProgressBar1.Maximum = 100;
            repositoryItemProgressBar1.Step = 1;
            repositoryItemProgressBar1.PercentView = true;
            repositoryItemProgressBar1.ShowTitle = true;

            barStaticLoading.Caption = "Ready";
            barStaticLoading.Visibility = BarItemVisibility.Never;

            barEditLoadingMarquee.Visibility = BarItemVisibility.Never;
            barEditLoadingPercent.Visibility = BarItemVisibility.Never;

            barEditLoadingPercent.EditValue = 0;
        }

        public void ShowLoadingMarquee(string message = "Loading...")
        {
            if (this.IsDisposed) return;

            if (InvokeRequired)
            {
                this.Invoke(new Action(() => ShowLoadingMarquee(message)));
                return;
            }

            barManager1.BeginUpdate();
            try
            {
                barStaticLoading.Caption = message;
                barStaticLoading.Visibility = BarItemVisibility.Always;

                barEditLoadingPercent.Visibility = BarItemVisibility.Never;
                barEditLoadingMarquee.Visibility = BarItemVisibility.Always;
            }
            finally
            {
                barManager1.EndUpdate();
            }

            this.Refresh();
        }

        public void ShowProgress(string message = "Processing...", int percent = 0)
        {
            if (this.IsDisposed) return;

            if (InvokeRequired)
            {
                this.Invoke(new Action(() => ShowProgress(message, percent)));
                return;
            }

            if (percent < 0) percent = 0;
            if (percent > 100) percent = 100;

            barManager1.BeginUpdate();
            try
            {
                barStaticLoading.Caption = message;
                barStaticLoading.Visibility = BarItemVisibility.Always;

                barEditLoadingMarquee.Visibility = BarItemVisibility.Never;
                barEditLoadingPercent.Visibility = BarItemVisibility.Always;
                barEditLoadingPercent.EditValue = percent;
            }
            finally
            {
                barManager1.EndUpdate();
            }

            this.Refresh();
        }

        public void SetProgressValue(int percent)
        {
            if (this.IsDisposed) return;

            if (InvokeRequired)
            {
                this.Invoke(new Action(() => SetProgressValue(percent)));
                return;
            }

            if (percent < 0) percent = 0;
            if (percent > 100) percent = 100;

            barEditLoadingPercent.EditValue = percent;
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

            barManager1.BeginUpdate();
            try
            {
                barStaticLoading.Visibility = BarItemVisibility.Never;
                barEditLoadingMarquee.Visibility = BarItemVisibility.Never;
                barEditLoadingPercent.Visibility = BarItemVisibility.Never;
                barEditLoadingPercent.EditValue = 0;
            }
            finally
            {
                barManager1.EndUpdate();
            }

            this.Refresh();
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

        private void DisableActionButton()
        {
            SetActionButton(false, false, false, false);
        }

        protected void SetActionButton(bool allowNew, bool allowEdit, bool allowDelete, bool allowPrint)
        {
            barBtnNew.Enabled = allowNew;
            barBtnEdit.Enabled = allowEdit;
            barBtnDelete.Enabled = allowDelete;
            barBtnPrint.Enabled = allowPrint;
            barBtnClose.Enabled = true;
        }
    }
}