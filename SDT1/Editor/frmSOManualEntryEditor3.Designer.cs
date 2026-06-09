namespace TIRASnDNet.PROCESS.SO.SOManualEntry
{
    partial class frmSOManualEntryEditor2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSOManualEntryEditor2));
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label16 = new System.Windows.Forms.Label();
            this.dgvPopUp = new System.Windows.Forms.DataGridView();
            this.g_iPCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.g_iTPRCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.g_iQtyOnSO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.g_iQtyOnBudget = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPopUp)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label16);
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(486, 47);
            this.groupBox2.TabIndex = 129;
            this.groupBox2.TabStop = false;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(9, 19);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(289, 14);
            this.label16.TabIndex = 122;
            this.label16.Text = "Terdapat PCode yang tidak mencukupi TPR Budget";
            // 
            // dgvPopUp
            // 
            this.dgvPopUp.AllowUserToAddRows = false;
            this.dgvPopUp.AllowUserToDeleteRows = false;
            this.dgvPopUp.AllowUserToResizeRows = false;
            this.dgvPopUp.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvPopUp.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPopUp.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.g_iPCode,
            this.g_iTPRCode,
            this.g_iQtyOnSO,
            this.g_iQtyOnBudget});
            this.dgvPopUp.Location = new System.Drawing.Point(12, 65);
            this.dgvPopUp.MultiSelect = false;
            this.dgvPopUp.Name = "dgvPopUp";
            this.dgvPopUp.ReadOnly = true;
            this.dgvPopUp.RowHeadersVisible = false;
            this.dgvPopUp.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvPopUp.Size = new System.Drawing.Size(484, 219);
            this.dgvPopUp.TabIndex = 130;
            // 
            // g_iPCode
            // 
            this.g_iPCode.HeaderText = "PCode";
            this.g_iPCode.Name = "g_iPCode";
            this.g_iPCode.ReadOnly = true;
            this.g_iPCode.Width = 120;
            // 
            // g_iTPRCode
            // 
            this.g_iTPRCode.HeaderText = "TPR Code";
            this.g_iTPRCode.Name = "g_iTPRCode";
            this.g_iTPRCode.ReadOnly = true;
            this.g_iTPRCode.Width = 120;
            // 
            // g_iQtyOnSO
            // 
            this.g_iQtyOnSO.HeaderText = "Qty On SO";
            this.g_iQtyOnSO.Name = "g_iQtyOnSO";
            this.g_iQtyOnSO.ReadOnly = true;
            // 
            // g_iQtyOnBudget
            // 
            this.g_iQtyOnBudget.HeaderText = "Qty On TPR Budget";
            this.g_iQtyOnBudget.Name = "g_iQtyOnBudget";
            this.g_iQtyOnBudget.ReadOnly = true;
            this.g_iQtyOnBudget.Width = 140;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(121, 305);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(189, 14);
            this.label1.TabIndex = 131;
            this.label1.Text = "apakah proses akan di lanjutkan ?";
            // 
            // btnCancel
            // 
            this.btnCancel.Image = global::TIRASnDNet.Properties.Resources._1490217733_3;
            this.btnCancel.Location = new System.Drawing.Point(409, 301);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(87, 23);
            this.btnCancel.TabIndex = 181;
            this.btnCancel.Text = "&NO";
            this.btnCancel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOk
            // 
            this.btnOk.Image = global::TIRASnDNet.Properties.Resources._1490217776_4;
            this.btnOk.Location = new System.Drawing.Point(316, 301);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(87, 23);
            this.btnOk.TabIndex = 180;
            this.btnOk.Text = "&YES";
            this.btnOk.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // frmSOManualEntryEditor2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(503, 341);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgvPopUp);
            this.Controls.Add(this.groupBox2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmSOManualEntryEditor2";
            this.Text = "Editor 2";
            this.Load += new System.EventHandler(this.frmSOManualEntryEditor2_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPopUp)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.DataGridView dgvPopUp;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.DataGridViewTextBoxColumn g_iPCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn g_iTPRCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn g_iQtyOnSO;
        private System.Windows.Forms.DataGridViewTextBoxColumn g_iQtyOnBudget;
    }
}