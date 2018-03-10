namespace Siebel_custom_report
{
    partial class Export
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
            this.loadInfo = new System.Windows.Forms.Label();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnOpenExport = new System.Windows.Forms.Button();
            this.openExportDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveExportDialog = new System.Windows.Forms.SaveFileDialog();
            this.chkDatumi = new System.Windows.Forms.CheckBox();
            this.lblVersion = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // loadInfo
            // 
            this.loadInfo.BackColor = System.Drawing.SystemColors.Control;
            this.loadInfo.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.loadInfo.Font = new System.Drawing.Font("Verdana", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.loadInfo.ForeColor = System.Drawing.Color.Red;
            this.loadInfo.Location = new System.Drawing.Point(12, 57);
            this.loadInfo.Name = "loadInfo";
            this.loadInfo.Size = new System.Drawing.Size(370, 32);
            this.loadInfo.TabIndex = 0;
            this.loadInfo.Text = "Nema podataka";
            this.loadInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnExport
            // 
            this.btnExport.Enabled = false;
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExport.Font = new System.Drawing.Font("Verdana", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnExport.Location = new System.Drawing.Point(12, 94);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(370, 95);
            this.btnExport.TabIndex = 1;
            this.btnExport.Text = "Napravi excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnReport_Click);
            // 
            // btnOpenExport
            // 
            this.btnOpenExport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnOpenExport.Font = new System.Drawing.Font("Verdana", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnOpenExport.Location = new System.Drawing.Point(12, 13);
            this.btnOpenExport.Name = "btnOpenExport";
            this.btnOpenExport.Size = new System.Drawing.Size(370, 41);
            this.btnOpenExport.TabIndex = 0;
            this.btnOpenExport.Text = "Učitaj export";
            this.btnOpenExport.UseVisualStyleBackColor = true;
            this.btnOpenExport.Click += new System.EventHandler(this.openExport_Click);
            // 
            // openExportDialog
            // 
            this.openExportDialog.DefaultExt = "xlsx";
            this.openExportDialog.Filter = "Excel files | *.xls; *.xlsx";
            // 
            // saveExportDialog
            // 
            this.saveExportDialog.DefaultExt = "xlsx";
            this.saveExportDialog.Filter = "Excel files | *.xlsx";
            // 
            // chkDatumi
            // 
            this.chkDatumi.AutoSize = true;
            this.chkDatumi.Enabled = false;
            this.chkDatumi.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.chkDatumi.Location = new System.Drawing.Point(20, 66);
            this.chkDatumi.Name = "chkDatumi";
            this.chkDatumi.Size = new System.Drawing.Size(72, 20);
            this.chkDatumi.TabIndex = 2;
            this.chkDatumi.Text = "Datumi";
            this.chkDatumi.UseVisualStyleBackColor = true;
            this.chkDatumi.CheckedChanged += new System.EventHandler(this.chkDatumi_CheckedChanged);
            // 
            // lblVersion
            // 
            this.lblVersion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblVersion.BackColor = System.Drawing.SystemColors.Control;
            this.lblVersion.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblVersion.Font = new System.Drawing.Font("Verdana", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.lblVersion.Location = new System.Drawing.Point(12, 192);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(370, 10);
            this.lblVersion.TabIndex = 3;
            this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Export
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(394, 205);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.chkDatumi);
            this.Controls.Add(this.btnOpenExport);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.loadInfo);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Export";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Stefex - prilagođen siebel export";
            this.Load += new System.EventHandler(this.Export_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label loadInfo;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnOpenExport;
        private System.Windows.Forms.OpenFileDialog openExportDialog;
        private System.Windows.Forms.SaveFileDialog saveExportDialog;
        private System.Windows.Forms.CheckBox chkDatumi;
        private System.Windows.Forms.Label lblVersion;

    }
}

