namespace readText
{
    partial class Form1
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
            this.cboSheet = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.cboPath = new System.Windows.Forms.ComboBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.lstFoodCalculator = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // cboSheet
            // 
            this.cboSheet.FormattingEnabled = true;
            this.cboSheet.Location = new System.Drawing.Point(264, 29);
            this.cboSheet.Name = "cboSheet";
            this.cboSheet.Size = new System.Drawing.Size(121, 21);
            this.cboSheet.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(220, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Sheet:";
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(12, 32);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 3;
            this.btnImport.Text = "&Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // cboPath
            // 
            this.cboPath.FormattingEnabled = true;
            this.cboPath.Location = new System.Drawing.Point(93, 32);
            this.cboPath.Name = "cboPath";
            this.cboPath.Size = new System.Drawing.Size(121, 21);
            this.cboPath.TabIndex = 4;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(404, 27);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 5;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // lstFoodCalculator
            // 
            this.lstFoodCalculator.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstFoodCalculator.FullRowSelect = true;
            this.lstFoodCalculator.GridLines = true;
            this.lstFoodCalculator.Location = new System.Drawing.Point(12, 62);
            this.lstFoodCalculator.Name = "lstFoodCalculator";
            this.lstFoodCalculator.Size = new System.Drawing.Size(572, 389);
            this.lstFoodCalculator.TabIndex = 6;
            this.lstFoodCalculator.UseCompatibleStateImageBehavior = false;
            this.lstFoodCalculator.View = System.Windows.Forms.View.Details;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(596, 463);
            this.Controls.Add(this.lstFoodCalculator);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.cboPath);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboSheet);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox cboSheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.ComboBox cboPath;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ListView lstFoodCalculator;
    }
}

