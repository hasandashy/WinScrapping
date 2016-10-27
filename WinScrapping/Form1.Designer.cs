namespace WinScrapping
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
            this.btnExecute = new System.Windows.Forms.Button();
            this.btnGetCategory = new System.Windows.Forms.Button();
            this.btnProductLink = new System.Windows.Forms.Button();
            this.btnProductData = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.cbxWebsite = new System.Windows.Forms.ComboBox();
            this.btnImageDownload = new System.Windows.Forms.Button();
            this.cbxCategory = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExecute
            // 
            this.btnExecute.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExecute.Location = new System.Drawing.Point(89, 274);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(217, 49);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Read and Save All data";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnGetCategory
            // 
            this.btnGetCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGetCategory.Location = new System.Drawing.Point(89, 125);
            this.btnGetCategory.Name = "btnGetCategory";
            this.btnGetCategory.Size = new System.Drawing.Size(217, 39);
            this.btnGetCategory.TabIndex = 2;
            this.btnGetCategory.Text = "Save Elk Brand Category";
            this.btnGetCategory.UseVisualStyleBackColor = true;
            this.btnGetCategory.Click += new System.EventHandler(this.btnGetCategory_Click);
            // 
            // btnProductLink
            // 
            this.btnProductLink.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProductLink.Location = new System.Drawing.Point(89, 170);
            this.btnProductLink.Name = "btnProductLink";
            this.btnProductLink.Size = new System.Drawing.Size(217, 45);
            this.btnProductLink.TabIndex = 2;
            this.btnProductLink.Text = "Save All Product Link";
            this.btnProductLink.UseVisualStyleBackColor = true;
            this.btnProductLink.Click += new System.EventHandler(this.btnProductLink_Click);
            // 
            // btnProductData
            // 
            this.btnProductData.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProductData.Location = new System.Drawing.Point(89, 221);
            this.btnProductData.Name = "btnProductData";
            this.btnProductData.Size = new System.Drawing.Size(217, 47);
            this.btnProductData.TabIndex = 2;
            this.btnProductData.Text = "Save All Product Data";
            this.btnProductData.UseVisualStyleBackColor = true;
            this.btnProductData.Click += new System.EventHandler(this.btnProductData_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Salmon;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.cbxWebsite);
            this.panel1.Controls.Add(this.btnImageDownload);
            this.panel1.Controls.Add(this.cbxCategory);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btnGetCategory);
            this.panel1.Controls.Add(this.btnExecute);
            this.panel1.Controls.Add(this.btnProductData);
            this.panel1.Controls.Add(this.btnProductLink);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(413, 392);
            this.panel1.TabIndex = 3;
            // 
            // cbxWebsite
            // 
            this.cbxWebsite.FormattingEnabled = true;
            this.cbxWebsite.Location = new System.Drawing.Point(11, 55);
            this.cbxWebsite.Name = "cbxWebsite";
            this.cbxWebsite.Size = new System.Drawing.Size(385, 21);
            this.cbxWebsite.TabIndex = 7;
            this.cbxWebsite.SelectedIndexChanged += new System.EventHandler(this.cbxWebsite_SelectedIndexChanged);
            // 
            // btnImageDownload
            // 
            this.btnImageDownload.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImageDownload.Location = new System.Drawing.Point(89, 329);
            this.btnImageDownload.Name = "btnImageDownload";
            this.btnImageDownload.Size = new System.Drawing.Size(217, 49);
            this.btnImageDownload.TabIndex = 6;
            this.btnImageDownload.Text = "Read and Save All data";
            this.btnImageDownload.UseVisualStyleBackColor = true;
            this.btnImageDownload.Click += new System.EventHandler(this.btnImageDownload_Click);
            // 
            // cbxCategory
            // 
            this.cbxCategory.FormattingEnabled = true;
            this.cbxCategory.Location = new System.Drawing.Point(11, 98);
            this.cbxCategory.Name = "cbxCategory";
            this.cbxCategory.Size = new System.Drawing.Size(385, 21);
            this.cbxCategory.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(4, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(316, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "All the downloaded data will be saved in Ms-Excel format";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(207, 16);
            this.label1.TabIndex = 3;
            this.label1.Text = "Lighting NewYork Data Read";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(437, 451);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.Button btnGetCategory;
        private System.Windows.Forms.Button btnProductLink;
        private System.Windows.Forms.Button btnProductData;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbxCategory;
        private System.Windows.Forms.Button btnImageDownload;
        private System.Windows.Forms.ComboBox cbxWebsite;
    }
}

