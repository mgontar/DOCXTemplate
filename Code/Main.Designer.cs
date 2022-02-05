namespace DOCXTemplate
{
    partial class Main
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
            this.components = new System.ComponentModel.Container();
            this.btnScanTemplates = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.clbTemplates = new System.Windows.Forms.CheckedListBox();
            this.btnToggleSelect = new System.Windows.Forms.Button();
            this.dgvVariables = new System.Windows.Forms.DataGridView();
            this.templateVariableBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.btnGenerate = new System.Windows.Forms.Button();
            this.ColumnName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvVariables)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.templateVariableBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // btnScanTemplates
            // 
            this.btnScanTemplates.Location = new System.Drawing.Point(13, 13);
            this.btnScanTemplates.Name = "btnScanTemplates";
            this.btnScanTemplates.Size = new System.Drawing.Size(114, 23);
            this.btnScanTemplates.TabIndex = 0;
            this.btnScanTemplates.Text = "Scan template folder";
            this.btnScanTemplates.UseVisualStyleBackColor = true;
            this.btnScanTemplates.Click += new System.EventHandler(this.btnScanTemplates_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel,
            this.toolStripProgressBar});
            this.statusStrip1.Location = new System.Drawing.Point(0, 428);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(777, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(39, 17);
            this.toolStripStatusLabel.Text = "Status";
            // 
            // toolStripProgressBar
            // 
            this.toolStripProgressBar.Name = "toolStripProgressBar";
            this.toolStripProgressBar.Size = new System.Drawing.Size(100, 16);
            // 
            // clbTemplates
            // 
            this.clbTemplates.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.clbTemplates.FormattingEnabled = true;
            this.clbTemplates.Location = new System.Drawing.Point(13, 42);
            this.clbTemplates.Name = "clbTemplates";
            this.clbTemplates.Size = new System.Drawing.Size(300, 379);
            this.clbTemplates.TabIndex = 2;
            this.clbTemplates.SelectedIndexChanged += new System.EventHandler(this.clbTemplates_SelectedIndexChanged);
            // 
            // btnToggleSelect
            // 
            this.btnToggleSelect.Location = new System.Drawing.Point(134, 13);
            this.btnToggleSelect.Name = "btnToggleSelect";
            this.btnToggleSelect.Size = new System.Drawing.Size(179, 23);
            this.btnToggleSelect.TabIndex = 3;
            this.btnToggleSelect.Text = "Toggle selection";
            this.btnToggleSelect.UseVisualStyleBackColor = true;
            this.btnToggleSelect.Click += new System.EventHandler(this.btnToggleSelect_Click);
            // 
            // dgvVariables
            // 
            this.dgvVariables.AllowUserToAddRows = false;
            this.dgvVariables.AllowUserToDeleteRows = false;
            this.dgvVariables.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvVariables.AutoGenerateColumns = false;
            this.dgvVariables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvVariables.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnName,
            this.ColumnValue});
            this.dgvVariables.DataSource = this.templateVariableBindingSource;
            this.dgvVariables.Location = new System.Drawing.Point(319, 42);
            this.dgvVariables.Name = "dgvVariables";
            this.dgvVariables.Size = new System.Drawing.Size(458, 379);
            this.dgvVariables.TabIndex = 4;
            // 
            // templateVariableBindingSource
            // 
            this.templateVariableBindingSource.DataSource = typeof(DOCXTemplate.TemplateVariable);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(320, 12);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(159, 23);
            this.btnGenerate.TabIndex = 5;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // ColumnName
            // 
            this.ColumnName.DataPropertyName = "Name";
            this.ColumnName.Frozen = true;
            this.ColumnName.HeaderText = "Name";
            this.ColumnName.Name = "ColumnName";
            this.ColumnName.ReadOnly = true;
            // 
            // ColumnValue
            // 
            this.ColumnValue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ColumnValue.DataPropertyName = "Value";
            this.ColumnValue.HeaderText = "Value";
            this.ColumnValue.Name = "ColumnValue";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(777, 450);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.dgvVariables);
            this.Controls.Add(this.btnToggleSelect);
            this.Controls.Add(this.clbTemplates);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.btnScanTemplates);
            this.Name = "Main";
            this.Text = "DOCX Template App";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvVariables)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.templateVariableBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnScanTemplates;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar;
        private System.Windows.Forms.CheckedListBox clbTemplates;
        private System.Windows.Forms.Button btnToggleSelect;
        private System.Windows.Forms.DataGridView dgvVariables;
        private System.Windows.Forms.BindingSource templateVariableBindingSource;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnValue;
    }
}

