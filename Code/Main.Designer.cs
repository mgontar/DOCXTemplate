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
            this.clbTemplates = new System.Windows.Forms.CheckedListBox();
            this.btnToggleSelect = new System.Windows.Forms.Button();
            this.dgvVariables = new System.Windows.Forms.DataGridView();
            this.ColumnName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.templateVariableBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.btnGenerate = new System.Windows.Forms.Button();
            this.btnTemplates = new System.Windows.Forms.Button();
            this.btnOutput = new System.Windows.Forms.Button();
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
            // clbTemplates
            // 
            this.clbTemplates.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.clbTemplates.FormattingEnabled = true;
            this.clbTemplates.Location = new System.Drawing.Point(13, 42);
            this.clbTemplates.Name = "clbTemplates";
            this.clbTemplates.Size = new System.Drawing.Size(300, 394);
            this.clbTemplates.TabIndex = 2;
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
            this.dgvVariables.Size = new System.Drawing.Size(458, 394);
            this.dgvVariables.TabIndex = 4;
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
            // templateVariableBindingSource
            // 
            this.templateVariableBindingSource.DataSource = typeof(DOCXTemplate.TemplateVariable);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(319, 13);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(101, 23);
            this.btnGenerate.TabIndex = 5;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // btnTemplates
            // 
            this.btnTemplates.Location = new System.Drawing.Point(426, 13);
            this.btnTemplates.Name = "btnTemplates";
            this.btnTemplates.Size = new System.Drawing.Size(100, 23);
            this.btnTemplates.TabIndex = 6;
            this.btnTemplates.Text = "Templates";
            this.btnTemplates.UseVisualStyleBackColor = true;
            this.btnTemplates.Click += new System.EventHandler(this.btnTemplates_Click);
            // 
            // btnOutput
            // 
            this.btnOutput.Location = new System.Drawing.Point(532, 13);
            this.btnOutput.Name = "btnOutput";
            this.btnOutput.Size = new System.Drawing.Size(109, 23);
            this.btnOutput.TabIndex = 7;
            this.btnOutput.Text = "Output";
            this.btnOutput.UseVisualStyleBackColor = true;
            this.btnOutput.Click += new System.EventHandler(this.btnOutput_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(777, 450);
            this.Controls.Add(this.btnOutput);
            this.Controls.Add(this.btnTemplates);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.dgvVariables);
            this.Controls.Add(this.btnToggleSelect);
            this.Controls.Add(this.clbTemplates);
            this.Controls.Add(this.btnScanTemplates);
            this.Name = "Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DOCX Template App";
            ((System.ComponentModel.ISupportInitialize)(this.dgvVariables)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.templateVariableBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnScanTemplates;
        private System.Windows.Forms.CheckedListBox clbTemplates;
        private System.Windows.Forms.Button btnToggleSelect;
        private System.Windows.Forms.DataGridView dgvVariables;
        private System.Windows.Forms.BindingSource templateVariableBindingSource;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnValue;
        private System.Windows.Forms.Button btnTemplates;
        private System.Windows.Forms.Button btnOutput;
    }
}

