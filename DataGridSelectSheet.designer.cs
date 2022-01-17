
namespace NSA.Revit.RibbonButton
{
    partial class DataGridSelectSheet
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
            this.viewSelectionDataGrid = new System.Windows.Forms.DataGridView();
            this.checkBox = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.okButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.displaySelectedLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.viewSelectionDataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // viewSelectionDataGrid
            // 
            this.viewSelectionDataGrid.AllowUserToAddRows = false;
            this.viewSelectionDataGrid.AllowUserToDeleteRows = false;
            this.viewSelectionDataGrid.AllowUserToResizeRows = false;
            this.viewSelectionDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.viewSelectionDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.viewSelectionDataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.checkBox,
            this.Column1,
            this.Column2});
            this.viewSelectionDataGrid.Location = new System.Drawing.Point(12, 44);
            this.viewSelectionDataGrid.Name = "viewSelectionDataGrid";
            this.viewSelectionDataGrid.ReadOnly = true;
            this.viewSelectionDataGrid.RowHeadersVisible = false;
            this.viewSelectionDataGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.viewSelectionDataGrid.Size = new System.Drawing.Size(750, 380);
            this.viewSelectionDataGrid.TabIndex = 0;
            this.viewSelectionDataGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.viewSelectionDataGrid_CellClick);
            this.viewSelectionDataGrid.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.viewSelectionDataGrid_CellContentClick);
            this.viewSelectionDataGrid.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.viewSelectionDataGrid_CellMouseClick);
            this.viewSelectionDataGrid.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.viewSelectionDataGrid_CellValueChanged);
            this.viewSelectionDataGrid.KeyDown += new System.Windows.Forms.KeyEventHandler(this.viewSelectionDataGrid_KeyDown);
            // 
            // checkBox
            // 
            this.checkBox.HeaderText = "";
            this.checkBox.Name = "checkBox";
            this.checkBox.ReadOnly = true;
            this.checkBox.Width = 30;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Width = 350;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Column2";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            this.Column2.Width = 350;
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okButton.Location = new System.Drawing.Point(606, 430);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 1;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // closeButton
            // 
            this.closeButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.closeButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.closeButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.closeButton.Location = new System.Drawing.Point(687, 430);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(75, 23);
            this.closeButton.TabIndex = 2;
            this.closeButton.Text = "CLOSE";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // searchTextBox
            // 
            this.searchTextBox.AccessibleDescription = "";
            this.searchTextBox.Location = new System.Drawing.Point(12, 12);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(750, 20);
            this.searchTextBox.TabIndex = 3;
            this.searchTextBox.TextChanged += new System.EventHandler(this.searchTextBox_TextChanged);
            this.searchTextBox.Enter += new System.EventHandler(this.searchTextBox_Enter);
            this.searchTextBox.Leave += new System.EventHandler(this.searchTextBox_Leave);
            // 
            // displaySelectedLabel
            // 
            this.displaySelectedLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.displaySelectedLabel.AutoSize = true;
            this.displaySelectedLabel.Location = new System.Drawing.Point(9, 430);
            this.displaySelectedLabel.Name = "displaySelectedLabel";
            this.displaySelectedLabel.Size = new System.Drawing.Size(103, 13);
            this.displaySelectedLabel.TabIndex = 4;
            this.displaySelectedLabel.Text = "0 Sheet(s) Selected.";
            // 
            // DataGridSelectSheet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.CancelButton = this.closeButton;
            this.ClientSize = new System.Drawing.Size(784, 461);
            this.Controls.Add(this.displaySelectedLabel);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.viewSelectionDataGrid);
            this.MinimumSize = new System.Drawing.Size(400, 250);
            this.Name = "DataGridSelectSheet";
            this.Padding = new System.Windows.Forms.Padding(30, 30, 30, 30);
            this.Text = "Select Sheets";
            ((System.ComponentModel.ISupportInitialize)(this.viewSelectionDataGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView viewSelectionDataGrid;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Label displaySelectedLabel;
        private System.Windows.Forms.DataGridViewCheckBoxColumn checkBox;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
    }
}