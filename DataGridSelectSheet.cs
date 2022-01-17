using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.Revit.DB;

namespace NSA.Revit.RibbonButton
{
    public partial class DataGridSelectSheet : System.Windows.Forms.Form
    {
        public List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();

        public List<ViewSheet> tempSheets = new List<ViewSheet>();
        public List<ViewSheet> selectedSheets = new List<ViewSheet>();

        public int checkedCount = 0;

        public DataGridSelectSheet(List<ViewSheet> viewSheetList)
        {
            InitializeComponent();
            formLoad(viewSheetList);
        }

        private void formLoad(List<ViewSheet> viewSheetList)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = viewSheetList;
            viewSelectionDataGrid.AutoGenerateColumns = false;
            viewSelectionDataGrid.DataSource = bs;
            searchTextBox.Text = "search sheets...";

            viewSelectionDataGrid.Columns[1].DataPropertyName = "SheetNumber";
            viewSelectionDataGrid.Columns[1].HeaderText = "Sheet Number";

            viewSelectionDataGrid.Columns[2].DataPropertyName = "Name";
            viewSelectionDataGrid.Columns[2].HeaderText = "Sheet Name";

            searchTextBox.Text = "search sheets...";
        }

        private void closeButton_Click(object sender, EventArgs e)
        {

        }

        private void okButton_Click(object sender, EventArgs e)
        {
            //viewSelectionDataGrid.AllowUserToAddRows = false;

            foreach (DataGridViewRow row in viewSelectionDataGrid.SelectedRows)
            {
                selectedSheets.Add((ViewSheet)row.DataBoundItem);
            }

            this.DialogResult = DialogResult.OK;
            Close();

            return;
        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[viewSelectionDataGrid.DataSource];
            currencyManager1.SuspendBinding();

            if (searchTextBox.Text != string.Empty && searchTextBox.Text != "search sheets...")
            {
                foreach (DataGridViewRow row in viewSelectionDataGrid.Rows)
                {
                    if (row.Cells[1].Value.ToString().Trim().Contains(searchTextBox.Text.Trim()) ||
                        row.Cells[2].Value.ToString().Trim().Contains(searchTextBox.Text.Trim()))
                    {
                        row.Visible = true;
                    }
                    else
                        row.Visible = false;
                }
            }
            else if (searchTextBox.Text == string.Empty)
            {
                foreach (DataGridViewRow row in viewSelectionDataGrid.Rows)
                {
                    row.Visible = true;
                }
            }
                
            currencyManager1.ResumeBinding();
        }

        private void viewSelectionDataGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            checkedCount = 0;
            foreach (DataGridViewRow row in this.viewSelectionDataGrid.SelectedRows)
            {
                checkedCount++;
            }
            displaySelectedLabel.Text = checkedCount + " Sheet(s) Selected.";
        }

        private void viewSelectionDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                selectedRows.Clear();
                foreach (DataGridViewRow row in this.viewSelectionDataGrid.Rows)
                {
                    row.Cells["checkBox"].Value = false;   
                }

                foreach (DataGridViewRow row in this.viewSelectionDataGrid.SelectedRows)
                {
                    row.Cells["checkBox"].Value = true;
                    selectedRows.Add(row);
                }
            }
        }

        private void searchTextBox_Enter(object sender, EventArgs e)
        {
            searchTextBox.Text = string.Empty;
        }

        private void searchTextBox_Leave(object sender, EventArgs e)
        {
            searchTextBox.Text = "search sheets...";
        }


        private void viewSelectionDataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Modifiers == Keys.Shift)
            {
                foreach (DataGridViewRow row in viewSelectionDataGrid.SelectedRows)
                {
                    if (!row.Visible)
                    {
                        row.Selected = false;
                        row.Cells["checkBox"].Value = false;
                    }
                }
            }
        }

        private void viewSelectionDataGrid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
        }

        private void viewSelectionDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}