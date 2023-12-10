using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace BaseData1
{
    internal class QueryManager
    {
        BaseData BaseData = new BaseData();

        public SqlDataAdapter sqlDataAdapter = null;
        public SqlCommandBuilder sqlCommandBuilder = null;
        public DataSet dataSet = null;
        public bool newRowAdding = false;
        public void queryEx(string query, DataGridView dgv)
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter(query, BaseData.GetConnection());
                sqlCommandBuilder = new SqlCommandBuilder(sqlDataAdapter);
                dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dgv.DataSource = dataSet.Tables[0];
                int ind = query.IndexOf("join");
                if (ind == -1)
                {
                    sqlCommandBuilder.GetInsertCommand();
                    sqlCommandBuilder.GetUpdateCommand();
                    sqlCommandBuilder.GetDeleteCommand();


                    for (int i = 0; i < dgv.Rows.Count; i++)
                    {
                        DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                        dgv[dgv.ColumnCount - 1, i] = linkCell;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void queryExFRep(string query, DataGridView dgv)
        {
            try
            {
                sqlDataAdapter = new SqlDataAdapter(query, BaseData.GetConnection());
                sqlCommandBuilder = new SqlCommandBuilder(sqlDataAdapter);
                dataSet = new DataSet();
                sqlDataAdapter.Fill(dataSet);
                dgv.DataSource = dataSet.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void DeleteRow(DataGridView dgv, int rowIndex)
        {
            try
            {
                dgv.Rows.RemoveAt(rowIndex);
                DataTable table = (DataTable)dgv.DataSource;
                dataSet.Tables.Clear();
                dataSet.Tables.Add(table);
                sqlDataAdapter.Update(dataSet, table.TableName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void InsertRow(DataGridView dgv)
        {
            try
            {
                if (newRowAdding == false)
                {
                    newRowAdding = true;

                    int lastRow = dgv.Rows.Count - 2;

                    DataGridViewRow row = dgv.Rows[lastRow];
                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dgv[dgv.ColumnCount - 1, lastRow] = linkCell;
                    row.Cells["Command"].Value = "Insert";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void UpdateRow(DataGridView dgv)
        {
            try
            {
                if (newRowAdding == false)
                {
                    int rowIndex = dgv.SelectedCells[0].RowIndex;

                    DataGridViewRow editingRow = dgv.Rows[rowIndex];

                    DataGridViewLinkCell linkCell = new DataGridViewLinkCell();
                    dgv[dgv.ColumnCount - 1, rowIndex] = linkCell;
                    editingRow.Cells["Command"].Value = "Update";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CreateRow(DataGridView dgv, int ErowIndex)
        {
            int rowIndex = dgv.Rows.Count - 2;

            DataRow row = dataSet.Tables[0].NewRow();

            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                row[dgv.Columns[i].HeaderText.ToString()] = dgv.Rows[rowIndex].Cells[dgv.Columns[i].HeaderText.ToString()].Value;
            }

            dataSet.Tables[0].Rows.Add(row);
            dataSet.Tables[0].Rows.RemoveAt(dataSet.Tables[0].Rows.Count - 1);
            dgv.Rows.RemoveAt(dgv.Rows.Count - 2);
            dgv.Rows[ErowIndex].Cells[dgv.ColumnCount - 1].Value = "Delete";
            sqlDataAdapter.Update(dataSet);
            newRowAdding = false;
        }

        public void CreateUPDRow(DataGridView dgv, int rowIndex, string table_name)
        {
            int r = rowIndex;

            DataTable table = (DataTable)dgv.DataSource;
            dataSet.Tables.Clear();
            dataSet.Tables.Add(table);
            sqlDataAdapter.Update(dataSet, table.TableName);

            dgv.Rows[r].Cells[dgv.ColumnCount - 1].Value = "Delete";
        }
        public void ReloadData(DataGridView dgv, string table_name)
        {
            dataSet.Tables.Clear();
            sqlDataAdapter.Fill(dataSet, table_name);
            dgv.DataSource = dataSet.Tables[table_name];

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                DataGridViewLinkCell linkCell = new DataGridViewLinkCell();

                dgv[dgv.ColumnCount - 1, i] = linkCell;
            }
        }
        public void ExportQueryToExc(DataGridView dgv)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;

            int i, j;
            for (i = 0; i <= dgv.RowCount - 2; i++)
            {
                for (j = 0; j <= dgv.ColumnCount - 1; j++)
                {
                    wsh.Cells[1, j + 1] = dgv.Columns[j].HeaderText.ToString();
                    wsh.Cells[i + 2, j + 1] = dgv[j, i].Value.ToString();
                }
            }
            exApp.Visible = true;
        }
    }
}
