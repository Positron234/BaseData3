using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace BaseData1
{
    public partial class Budget : Form
    {
        QueryManager QR = new QueryManager();
        string query;
        string tab;
        public Budget(string que,string table)
        {
            query = que;
            tab = table;
            InitializeComponent();
        }
        private void Budget_click_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int i, j;
            for (i = 0; i <= dataGridView1.RowCount - 2; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    wsh.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                    wsh.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                }
            }
            exApp.Visible = true;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
                try
                {
                    if (e.ColumnIndex == dataGridView1.ColumnCount - 1)
                    {
                        string task = dataGridView1.Rows[e.RowIndex].Cells[dataGridView1.ColumnCount - 1].Value.ToString();
                        if (task == "Insert")
                        {
                            int ErowIndex = e.RowIndex;
                            QR.CreateRow(dataGridView1, ErowIndex);
                        }
                        else if (task == "DELETE")
                        {
                            if (MessageBox.Show("Удалить эту строку?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                int rowIndex = e.RowIndex;

                                QR.DeleteRow(dataGridView1, rowIndex);
                            }
                        }
                        else if (task == "Update")
                        {
                            int r = e.RowIndex;
                            QR.CreateUPDRow(dataGridView1, r, tab);
                        }

                        QR.ReloadData(dataGridView1, tab);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            
        }

        private void Budget_Load(object sender, EventArgs e)
        {
            QR.queryEx(query, dataGridView1);
        }

       

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            QR.UpdateRow(dataGridView1);
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            QR.InsertRow(dataGridView1);
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);
        }
        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }
    }
}
