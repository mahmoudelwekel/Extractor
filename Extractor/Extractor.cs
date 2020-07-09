using ClosedXML.Excel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Extractor
{
    public partial class Extractor : Form
    {
        public Extractor()
        {
            InitializeComponent();
        }

        private void Attach_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.Filter = "Images Files|*.MDF";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    DatabaseTextBox.Text = openFileDialog1.FileName;
                }
            }
            catch (Exception errortext)
            {
                MessageBox.Show(errortext.ToString());
            }
        }

        private void DatabaseTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string constr = @"Data Source=(LocalDB)\.;AttachDbFilename=" + DatabaseTextBox.Text + ";Integrated Security=True;Packet Size=32767";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM INFORMATION_SCHEMA.TABLES", con))
                    {
                        //Fill the DataTable with records from Table.
                        DataTable dt = new DataTable();
                        sda.Fill(dt);

                        //Insert the Default Item to DataTable.
                        //DataRow row = dt.NewRow();
                        //row[0] = 0;
                        //row[1] = "Please select";
                        //dt.Rows.InsertAt(row, 0);

                        //Assign DataTable as DataSource.
                        TablesComboBox.DataSource = dt;
                        TablesComboBox.DisplayMember = "TABLE_NAME";
                        TablesComboBox.ValueMember = "TABLE_NAME";
                    }
                }

            }
            catch (Exception errortext)
            {
                MessageBox.Show(errortext.ToString());
            }
        }

        private void ExtractBtn_Click(object sender, EventArgs e)
        {
            try
            {
                string constr = @"Data Source=(LocalDB)\.;AttachDbFilename=" + DatabaseTextBox.Text + ";Integrated Security=True;Packet Size=32767";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM " + TablesComboBox.SelectedValue, con))
                    {
                        //Fill the DataTable with records from Table.
                        DataTable dt = new DataTable();
                        sda.Fill(dt);

                        dt.TableName = TablesComboBox.SelectedValue.ToString();

                        CreateExcel(dt);
                    }
                }

            }
            catch (Exception errortext)
            {
                MessageBox.Show(errortext.ToString());
            }
        }


        public void CreateExcel(DataTable datatable)
        {
            try
            {

                saveFileDialog1.Filter = "Excel |*.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new XLWorkbook();
                    string name = saveFileDialog1.FileName;

                    workbook.Worksheets.Add(datatable);

                    workbook.SaveAs(name);

                    MessageBox.Show("saved successfully", DateTime.Now.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception errortext)
            {
                MessageBox.Show(errortext.ToString());
            }
        }











    }
}
