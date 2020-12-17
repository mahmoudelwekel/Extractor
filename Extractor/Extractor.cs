using ClosedXML.Excel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Extractor
{
    public partial class Extractor : Form
    {
        string constr;

        public Extractor()
        {
            InitializeComponent();
            AttachDatabase();
        }

        void AttachDatabase()
        {
            try
            {
                openFileDialog1.Filter = "Images Files|*.MDF";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    DatabaseTextBox.Text = openFileDialog1.FileName;
                    constr = @"Data Source=(LocalDB)\.;AttachDbFilename=" + DatabaseTextBox.Text + ";Integrated Security=True;Packet Size=32767";

                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        using (SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM INFORMATION_SCHEMA.TABLES order by TABLE_TYPE,TABLE_NAME", con))
                        {
                            DataTable dt = new DataTable();
                            sda.Fill(dt);

                            TablesComboBox.DisplayMember = "TABLE_NAME";
                            TablesComboBox.ValueMember = "TABLE_NAME";
                            TablesComboBox.DataSource = dt;
                        }
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

        private void Attach_Click(object sender, EventArgs e)
        {
            try
            {
                AttachDatabase();
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

        private void TablesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM " + TablesComboBox.SelectedValue, con))
                    {
                        //Fill the DataTable with records from Table.
                        DataTable dt = new DataTable();
                        sda.Fill(dt);

                        dataGridView1.DataSource = dt;

                    }
                }

            }
            catch (Exception errortext)
            {
                MessageBox.Show(errortext.ToString());
            }
        }
    }
}
