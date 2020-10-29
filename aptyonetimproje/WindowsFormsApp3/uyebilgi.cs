using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp3
{
    public partial class uyebilgi : Form
    {
        private BindingSource bindingSource1 = new BindingSource();
        SqlConnection con1;
        SqlDataReader sdr;
        SqlCommand skomut;
        public void bilgileriGetir(string param)
        {
            con1.Open();
            //musteriler tablosundaki tüm kayıtları çekecek olan sql sorgusu.
            SqlDataAdapter da = new SqlDataAdapter(param, con1);
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            this.dataGridView1.Columns["blok_no"].HeaderText = "Blok Numarası";
            this.dataGridView1.Columns["kapi_no"].HeaderText = "Daire Numarası";
            this.dataGridView1.Columns["kat"].HeaderText = "Kat";
            this.dataGridView1.Columns["daire_tipi"].HeaderText = "Daire Tipi";
            this.dataGridView1.Columns["adi_soyadi"].HeaderText = "Adı Soyadı";
            this.dataGridView1.Columns["cep_telefonu"].HeaderText = "Cep Telefonu";
            this.dataGridView1.Columns["is_telefonu"].HeaderText = "İş Telefonu";
            this.dataGridView1.Columns["ev_telefonu"].HeaderText = "Ev Telefonu";
            this.dataGridView1.Columns["mail_adresi"].HeaderText = "Mail Adresi";
            this.dataGridView1.Columns["adres"].HeaderText = "Adres";
            dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }
        public uyebilgi()
        {
            InitializeComponent();
            con1 = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True;User Instance=true;");
        }

        private void uyebilgi_Load(object sender, EventArgs e)
        {
            label13.Text = "Kayıt güncellemek için öncelikle tablodan" + "\n" + "güncellemek istediğiniz daireyi seçin.";
            string query = "SELECT * FROM uyebilgiler";
            bilgileriGetir(query);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string sil = "INSERT INTO uyebilgiler (blok_no,kapi_no,kat,daire_tipi,adi_soyadi,cep_telefonu,is_telefonu,ev_telefonu,mail_adresi,adres) VALUES ('" + comboBox1.Text + "','" + textBox1.Text + "','" + textBox7.Text + "','" + comboBox2.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + richTextBox1.Text + "')";
                SqlCommand da = new SqlCommand(sil, con1);
                con1.Open();
                da.ExecuteNonQuery();
                con1.Close();
                string query = "SELECT * FROM uyebilgiler";
                bilgileriGetir(query);
            }
            catch (SqlException)
            {
                MessageBox.Show("Aynı veri girilemez");
                con1.Close();
            }
         }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                //gets a collection that contains all the rows
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                //populate the textbox from specific value of the coordinates of column and row.
                textBox1.Text = row.Cells[0].Value.ToString();               
                textBox2.Text = row.Cells[1].Value.ToString();
                textBox3.Text = row.Cells[2].Value.ToString();
                textBox4.Text = row.Cells[3].Value.ToString();
                textBox5.Text = row.Cells[4].Value.ToString();
                textBox6.Text = row.Cells[5].Value.ToString();
                textBox7.Text = row.Cells[7].Value.ToString();
                richTextBox1.Text = row.Cells[6].Value.ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string sil = "DELETE FROM uyebilgiler " + "WHERE kapi_no = '" + textBox1.Text + "' ";
            SqlCommand da = new SqlCommand(sil, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            string query = "SELECT * FROM uyebilgiler";
            bilgileriGetir(query);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string guncellex = "UPDATE uyebilgiler SET " + "kapi_no='" + textBox1.Text + "',adi_soyadi='" + textBox2.Text + "',daire_tipi='" + comboBox2.Text + "',cep_telefonu='" + textBox3.Text + "',is_telefonu='" + textBox4.Text + "',ev_telefonu='" + textBox5.Text + "',mail_adresi='" + textBox6.Text + "',adres='" + richTextBox1.Text + "',kat='" + textBox7.Text + "'" + "WHERE kapi_no = '" + textBox1.Text + "' ";
            SqlCommand da = new SqlCommand(guncellex, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            string query = "SELECT * FROM uyebilgiler";
            bilgileriGetir(query);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                string kayit = "INSERT INTO uyebilgiler (blok_no,kapi_no,kat,daire_tipi,adi_soyadi,cep_telefonu,is_telefonu,ev_telefonu,mail_adresi,adres) VALUES ('" + comboBox1.Text + "','" + textBox1.Text + "','" + textBox7.Text + "','" + comboBox2.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + richTextBox1.Text + "')";
                SqlCommand da = new SqlCommand(kayit, con1);
                con1.Open();
                da.ExecuteNonQuery();
                con1.Close();
                string query = "SELECT * FROM uyebilgiler";
                bilgileriGetir(query);
            }
            catch (SqlException)
            {
                MessageBox.Show("Aynı veri girilemez");
                con1.Close();
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            string guncellex = "UPDATE uyebilgiler SET " + "blok_no='" + comboBox1.Text + "',kapi_no='" + textBox1.Text + "',kat='" + textBox7.Text + "',daire_tipi='"
                + comboBox2.Text + "',adi_soyadi='" + textBox2.Text + "',cep_telefonu='"
                + textBox3.Text + "',is_telefonu='" + textBox4.Text + "',ev_telefonu='"
                + textBox5.Text + "',mail_adresi='" + textBox6.Text + "',adres='"
                + richTextBox1.Text + "' WHERE kapi_no = '" + textBox1.Text + "' AND blok_no = '" + comboBox1.Text + "' AND adi_soyadi = '" + uyeisimsql.Text + "'  ";
            SqlCommand da = new SqlCommand(guncellex, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            string query = "SELECT * FROM uyebilgiler";
            bilgileriGetir(query);
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            string sil = "DELETE FROM uyebilgiler " + "WHERE kapi_no = '" + textBox1.Text + "' AND adi_soyadi = '" + textBox2.Text + "' ";
            SqlCommand da = new SqlCommand(sil, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            string query = "SELECT * FROM uyebilgiler";
            bilgileriGetir(query);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string query = "SELECT * FROM uyebilgiler WHERE blok_no LIKE '" + comboBox4.Text + "'";
            bilgileriGetir(query);
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            string query = "SELECT * FROM uyebilgiler WHERE adi_soyadi LIKE '" + textBox10.Text + "'";
            bilgileriGetir(query);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            string query = "SELECT * FROM uyebilgiler";
            bilgileriGetir(query);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string query = "SELECT * FROM uyebilgiler WHERE daire_tipi LIKE '" + comboBox3.Text + "'";
            bilgileriGetir(query);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            sheet1.Name = "Apartman Sakinleri";
            sheet1.Columns[1].ColumnWidth = 4;
            sheet1.Columns[2].ColumnWidth = 4;
            sheet1.Columns[3].ColumnWidth = 4;
            sheet1.Columns[4].ColumnWidth = 8;
            sheet1.Columns[5].ColumnWidth = 15;
            sheet1.Columns[6].ColumnWidth = 10;
            sheet1.Columns[7].ColumnWidth = 10;
            sheet1.Columns[8].ColumnWidth = 10;
            sheet1.Columns[9].ColumnWidth = 15;
            sheet1.Columns[10].ColumnWidth = 20;
            sheet1.Rows.HorizontalAlignment = HorizontalAlignment.Center;
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();


                }
            }
        }

        private void uyebilgi_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                //gets a collection that contains all the rows
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                //populate the textbox from specific value of the coordinates of column and row.
                comboBox1.Text = row.Cells[0].Value.ToString();
                textBox1.Text = row.Cells[1].Value.ToString();
                textBox7.Text = row.Cells[2].Value.ToString(); //kat
                comboBox2.Text = row.Cells[3].Value.ToString(); // daire tipi
                textBox2.Text = row.Cells[4].Value.ToString();
                uyeisimsql.Text = row.Cells[4].Value.ToString();
                textBox3.Text = row.Cells[5].Value.ToString();
                textBox4.Text = row.Cells[6].Value.ToString();
                textBox5.Text = row.Cells[7].Value.ToString();
                textBox6.Text = row.Cells[8].Value.ToString();
                richTextBox1.Text = row.Cells[9].Value.ToString();
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
