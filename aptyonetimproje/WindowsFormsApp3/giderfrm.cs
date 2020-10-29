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
    public partial class giderfrm : Form
    {
        SqlConnection con1;
        SqlDataReader sdr;
        SqlCommand skomut;
        //SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True");

        public void guncelle()
        {
            con1.Open();
            string kayit = "SELECT * from giderbilgi";
            SqlDataAdapter da = new SqlDataAdapter(kayit, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            con1.Close();

        }
        public giderfrm()
        {
            InitializeComponent();
            con1 = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True;User Instance=true;");
        }

        private void giderfrm_Load(object sender, EventArgs e)
        {
            con1.Open();
            string kayit = "SELECT * from giderbilgi";
            //musteriler tablosundaki tüm kayıtları çekecek olan sql sorgusu.
            SqlDataAdapter da = new SqlDataAdapter(kayit, con1);
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            this.dataGridView1.Columns["Id"].HeaderText = "İşlem No";
            this.dataGridView1.Columns["gider_turu"].HeaderText = "Gider Türü";
            this.dataGridView1.Columns["makbuz"].HeaderText = "Makbuz No";
            this.dataGridView1.Columns["tarih"].HeaderText = "Tarih";
            this.dataGridView1.Columns["tutar"].HeaderText = "Tutar";
            this.dataGridView1.Columns["aciklama"].HeaderText = "Açıklama";
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string sil = "INSERT INTO giderbilgi (gider_turu,makbuz,tarih,tutar,aciklama) VALUES ('" + comboBox1.Text + "','" + textBox2.Text + "','" + dateTimePicker1.Text + "','" + textBox1.Text + "','" + richTextBox1.Text + "')";
                //string sil = "INSERT INTO gelirbilgi (gelir_turu,kapi_no,tutar,aciklama) VALUES ('" + comboBox1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "','" + richTextBox1.Text + "')";
                SqlCommand da = new SqlCommand(sil, con1);
                con1.Open();
                da.ExecuteNonQuery();
                con1.Close();
                guncelle();
            }
            catch (SqlException)
            {
                MessageBox.Show("Hata Var");
                con1.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string guncellex = "UPDATE giderbilgi SET " + "gider_turu='" + comboBox1.Text + "',makbuz='" + textBox2.Text + "',tutar='" + textBox1.Text + "',aciklama='" + richTextBox1.Text + "'" + "WHERE Id = '" + textBox3.Text + "' ";
            SqlCommand da = new SqlCommand(guncellex, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            guncelle();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
                if (e.RowIndex >= 0)
                {
                    //gets a collection that contains all the rows
                    DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                    //populate the textbox from specific value of the coordinates of column and row.
                    textBox3.Text = row.Cells[0].Value.ToString();
                    comboBox1.Text = row.Cells[1].Value.ToString();
                    textBox2.Text = row.Cells[2].Value.ToString();
                textBox1.Text = row.Cells[4].Value.ToString();
                richTextBox1.Text = row.Cells[5].Value.ToString();
                }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sil = "DELETE FROM giderbilgi " + "WHERE Id = '" + textBox3.Text + "' ";
            SqlCommand da = new SqlCommand(sil, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            guncelle();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string arama = "SELECT * from giderbilgi WHERE tarih BETWEEN '" + dateTimePicker4.Text + "' AND '" + dateTimePicker5.Text + "' AND gider_turu = '" + comboBox2.Text + "' " ;
            SqlDataAdapter sda = new SqlDataAdapter(arama, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            guncelle();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string arama = "SELECT * from giderbilgi WHERE tarih BETWEEN '" + dateTimePicker2.Text + "' AND '" + dateTimePicker3.Text + "'";
            MessageBox.Show(arama);
            SqlDataAdapter sda = new SqlDataAdapter(arama, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Missing);
            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
            sheet1.Name = "Gider Listesi";
            sheet1.Columns[1].ColumnWidth = 4;
            sheet1.Columns[2].ColumnWidth = 14;
            sheet1.Columns[3].ColumnWidth = 14;
            sheet1.Columns[4].ColumnWidth = 10;
            sheet1.Columns[5].ColumnWidth = 15;
            sheet1.Columns[6].ColumnWidth = 20;
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

        private void giderfrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
