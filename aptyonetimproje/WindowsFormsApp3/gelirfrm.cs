using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp3
{
    public partial class gelirfrm : Form
    {
        SqlConnection con1;
        SqlDataReader sdr;
        SqlCommand skomut;
        //SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True");
        public void dairelericek()
        {
            skomut = new SqlCommand();
            skomut.CommandText = "SELECT adi_soyadi,blok_no,kapi_no FROM uyebilgiler";
            skomut.Connection = con1;
            skomut.CommandType = CommandType.Text;
            con1.Open();
            sdr = skomut.ExecuteReader();
            while (sdr.Read())
            {
                comboBox3.Items.Add(sdr["adi_soyadi"]);
                comboBox5.Items.Add(sdr["adi_soyadi"]);
            }
            con1.Close();

        }
        public void guncelle()
        {
            con1.Open();
            string kayit = "SELECT * from gelirbilgi";
            SqlDataAdapter da = new SqlDataAdapter(kayit, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            con1.Close();

        }
        public gelirfrm()
        {
            InitializeComponent();
            con1 = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True;User Instance=true;");
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void gelirfrm_Load(object sender, EventArgs e)
        {
            dairelericek();
            con1.Open();
            string kayit = "SELECT * from gelirbilgi";
            //musteriler tablosundaki tüm kayıtları çekecek olan sql sorgusu.
            SqlDataAdapter da = new SqlDataAdapter(kayit, con1);
            //Sorgumuzu ve baglantimizi parametre olarak alan bir SqlCommand nesnesi oluşturuyoruz.
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            this.dataGridView1.Columns["Id"].HeaderText = "İşlem No";
            this.dataGridView1.Columns["gelir_turu"].HeaderText = "Gelir Türü";
            this.dataGridView1.Columns["blok_no"].HeaderText = "Blok No";
            this.dataGridView1.Columns["kapi_no"].HeaderText = "Daire No";
            this.dataGridView1.Columns["adi_soyadi"].HeaderText = "Adı Soyadı";
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
                string sil = "INSERT INTO gelirbilgi (gelir_turu,adi_soyadi,blok_no,kapi_no,tarih,tutar,aciklama) VALUES ('" + comboBox1.Text + "','" + comboBox3.Text + "','" + comboBox4.Text + "'," +
                    "'" + comboBox2.Text + "','" + dateTimePicker1.Text + "','" + textBox1.Text + "','" + richTextBox1.Text + "')";
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

        private void button3_Click(object sender, EventArgs e)
        {
            string sil = "DELETE FROM gelirbilgi " + "WHERE Id = '" + textBox2.Text + "' ";
            SqlCommand da = new SqlCommand(sil, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            guncelle();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string guncellex = "UPDATE gelirbilgi SET " + "gelir_turu='" + comboBox1.Text + "',kapi_no='" + comboBox2.Text + "',tutar='" + textBox1.Text + "',aciklama='" + richTextBox1.Text + "'" + "WHERE Id = '" + textBox2.Text + "' ";
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
                textBox2.Text = row.Cells[0].Value.ToString();
                comboBox1.Text = row.Cells[1].Value.ToString();
                comboBox2.Text = row.Cells[2].Value.ToString();
                textBox1.Text = row.Cells[4].Value.ToString();
                richTextBox1.Text = row.Cells[5].Value.ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            sheet1.Name = "Borç Listesi";
            sheet1.Columns[1].ColumnWidth = 4;
            sheet1.Columns[2].ColumnWidth = 8;
            sheet1.Columns[3].ColumnWidth = 5;
            sheet1.Columns[4].ColumnWidth = 5;
            sheet1.Columns[5].ColumnWidth = 15;
            sheet1.Columns[6].ColumnWidth = 15;
            sheet1.Columns[7].ColumnWidth = 8;
            sheet1.Columns[8].ColumnWidth = 18;
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

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            skomut = new SqlCommand();
            skomut.CommandText = "SELECT blok_no,kapi_no FROM uyebilgiler WHERE adi_soyadi = '" + comboBox3.Text + "'";
            skomut.Connection = con1;
            skomut.CommandType = CommandType.Text;
            con1.Open();
            sdr = skomut.ExecuteReader();
            while (sdr.Read())
            {
                string smd1 = (string)sdr["blok_no"].ToString();
                string smd2 = (string)sdr["kapi_no"].ToString();
                comboBox4.Text = smd1;
                comboBox2.Text = smd2;
            }
            con1.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            con1.Open();
            string kayit = "SELECT * from gelirbilgi WHERE tarih BETWEEN '" + dateTimePicker2.Text + "' AND '" + dateTimePicker3.Text + "' AND adi_soyadi = '" + comboBox5.Text + "'" +
                " AND kapi_no = '" + comboBox6.Text + "'";
            MessageBox.Show(kayit);
            SqlDataAdapter sda = new SqlDataAdapter(kayit, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }

        private void gelirfrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            skomut = new SqlCommand();
            skomut.CommandText = "SELECT kapi_no FROM uyebilgiler WHERE adi_soyadi = '" + comboBox5.Text + "'";
            skomut.Connection = con1;
            skomut.CommandType = CommandType.Text;
            con1.Open();
            sdr = skomut.ExecuteReader();
            while (sdr.Read())
            {
                string smd2 = (string)sdr["kapi_no"].ToString();
                comboBox6.Text = smd2;
            }
            con1.Close();
        }
    }
}
