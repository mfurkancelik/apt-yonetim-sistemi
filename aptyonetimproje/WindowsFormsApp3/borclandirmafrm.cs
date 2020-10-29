using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace WindowsFormsApp3
{
    public partial class borclandirmafrm : Form
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
                comboBox2.Items.Add(sdr["adi_soyadi"]);
            }
            con1.Close();

        }
        public void guncelle()
        {
            con1.Open();
            string kayit = "SELECT * from borclandirmafrm";
            SqlDataAdapter sda = new SqlDataAdapter(kayit, con1);
            //SqlDataAdapter da = new SqlDataAdapter(komut);
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            dataGridView1.DataSource = dt;
            con1.Close();
        }

        public borclandirmafrm()
        {
            InitializeComponent();
            con1 = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True;User Instance=true;");
        }
        public void borclandirmaVerileriCek()
        {
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker2.Value = DateTime.Today;
            dateTimePicker3.Value = DateTime.Today;
            con1.Open();
            string kayit = "SELECT * from borclandirmafrm";
            //musteriler tablosundaki tüm kayıtları çekecek olan sql sorgusu.
            SqlDataAdapter sda = new SqlDataAdapter(kayit, con1);
            //SqlCommand komut = new SqlCommand(kayit, con);
            //Sorgumuzu ve baglantimizi parametre olarak alan bir SqlCommand nesnesi oluşturuyoruz.
            //SqlDataAdapter da = new SqlDataAdapter(komut);
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            this.dataGridView1.Columns["Id"].HeaderText = "İşlem No";
            this.dataGridView1.Columns["blok_no"].HeaderText = "Blok No";
            this.dataGridView1.Columns["kapi_no"].HeaderText = "Daire No";
            this.dataGridView1.Columns["adi_soyadi"].HeaderText = "Adı Soyadı";
            this.dataGridView1.Columns["borcturu"].HeaderText = "Borç Türü";
            this.dataGridView1.Columns["tarih"].HeaderText = "Tarih";
            this.dataGridView1.Columns["tutar"].HeaderText = "Tutar";
            this.dataGridView1.Columns["odenen_tutar"].HeaderText = "Ödenen Tutar";
            this.dataGridView1.Columns["kalan_tutar"].HeaderText = "Kalan Tutar";
            this.dataGridView1.Columns["aciklama"].HeaderText = "Açıklama";
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }
        private void borclandirmafrm_Load(object sender, EventArgs e)
        {
            //this.Height = Screen.PrimaryScreen.Bounds.Height;
            //this.Width = Screen.PrimaryScreen.Bounds.Width;
            dairelericek();
            borclandirmaVerileriCek();


        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string sil = "INSERT INTO borclandirmafrm (blok_no,kapi_no,adi_soyadi,borcturu,tarih,tutar,aciklama) VALUES ('" + comboBox3.Text + "','" + comboBox4.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "','" + dateTimePicker1.Text + "','" + textBox1.Text + "','" + richTextBox1.Text + "')";
                SqlCommand sda = new SqlCommand(sil, con1);
                con1.Open();
                sda.ExecuteNonQuery();
                con1.Close();
                guncelle();
            }
            catch (SqlException)
            {
                MessageBox.Show("Aynı veri girilemez");
                con1.Close();
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            con1.Open();
            string kayit = "SELECT * from borclandirmafrm where " + "kapi_no = '" + textBox2.Text + "' AND blok_no = '" + textBox4.Text + "' AND tarih BETWEEN '" + dateTimePicker5.Text + "' AND '" + dateTimePicker4.Text + "'";
            //musteriler tablosundaki tüm kayıtları çekecek olan sql sorgusu.
            MessageBox.Show(kayit);
            SqlDataAdapter sda = new SqlDataAdapter(kayit, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            con1.Open();
            string kayit = "SELECT * from borclandirmafrm where " + "borcturu = '" + comboBox3.Text +
            "' AND tarih BETWEEN '" + dateTimePicker5.Text + "' AND '" + dateTimePicker4.Text + "'";
            //musteriler tablosundaki tüm kayıtları çekecek olan sql sorgusu.
            SqlDataAdapter sda = new SqlDataAdapter(kayit, con1);
            //SqlCommand komut = new SqlCommand(kayit, con1);
            //Sorgumuzu ve baglantimizi parametre olarak alan bir SqlCommand nesnesi oluşturuyoruz.
            //SqlDataAdapter da = new SqlDataAdapter(komut);
            //SqlDataAdapter sınıfı verilerin databaseden aktarılması işlemini gerçekleştirir.
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            con1.Open();
            string kayit = "SELECT * from borclandirmafrm WHERE tarih BETWEEN '" + dateTimePicker2.Text + "' AND '" + dateTimePicker3.Text + "'";
            SqlDataAdapter sda = new SqlDataAdapter(kayit, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            sda.Fill(dt);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt;
            //Formumuzdaki DataGridViewin veri kaynağını oluşturduğumuz tablo olarak gösteriyoruz.
            con1.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sil = "DELETE FROM borclandirmafrm " + "WHERE Id = '" + textBox5.Text + "' ";
            SqlCommand da = new SqlCommand(sil, con1);
            con1.Open();
            da.ExecuteNonQuery();
            con1.Close();
            guncelle();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string guncellex = "UPDATE borclandirmafrm SET " + "borcturu='" + comboBox1.Text + "',kapi_no='" + comboBox4.Text + "',blok_no='" + comboBox3.Text + "',adi_soyadi='" + comboBox2.Text + "',tutar='" + textBox1.Text + "',aciklama='" + richTextBox1.Text + "'" + "WHERE Id = '" + textBox5.Text + "' ";
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
                //gets a collection that con1tains all the rows
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                //populate the textbox from specific value of the coordinates of column and row.
                textBox5.Text = row.Cells[0].Value.ToString();
                comboBox3.Text = row.Cells[1].Value.ToString();
                comboBox4.Text = row.Cells[2].Value.ToString();
                comboBox2.Text = row.Cells[3].Value.ToString();
                comboBox1.Text = row.Cells[4].Value.ToString();
                textBox1.Text = row.Cells[6].Value.ToString();
                richTextBox1.Text = row.Cells[7].Value.ToString();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            guncelle();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            skomut = new SqlCommand();
            skomut.CommandText = "SELECT blok_no,kapi_no FROM uyebilgiler WHERE adi_soyadi = '" + comboBox2.Text + "'";
            skomut.Connection = con1;
            skomut.CommandType = CommandType.Text;
            con1.Open();
            sdr = skomut.ExecuteReader();
            while (sdr.Read())
            {
                string smd1 = (string)sdr["blok_no"].ToString();
                string smd2 = (string)sdr["kapi_no"].ToString();
                comboBox3.Text = smd1;
                comboBox4.Text = smd2;
            }
            con1.Close();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            sheet1.Name = "Borç Listesi";
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

        private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            borclandirmaVerileriCek();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void borclandirmafrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                comboBox6.Enabled = true;
            }
            else
            {
                comboBox6.Enabled = false;
            }
        }

        private void comboBox6_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            switch (comboBox6.Text)
            {
                case "1":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak " + comboBox1.Text;
                    break;
                case "2":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat " + comboBox1.Text;
                    break;
                case "3":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart " + comboBox1.Text;
                    break;
                case "4":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan " + comboBox1.Text;
                    break;
                case "5":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs " + comboBox1.Text;
                    break;
                case "6":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs-Haziran " + comboBox1.Text;
                    break;
                case "7":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs-Haziran-Temmuz " + comboBox1.Text;
                    break;
                case "8":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs-Haziran-Temmuz-Ağustos " + comboBox1.Text;
                    break;
                case "9":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs-Haziran-Temmuz-Ağustos-Eylül " + comboBox1.Text;
                    break;
                case "10":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs-Haziran-Temmuz-Ağustos-Eylül-Ekim " + comboBox1.Text;
                    break;
                case "11":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs-Haziran-Temmuz-Ağustos-Eylül-Ekim-Kasım " + comboBox1.Text;
                    break;
                case "12":
                    richTextBox1.Clear();
                    richTextBox1.Text = "Ocak-Şubat-Mart-Nisan-Mayıs-Haziran-Temmuz-Ağustos-Eylül-Ekim-Kasım-Aralık " + comboBox1.Text;
                    break;
            }
        }
    }
}
