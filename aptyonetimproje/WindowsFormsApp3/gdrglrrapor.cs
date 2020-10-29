using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp3
{

    public partial class gdrglrrapor : Form
    {
        SqlConnection con1;
        SqlDataReader sdr;
        SqlCommand skomut;

        //SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True");
        
        public gdrglrrapor()
        {
            InitializeComponent();
            con1 = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True;User Instance=true;");
        }

        private void gdrglrrapor_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Text = DateTime.Now.ToString("MM/dd/yyyy");
            dateTimePicker2.Text = DateTime.Now.ToString("MM/dd/yyyy");
        }

       
        private void button1_Click(object sender, EventArgs e)
        {
            string giderbilgi = "SELECT * FROM giderbilgi WHERE tarih BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "'";
            string gelirbilgi = "SELECT * FROM gelirbilgi WHERE tarih BETWEEN '" + dateTimePicker1.Text + "' AND '" + dateTimePicker2.Text + "'";
            MessageBox.Show(giderbilgi);
            MessageBox.Show(gelirbilgi);
            con1.Open();
            SqlDataAdapter da = new SqlDataAdapter(giderbilgi, con1);
            SqlDataAdapter da2 = new SqlDataAdapter(gelirbilgi, con1);
            System.Data.DataTable dt = new System.Data.DataTable();
            System.Data.DataTable dt2 = new System.Data.DataTable();
            da.Fill(dt);
            da2.Fill(dt2);
            //Bir DataTable oluşturarak DataAdapter ile getirilen verileri tablo içerisine dolduruyoruz.
            dataGridView1.DataSource = dt2;
            dataGridView4.DataSource = dt;
            this.dataGridView4.Columns["Id"].HeaderText = "İşlem No";
            this.dataGridView4.Columns["gider_turu"].HeaderText = "Gider Türü";
            this.dataGridView4.Columns["makbuz"].HeaderText = "Makbuz No";
            this.dataGridView4.Columns["tarih"].HeaderText = "Tarih";
            this.dataGridView4.Columns["tutar"].HeaderText = "Tutar";
            this.dataGridView4.Columns["aciklama"].HeaderText = "Açıklama";
            //
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

            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            sheet1.Name = "Gelir-Gider Tablosu";
            sheet1.Columns[1].ColumnWidth = 8;
            sheet1.Columns[2].ColumnWidth = 8;
            sheet1.Columns[3].ColumnWidth = 10;
            sheet1.Columns[4].ColumnWidth = 14;
            sheet1.Rows.HorizontalAlignment = HorizontalAlignment.Center;
            int StartCol = 1;
            int startCol2 = 10;
            int StartRow = 1;
            int StartRow2 = 1;
            float gelirToplam = 0;
            float giderToplam = 0;
            for (int j = 0; j < dataGridView4.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView4.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView4.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView4[j, i].Value == null ? "" : dataGridView4[j, i].Value;
                    myRange.Select();


                }
            }
            //
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow2, startCol2 + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow2++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow2 + i, startCol2 + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();


                }
            }
            for (int i = 0; i< dataGridView4.Columns.Count; i++)
            {
                try {
                    giderToplam += Convert.ToInt64(dataGridView1.Rows[4].Cells[i].Value);
                }
                catch(ArgumentOutOfRangeException)
                {
                    MessageBox.Show("Tarih Aralığı Boş");
                }
            }
            MessageBox.Show(Convert.ToString(giderToplam));
            
        }

        private void gdrglrrapor_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void gdrglrrapor_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
