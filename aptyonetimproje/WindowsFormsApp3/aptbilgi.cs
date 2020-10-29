using System;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace WindowsFormsApp3
{
    public partial class aptbilgi : Form
    {
        //SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True");
        //public MySqlConnection con= new MySqlConnection("Server=remotemysql.com;Database=93Z8mvMd3D;Uid=93Z8mvMd3D;Pwd='anWk0jOjUk';Encrypt=false;AllowUserVariables=True;UseCompression=True");
        SqlConnection con1;
        SqlDataReader sdr;
        SqlCommand skomut;
        public aptbilgi()
        {
            InitializeComponent();
            con1 = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database4.mdf;Integrated Security=True;User Instance=true;");
        }

        private void aptbilgi_Load(object sender, EventArgs e)
        {
            skomut = new SqlCommand();
            con1.Open();
            skomut.Connection = con1;
            skomut.CommandText = "Select * from aptbilgileri";
            sdr = skomut.ExecuteReader();
            if (sdr.Read())
            {
                textBox1.Text = sdr["apartmanismi"].ToString().Trim();
                richTextBox1.Text = sdr["adres"].ToString().Trim();
                textBox2.Text = sdr["hesapbanka"].ToString().Trim();
                textBox3.Text = sdr["hesapsube"].ToString().Trim();
                textBox4.Text = sdr["hesapno"].ToString().Trim();
                textBox5.Text = sdr["ibanNo"].ToString().Trim();
                textBox6.Text = sdr["aktifdonem"].ToString().Trim();
                textBox7.Text = sdr["yonetici"].ToString().Trim();
                textBox8.Text = sdr["yoneticiyrd"].ToString().Trim();
                textBox9.Text = sdr["yoneticiyrd2"].ToString().Trim();
                textBox10.Text = sdr["denetici1"].ToString().Trim();
                textBox11.Text = sdr["denetici2"].ToString().Trim();
                textBox12.Text = sdr["denetici3"].ToString().Trim();
                textBox13.Text = sdr["denetici4"].ToString().Trim();
                textBox14.Text = sdr["denetici5"].ToString().Trim();
            }
            else
            {
                MessageBox.Show("Veritabanı ile bağlantı kurulamadı.");
            }
            con1.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            con1.Open();
            skomut.Connection = con1;
            skomut.CommandText = "UPDATE aptbilgileri SET apartmanismi=@apartmanismi,adres=@adres,hesapbanka=@hesapbanka,hesapsube=@hesapsube,hesapno=@hesapno,ibanNo=@ibanNo,aktifdonem=@aktifdonem,yonetici=@yonetici,yoneticiyrd=@yoneticiyrd,yoneticiyrd2=@yoneticiyrd2,denetici1=@denetici1,denetici2=@denetici2,denetici3=@denetici3,denetici4=@denetici4,denetici5=@denetici5";
            skomut.Parameters.AddWithValue("@apartmanismi", textBox1.Text.Trim());
            skomut.Parameters.AddWithValue("@adres", richTextBox1.Text.Trim());
            skomut.Parameters.AddWithValue("@hesapbanka", textBox2.Text.Trim());
            skomut.Parameters.AddWithValue("@hesapsube", textBox3.Text.Trim());
            skomut.Parameters.AddWithValue("@hesapno", textBox4.Text.Trim());
            skomut.Parameters.AddWithValue("@ibanNo", textBox5.Text.Trim());
            skomut.Parameters.AddWithValue("@aktifdonem", textBox6.Text.Trim());
            skomut.Parameters.AddWithValue("@yonetici", textBox7.Text.Trim());
            skomut.Parameters.AddWithValue("@yoneticiyrd", textBox8.Text.Trim());
            skomut.Parameters.AddWithValue("@yoneticiyrd2", textBox9.Text.Trim());
            skomut.Parameters.AddWithValue("@denetici1", textBox10.Text.Trim());
            skomut.Parameters.AddWithValue("@denetici2", textBox11.Text.Trim());
            skomut.Parameters.AddWithValue("@denetici3", textBox12.Text.Trim());
            skomut.Parameters.AddWithValue("@denetici4", textBox13.Text.Trim());
            skomut.Parameters.AddWithValue("@denetici5", textBox14.Text.Trim());
            skomut.ExecuteNonQuery();
            MessageBox.Show("Bilgiler Güncellendi");
            con1.Close();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
