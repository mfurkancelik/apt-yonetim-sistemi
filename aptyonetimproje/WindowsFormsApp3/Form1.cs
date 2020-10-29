using System;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WindowsFormsApp3

{
    public partial class Form1 : Form
    {

        SqlConnection con = new SqlConnection(@"SERVER = remotemysql.com; DATABASE = 93Z8mvMd3D; User ID = 93Z8mvMd3D; PASSWORD = anWk0jOjUk;");
        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void button2_Click(object sender, EventArgs e)
        {
            aptbilgi aptbilgi = new aptbilgi();
            aptbilgi.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            uyebilgi uyebilgi = new uyebilgi();
            uyebilgi.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            borclandirmafrm borclandirmafrm = new borclandirmafrm();
            borclandirmafrm.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            gelirfrm gelirfrm = new gelirfrm();
            gelirfrm.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            giderfrm giderfrm = new giderfrm();
            giderfrm.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            gdrglrrapor gdrglrrapor = new gdrglrrapor();
            gdrglrrapor.ShowDialog();
        }

    }
}
