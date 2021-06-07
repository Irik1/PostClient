using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lab2
{
    public partial class NewUser : Form
    {
        public NewUser()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;" +
              @"AttachDbFilename=|DataDirectory|\Database.mdf;
                Integrated Security=True;
                Connect Timeout=30;";
            SqlConnection Con = new SqlConnection(sqlCon);
            Con.Open();
            SqlDataReader sqlReader = null;

            SqlCommand command = new SqlCommand("Insert into [Users] values ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text  + "')", Con);

            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                Form1 main = Owner as Form1;
                if (main != null)
                {
                    main.Form1_Load();
                }
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
            Con.Close();
        }
    }
}
