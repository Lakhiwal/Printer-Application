using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public static string roll = "";
        public static string name = "";
        public static int pages = 0;
        OleDbConnection conn = new OleDbConnection();

        public Form1()
        {
            conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Project\Project 7th Sem Printer App\users1.mdb";
            InitializeComponent();
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(textBox1.Text.Trim()))
                {

                    textBox1.BackColor = System.Drawing.Color.Yellow;

                    MessageBox.Show("Please enter username.");

                    return;
                }
                else if(string.IsNullOrEmpty(textBox2.Text.Trim()))
                    {

                        textBox2.BackColor = System.Drawing.Color.Yellow;

                        MessageBox.Show("Please enter Password.");

                        return;
                    }
                conn.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = conn;
                command.CommandText = "Select student_roll, student_name, pages from students where student_roll='" + textBox1.Text + "' and student_password='" + textBox2.Text + "'";
                OleDbDataReader reader = command.ExecuteReader();
                int count = 0;
                while (reader.Read())
                {
                    roll = reader[0].ToString();
                    name = reader[1].ToString();
                    String val = reader[2].ToString();
                    Int32.TryParse(val, out pages);
                    count =count + 1;
                }
                if (count == 1)
                {
                    conn.Close();
                    conn.Dispose();
                    this.Hide();
                    Form2 f2 = new Form2();
                    f2.ShowDialog();
                }
                else
                {
                    conn.Close();
                    MessageBox.Show("Invalid Roll no. or Password");
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
