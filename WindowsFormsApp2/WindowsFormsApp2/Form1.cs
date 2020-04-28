using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        OleDbConnection conn = new OleDbConnection();
        public Form1()
        {
            conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kunal\Documents\users1.mdb";
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int check = 1;
            while (check == 1)
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = conn;
                command.CommandText = "Select Top 1 print_id, student_roll, print_copy, print_document from requests where print_status=0 ";
                OleDbDataReader reader = command.ExecuteReader();
                int count = 0;
                int print_copy=1;
                string student_roll;
                string print_document="";
                int print_id=0;
                int copy = 1;
                while (reader.Read())
                {
                    string val1 = reader[0].ToString();
                    string val2 = reader[2].ToString();
                    student_roll = reader[1].ToString();
                    print_document = reader[3].ToString();
                    Int32.TryParse(val1, out print_id);
                    Int32.TryParse(val2, out print_copy);
                    count = count + 1;
                }
                conn.Close();
                if (print_id != 0)
                {
                    while (copy <= print_copy)
                    {
                        label1.Text = "Printing " + print_id;
                        ProcessStartInfo info = new ProcessStartInfo(print_document);

                        info.Verb = "Print";

                        info.CreateNoWindow = true;

                        info.WindowStyle = ProcessWindowStyle.Hidden;

                        Process.Start(info);

                        copy++;
                    }

                    conn.Open();
                    OleDbCommand command1 = new OleDbCommand();
                    command.Connection = conn;
                    command.CommandText = "update requests set print_status = 1 where print_id =" + print_id;
                    command.ExecuteNonQuery();
                    conn.Close();
                    
                }
                else
                {
                    label1.Text = "Idle";
                }

                Application.DoEvents();

            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            this.Close();
        }

        private void Button3_Click(object sender, EventArgs e)
        {

        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
