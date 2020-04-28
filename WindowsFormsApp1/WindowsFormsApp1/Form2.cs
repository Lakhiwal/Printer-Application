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
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using Spire.Pdf;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        OleDbConnection conn = new OleDbConnection();
        public Form2()
        {
            conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kunal\Documents\users1.mdb";
            InitializeComponent();
        }
        static string CalculateStatistics(string path)
        {
            // Load a DOCX file.
            string filePath = path;

            DocumentCore dc = DocumentCore.Load(filePath);

            // Update and count the number of words and pages in the file.
            dc.CalculateStats();

            // Show statistics.
            return dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages];
        }


        private void Form2_Load(object sender, EventArgs e)
        {
            label1.Text = "Name: " +Form1.name;
            label2.Text = "Pages: " +Form1.pages + "";
            label3.Text = "Roll: " + Form1.roll;
            
        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            Form1.roll = "";
            Form1.name = "";
            Form1.pages = 0;
            this.Hide();
            Form1 f1 = new Form1();
            f1.ShowDialog();
        }

        OpenFileDialog ofd = new OpenFileDialog();

        private void Button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = ofd.ShowDialog();

            string[] s = ofd.FileName.Split('.');

            if (dr.ToString() == "OK")

            {

                if (s.Length > 1)

                    if (s[1] == "doc" || s[1] == "docx" || s[1] == "jpg" || s[1] == "pdf" || s[1] == "jpeg" || s[1] == "png")

                        textBox1.Text = ofd.FileName;

                    else

                        MessageBox.Show("Please select doc,docx,jpeg or pdf file !!");

            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text.Trim()))
            {

                textBox1.BackColor = System.Drawing.Color.Yellow;

                MessageBox.Show("Please Select file.");

                return;
            }
            if (string.IsNullOrEmpty(textBox2.Text.Trim()))
            {

                textBox2.BackColor = System.Drawing.Color.Yellow;

                MessageBox.Show("Please enter no. of copies.");

                return;
            }
            string student_roll = Form1.roll;
            string temp_name = ofd.FileName;
            string print_document = "";
            int print_status = 0;
            int print_copy = 0;

            if (Int32.TryParse(textBox2.Text, out print_copy))
            {
                if (print_copy < 1)
                {
                    MessageBox.Show("No. of copies should be 1 or Greater.");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Enter valid number of copies.");
                return;
            }

            string[] s = ofd.FileName.Split('.');
            int page_count=0;
            int total_count = 0;
            if (s[1] == "pdf")
            {
                PdfDocument document = new PdfDocument();

                document.LoadFromFile(temp_name);
                page_count = document.Pages.Count;
                document.Close();

            }
            else if (s[1] == "doc" || s[1] == "docx")
            {
                Int32.TryParse(CalculateStatistics(temp_name), out page_count);
            }
            else if (s[1] == "jpg" || s[1] == "png" || s[1] == "jpeg")
            {
                page_count = 1;
            }
            total_count = page_count * print_copy;
            int new_pages = Form1.pages - total_count;
            if (new_pages < 0)
            {
                MessageBox.Show("Page limit exceeded.");
                return;
            }
            string[] c = ofd.FileName.Split('\\');
            string only_name = c.Last();
            //string path = Application.StartupPath.Substring(0, (Application.StartupPath.Length - 10));
            print_document = "G:\\Documents\\" + only_name;
            //MessageBox.Show(path+"\\Documents\\"+only_name);
            System.IO.File.Copy(temp_name, print_document);
            

            try
            {
                conn.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = conn;
                command.CommandText = "Insert into requests(student_roll,print_status,print_document,print_copy) values('"+student_roll+"',"+print_status+",'"+print_document+"',"+print_copy+")";
                command.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Printing " + total_count +" pages.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }

            if (new_pages >= 0)
            {
                try
                {
                    conn.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = conn;
                    command.CommandText = "update students set pages = " + new_pages + " where student_roll = '" + Form1.roll + "'";
                    command.ExecuteNonQuery();
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error" + ex);
                }
                Form1.pages = new_pages;

                label2.Text = "Pages: "+new_pages;

            }

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Label4_Click(object sender, EventArgs e)
        {

        }
    }
}



