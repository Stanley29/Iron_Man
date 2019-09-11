using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace taskt.UI.Forms
{
    public partial class Login : Form
    {
        public string Status { get; set; }
        public DateTime expirationDate = new DateTime(2019, 9, 20, 18, 30, 25);
        // год - месяц - день - час - минута - секунда

        public Login()
        {
            InitializeComponent();
            Status = "Undefined";
        }

        private void enterButton_Click(object sender, EventArgs e)
        {
            string login = loginBox.Text;
            string passw = passwordBox.Text;
            // DateTime today = DateTime.Now;


            if (login.Length > 0 && passw.Length > 0)
            {
                //  passw = GetMd5Hash(passw);
                /* string query = "SELECT login, passw, status ";
                 query += "FROM Users WHERE ";
                 query += $"login='{login}' AND passw='{passw}'";

                 string conn_str =
                     ConfigurationManager.ConnectionStrings["server1"].ConnectionString;
                 SqlConnection conn = new SqlConnection(conn_str);
                 conn.Open();
                 SqlCommand cmd = new SqlCommand(query, conn);
                 SqlDataReader reader = cmd.ExecuteReader();
                 //    if (reader.Read())*/
                if ((passw == "IronMan#3579@" && login == "admin") /*&& (validTime() == true)*/)
                {
                    // Status = (string)reader["status"];
                    Status = "admin";
                    //  conn.Close();
                    this.DialogResult = DialogResult.OK;
                    //if (validTime() == true)
                    //{
                    //    this.DialogResult = DialogResult.OK;
                    //}

                }
                else if ((passw == "ironManUser@579#" && login == "user") && (validTime() == true))
                {
                    Status = "user";
                    // conn.Close();
                    this.DialogResult = DialogResult.OK;
                    //if (validTime() == true)
                    //{
                    //    this.DialogResult = DialogResult.OK;
                    //}
                }
                else
                {
                    this.DialogResult = DialogResult.None;
                    //MessageBox.Show("Wrong login or password");
                    if (validTime() == false)
                    {
                        MessageBox.Show("License expired");
                    }
                    else
                    {
                        MessageBox.Show("Wrong login or password");
                    }
                }
            }
        }

        private void resetButton_Click(object sender, EventArgs e)
        {
            loginBox.Clear();
            passwordBox.Clear();
        }

        public bool validTime()
        {
            //if (expirationDate > DateTime.Now )
            //{
            //    return true;
            //}else
            //{
            //    return false;
            //}     
            int value = DateTime.Compare(expirationDate, DateTime.Now);

            // checking 
            if (value > 0)
                return true;
            //Console.Write("expirationDate is later than date2. ");
            else /*if (value < 0)*/
                return false;
            //Console.Write("expirationDate is earlier than date2. ");
            //else
            //    Console.Write("expirationDate is the same as date2. ");

        }
    }
}
