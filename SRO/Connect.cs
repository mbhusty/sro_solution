using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.InteropServices;

namespace SRO
{
    public partial class Connect : Form
    {
        public Connect()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private void button2_Click(object sender, EventArgs e)
        {
            if (tbLogin.Text != "" && tbPassword.Text != "" && tbPhoneOper.Text != "")
            {
                string Login = "", Password = "";

                Login = tbLogin.Text;
                Password = tbPassword.Text;
                Config.phone_operator = tbPhoneOper.Text;

                Config.connectionString = @"Data Source=192.168.179.205, 49172; Initial Catalog = CC;
Network Library=DBMSSOCN; User ID=" + Login + ";Password=" + Password + ";";

                bool connectNorm = false;
                // int id_operator = 0;

                //проверка соединения
                connectNorm = WorkDB.isConnectNorm(Config.connectionString);

                if (connectNorm)
                {

                    SqlCommand scPhoneOper = new SqlCommand();
                    scPhoneOper.CommandType = CommandType.StoredProcedure;
                    scPhoneOper.CommandText = "SearchIdOpeartor";
                    scPhoneOper.Parameters.Add("@PhoneOper", SqlDbType.NVarChar);
                    scPhoneOper.Parameters["@PhoneOper"].Value = tbPhoneOper.Text;

                    Config.id_operator = WorkDB.searchId(scPhoneOper);

                    if (Config.id_operator != 0)
                    {
                        //запись в файл телефон..
                        
                        /*try
                        {
                            File.WriteAllText("config.ini", tbPhoneOper.Text);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка записи в файл: " + ex.Message);
                        }*/

                        //изменение агента
                        string s1 = "Не готов - Cisco Agent Desktop";


                        int width,height;
                        width = Screen.PrimaryScreen.Bounds.Width;
                        height = Screen.PrimaryScreen.Bounds.Height;

                        int agentHeight = 200;

                        IntPtr hWindow = FindWindow(null, s1);

                        MoveWindow(hWindow, 0, 0, width, agentHeight, true);

                        ShowWindow(hWindow, 6);
                        ShowWindow(hWindow, 1);
                        //
                        Hide();
                        SRO sro = new SRO();
                        sro.Width = width;
                        sro.Height = height - agentHeight * 2;

                        //sro.Location = new Point(0, 100);
                        sro.Owner = this;
                        sro.ShowDialog();
                        
                    }
                    else
                    {
                        MessageBox.Show("Телефон " + Config.phone_operator + " не зарегистрирован в программе \nОбратитесь к администратору");
                    }
                }
            }
            else
            {
                MessageBox.Show("Проверьте введенные данные");
            }
        }

        private void sqlConnection1_InfoMessage(object sender, System.Data.SqlClient.SqlInfoMessageEventArgs e)
        {
            
        }

        private void tbPhoneOper_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                button2_Click(sender, e);
            }
        }

        private void tbPhoneOper_TextChanged(object sender, EventArgs e)
        {
            //управление цветом
            if (tbPhoneOper.Text != "")
            {
                tbPhoneOper.BackColor = Color.White;
            }
            else
            {
                tbPhoneOper.BackColor = Color.Yellow;
            }
        }

        private void tbLogin_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                button2_Click(sender, e);
            }
        }

        private void tbPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                button2_Click(sender, e);
            }
        }

        private void Connect_Load(object sender, EventArgs e)
        {
           //Чтение Телефона из файла
            /*if (File.Exists("config.ini"))
            {
                string strPhone = File.ReadAllText("config.ini");
                tbPhoneOper.Text = strPhone;
            }*/
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

 
    }
}
