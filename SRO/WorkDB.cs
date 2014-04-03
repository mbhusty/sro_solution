using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace SRO
{
    class WorkDB
    {
        //процедура проверки соединения с базой
        static public bool isConnectNorm(string ConnectionString)
        {
            SqlConnection sConn = new SqlConnection(ConnectionString);
            try
            {
                sConn.Open();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nОбратитесь к администратору");
                return false;
            }
            finally
            {
                sConn.Close();
            }
        }


        //процедура проверки соединения с базой
        static public int searchId(SqlCommand sc)
        {
            int id = 0;
            SqlConnection sConn = new SqlConnection(Config.connectionString);
            try
            {
                sConn.Open();
                sc.Connection = sConn;

                SqlDataReader dr = sc.ExecuteReader();
                while (dr.Read())
                {
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        id = Convert.ToInt32(dr.GetValue(i));
                    }

                }
                dr.Close();
                return id;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return id;
            }
            finally
            {
                sConn.Close();
            }
        }


        /// <summary>
        /// заполнение комббокса из базы данных
        /// </summary>
        /// <param name="cb">комбобокс, который заполняем</param>
        /// <param name="sc">в этом компоненте запрос или хранимая процедура с параметрами</param>
        static public void fillComboBox(ComboBox cb, SqlCommand sc)
        {
            SqlConnection sConn = new SqlConnection(Config.connectionString);
            try
            {
                cb.Text = "";
                cb.Items.Clear();
                sConn.Open();

                sc.Connection = sConn;

                SqlDataReader dr = sc.ExecuteReader();
                while (dr.Read())
                {
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        cb.Items.Add(dr.GetValue(i));
                    }
                }
                dr.Close();

                //cb.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sConn.Close();
            }
        }



        /// <summary>
        /// заполнение DataGridView из базы данных
        /// </summary>
        /// <param name="cb">DataGridView, который заполняем</param>
        /// <param name="sc">в этом компоненте запрос или хранимая процедура с параметрами</param>
        static public void fillDataGridView(DataGridView dgv, SqlCommand sc)
        {
            SqlConnection sConn = new SqlConnection(Config.connectionString);
            try
            {
                dgv.Rows.Clear();
                sConn.Open();

                sc.Connection = sConn;

                SqlDataReader dr = sc.ExecuteReader();

                dgv.Columns.Clear();

                //столбцы
                for (int i = 0; i < dr.FieldCount; i++)
                {
                    dgv.Columns.Add(dr.GetName(i), dr.GetName(i));
                }

                while (dr.Read())
                {  
                    //сами данны в строчки
                    string[] str = new string[dr.FieldCount];

                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        str[i] = dr.GetValue(i).ToString();
                        //[] ii = dr.GetValue(i);
                    }

                    dgv.Rows.Add(str);

                    //    dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.ColumnHeader);

                }
                dr.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sConn.Close();
            }
        }

        //добавление записи в таблицу с помощью хранимой процедуры
        static public void insert(SqlCommand sc)
        {
            SqlConnection sConn = new SqlConnection(Config.connectionString);
            try
            {
                sConn.Open();
                sc.Connection = sConn;

                int i = sc.ExecuteNonQuery();

                //MessageBox.Show("Запись успешно добавлена " + i); ////////
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sConn.Close();
            }
        }

        //обновление записи в таблице с помощью хранимой процедуры
        static public void update(SqlCommand sc)
        {
            SqlConnection sConn = new SqlConnection(Config.connectionString);
            try
            {
                sConn.Open();
                sc.Connection = sConn;

                int i = sc.ExecuteNonQuery();

                //MessageBox.Show("Запись успешно добавлена " + i); ////////
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                sConn.Close();
            }
        }

        //добавляет строчку в базу и выбирает id данной записи
        static public int insertANDid(SqlCommand sc)
        {
            int id_client = 0;
            SqlConnection sConn = new SqlConnection(Config.connectionString);
            try
            {
                sConn.Open();
                sc.Connection = sConn;

                SqlDataReader dr = sc.ExecuteReader();
                while (dr.Read())
                {

                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        id_client = dr.GetInt32(i);
                        //[] ii = dr.GetValue(i);
                    }



                }
                dr.Close();
                return id_client;
                //MessageBox.Show("Запись успешно добавлена " + i); ////////
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return id_client;
            }
            finally
            {
                sConn.Close();

            }
        }
    }
}
