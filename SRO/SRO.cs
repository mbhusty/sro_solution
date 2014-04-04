using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace SRO
{
    public partial class SRO : Form
    {
        public SRO()
        {
            InitializeComponent();
        }

        private void ClearClient()
        {
            //Config.id_client = 0;
            


            tbSurname.Text = "";
            tbName.Text = "";
            tbOtchestvo.Text = "";
            //tbPhone.Text = "";//
            tbMail.Text = "";
            //cbRegions.Text = "";
            // cbWhereKnow.Text = "";
        }
        private void ClearHandling()
        {
            cbClientType.Text = "";
            cbRegions.Text = "";
            cbWhereKnow.Text = "";
            cbOperator_action.Text = "";
            //Controls.Add(label16);
            cbProducts.Text = "";
            cbSubProducts.Text = "";
            tbComment.Text = "";
        }

        private void SearchClient()
        {
            SqlCommand scSearchClient = new SqlCommand();
            scSearchClient.CommandType = CommandType.StoredProcedure;
            scSearchClient.CommandText = "SearchClient";
            scSearchClient.Parameters.Add("@ph_Mob", SqlDbType.NVarChar);
            scSearchClient.Parameters["@ph_Mob"].Value = tbPhone.Text;
            WorkDB.fillDataGridView(dgvSearchClient, scSearchClient);

            dgvSearchClient.ClearSelection();
            foreach (DataGridViewRow dgvr in dgvSearchClient.Rows)
            {
                if (Convert.ToInt64(dgvr.Cells[0].Value) == Config.id_client)
                {
                    dgvr.Selected = true;
                    break;
                }
            }
        }

        private void SearchHandling(long id_client = 1)
        {
            SqlCommand scSearchHandling = new SqlCommand();
            scSearchHandling.CommandType = CommandType.StoredProcedure;
            scSearchHandling.CommandText = "SearchHandling";
            scSearchHandling.Parameters.Add("@id_client", SqlDbType.BigInt);
            scSearchHandling.Parameters["@id_client"].Value = id_client;
            WorkDB.fillDataGridView(dgvHistoryHandling, scSearchHandling);
        }

        //копирование клиента из DataGridView в ТекстБоксы
        private void CopyClientFromDGV(int i)
        {
            tbSurname.Text = dgvSearchClient.Rows[i].Cells[1].Value.ToString();
            tbName.Text = dgvSearchClient.Rows[i].Cells[2].Value.ToString();
            tbOtchestvo.Text = dgvSearchClient.Rows[i].Cells[3].Value.ToString();
            //tbPhone.Text = dgvSearchClient.Rows[i].Cells[4].Value.ToString();
            tbMail.Text = dgvSearchClient.Rows[i].Cells[5].Value.ToString();
        }

        private void CopyHandlingFromDGV(int i)
        {
            if (dgvHistoryHandling.RowCount > 0)
            {
                cbClientType.Text = dgvHistoryHandling.Rows[i].Cells[3].Value.ToString();
                cbRegions.Text = dgvHistoryHandling.Rows[i].Cells[7].Value.ToString();
                cbWhereKnow.Text = dgvHistoryHandling.Rows[i].Cells[8].Value.ToString();
                
            }

        }

        private void SRO_Load(object sender, EventArgs e)
        {
            Top = 200;

            //Типы клиентов
            SqlCommand scAllClientType = new SqlCommand();
            scAllClientType.CommandType = CommandType.StoredProcedure;
            scAllClientType.CommandText = "Name_Client_Type";
            WorkDB.fillComboBox(cbClientType, scAllClientType);
            cbClientType.SelectedIndex = 0;

            //Действия оператора
            SqlCommand scAll_Operator_action = new SqlCommand();
            scAll_Operator_action.CommandType = CommandType.StoredProcedure;
            scAll_Operator_action.CommandText = "Name_Operator_action";
            WorkDB.fillComboBox(cbOperator_action, scAll_Operator_action);
            cbOperator_action.SelectedIndex = 0;

            //Регион
            SqlCommand scName_Region = new SqlCommand();
            scName_Region.CommandType = CommandType.StoredProcedure;
            scName_Region.CommandText = "Name_Regions";
            scName_Region.Parameters.Add("@isBank", SqlDbType.Bit);
            scName_Region.Parameters["@isBank"].Value = chbIsBank.Checked;
            WorkDB.fillComboBox(cbRegions, scName_Region);

            //Откуда узнали
            SqlCommand scName_Where = new SqlCommand();
            scName_Where.CommandType = CommandType.StoredProcedure;
            scName_Where.CommandText = "Name_Where";
            WorkDB.fillComboBox(cbWhereKnow, scName_Where);

            //время для отчетов
            dtpDateBeg.Value = DateTime.Now.Date;
            dtpDateEnd.Value = DateTime.Now.Date;

            //название формы
            this.Text = this.Text + " - " + Config.phone_operator;
        }



        private void SRO_FormClosed(object sender, FormClosedEventArgs e)
        {

            Owner.Visible = true;
        }

        private void cbClientType_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlCommand scName_Product = new SqlCommand();
            scName_Product.CommandType = CommandType.StoredProcedure;
            scName_Product.CommandText = "Name_Products";
            scName_Product.Parameters.Add("@type_client", SqlDbType.NVarChar);
            scName_Product.Parameters["@type_client"].Value = cbClientType.Text;
            WorkDB.fillComboBox(cbProducts, scName_Product);

            cbSubProducts.Items.Clear();
        }

        private void cbProducts_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlCommand scName_SubProduct = new SqlCommand();
            scName_SubProduct.CommandType = CommandType.StoredProcedure;
            scName_SubProduct.CommandText = "Name_SubProducts";
            scName_SubProduct.Parameters.Add("@Name_product", SqlDbType.NVarChar);
            scName_SubProduct.Parameters["@Name_product"].Value = cbProducts.Text;
            WorkDB.fillComboBox(cbSubProducts, scName_SubProduct);
            //Привязка выбранного пункта из COMBOBOX в  Label
            //label16.Text = (sender as ComboBox).Text;
            //button2.Text = (sender as ComboBox).Text;
            #region Кейсы для скрытия лэйблов
            switch ((string)cbProducts.SelectedItem)
            {

                case "Ипотека":
                    LabelGroup1Visible(true);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                    LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "Вклады":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(true);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                    LabelGroup8Visible(true);
                     LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "Адреса и номера телефонов":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(true);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                    LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "Перевод":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(true);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                    LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "Известная проблема":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(true);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                   LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "МСБ":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(true);
                    LabelGroup7Visible(false);
                   LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "Пластиковые карты":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(true);
                   LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "Чемпионат снегоход":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                   LabelGroup8Visible(true);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    
                    break;
                case "Информация о банке и руководстве":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                  LabelGroup8Visible(false);
                    LabelGroup9Visible(true);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    break;
                case "Курсы валют":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                   LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(true);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    break;
                case "Срочные денежные переводы":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                   LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(true);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    break;
                case "РКО":
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                   LabelGroup8Visible(false);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(true);
                    LabelGroup13Visible(false);
                                        break;
                    
                case "Потребительские кредиты":
                      LabelGroup1Visible(false);
                                        LabelGroup2Visible(false);
                                        LabelGroup3Visible(false);
                                        LabelGroup4Visible(false);
                                        LabelGroup5Visible(false);
                                        LabelGroup6Visible(false);
                                        LabelGroup7Visible(false);
                                        LabelGroup8Visible(false);
                                        LabelGroup9Visible(false);
                                        LabelGroup10Visible(false);
                                        LabelGroup11Visible(false);
                                        LabelGroup12Visible(false);
                                        LabelGroup13Visible(true);
                                        break;
                default:
                    LabelGroup1Visible(false);
                    LabelGroup2Visible(false);
                    LabelGroup3Visible(false);
                    LabelGroup4Visible(false);
                    LabelGroup5Visible(false);
                    LabelGroup6Visible(false);
                    LabelGroup7Visible(false);
                   LabelGroup8Visible(true);
                    LabelGroup9Visible(false);
                    LabelGroup10Visible(false);
                    LabelGroup11Visible(false);
                    LabelGroup12Visible(false);
                    LabelGroup13Visible(false);
                    break;
            #endregion

            }
        }

        void LabelGroup1Visible(bool visible)//Ипотека
        {
            linkLabel1.Visible = visible;
            linkLabel2.Visible = visible;
            linkLabel3.Visible = visible;
            linkLabel4.Visible = visible;
            linkLabel5.Visible = visible;
            linkLabel14.Visible = visible;
            linkLabel16.Visible = visible;
        }

        void LabelGroup2Visible(bool visible)//Вклады
        {
            linkLabel6.Visible = visible;
        }

        void LabelGroup3Visible(bool visible)//Адреса и номера телефонов
        {
            linkLabel8.Visible = visible;
            linkLabel25.Visible = visible;
        }
        void LabelGroup4Visible(bool visible)//Перевод
        {
            linkLabel9.Visible = visible;
        }
        void LabelGroup5Visible(bool visible)//Известная проблема
        {
            linkLabel10.Visible = visible;
        }
        void LabelGroup6Visible(bool visible)//мсб
        {
            linkLabel11.Visible = visible;
            linkLabel17.Visible = visible;
            linkLabel26.Visible = visible;
        }
        void LabelGroup7Visible(bool visible)//Пл.Карты
        {
            linkLabel12.Visible = visible;
            linkLabel13.Visible = visible;
            linkLabel30.Visible = visible;
        }

        void LabelGroup9Visible(bool visible)//о банке
        {
            linkLabel18.Visible = visible;
            linkLabel19.Visible = visible;
            linkLabel20.Visible = visible;
            linkLabel28.Visible = visible;
            linkLabel29.Visible = visible;
            linkLabel31.Visible = visible;
            linkLabel32.Visible = visible;
        }
        void LabelGroup10Visible(bool visible)//Валюта
        {
            linkLabel21.Visible = visible;
            linkLabel27.Visible = visible;
        }
        void LabelGroup11Visible(bool visible)//СДП
        {
            linkLabel22.Visible = visible;
        }
        void LabelGroup12Visible(bool visible)//РКО
        {
            linkLabel23.Visible = visible;
            linkLabel24.Visible = visible;
        }

        void LabelGroup8Visible(bool visible)//другое
        {
            linkLabel7.Visible = visible;
            linkLabel15.Visible = visible;
        }

        void LabelGroup13Visible(bool visible)//потреб.кр.
        {
            linkLabel33.Visible = visible;
        }

        private void chbIsBank_CheckedChanged(object sender, EventArgs e)
        {
            //Регион            
            SqlCommand scName_Region = new SqlCommand();
            scName_Region.CommandType = CommandType.StoredProcedure;
            scName_Region.CommandText = "Name_Regions";
            scName_Region.Parameters.Add("@isBank", SqlDbType.Bit);
            scName_Region.Parameters["@isBank"].Value = chbIsBank.Checked;
            WorkDB.fillComboBox(cbRegions, scName_Region);
        }

        private void butSaveClient_Click(object sender, EventArgs e)
        {

        }

        private void tbPhone_TextChanged(object sender, EventArgs e)
        {

            //управление цветом
            if (tbPhone.Text != "")
            {
                tbPhone.BackColor = Color.White;
                ////////////////////////////////////////////
                Config.id_client = 0;

                dgvHistoryHandling.Rows.Clear();
                //поиск клиента
                SearchClient();

                //если больше одного клиента
                if (dgvSearchClient.RowCount > 0)
                {
                    //то вероятней всего клиент нашего банка, следовательно убираем галочку
                    chbNewClient.Checked = false;

                    //если клиент найден один
                    if (dgvSearchClient.RowCount == 1)
                    {
                        dgvSearchClient.Rows[0].Selected = true;
                        Config.id_client = Convert.ToInt64(dgvSearchClient.Rows[0].Cells[0].Value);
                        CopyClientFromDGV(0);
                        //ищем обращения, так как клиент один и он выбран
                        SearchHandling(Config.id_client);
                        //если обращений больше 0, то копируем самое верхнее
                        if (dgvHistoryHandling.RowCount > 0)
                        {
                            CopyHandlingFromDGV(0);
                        }
                    }

                }
                //клиентов не найдено, то скорее всего это новый клиент
                else
                {
                    Config.id_client = 0;
                    Config.isNewPhone = true;
                    chbNewClient.Checked = true;
                    ClearHandling();
                }
                ///////////////////////////////////////////
                tbPhone.BackColor = Color.White;
            }
            else
            {
                tbPhone.BackColor = Color.Yellow;
            }
        }



        private void butNewClient_Click(object sender, EventArgs e)
        {
            dgvSearchClient.Rows.Clear();
            dgvHistoryHandling.Rows.Clear();
            ClearClient();
            cbClientType.SelectedIndex = 0;
            cbRegions.SelectedIndex = -1;
            cbWhereKnow.SelectedIndex = -1;
            cbOperator_action.SelectedIndex = -1;
            cbProducts.SelectedIndex = -1;
            cbSubProducts.SelectedIndex = -1;
                    //          //          //         //
            tbPhone.Text = "";
            ClearHandling();
            chbNewClient.Checked = true;
            Config.isNewPhone = true;

        }

        private void dgvSearchClient_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            if (i >= 0)
            {
                //2 раза нажали на ячейку и клиент 1
                Config.id_client = Convert.ToInt64(dgvSearchClient.Rows[i].Cells[0].Value);

                dgvSearchClient.ClearSelection();
                dgvSearchClient.Rows[i].Selected = true;

                SearchHandling(Config.id_client);

                ClearClient();
                CopyClientFromDGV(i);

                ClearHandling();
                CopyHandlingFromDGV(0);
            }
        }

        private void butHandClear_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Вы действительно хотите очистить данные об обращении клиента", "Подтверждение очистки",
                MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                ClearHandling();
                cbClientType.SelectedIndex = -1;
                cbRegions.SelectedIndex = -1;
                cbWhereKnow.SelectedIndex = -1;
                cbOperator_action.SelectedIndex = -1;
                cbProducts.SelectedIndex = -1;
                cbSubProducts.SelectedIndex = -1;
            }

        }

        private void butHandAdd_Click(object sender, EventArgs e)
        {
            if (tbPhone.Text != "")
            {
                //если клиент не определен или клиенты не найдены и галочка "Новый клиент" не стоит
                if ((Config.id_client == 0 || dgvSearchClient.Rows.Count == 0) && chbNewClient.Checked == false)
                {
                    MessageBox.Show("Выберите клиента, либо поставьте галочку новый клиент");
                }
                else
                {
                    //если стоит галочка "Новый клиент, то добавляем клиента
                    if (chbNewClient.Checked == true)
                    {
                        //Добавление клиента     

                        SqlCommand scInsertClient = new SqlCommand();
                        scInsertClient.CommandType = CommandType.StoredProcedure;
                        scInsertClient.CommandText = "InsertClient";
                        scInsertClient.Parameters.Add("@Name", SqlDbType.NVarChar);
                        scInsertClient.Parameters["@Name"].Value = tbName.Text;
                        scInsertClient.Parameters.Add("@Surname", SqlDbType.NVarChar);
                        scInsertClient.Parameters["@Surname"].Value = tbSurname.Text;
                        scInsertClient.Parameters.Add("@Otchestvo", SqlDbType.NVarChar);
                        scInsertClient.Parameters["@Otchestvo"].Value = tbOtchestvo.Text;
                        scInsertClient.Parameters.Add("@ph_Mob", SqlDbType.NVarChar);
                        scInsertClient.Parameters["@ph_Mob"].Value = tbPhone.Text;
                        scInsertClient.Parameters.Add("@ph_Home", SqlDbType.NVarChar);
                        scInsertClient.Parameters["@ph_Home"].Value = null;
                        scInsertClient.Parameters.Add("@ph_Work", SqlDbType.NVarChar);
                        scInsertClient.Parameters["@ph_Work"].Value = null;
                        scInsertClient.Parameters.Add("@Mail", SqlDbType.NVarChar);
                        scInsertClient.Parameters["@Mail"].Value = tbMail.Text;

                        Config.id_client = WorkDB.insertANDid(scInsertClient);
                        chbNewClient.Checked = false;
                    }
                    //номер обращения
                    int n1 = 1, n2 = 1;
                    string N;

                    /*if (Config.isNewPhone)
                        n2 = 1;*/

                    if (dgvHistoryHandling.Rows.Count > 0)
                    {
                        string str1, str2;
                        string[] strN;

                        strN = dgvHistoryHandling.Rows[0].Cells[1].Value.ToString().Split('.');
                        str1 = strN[0];
                        str2 = strN[1];

                        n1 = Convert.ToInt32(str1);
                        n2 = Convert.ToInt32(str2);

                        if (Config.isNewPhone)
                        { n1++; n2 = 1; }
                        else
                            n2++;

                    }
                    N = n1.ToString() + "." + n2.ToString();

                    if (chbNewClient.Checked == true)
                    {
                        N = "1.1";
                    }

                    //Добавление Обращения           
                    SqlCommand scInsertHandling = new SqlCommand();
                    scInsertHandling.CommandType = CommandType.StoredProcedure;
                    scInsertHandling.CommandText = "InsertHandling";
                    scInsertHandling.Parameters.Add("@N", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@N"].Value = N;
                    scInsertHandling.Parameters.Add("@id_client", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@id_client"].Value = Config.id_client;
                    scInsertHandling.Parameters.Add("@Client_type", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@Client_type"].Value = cbClientType.Text;
                    scInsertHandling.Parameters.Add("@product", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@product"].Value = cbProducts.Text;
                    scInsertHandling.Parameters.Add("@subproduct", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@subproduct"].Value = cbSubProducts.Text;
                    scInsertHandling.Parameters.Add("@action", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@action"].Value = cbOperator_action.Text;
                    scInsertHandling.Parameters.Add("@region", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@region"].Value = cbRegions.Text;
                    scInsertHandling.Parameters.Add("@where", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@where"].Value = cbWhereKnow.Text;
                    scInsertHandling.Parameters.Add("@id_operator", SqlDbType.Int);
                    scInsertHandling.Parameters["@id_operator"].Value = Config.id_operator;
                    scInsertHandling.Parameters.Add("@Comment", SqlDbType.NVarChar);
                    scInsertHandling.Parameters["@Comment"].Value = tbComment.Text + " - СКП";

                    WorkDB.insert(scInsertHandling);

                    SearchClient();
                    SearchHandling(Config.id_client);

                    //устанавливаем, что регистрируем в рамках одного звонка
                    Config.isNewPhone = false;

                    //стираем тему, подтему
                }
            }
            else
            {
                MessageBox.Show("Заполните поле телефон");
                tbPhone.Focus();
            }
        }



        private void tbPhone_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyData == Keys.Enter)
            {
                if (tbPhone.Text != "")
                {
                    tbPhone.BackColor = Color.White;

                    Config.id_client = 0;

                    dgvHistoryHandling.Rows.Clear();
                    //поиск клиента
                    SearchClient();

                    //если больше одного клиента
                    if (dgvSearchClient.RowCount > 0)
                    {
                        //то вероятней всего клиент нашего банка, следовательно убираем галочку
                        chbNewClient.Checked = false;

                        //если клиент найден один
                        if (dgvSearchClient.RowCount == 1)
                        {
                            dgvSearchClient.Rows[0].Selected = true;
                            Config.id_client = Convert.ToInt64(dgvSearchClient.Rows[0].Cells[0].Value);
                            CopyClientFromDGV(0);
                            //ищем обращения, так как клиент один и он выбран
                            SearchHandling(Config.id_client);
                            //если обращений больше 0, то копируем самое верхнее
                            if (dgvHistoryHandling.RowCount > 0)
                            {
                                CopyHandlingFromDGV(0);
                            }
                        }

                    }
                    //клиентов не найдено, то скорее всего это новый клиент
                    else
                    {
                        Config.id_client = 0;
                        Config.isNewPhone = true;
                        chbNewClient.Checked = true;
                        ClearHandling();
                    }
                }
                else
                {
                    tbPhone.BackColor = Color.Yellow;
                }
            }
        }

        /* private void button1_Click(object sender, EventArgs e)
         {
             string url = "www.itb.ru";
             DateTime now = DateTime.Now.Date;

             switch (cbRequest.Text)
             {
                 case "Потреб":
                     {
                         url = "http://www.itb.ru/bitrix/admin/form_result_edit.php?login=yes&lang=ru&WEB_FORM_ID=31";

                         webBrowser1.Navigate(url);

                       
                         tcHandlings.SelectTab(1);
                         if (webBrowser1.)
                         {
                             webBrowser1.Document.GetElementById("form_date_4105").InnerText = now.Date.ToShortDateString();
                             // webBrowser1.Document.GetElementById("form_dropdown_CC_STATUS").InnerText = "Направлено в Филиал/ДО";
                             webBrowser1.Document.GetElementById("form_text_3309").InnerText = tbSurname.Text;
                             webBrowser1.Document.GetElementById("form_text_3310").InnerText = tbName.Text;
                             webBrowser1.Document.GetElementById("form_text_3311").InnerText = tbOtchestvo.Text;
                             webBrowser1.Document.GetElementById("form_text_3319").InnerText = tbPhone.Text;
                         }
                        
                         break;
                     }
                 case "Ипотека":
                     {
                         url = "http://www.itb.ru/bitrix/admin/form_result_list.php?lang=ru&WEB_FORM_ID=29";
                         break;
                     }

                 case "МСБ":
                     {
                         url = "http://www.itb.ru/bitrix/admin/form_result_list.php?lang=ru&WEB_FORM_ID=15";
                         break;
                     }

                 case "Претензия":
                     {
                         url = "http://192.168.179.210/skp/clients/";
                         break;
                     }
             }




         }*/

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage.Text == "Отчеты")
            {
                tcClients.Dock = DockStyle.Fill;
                tcHandlings.Visible = false;
                dgvHistoryHandling.Visible = false;

            }
            if (e.TabPage.Text == "Данные по клиенту")
            {
                tcClients.Dock = DockStyle.Top;
                tcHandlings.Visible = true;
                dgvHistoryHandling.Visible = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime dt_begin, dt_end;
            dt_begin = dtpDateBeg.Value.Date + dtpTimeBeg.Value.TimeOfDay;
            dt_end = dtpDateEnd.Value.Date + dtpTimeEnd.Value.TimeOfDay;

            /*SqlCommand scSearchHandlingDate = new SqlCommand();
            scSearchHandlingDate.CommandType = CommandType.StoredProcedure;
            scSearchHandlingDate.CommandText = "SearchHandlingDate";
            scSearchHandlingDate.Parameters.Add("@begin_date", SqlDbType.DateTime);
            scSearchHandlingDate.Parameters["@begin_date"].Value = dt_begin;
            scSearchHandlingDate.Parameters.Add("@end_date", SqlDbType.DateTime);
            scSearchHandlingDate.Parameters["@end_date"].Value = dt_end;
            WorkDB.fillDataGridView(dgvReports, scSearchHandlingDate);*/

            SqlCommand scSearchHandlingDate = new SqlCommand();
            scSearchHandlingDate.CommandType = CommandType.StoredProcedure;
            scSearchHandlingDate.CommandText = "SearchHandlingDateTest";
            scSearchHandlingDate.Parameters.Add("@begin_date", SqlDbType.DateTime);
            scSearchHandlingDate.Parameters["@begin_date"].Value = dt_begin;
            scSearchHandlingDate.Parameters.Add("@end_date", SqlDbType.DateTime);
            scSearchHandlingDate.Parameters["@end_date"].Value = dt_end;

            if (chbOper.Checked == true)
            {
            scSearchHandlingDate.Parameters.Add("@id_oper", SqlDbType.Int);

                scSearchHandlingDate.Parameters["@id_oper"].Value = Config.id_operator;
            }
            WorkDB.fillDataGridView(dgvReports, scSearchHandlingDate);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
           // Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
           //ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelApp.Columns.ColumnWidth = 15;

            //выравнивание
            //ExcelApp.Range["A1", "C2"].HorizontalAlignment = XlHAlign.xlHAlignCenter;


            //добавление названия столбцов
            for (int j = 0; j < dgvReports.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j + 1] = dgvReports.Columns[j].HeaderText;
            }

            //заполнение ячеек
            for (int i = 0; i < dgvReports.Rows.Count; i++)
            {
                for (int j = 0; j < dgvReports.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dgvReports.Rows[i].Cells[j].Value;
                }
            }




            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }


        private void cbRegions_TextChanged(object sender, EventArgs e)
        {

        }

        private void cbRegions_SelectedIndexChanged(object sender, EventArgs e)
        {
            //cbRegions.AutoCompleteMode = AutoCompleteMode.Suggest;
        }

        /*private void button4_Click(object sender, EventArgs e)
        {DateTime now = DateTime.Now;
            webBrowser1.Document.GetElementById("form_date_4105").InnerText = now.Date.ToShortDateString();
            // webBrowser1.Document.GetElementById("form_dropdown_CC_STATUS").InnerText = "Направлено в Филиал/ДО";
            webBrowser1.Document.GetElementById("form_text_3309").InnerText = tbSurname.Text;
            webBrowser1.Document.GetElementById("form_text_3310").InnerText = tbName.Text;
            webBrowser1.Document.GetElementById("form_text_3311").InnerText = tbOtchestvo.Text;
            webBrowser1.Document.GetElementById("form_text_3319").InnerText = tbPhone.Text;
        }*/

        private void tcHandlings_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage.Text == "Заявка")
            {
                tcHandlings.Dock = DockStyle.Fill;
                dgvHistoryHandling.Visible = false;
                /* tcHandlings.Visible = false;
                 dgvHistoryHandling.Visible = false;*/

            }
            if (e.TabPage.Text == "Регистрация обращения клиента")
            {
                //dgvHistoryHandling.Visible = true;
                tcHandlings.Dock = DockStyle.Top;/*
                tcHandlings.Visible = true;*/
                dgvHistoryHandling.Visible = true;
            }
        }

        private void cbRequest_SelectedIndexChanged(object sender, EventArgs e)
        {
            string url = "www.itb.ru";
            try
            {
                switch (cbRequest.Text)
                {
                    case "Потреб":
                        {
                            url = "http://www.itb.ru/bitrix/admin/form_result_edit.php?login=yes&lang=ru&WEB_FORM_ID=31";
                            break;
                        }
                    case "Ипотека":
                        {
                            url = "http://www.itb.ru/bitrix/admin/form_result_edit.php?lang=ru&WEB_FORM_ID=29";
                            break;
                        }

                    case "МСБ":
                        {
                            url = "http://www.itb.ru/bitrix/admin/form_result_edit.php?lang=ru&WEB_FORM_ID=15";
                            break;
                        }

                    case "Претензия":
                        {
                            url = "http://192.168.179.210/skp/clients/";
                            break;
                        }

                } webBrowser1.Navigate(url);
                butRequest.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        //вставка в браузер
        private void butRequest_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now.Date;
            try
            {
                switch (cbRequest.Text)
                {
                    case "Потреб":
                        {
                            webBrowser1.Document.GetElementById("form_date_4105").InnerText = now.Date.ToShortDateString();
                            webBrowser1.Document.GetElementById("form_text_3309").InnerText = tbSurname.Text;
                            webBrowser1.Document.GetElementById("form_text_3310").InnerText = tbName.Text;
                            webBrowser1.Document.GetElementById("form_text_3311").InnerText = tbOtchestvo.Text;
                            webBrowser1.Document.GetElementById("form_text_3319").InnerText = tbPhone.Text;
                            webBrowser1.Document.GetElementById("form_email_3320").InnerText = tbMail.Text;
                            break;
                        }
                    case "Ипотека":
                        {
                            webBrowser1.Document.GetElementById("form_date_4104").InnerText = now.Date.ToShortDateString();
                            webBrowser1.Document.GetElementById("form_text_3185").InnerText = tbSurname.Text;
                            webBrowser1.Document.GetElementById("form_text_3186").InnerText = tbName.Text;
                            webBrowser1.Document.GetElementById("form_text_3187").InnerText = tbOtchestvo.Text;
                            webBrowser1.Document.GetElementById("form_text_3195").InnerText = tbPhone.Text;
                            webBrowser1.Document.GetElementById("form_email_3196").InnerText = tbMail.Text;
                            break;
                        }

                    case "МСБ":
                        {
                            //webBrowser1.Document.GetElementById("form_date_4090").InnerText = now.Date.ToShortDateString();
                            webBrowser1.Document.GetElementById("form_text_1428").InnerText =
                                tbSurname.Text + " " + tbName.Text + " " + tbOtchestvo.Text;
                            webBrowser1.Document.GetElementById("form_text_1430").InnerText = tbPhone.Text;
                            webBrowser1.Document.GetElementById("form_email_1431").InnerText = tbMail.Text;
                            break;
                        }

                    case "Претензия":
                        {
                            //webBrowser1.Document.GetElementById("form_date_LVcYrboU").InnerText = now.Date.ToString();
                            webBrowser1.Document.GetElementById("test").InnerText =
                                tbSurname.Text + " " + tbName.Text + " " + tbOtchestvo.Text;
                            webBrowser1.Document.GetElementById("form_text_73").InnerText = tbPhone.Text;
                            webBrowser1.Document.GetElementById("form_text_108").InnerText = tbMail.Text;
                            webBrowser1.Document.GetElementById("form_textarea_75").InnerText = tbComment.Text;
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //количество зарегистрированных вызовов
        private void button1_Click(object sender, EventArgs e)
        {
            DateTime dt_begin, dt_end;
            dt_begin = dtpDateBeg.Value.Date + dtpTimeBeg.Value.TimeOfDay;
            dt_end = dtpDateEnd.Value.Date + dtpTimeEnd.Value.TimeOfDay;

            SqlCommand scOperRegCount = new SqlCommand();
            scOperRegCount.CommandType = CommandType.StoredProcedure;
            scOperRegCount.CommandText = "OperRegCount";
            scOperRegCount.Parameters.Add("@begin_date", SqlDbType.DateTime);
            scOperRegCount.Parameters["@begin_date"].Value = dt_begin;
            scOperRegCount.Parameters.Add("@end_date", SqlDbType.DateTime);
            scOperRegCount.Parameters["@end_date"].Value = dt_end;
            //scOperRegCount.Parameters.Add("@action", SqlDbType.NVarChar);
            //scOperRegCount.Parameters["@action"].Value = cbOperator_action.Text;
            WorkDB.fillDataGridView(dgvReports, scOperRegCount);
        }

        private void butUpdateClient_Click(object sender, EventArgs e)
        {
            if (Config.id_client != 0)
            {
                if (MessageBox.Show("Вы действительно хотите ОБНОВИТЬ информацию о клиенте", "Подтверждение обновления",
                MessageBoxButtons.OKCancel) == DialogResult.OK)
                {


                    //Обновление информации о клиенте          
                    SqlCommand scUpdateClient = new SqlCommand();
                    scUpdateClient.CommandType = CommandType.StoredProcedure;
                    scUpdateClient.CommandText = "UpdateClient";
                    scUpdateClient.Parameters.Add("@id_client", SqlDbType.BigInt);
                    scUpdateClient.Parameters["@id_client"].Value = Config.id_client;
                    scUpdateClient.Parameters.Add("@Surname", SqlDbType.NVarChar);
                    scUpdateClient.Parameters["@Surname"].Value = tbSurname.Text;
                    scUpdateClient.Parameters.Add("@Name", SqlDbType.NVarChar);
                    scUpdateClient.Parameters["@Name"].Value = tbName.Text;
                    scUpdateClient.Parameters.Add("@Otchestvo", SqlDbType.NVarChar);
                    scUpdateClient.Parameters["@Otchestvo"].Value = tbOtchestvo.Text;
                    scUpdateClient.Parameters.Add("@Mail", SqlDbType.NVarChar);
                    scUpdateClient.Parameters["@Mail"].Value = tbMail.Text;

                    WorkDB.update(scUpdateClient);

                    SearchClient();
                }
            }

            else
            {
                MessageBox.Show("Выберите клиента, информацию о котором хотите обновить");
            }
        }

        #region Ненужная часть
        //private void label16_Click(object sender, EventArgs e)
        //{

        //    // testcode
        //    //MessageBox.Show(cbProducts.Items[cbProducts.SelectedIndex].ToString());
        //    //System.Diagnostics.Process.Start(@"D:\script.docx");
        //    //cbProducts.Text = (sender as ComboBox).Text;

        //}

        //private void button2_Click_1(object sender, EventArgs e)
        //{
            //try
            //{
            //    string str = @"\\cc001\Soft\scr"; //указываем адрес вплоть до полного имени
            //    str += cbProducts.SelectedIndex;
            //    str += ".doc";
            //  //StreamReader Wrr = new StreamReader(str);//не открывает
            //    System.Diagnostics.Process.Start(str);
            //}
            //catch (Win32Exception)
            //{
            //    MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

        //}

        //private void label18_Click(object sender, EventArgs e)
        //{

        //}

        //private void label16_Click(object sender, EventArgs e)  //Обработка открытия документа по двойному клику на Label16
        //{
        // System.Diagnostics.Process.Start(@"D:\scr0.docx");
        //}
        #endregion
        #region Открытие файлов при выборе пункта Ипотека
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Ипотека\Лица сопровождающие закладные.xlsx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Ипотека\Наш банк выкупил закладную.doc");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Ипотека\Наш банк продал закладную.docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Ипотека\Скрипт Ипотека+Карта.docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

            }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Ипотека\спец программа ипотеки Горки 8.docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void linkLabel14_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Ипотека\Материнский капитал.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        #endregion
        #region при выборе пункта Вклады
        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Вклады\заказ дс по вкладу.docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        #endregion
        #region при выборе Адреса и филиалы
        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Адреса и филиалы\Как добраться до офисов.xlsx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

         }
        #endregion
        #region при выборе Перевод на сотрудника
        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Перевод на сотрудника\перевод на кл ОМРК.docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        #endregion
        #region при выборе Известная проблема
        private void linkLabel10_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Известная проблема\Скрипт по Банку Москвы.docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
        #region при выборе МСБ
        private void linkLabel11_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\МСБ\преимущества.docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void linkLabel17_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\МСБ\ПРИГЛАШЕНИЕ НА СЕМИНАР МСБ.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
        #region при выбор СЧК Карты
        private void linkLabel12_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\СЧК Карты\Инструкция Visa Infinite.pdf");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void linkLabel13_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
            System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\СЧК Карты\памятка виза инфинит (2).docx");
            }
            catch (Win32Exception)
            {
            MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
       
       #region Открытие файлов
        private void cbOperator_action_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel16_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Ипотека\Этапы оформления правособственности.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel18_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Информация о банке и руковдстве\О БАНКЕ.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel19_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Информация о банке и руковдстве\Рейтинги Банка.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel20_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Информация о банке и руковдстве\Скрипт по вхождению ИТБ.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel21_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Валюта\Возможные операции с валютой.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel22_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\СДП\тариф КИРГИЗИЯ.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel23_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\РКО\Разъяснения по РКО.pdf");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel24_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\РКО\Разъяснения по тарифам на РКО физ лиц.doc");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion 

        private void button2_Click_1(object sender, EventArgs e)
        {
            //DateTime dt_begin, dt_end;
            //dt_begin = dtpDateBeg.Value.Date + dtpTimeBeg.Value.TimeOfDay;
            //dt_end = dtpDateEnd.Value.Date + dtpTimeEnd.Value.TimeOfDay;

            //SqlCommand scOperRegCount = new SqlCommand();
            //scOperRegCount.CommandType = CommandType.StoredProcedure;
            //scOperRegCount.CommandText = "OperRegCount";
            //scOperRegCount.Parameters.Add("@begin_date", SqlDbType.DateTime);
            //scOperRegCount.Parameters["@begin_date"].Value = dt_begin;
            //scOperRegCount.Parameters.Add("@end_date", SqlDbType.DateTime);
            //scOperRegCount.Parameters["@end_date"].Value = dt_end;
            //scOperRegCount.Parameters.Add("@action", SqlDbType.NVarChar);
            //scOperRegCount.Parameters["@action"].Value = cbOperator_action.Text;
            //WorkDB.fillDataGridView(dgvReports, scOperRegCount);

        }

        private void linkLabel25_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Адреса и филиалы\Новые офисы адреса.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click_2(object sender, EventArgs e)
        {

            DateTime dt_begin, dt_end;
            dt_begin = dtpDateBeg.Value.Date + dtpTimeBeg.Value.TimeOfDay;
            dt_end = dtpDateEnd.Value.Date + dtpTimeEnd.Value.TimeOfDay;

            SqlCommand scOperWebCount = new SqlCommand();
            scOperWebCount.CommandType = CommandType.StoredProcedure;
            scOperWebCount.CommandText = "OperWebCount";
            scOperWebCount.Parameters.Add("@begin_date", SqlDbType.DateTime);
            scOperWebCount.Parameters["@begin_date"].Value = dt_begin;
            scOperWebCount.Parameters.Add("@end_date", SqlDbType.DateTime);
            scOperWebCount.Parameters["@end_date"].Value = dt_end;
            //scOperRegCount.Parameters.Add("@action", SqlDbType.NVarChar);
            //scOperRegCount.Parameters["@action"].Value = cbOperator_action.Text;
            WorkDB.fillDataGridView(dgvReports, scOperWebCount);
        }

        private void linkLabel26_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\МСБ\Акция МСБ Компании продаж.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel27_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Валюта\Скрипт вклад Срочный.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel28_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Информация о банке и руковдстве\Рейтинг банка_по продуктам.pdf");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel29_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Информация о банке и руковдстве\Негативная информация о банке.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel30_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\СЧК Карты\Активные СМС-сервисы.doc");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Вклады\Акция для пенсионеров.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel15_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Вклады\Акция по вкладам с 01.04.2014.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel32_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Информация о банке и руковдстве\Негативная информация о Банке от РА Эксперт.pdf");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel31_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Информация о банке и руковдстве\Совет Директоров и Правление.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void linkLabel33_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"\\CC-SERV\Users\Administrator\Documents\01-Fileserver\Scripts\Потреб. кредиты\ТП Корпоративный.docx");
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Ничего не выбрано или\nнеобходимый файл отсутствует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
