using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SRO
{
    class Config
    {//строка подключения
        static public string connectionString = "";
        //ид клиента
        static public long id_client = 0; 
        //ид оператора
        static public long id_operator = 0;
        //телефон оператора
        static public string  phone_operator = "";
        
        static public bool isNewPhone = true;
    }
}
