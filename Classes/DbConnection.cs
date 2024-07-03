using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FirebirdSql.Data.FirebirdClient;

namespace SGmaster.Classes
{
    public static class FirebirdConnectionHelper
    {
        private static string connectionString = "DataSource=localhost;Database=D:\\Programação\\Projeto SGmaster\\SGmaster\\BD\\BASESGMASTERzerada.FDB;Port=3050;User=SYSDBA;Password=masterkey;Charset=UTF8;Dialect=3;Connection lifetime=15;PacketSize=8192;ServerType=0;Unicode=True;Max Pool Size=1000";
        private static FbConnection connection;

        public static FbConnection GetConnection()
        {
            if (connection == null)
            {
                connection = new FbConnection(connectionString);
                connection.Open();
            }
            else if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }

            return connection;
        }

        public static void CloseConnection()
        {
            if (connection != null && connection.State == ConnectionState.Open)
            {
                connection.Close();
            }
        }
    }
}
