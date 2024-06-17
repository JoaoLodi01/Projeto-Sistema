using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGmaster
{
    public class Conexao
    {
        private FbConnection _conexao;

        private string _stringConexao = "Database=D:\\Programação\\Projeto SGmaster\\SGmaster\\BASESGMASTERpopulada.FDB;DataSource=localhost;Port=3050;User=SYSDBA;Password=masterkey;Dialect=3;Charset=UTF8;Pooling=true;MinPoolSize=0;MaxPoolSize=50;ConnectionLifeTime=15;";

        public Conexao()
        {
            _conexao = new FbConnection(_stringConexao);
        }

        public DataTable ExecutarConsulta(string query)
        {
            DataTable dataTable = new DataTable();

            try
            {
                _conexao.Open();

                FbDataAdapter dataAdapter = new FbDataAdapter(query, _conexao);
                dataAdapter.Fill(dataTable);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro ao executar consulta: " + ex.Message);
                throw;
            }
            finally
            {
                _conexao.Close();
            }

            return dataTable;
        }
    }
}
