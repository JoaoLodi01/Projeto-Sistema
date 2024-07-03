using FirebirdSql.Data.FirebirdClient;
using OfficeOpenXml;
using Serilog;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SGmaster.Classes
{
    public class DataHelper
    {
        private FbConnection connection;

        public DataHelper(FbConnection connection)
        {
            this.connection = connection;
        }

        public async Task<DataTable> GetDataAsync(string module)
        {
            DataTable dataTable = new DataTable();

            try
            {
                string query = GetQueryForModule(module);

                using (FbCommand command = new FbCommand(query, connection))
                {
                    using (FbDataAdapter adapter = new FbDataAdapter(command))
                    {
                        await Task.Run(() =>
                        {
                            adapter.Fill(dataTable);
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Erro ao obter dados do módulo {Module}", module);
                throw;
            }

            return dataTable;
        }

        

        private string GetQueryForModule(string module)
        {
            switch (module.ToLower())
            {
                case "clientes":
                    return @"SELECT FIRST 100 * FROM TCLIENTE WHERE ativo = 'SIM' ORDER BY controle ASC";
                case "estoque":
                    return @"SELECT FIRST 100 * FROM TESTOQUE ORDER BY codigo ASC";
                case "fornecedores":
                    return @"SELECT FIRST 100 * FROM TFORNECEDOR ORDER BY codfornecedor ASC";
                default:
                    throw new ArgumentException("Módulo não suportado para consulta de dados.");
            }
        }

        public async Task<DataTable> ImportData(string filePath)
        {
            Log.Information("Iniciando importação de dados do arquivo {FilePath}", filePath);
            DataTable dataTable = new DataTable();

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        throw new Exception("O arquivo não contém planilhas.");
                    }

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    if (worksheet.Dimension == null || worksheet.Dimension.End.Row < 1)
                    {
                        throw new Exception("A planilha está vazia ou não tem dimensões definidas.");
                    }

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string columnName = worksheet.Cells[1, col].Value?.ToString();
                        if (!string.IsNullOrEmpty(columnName))
                        {
                            dataTable.Columns.Add(columnName);
                        }
                    }

                    for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                    {
                        DataRow row = dataTable.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            string columnName = worksheet.Cells[1, col].Value?.ToString();
                            var cellValue = worksheet.Cells[rowNum, col].Value;
                            row[columnName] = cellValue ?? DBNull.Value;
                        }
                        dataTable.Rows.Add(row);
                    }
                }
                Log.Information("Importação de dados concluída com sucesso.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Erro ao importar dados do arquivo {FilePath}", filePath);
                throw;
            }

            return dataTable;
        }
        public async Task ImportToDatabaseAsync(string filePath, string module)
        {
            FbConnection connection = null;

            Log.Information("Iniciando importação de dados do arquivo {FilePath}", filePath);
            DataTable dataTable = await ImportData(filePath);
            Log.Information("Importação de dados concluída com sucesso. Total de colunas importadas: {ColumnCount}", dataTable.Columns.Count);

            connection = FirebirdConnectionHelper.GetConnection();

            if (connection.State != ConnectionState.Open)
            {
                await connection.OpenAsync();
            }

            foreach (DataRow row in dataTable.Rows)
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = GetInsertCommandText(module);
                    AddParameters(command, row, module);
                    await command.ExecuteNonQueryAsync();
                }
            }
        
            /*finally
            {
                if (connection != null && connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }*/
        }

        private string GetInsertCommandText(string module)
        {
            switch (module.ToLower())
            {
                case "clientes":
                    return @"INSERT INTO TCLIENTE (
                           CLIENTE, ENDERECO, COMPLEMENTO, BAIRRO, CIDADE, UF, CEP, 
                           RG, CPF, CNPJ, IE,TELEFONE, CELULAR, EMAIL,FANTASIA,NUMERO
                           ) VALUES (
                           @CLIENTE, @ENDERECO, @COMPLEMENTO, @BAIRRO, @CIDADE, @UF, @CEP, 
                           @RG, @CPF, @CNPJ, @IE, @TELEFONE, @CELULAR, @EMAIL,@FANTASIA, @NUMERO
                           )";
                case "estoque":
                    return @"INSERT INTO TESTOQUE (
                           PRODUTO, CODBARRAS, PRECOCUSTO, PRECOVENDA,
                           DATAULTIMACOMPRA, DATAULTIMAVENDA, DATAHORACADASTRO, 
                           QTDE, ATIVO, NCM, CODTRIBUTACAOIPI, CODTRIBUTACAOPIS, CODTRIBUTACAOCOFINS
                           ) VALUES (
                           @PRODUTO, @CODBARRAS, @PRECOCUSTO, @PRECOVENDA,
                           @DATAULTIMACOMPRA, @DATAULTIMAVENDA, @DATAHORACADASTRO, 
                           @QTDE, @ATIVO, @NCM, @CODTRIBUTACAOIPI, @CODTRIBUTACAOPIS, @CODTRIBUTACAOCOFINS
                           )";
                case "fornecedores":
                    return @"INSERT INTO TFORNECEDOR (
                           CLIENTE, ENDERECO, COMPLEMENTO, BAIRRO, CIDADE, UF, CEP, 
                           RG, CPF, CNPJ, IE,TELEFONE, CELULAR, EMAIL,FANTASIA,NUMERO
                           ) VALUES (
                           @CLIENTE, @ENDERECO, @COMPLEMENTO, @BAIRRO, @CIDADE, @UF, @CEP, 
                           @RG, @CPF, @CNPJ, @IE, @TELEFONE, @CELULAR, @EMAIL,@FANTASIA, @NUMERO
                           )";
                default:
                    throw new ArgumentException("Módulo não suportado para inserção de dados.");
            }
        }

        private async Task AddParameters(FbCommand command, DataRow row, string module)
        {
            DateTime dataHoraCadastro = DateTime.Now;
            
                command.Parameters.Clear(); // Limpar parâmetros antes de configurar novos

            switch (module.ToLower())
            {
                case "clientes":
                    string queryClientes = @"INSERT INTO TCLIENTE (
                       CLIENTE, ENDERECO, COMPLEMENTO, BAIRRO, CIDADE, UF, CEP, 
                       RG, CPF, CNPJ, IE, TELEFONE, CELULAR, EMAIL, FANTASIA, NUMERO, 
                       DATAHORACADASTRO, ATIVO
                       ) VALUES (
                       @CLIENTE, @ENDERECO, @COMPLEMENTO, @BAIRRO, @CIDADE, @UF, @CEP, 
                       1234567, @CPF, @CNPJ, @IE, @TELEFONE, @CELULAR, @EMAIL, @FANTASIA, @NUMERO, 
                       @DATAHORACADASTRO, 'SIM'
                       )";
                    command.CommandText = queryClientes;

                    command.Parameters.AddWithValue("@CLIENTE", row["Cliente/Fornecedor"].ToString());
                    command.Parameters.AddWithValue("@FANTASIA", row["Fantasia"].ToString());

                    string cpf = row["CPF"].ToString();
                    string cnpj = row["CNPJ"].ToString();
                    string rg = row["RG/IE"].ToString();

                    if (!string.IsNullOrEmpty(cpf) && cpf.Length == 11)
                    {
                        string formattedCpf = $"{cpf.Substring(0, 3)}.{cpf.Substring(3, 3)}.{cpf.Substring(6, 3)}-{cpf.Substring(9, 2)}";
                        command.Parameters.AddWithValue("@CPF", formattedCpf);
                        //command.Parameters.AddWithValue("@RG", row["RG/IE"].ToString());
                        command.Parameters.AddWithValue("@CNPJ", DBNull.Value);
                        command.Parameters.AddWithValue("@IE", DBNull.Value);
                    }
                    else if (!string.IsNullOrEmpty(cnpj) && cnpj.Length == 14)
                    {
                        string formattedCnpj = $"{cnpj.Substring(0, 2)}.{cnpj.Substring(2, 3)}.{cnpj.Substring(5, 3)}/{cnpj.Substring(8, 4)}-{cnpj.Substring(12, 2)}";
                        command.Parameters.AddWithValue("@CNPJ", formattedCnpj);
                        //command.Parameters.AddWithValue("@IE", row["RG/IE"].ToString());
                        command.Parameters.AddWithValue("@CPF", DBNull.Value);
                        command.Parameters.AddWithValue("@RG", DBNull.Value);
                    }

                    command.Parameters.AddWithValue("@ENDERECO", row["Endereço"].ToString());
                    command.Parameters.AddWithValue("@COMPLEMENTO", row["Complemento"].ToString());
                    command.Parameters.AddWithValue("@BAIRRO", row["Bairro"].ToString());
                    command.Parameters.AddWithValue("@CEP", row["CEP"].ToString());
                    command.Parameters.AddWithValue("@NUMERO", row["Número"].ToString());
                    command.Parameters.AddWithValue("@CIDADE", row["Município"].ToString());
                    command.Parameters.AddWithValue("@UF", row["UF"].ToString());
                    command.Parameters.AddWithValue("@TELEFONE", row["Telefone/Celular"].ToString());
                    command.Parameters.AddWithValue("@CELULAR", row["Telefone/Celular"].ToString());
                    command.Parameters.AddWithValue("@EMAIL", row["E-mail"].ToString());
                    command.Parameters.AddWithValue("@DATAHORACADASTRO", dataHoraCadastro);
                    break;

                case "fornecedores":
                    string queryFornecedores = @"INSERT INTO TFORNECEDOR (
                           RAZAOSOCIAL, ENDERECO, COMPLEMENTO, BAIRRO, CIDADE, UF, CEP, 
                           RG, CPF, CNPJ, IE, TELEFONE, CELULAR, EMAIL, NOMEFANTASIA, NUMERO, 
                           DATAHORACADASTRO, ATIVO
                           ) VALUES (
                           @RAZAOSOCIAL, @ENDERECO, @COMPLEMENTO, @BAIRRO, @CIDADE, @UF, @CEP, 
                           @RG, @CPF, @CNPJ, @IE, @TELEFONE, @CELULAR, @EMAIL, @NOMEFANTASIA, @NUMERO, 
                           @DATAHORACADASTRO, 'SIM'
                           )";
                    command.CommandText = queryFornecedores;

                    command.Parameters.AddWithValue("@RAZAOSOCIAL", row["Cliente/Fornecedor"].ToString());
                    command.Parameters.AddWithValue("@NOMEFANTASIA", row["Fantasia"].ToString());

                    string cpfFornecedor = row["CPF"].ToString();
                    string cnpjFornecedor = row["CNPJ"].ToString();

                    if (!string.IsNullOrEmpty(cpfFornecedor) && cpfFornecedor.Length == 11)
                    {
                        command.Parameters.AddWithValue("@CPF", cpfFornecedor);
                        command.Parameters.AddWithValue("@RG", row["RG/IE"].ToString());
                        command.Parameters.AddWithValue("@CNPJ", DBNull.Value);
                        command.Parameters.AddWithValue("@IE", DBNull.Value);
                    }
                    else if (!string.IsNullOrEmpty(cnpjFornecedor) && cnpjFornecedor.Length == 14)
                    {
                        command.Parameters.AddWithValue("@CNPJ", cnpjFornecedor);
                        command.Parameters.AddWithValue("@IE", row["RG/IE"].ToString());
                        command.Parameters.AddWithValue("@CPF", DBNull.Value);
                        command.Parameters.AddWithValue("@RG", DBNull.Value);
                    }
                    
                    command.Parameters.AddWithValue("@ENDERECO", row["Endereço"].ToString());
                    command.Parameters.AddWithValue("@COMPLEMENTO", row["Complemento"].ToString());
                    command.Parameters.AddWithValue("@BAIRRO", row["Bairro"].ToString());
                    command.Parameters.AddWithValue("@CEP", row["CEP"].ToString());
                    command.Parameters.AddWithValue("@NUMERO", row["Número"].ToString());
                    command.Parameters.AddWithValue("@CIDADE", row["Município"].ToString());
                    command.Parameters.AddWithValue("@UF", row["UF"].ToString());
                    command.Parameters.AddWithValue("@TELEFONE", row["Telefone/Celular"].ToString());
                    command.Parameters.AddWithValue("@CELULAR", row["Telefone/Celular"].ToString());
                    command.Parameters.AddWithValue("@EMAIL", row["E-mail"].ToString());
                    command.Parameters.AddWithValue("@DATAHORACADASTRO", dataHoraCadastro);
                    break;

                case "estoque":
                    string queryEstoque = @"INSERT INTO TESTOQUE (
                           PRODUTO, CODBARRAS, PRECOCUSTO, PRECOVENDA, CSOSN, ALIQUOTAICMSECF, OBS,
                           DATAULTIMACOMPRA, DATAULTIMAVENDA, DATAHORACADASTRO, 
                           QTDE, ATIVO, NCM, CODTRIBUTACAOIPI, CODTRIBUTACAOPIS, CODTRIBUTACAOCOFINS
                           ) VALUES (
                           @PRODUTO, @CODBARRAS, @PRECOCUSTO, @PRECOVENDA, @CSOSN, @ALIQUOTAICMSECF, @OBS,
                           @DATAULTIMACOMPRA, @DATAULTIMAVENDA, @DATAHORACADASTRO, 
                           @QTDE, 'SIM', @NCM, @CODTRIBUTACAOIPI, @CODTRIBUTACAOPIS, @CODTRIBUTACAOCOFINS
                           )";
                    command.CommandText = queryEstoque;

                    command.Parameters.AddWithValue("@PRODUTO", row["Descrição"].ToString());
                    command.Parameters.AddWithValue("@CODBARRAS", row["Código de Barras"].ToString());
                    command.Parameters.AddWithValue("@PRECOCUSTO", Convert.ToDecimal(row["Preço de Custo"]));
                    command.Parameters.AddWithValue("@PRECOVENDA", Convert.ToDecimal(row["Preço de Venda"]));
                    command.Parameters.AddWithValue("@CSOSN", row["CSOSN/CST"].ToString());
                    command.Parameters.AddWithValue("@ALIQUOTAICMSECF", Convert.ToDecimal(row["Alíquota ECF"]));
                    command.Parameters.AddWithValue("@OBS", row["Referência"].ToString());
                    command.Parameters.AddWithValue("@DATAULTIMACOMPRA", DBNull.Value);
                    command.Parameters.AddWithValue("@DATAULTIMAVENDA", DBNull.Value);
                    command.Parameters.AddWithValue("@DATAHORACADASTRO", dataHoraCadastro);
                    command.Parameters.AddWithValue("@QTDE", Convert.ToInt32(row["Qtde inicial"]));
                    command.Parameters.AddWithValue("@NCM", row["NCM"].ToString());
                    command.Parameters.AddWithValue("@CODTRIBUTACAOIPI", row["CST IPI"].ToString());
                    command.Parameters.AddWithValue("@CODTRIBUTACAOPIS", row["CST PIS"].ToString());
                    command.Parameters.AddWithValue("@CODTRIBUTACAOCOFINS", row["CST COFINS"].ToString());
                    break;

                default:
                    MessageBox.Show("Não foi possível realizar a importação.", "Erro de Importação", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }            
        }
    }
}
