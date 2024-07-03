using FirebirdSql.Data.FirebirdClient;
using OfficeOpenXml;
using SGmaster.Classes;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace Sistema_SGmaster
{
    public partial class FornecedoresForm : Form
    {
        private DataHelper dataHelper;
        private FbConnection connection;

        public FornecedoresForm()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;

            connection = FirebirdConnectionHelper.GetConnection();
            dataHelper = new DataHelper(connection);

            dataGridView1.ContextMenuStrip = contextMenuStrip1;
        }

        private async void FornecedoresForm_Load(object sender, EventArgs e)
        {
            await CarregarDadosAsync();
        }

        private async Task CarregarDadosAsync()
        {
            string query = $@"SELECT FIRST 100 *
                              FROM TFORNECEDOR 
                              WHERE ativo = 'SIM' 
                              ORDER BY controle ASC";

            try
            {
                using (FbConnection connection = FirebirdConnectionHelper.GetConnection())
                {
                    using (FbCommand command = new FbCommand(query, connection))
                    {
                        using (FbDataReader reader = await command.ExecuteReaderAsync())
                        {
                            DataTable dataTable = new DataTable();
                            dataTable.Load(reader);
                            dataGridView1.Invoke((Action)(() =>
                            {
                                dataGridView1.DataSource = dataTable;
                            }));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar dados: " + ex.Message);
            }
            finally
            {
                FirebirdConnectionHelper.CloseConnection();
            }
        }
        private async void importarPlanilhaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Selecionar arquivo Excel";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    try
                    {
                        string module = "fornecedores";
                        DataTable importedDataTable = await dataHelper.ImportData(filePath);

                        await dataHelper.ImportToDatabaseAsync(filePath, module);

                        await CarregarDadosAsync();

                        MessageBox.Show("Dados importados com sucesso.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao importar dados: " + ex.Message);
                    }
                }
            }
        }
        private void exportarParaPlanilhaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dataTable)
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Salvar como Excel";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcel(dataTable, saveFileDialog.FileName);
                        MessageBox.Show("Dados exportados com sucesso.");
                    }
                }
            }
        }

        public void ExportToExcel(DataTable dataTable, string filePath)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Data");

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }

                FileInfo fileInfo = new FileInfo(filePath);
                package.SaveAs(fileInfo);
            }
        }
    }
}

