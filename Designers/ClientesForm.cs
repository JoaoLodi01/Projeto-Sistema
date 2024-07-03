using FirebirdSql.Data.FirebirdClient;
using OfficeOpenXml;
using SGmaster.Classes;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sistema_SGmaster
{
    public partial class ClientesForm : Form
    {
        private DataHelper dataHelper;
        private FbConnection connection;

        public ClientesForm()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;

            connection = FirebirdConnectionHelper.GetConnection();
            dataHelper = new DataHelper(connection);

            dataGridView1.ContextMenuStrip = contextMenuStrip1;
        }

        private async void ClientesForm_Load(object sender, EventArgs e)
        {
            await CarregarDadosAsync();
        }

        private async Task CarregarDadosAsync()
        {
            try
            {
                string module = "clientes";
                DataTable dataTable = await dataHelper.GetDataAsync(module);

                dataGridView1.Invoke((Action)(() =>
                {
                    dataGridView1.DataSource = dataTable;
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar dados: " + ex.Message);
            }
        }

        private async void importarToolStripMenuItem_Click(object sender, EventArgs e)
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
                        string module = "clientes";
                        DataTable importedDataTable = await dataHelper.ImportData(filePath);

                        await dataHelper.ImportToDatabaseAsync(filePath, module);

                        await CarregarDadosAsync();

                        MessageBox.Show("Dados importados com sucesso.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao importar dados: " + ex.Message, "Erro de Importação", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        private void exportarToolStripMenuItem_Click(object sender, EventArgs e)
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
