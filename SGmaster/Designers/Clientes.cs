using FirebirdSql.Data.FirebirdClient;
using SGmaster;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sistema_SGmaster
{
    public partial class Clientes : Form
    {

        private Conexao conexao = new Conexao();

        public Clientes()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;
        }

        private void Clientes_Load(object sender, EventArgs e)
        {
            string query = "SELECT * FROM TCLIENTES";

            DataTable dataTable = conexao.ExecutarConsulta(query);

            dataGridView1.DataSource = dataTable;
        }

    }
}


