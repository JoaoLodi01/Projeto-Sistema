using FirebirdSql.Data.FirebirdClient;
using SGmaster.Designers;
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
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
            AlignForm("top-left");

        }

        private void AlignForm(string alignment)
        {
            this.StartPosition = FormStartPosition.Manual;
            Rectangle screen = Screen.PrimaryScreen.WorkingArea;

            switch (alignment.ToLower())
            {
                case "top-left":
                    this.Location = new Point(0, 0);
                    break;
                case "top-right":
                    this.Location = new Point(screen.Width - this.Width, 0);
                    break;
                case "bottom-left":
                    this.Location = new Point(0, screen.Height - this.Height);
                    break;
                case "bottom-right":
                    this.Location = new Point(screen.Width - this.Width, screen.Height - this.Height);
                    break;
                case "center":
                    this.Location = new Point((screen.Width - this.Width) / 2, (screen.Height - this.Height) / 2);
                    break;
                case "left":
                    this.Location = new Point(0, (screen.Height - this.Height) / 2);
                    break;
                case "right":
                    this.Location = new Point(screen.Width - this.Width, (screen.Height - this.Height) / 2);
                    break;
                case "top":
                    this.Location = new Point((screen.Width - this.Width) / 2, 0);
                    break;
                case "bottom":
                    this.Location = new Point((screen.Width - this.Width) / 2, screen.Height - this.Height);
                    break;
                default:
                    throw new ArgumentException("Invalid alignment specified");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ClientesForm Clientes = new ClientesForm();
            Clientes.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EstoqueForm Estoque = new EstoqueForm();
            Estoque.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FornecedoresForm Fornecedores = new FornecedoresForm();
            Fornecedores.Show();
        }

        private void clienteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClientesForm Clientes = new ClientesForm();
            Clientes.Show();
        }

        private void estoqueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EstoqueForm Estoque = new EstoqueForm();
            Estoque.Show();
        }

        private void fornecedoresToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FornecedoresForm Fornecedores = new FornecedoresForm();
            Fornecedores.Show();
        }
    }
}

