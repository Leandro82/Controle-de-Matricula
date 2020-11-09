using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Matricula
{
    public partial class frm_Principal : Form
    {
        public frm_Principal()
        {
            InitializeComponent();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            
        }

        private void cadastrarResultadoToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void listas1ªOpçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_Lista();
            if (Application.OpenForms.OfType<frm_Lista>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void listas2ªOpçãoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var peq = new frm_ListaSegOp();
            if (Application.OpenForms.OfType<frm_ListaSegOp>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void matriculadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_Matriculados();
            if (Application.OpenForms.OfType<frm_Matriculados>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void listãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_Listao();
            if (Application.OpenForms.OfType<frm_Listao>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void pesquisaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_Pesquisa();
            if (Application.OpenForms.OfType<frm_Pesquisa>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void resultadoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new Form1();
            if (Application.OpenForms.OfType<Form1>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void vagasChamadasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_DadosVest();
            if (Application.OpenForms.OfType<frm_DadosVest>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void listasFábioEmailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_Email();
            if (Application.OpenForms.OfType<frm_Email>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void estatísticasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_Estatisticas();
            if (Application.OpenForms.OfType<frm_Estatisticas>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void ocultarVestibulinhosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_Vestibulinhos();
            if (Application.OpenForms.OfType<frm_Vestibulinhos>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void nãoMatriculadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var peq = new frm_NaoMatriculado();
            if (Application.OpenForms.OfType<frm_NaoMatriculado>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }
    }
}
