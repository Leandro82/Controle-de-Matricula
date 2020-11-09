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
    public partial class frm_Vestibulinhos : Form
    {
        ConectaVagas cv = new ConectaVagas();
        Vagas vg = new Vagas();
        public frm_Vestibulinhos()
        {
            InitializeComponent();
        }

        private void frm_Vestibulinhos_Load(object sender, EventArgs e)
        {
            foreach (DataRow linha in cv.VestibulinhoTodos().Rows)
            {
                int n = dataGridView1.Rows.Add();
                if (linha["ocultar"].ToString() == "Ok")
                {
                    dataGridView1.Rows[n].Cells[0].Value = true;
                }
                dataGridView1.Rows[n].Cells[1].Value = linha["vestibulinho"].ToString();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns[0].Index)
            {
                dataGridView1.EndEdit();  //Stop editing of cell.
                string vest = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                if ((bool)dataGridView1.Rows[e.RowIndex].Cells[0].Value)
                {
                    vg.Vestibulinho = vest;
                    vg.Ocultar = "Ok";
                    cv.ocultarVestibulinho(vg);
                    string msg = "OS DADOS DO VESTIBULINHO "+vest+" FORAM OCULTOS";
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                }
                else
                {
                    vg.Vestibulinho = vest;
                    vg.Ocultar = "";
                    cv.ocultarVestibulinho(vg);
                    string msg = "VOCÊ PODERÁ VER OS DADOS DO VESTIBULINHO " + vest + " NOVAMENTE";
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                }
            }
        }
    }
}
