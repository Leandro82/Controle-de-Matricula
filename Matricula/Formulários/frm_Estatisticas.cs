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
    public partial class frm_Estatisticas : Form
    {
        ConectaVagas cv = new ConectaVagas();
        ConectaGeral cg = new ConectaGeral();
        ConectaChamada ch = new ConectaChamada();
        Vagas vg = new Vagas();
        int masc = 0, fem = 0, total = 0;
        public frm_Estatisticas()
        {
            InitializeComponent();
        }

        private void frm_Estatisticas_Load(object sender, EventArgs e)
        {
            int vest = cv.Vestibulinho().Rows.Count;
            for (int i = 0; i < vest; i++)
            {
                comboBox1.Items.Add(cv.Vestibulinho().Rows[i]["vestibulinho"].ToString());
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            masc = 0;
            fem = 0;
            total = 0;
            vg.Vestibulinho = comboBox1.Text;
            dataGridView1.Rows.Clear();
            foreach (DataRow item in cv.CursoCad(vg).Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = item["HABILITACAO"].ToString() +" - " + item["ESCOLA"].ToString();
                vg.Curso = item["HABILITACAO"].ToString();
                vg.Periodo = item["PERIODO"].ToString();
                vg.Escola = item["ESCOLA"].ToString();
                dataGridView1.Rows[n].Cells[1].Value = ch.TotalMasculino(vg).Rows[0][0].ToString();
                masc = Convert.ToInt32(ch.TotalMasculino(vg).Rows[0][0].ToString()) + masc;
                dataGridView1.Rows[n].Cells[2].Value = ch.TotalFeminino(vg).Rows[0][0].ToString();
                fem = Convert.ToInt32(ch.TotalFeminino(vg).Rows[0][0].ToString()) + fem;
                dataGridView1.Rows[n].Cells[3].Value = cg.SelecionaMatriculados(vg).Rows.Count;
                total = cg.SelecionaMatriculados(vg).Rows.Count + total;
            }
            int p = dataGridView1.Rows.Add();
            dataGridView1.Rows[p].Cells[0].Value = "TOTAL";
            dataGridView1.Rows[p].Cells[1].Value = masc;
            dataGridView1.Rows[p].Cells[2].Value = fem;
            dataGridView1.Rows[p].Cells[3].Value = total;
        }
    }
}
