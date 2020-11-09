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
    public partial class frm_DadosVest : Form
    {
        ConectaGeral cg = new ConectaGeral();
        ConectaVagas cv = new ConectaVagas();
        Vagas vg = new Vagas();
        Geral gr = new Geral();
        string aux;

        public frm_DadosVest()
        {
            InitializeComponent();
        }

        private void frm_DadosVest_Load(object sender, EventArgs e)
        {
            button3.Enabled = false;
            int vest = cv.Vestibulinho().Rows.Count;
            for (int i = 0; i < vest; i++)
            {
                comboBox1.Items.Add(cv.Vestibulinho().Rows[i]["vestibulinho"].ToString());
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            vg.Vestibulinho = comboBox1.Text;
            int quant = cv.CursoVag(vg).Rows.Count;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            if (quant == 0)
            {
                foreach (DataRow item in cv.CursoCad(vg).Rows)
                {
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = false;
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[1].Value = item["HABILITACAO"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["PERIODO"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["ESCOLA"].ToString();
                }
            }
            else
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = true;
                foreach (DataRow item in cv.CursoVag(vg).Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = item["COD"].GetHashCode();
                    dataGridView1.Rows[n].Cells[1].Value = item["CURSO"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["PERIODO"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["ESCOLA"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["VAGAS"].ToString();
                }

                foreach (DataRow item in cv.Chamada(vg).Rows)
                {
                    int n = dataGridView2.Rows.Add();
                    dataGridView2.Rows[n].Cells[0].Value = item["COD"].GetHashCode();
                    dataGridView2.Rows[n].Cells[1].Value = item["CHAMADA"].ToString();
                    dataGridView2.Rows[n].Cells[2].Value = Convert.ToDateTime(item["dtInicial"].ToString()).ToString("dd/MM/yyyy");
                    dataGridView2.Rows[n].Cells[3].Value = Convert.ToDateTime(item["dtFinal"].ToString()).ToString("dd/MM/yyyy");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int totalDias = (DateTime.Parse(dateTimePicker2.Value.ToString("dd/MM/yyyy")).Subtract(DateTime.Parse(dateTimePicker1.Value.ToString("dd/MM/yyyy")))).Days;

            if (totalDias == 0)
            {
                string msg = "DATA INICIAL IGUAL A DATA FINAL";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
            }
            else if (totalDias < 0)
            {
                string msg = "DATA INICIAL MAIOR QUE A DATA FINAL";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
            }
            else
            {
                int n = dataGridView2.Rows.Add();
                dataGridView2.Rows[n].Cells[1].Value = ((n + 1) + "ª chamada").ToString();
                dataGridView2.Rows[n].Cells[2].Value = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                dataGridView2.Rows[n].Cells[3].Value = dateTimePicker2.Value.ToString("dd/MM/yyyy");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int dt1 = dataGridView1.Rows.Count;
            int dt2 = dataGridView2.Rows.Count;

            if (dt1 == 0 || dt2 == 0)
            {
                string msg = "VERIFIQUE AS VAGAS OU AS CHAMADAS";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
            }
            else
            {
                for (int i = 0; i < dt1; i++)
                {
                    if (String.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[4].Value))
                    {
                        aux = "Ok";
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                    else
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                    }
                }

                if (aux == "Ok")
                {
                    string msg = "INFORME A QUANTIDADE DE VAGAS NOS CAMPOS AMARELOS";
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                    aux = "";
                }
                else
                {
                    for (int i = 0; i < dt1; i++)
                    {
                        vg.Curso = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        vg.Periodo = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        vg.Escola = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        vg.Vaga = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        vg.Vestibulinho = comboBox1.Text;
                        cv.cadastroVagas(vg);
                    }

                    for (int i = 0; i < dt2; i++)
                    {
                        vg.Chamada = dataGridView2.Rows[i].Cells[1].Value.ToString();
                        vg.DtInicio = Convert.ToDateTime(dataGridView2.Rows[i].Cells[2].Value.ToString());
                        vg.DtFim = Convert.ToDateTime(dataGridView2.Rows[i].Cells[3].Value.ToString());
                        vg.Vestibulinho = comboBox1.Text;
                        cv.cadastroDatas(vg);
                    }
                    string msg = "DADOS CADASTRADOS COM SUCESSO!!";
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int vagas = dataGridView1.Rows.Count;
            int cham = dataGridView2.Rows.Count;

            for (int i = 0; i < vagas; i++)
            {
                vg.Codigo = dataGridView1.Rows[i].Cells[0].Value.GetHashCode();
                vg.Vaga = dataGridView1.Rows[i].Cells[4].Value.ToString();
                cv.atualizarVagas(vg);
            }

            for (int j = 0; j < cham; j++)
            {
                vg.Codigo = dataGridView2.Rows[j].Cells[0].Value.GetHashCode();
                vg.DtInicio = Convert.ToDateTime(dataGridView2.Rows[j].Cells[2].Value.ToString());
                vg.DtFim = Convert.ToDateTime(dataGridView2.Rows[j].Cells[3].Value.ToString());
                cv.atualizarChamadas(vg);
            }
            string msg = "DADOS ALTERADOS COM SUCESSO!!";
            frm_Mensagem mg = new frm_Mensagem(msg);
            mg.ShowDialog();
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }
    }
}
    

