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
    public partial class frm_Pesquisa : Form
    {
        ConectaVagas cv = new ConectaVagas();
        ConectaChamada co = new ConectaChamada();
        ConectaGeral cg = new ConectaGeral();
        Geral gr = new Geral();
        Vagas vg = new Vagas();

        public frm_Pesquisa()
        {
            InitializeComponent();
        }

        public static int Idade(DateTime dtNascimento)
        {
            int idade = DateTime.Now.Year - dtNascimento.Year;
            if (DateTime.Now.Month < dtNascimento.Month || (DateTime.Now.Month == dtNascimento.Month && DateTime.Now.Day < dtNascimento.Day))
                idade--;
            return idade;
        }

        public void Pesquisa()
        {
            vg.Vestibulinho = comboBox1.Text;
            vg.Periodo = comboBox4.Text;
            vg.Nome = textBox1.Text;
            dataGridView1.Rows.Clear();
            if (comboBox3.Text == "")
            {
                string msg = "INFORME O TIPO DE PESQUISA";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
                comboBox1.Focus();
            }
            else
            {
                if (comboBox3.Text == "Curso")
                {
                    if (comboBox1.Text == "")
                    {
                        string msg = "INFORME O SEMESTRE E O ANO DO VESTIBULINHO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                        comboBox1.Focus();
                    }
                    else if (comboBox2.Text == "")
                    {
                        string msg = "INFORME O CURSO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                        comboBox2.Focus();
                    }
                    else if (comboBox4.Text == "")
                    {
                        string msg = "INFORME O PERÍODO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                        comboBox4.Focus();
                    }
                    else
                    {
                        string curso;
                        curso = comboBox2.Text;
                        vg.Curso = curso.Remove(curso.Length - 10);
                        vg.Escola = comboBox2.Text.Substring(comboBox2.Text.Length - 7);
                        foreach (DataRow item in co.Pesquisa(vg).Rows)
                        {
                            if (item["CLAS"].ToString() != "")
                            {
                                int n = dataGridView1.Rows.Add();
                                dataGridView1.Rows[n].Cells[0].Value = item["CLAS"].ToString();
                                dataGridView1.Rows[n].Cells[1].Value = item["NOTA"].ToString();
                                dataGridView1.Rows[n].Cells[2].Value = item["NOME"].ToString();
                                dataGridView1.Rows[n].Cells[3].Value = item["ENDERECO"].ToString();
                                dataGridView1.Rows[n].Cells[4].Value = item["TELEFONE"].ToString();
                                dataGridView1.Rows[n].Cells[5].Value = item["CELULAR"].ToString();
                                dataGridView1.Rows[n].Cells[6].Value = item["HABILITACAO"].ToString();
                                int idade = Idade(Convert.ToDateTime(item["dtNasc"].ToString()));
                                if ((item["HABILITACAO"].ToString() == "ENSINO MÉDIO" || item["HABILITACAO"].ToString() == "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" || item["HABILITACAO"].ToString().Contains("NOVOTEC") == true) && (idade <= 13))
                                {
                                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                                }
                                else if ((item["HABILITACAO"].ToString() != "ENSINO MÉDIO" && item["HABILITACAO"].ToString() != "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" && item["HABILITACAO"].ToString().Contains("NOVOTEC") == false) && (idade <= 14))
                                {
                                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                                }
                                dataGridView1.Rows[n].Cells[7].Value = item["ESCOL"].ToString();
                                if (item["matriculado"].ToString() == "Sim")
                                {
                                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.GreenYellow;
                                }
                                if (item["chamada"].ToString() == "2ª Opção" && item["matriculado"].ToString() == "Sim")
                                {
                                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Orange;
                                }
                            }
                        }
                    }
                }
                else if (comboBox3.Text == "Nome")
                {
                    if (comboBox1.Text == "")
                    {
                        string msg = "INFORME O SEMESTRE E O ANO DO VESTIBULINHO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                        comboBox1.Focus();
                    }
                    else
                    {
                        foreach (DataRow item in co.ListaoPorNome(vg).Rows)
                        {
                            int n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = item["CLAS"].ToString();
                            dataGridView1.Rows[n].Cells[1].Value = item["NOTA"].ToString();
                            dataGridView1.Rows[n].Cells[2].Value = item["NOME"].ToString();
                            dataGridView1.Rows[n].Cells[3].Value = item["ENDERECO"].ToString();
                            dataGridView1.Rows[n].Cells[4].Value = item["TELEFONE"].ToString();
                            dataGridView1.Rows[n].Cells[5].Value = item["CELULAR"].ToString();
                            dataGridView1.Rows[n].Cells[6].Value = item["HABILITACAO"].ToString();
                            int idade = Idade(Convert.ToDateTime(item["dtNasc"].ToString()));
                            if ((item["HABILITACAO"].ToString() == "ENSINO MÉDIO" || item["HABILITACAO"].ToString() == "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" || item["HABILITACAO"].ToString().Contains("NOVOTEC") == true) && (idade <= 13))
                            {
                                dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                            }
                            else if ((item["HABILITACAO"].ToString() != "ENSINO MÉDIO" && item["HABILITACAO"].ToString() != "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" && item["HABILITACAO"].ToString().Contains("NOVOTEC") == false) && (idade <= 14))
                            {
                                dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                            }
                            dataGridView1.Rows[n].Cells[7].Value = item["ESCOL"].ToString();
                            if (item["matriculado"].ToString() == "Sim")
                            {
                                dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.GreenYellow;
                            }
                            if (item["chamada"].ToString() == "2ª Opção" && item["matriculado"].ToString() == "Sim")
                            {
                                dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Orange;
                            }
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Pesquisa();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text == "Curso")
            {
                textBox1.Text = "";
                textBox1.Enabled = false;
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
                comboBox4.Enabled = true;
            }
            else
            {
                comboBox2.Text = "";
                comboBox4.Text = "";
                comboBox2.Enabled = false;
                comboBox4.Enabled = false;
                textBox1.Enabled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            vg.Vestibulinho = comboBox1.Text;
            int curso = cv.SelecionaCurso(vg).Rows.Count;
            int cham = cv.SelecionaChamada(vg).Rows.Count;
            comboBox2.Items.Clear();
            for (int i = 0; i < curso; i++)
            {
                comboBox2.Items.Add(cv.SelecionaCurso(vg).Rows[i]["curso"].ToString() + " - " + cv.SelecionaCurso(vg).Rows[i]["escola"].ToString());
            }

            comboBox4.Text = "";
            comboBox4.Items.Clear();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string curso;
            vg.Vestibulinho = comboBox1.Text;
            curso = this.comboBox2.Text;
            vg.Curso = curso.Remove(curso.Length - 10);
            int periodo = cv.SelecionaVagas(vg).Rows.Count;
            comboBox4.Items.Clear();
            comboBox4.Text = "";
            for (int i = 0; i < periodo; i++)
            {
                comboBox4.Items.Add(cv.SelecionaVagas(vg).Rows[i]["PERIODO"].ToString());
            }
        }

        private void frm_Pesquisa_Load(object sender, EventArgs e)
        {
            int vest = cv.Vestibulinho().Rows.Count;
            for (int i = 0; i < vest; i++)
            {
                comboBox1.Items.Add(cv.Vestibulinho().Rows[i]["vestibulinho"].ToString());
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                Pesquisa();
            }
        }
    }
}
