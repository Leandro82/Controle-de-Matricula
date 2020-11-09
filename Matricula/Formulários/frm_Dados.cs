using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Matricula
{
    public partial class Form1 : Form
    {
        string caminho, bt, unidade, vest, aux;
        Geral gr = new Geral();
        ConectaGeral cg = new ConectaGeral();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(delegate()
            {
                Carrega();
            }).Start();

            
        }

        public void Carrega()
        {
             System.Threading.Thread arquivo1 = new System.Threading.Thread(new System.Threading.ThreadStart(() =>
             {
                 OpenFileDialog arquivo = new OpenFileDialog();

            if (arquivo.ShowDialog() == DialogResult.OK)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    caminho = arquivo.FileName;

                    int n = dataGridView1.Rows.Add();
                    string email;
                    dataGridView1.Rows[n].Cells[0].Value = caminho;
                    dataGridView1.Rows[n].Cells[1].Value = "X";

                    OleDbConnection conexao = new OleDbConnection();
                    conexao.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + caminho + "'";
                    string comandoSQL = "SELECT * FROM VESTIBULINHO";
                    OleDbCommand commando = new OleDbCommand(comandoSQL, conexao);

                    try
                    {
                        conexao.Open();
                        OleDbDataReader dados = commando.ExecuteReader();
                        dataGridView2.Rows.Clear();
                        while (dados.Read())
                        {
                            int m = dataGridView2.Rows.Add();
                            dataGridView2.Rows[m].Cells[0].Value = dados["CLASS"].ToString();
                            dataGridView2.Rows[m].Cells[1].Value = dados["NOTA"].ToString();
                            dataGridView2.Rows[m].Cells[2].Value = dados["HABILITACAO"].ToString();
                            dataGridView2.Rows[m].Cells[3].Value = dados["NOME"].ToString();
                            dataGridView2.Rows[m].Cells[4].Value = dados["SEXO"].ToString();
                            dataGridView2.Rows[m].Cells[5].Value = dados["ENDERECO"].ToString() + " - " + dados["NUMERO"].ToString() + " - " + dados["BAIRRO"].ToString();
                            dataGridView2.Rows[m].Cells[6].Value = "(" + dados["DDD"].ToString() + ")" + " " + dados["TELEFONE"].ToString();
                            dataGridView2.Rows[m].Cells[7].Value = "(" + dados["DDD2"].ToString() + ")" + " " + dados["TELEFONE2"].ToString();
                            dataGridView2.Rows[m].Cells[8].Value = dados["CIDADE"].ToString().Replace("  ", "") + "/" + dados["UF"].ToString();
                            dataGridView2.Rows[m].Cells[9].Value = dados["DT_NASCIMENTO"].ToString();
                            email = dados["EMAIL"].ToString();
                            dataGridView2.Rows[m].Cells[10].Value = email.ToLower();
                            dataGridView2.Rows[m].Cells[11].Value = dados["AFRO_DESC"].ToString();
                            dataGridView2.Rows[m].Cells[12].Value = dados["ESCOLARIDADE"].ToString();
                            dataGridView2.Rows[m].Cells[13].Value = dados["PERIODO"].ToString();
                            dataGridView2.Rows[m].Cells[14].Value = dados["SITUACAO"].ToString();
                            dataGridView2.Rows[m].Cells[15].Value = dados["COD_ETE"].ToString();
                            dataGridView2.Rows[m].Cells[16].Value = dados["HABILITACAO2"].ToString();
                            dataGridView2.Rows[m].Cells[17].Value = dados["PERIODO_2"].ToString();
                            dataGridView2.Rows[m].Cells[18].Value = dados["CLASS2"].ToString();
                        }
                        conexao.Close();
                    }
                    catch (Exception exc)
                    {
                        throw new Exception(exc.Message);
                    }
                });
            }
            
            if (textBox1.Text != "" && comboBox1.Text != "")
            {
                button2.Enabled = true;
            }
            button1.Enabled = false;
             }));
             arquivo1.SetApartmentState(System.Threading.ApartmentState.STA);
             arquivo1.IsBackground = false;
             arquivo1.Start();  
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button2.Enabled = false;
            progressBar1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            aux = "";
            foreach (DataRow item in cg.Verificar(gr).Rows)
            {
                unidade = item["escola"].ToString();
                vest = item["vestibulinho"].ToString();
                if (unidade == dataGridView2.Rows[1].Cells[14].Value.ToString() && vest == comboBox1.Text + " " + textBox1.Text)
                {
                    aux = "Ok";
                }
            }
                
                if (aux == "Ok")
                {
                    string msg = "DADOS JÁ CADASTRADOS NO SISTEMA";
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                }
                else
                {
                    if (dataGridView1.Rows.Count == 0)
                    {
                        string msg = "SELECIONAR UM ARQUIVO COM O BANCO DE DADOS";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                    }
                    else if (comboBox1.Text == "" || textBox1.Text == "")
                    {
                        string msg = "INFORME O SEMESTRE E O ANO DO VESTIBULINHO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                    }
                    else
                    {
                        int quant = dataGridView2.Rows.Count;
                        progressBar1.Visible = true;
                        progressBar1.Maximum = quant;

                        for (int i = 0; i < quant; i++)
                        {
                            gr.Classificacao = dataGridView2.Rows[i].Cells[0].Value.ToString();
                            gr.Nota = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            gr.Habilitacao = dataGridView2.Rows[i].Cells[2].Value.ToString();
                            gr.Nome = dataGridView2.Rows[i].Cells[3].Value.ToString();
                            gr.Sexo = dataGridView2.Rows[i].Cells[4].Value.ToString();
                            gr.Endereco = dataGridView2.Rows[i].Cells[5].Value.ToString();
                            gr.Telefone = dataGridView2.Rows[i].Cells[6].Value.ToString();
                            gr.Celular = dataGridView2.Rows[i].Cells[7].Value.ToString();
                            gr.Cidade = dataGridView2.Rows[i].Cells[8].Value.ToString();
                            gr.DtNascimento = dataGridView2.Rows[i].Cells[9].Value.ToString();
                            gr.Email = dataGridView2.Rows[i].Cells[10].Value.ToString();
                            gr.Afrodescendente = dataGridView2.Rows[i].Cells[11].Value.ToString();
                            gr.Escolaridade = dataGridView2.Rows[i].Cells[12].Value.ToString();
                            gr.Periodo = dataGridView2.Rows[i].Cells[13].Value.ToString();
                            gr.Situacao = dataGridView2.Rows[i].Cells[14].Value.ToString();
                            gr.Vestibulinho = comboBox1.Text + " " + textBox1.Text;
                            gr.Escola = dataGridView2.Rows[i].Cells[15].Value.ToString();
                            gr.Habilitacao2 = dataGridView2.Rows[i].Cells[16].Value.ToString();
                            gr.Periodo2 = dataGridView2.Rows[i].Cells[17].Value.ToString();
                            gr.Classificacao2 = dataGridView2.Rows[i].Cells[18].Value.ToString();
                            cg.cadastro(gr);
                            progressBar1.Value++;
                        }
                        progressBar1.Visible = false;
                        progressBar1.Value = 0;
                        string msg = "DADOS CADASTRADOS COM SUCESSO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();

                        comboBox1.Text = "";
                        textBox1.Text = "";
                    }
                    button2.Enabled = false;
                }
            
        }

        private void comboBox1_Leave(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && textBox1.Text != "")
            {
                button2.Enabled = true;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && comboBox1.Text != "" && textBox1.Text != "")
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            var senderGrid2 = (DataGridView)sender;

            if (dataGridView1.Rows[e.RowIndex].Cells[1].Selected)
            {
                bt = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            }

            if (bt == "X")
            {

                if (senderGrid2.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                    e.RowIndex >= 0)
                {
                    dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
                    button1.Enabled = true;
                }
            }
        }
    }
}
