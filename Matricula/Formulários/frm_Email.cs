using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Matricula
{
    public partial class frm_Email : Form
    {
        ConectaVagas cv = new ConectaVagas();
        ConectaChamada co = new ConectaChamada();
        ConectaGeral cg = new ConectaGeral();
        Geral gr = new Geral();
        Vagas vg = new Vagas();
        string aux = "", chamada = "";
        int[] pos = new int[10];
        int sob, dtIn, dtFi;
        SaveFileDialog salvarArquivo = new SaveFileDialog(); // novo
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        public frm_Email()
        {
            InitializeComponent();
        }

        private void frm_Email_Load(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            int vest = cv.Vestibulinho().Rows.Count;
            for (int i = 0; i < vest; i++)
            {
                comboBox1.Items.Add(cv.Vestibulinho().Rows[i]["vestibulinho"].ToString());
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            vg.Vestibulinho = comboBox1.Text;
            int j = 0;
            int curso = cv.SelecionaCurso(vg).Rows.Count;
            int cham = cv.SelecionaChamada(vg).Rows.Count;
            comboBox2.Items.Clear();
            for (int i = 0; i < curso; i++)
            {
                comboBox2.Items.Add(cv.SelecionaCurso(vg).Rows[i]["curso"].ToString() + " - " + cv.SelecionaCurso(vg).Rows[i]["escola"].ToString());
            }

            comboBox3.Text = "";
            comboBox3.Items.Clear();

            for (int i = 0; i < cham; i++)
            {
                dtIn = (DateTime.Parse(Convert.ToDateTime(cv.SelecionaChamada(vg).Rows[i]["dtInicial"].ToString()).ToString("dd/MM/yyyy")).Subtract(DateTime.Today)).Days;
                dtFi = (DateTime.Parse(Convert.ToDateTime(cv.SelecionaChamada(vg).Rows[i]["dtFinal"].ToString()).ToString("dd/MM/yyyy")).Subtract(DateTime.Today)).Days;
                if ((dtIn <= 0 && dtFi >= 0) || (dtIn < 0 && dtFi < 0))
                {
                    aux = "Ok";
                    pos[j] = i;
                    j = j + 1;
                }
            }


            if (aux == "Ok")
            {
                for (int i = 0; i < j; i++)
                {
                    comboBox3.Items.Add(cv.SelecionaChamada(vg).Rows[i]["chamada"].ToString());
                }
                if (dtFi < 0)
                {
                    comboBox3.Items.Add("Pós Chamadas");
                }
            }
            else
            {
                int vagas = cv.CursoVag(vg).Rows.Count;

                if (vagas == 0)
                {
                    string msg = "PRIMEIRO CADASTRE AS VAGAS E AS CHAMADAS NOS SISTEMA";
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                    comboBox2.Enabled = false;
                    comboBox3.Enabled = false;
                    comboBox4.Enabled = false;
                }
                else
                {
                    comboBox3.Items.Add(cv.SelecionaChamada(vg).Rows[0]["chamada"].ToString());
                }
            }
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
            dataGridView1.Rows.Clear();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string curso;
            vg.Vestibulinho = comboBox1.Text;
            curso = comboBox2.Text;
            vg.Curso = curso.Remove(curso.Length - 10);
            vg.Vaga = cv.SelecionaVagas(vg).Rows[0]["VAGAS"].ToString();
        }

        public static int Idade(DateTime dtNascimento)
        {
            int idade = DateTime.Now.Year - dtNascimento.Year;
            if (DateTime.Now.Month < dtNascimento.Month || (DateTime.Now.Month == dtNascimento.Month && DateTime.Now.Day < dtNascimento.Day))
                idade--;

            return idade;

        }

        private void button1_Click(object sender, EventArgs e)
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
            else if (comboBox3.Text == "")
            {
                string msg = "SELECIONAR A CHAMADA";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
                comboBox3.Focus();
            }
            else
            {
                //MessageBox.Show(""+sob);
                dataGridView1.Rows.Clear();
                if (comboBox3.Text == "1ª chamada")
                {
                    string curso;
                    //chamada = Convert.ToInt32((comboBox3.Text).Remove((comboBox3.Text).Length - 9)) - 1;
                    vg.Banco = comboBox3.Text;
                    vg.Escola = comboBox2.Text.Substring(comboBox2.Text.Length - 7);
                    vg.Vestibulinho = comboBox1.Text;
                    curso = comboBox2.Text;
                    vg.Curso = curso.Remove(curso.Length - 10);
                    string vaga = cv.SelecionaVagas(vg).Rows[0]["VAGAS"].ToString();
                    vg.Periodo = comboBox4.Text;
                    vg.Chamada = comboBox3.Text;
                    vg.Banco = comboBox3.Text;
                    vg.Vaga = vaga;
                    sob = 0;
                    progressBar1.Visible = true;
                    progressBar1.Maximum = cv.SelecionaPorChamada(vg).Rows.Count;
                    foreach (DataRow item in cv.SelecionaPorChamada(vg).Rows)
                    {
                        int n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = item["CLAS"].ToString();
                        dataGridView1.Rows[n].Cells[1].Value = item["NOME"].ToString();
                        dataGridView1.Rows[n].Cells[2].Value = item["TELEFONE"].ToString();
                        dataGridView1.Rows[n].Cells[3].Value = item["CELULAR"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = item["EMAIL"].ToString();

                        int idade = Idade(Convert.ToDateTime(item["dtNasc"].ToString()));

                        if ((item["HABILITACAO"].ToString() == "ENSINO MÉDIO" || item["HABILITACAO"].ToString() == "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" || item["HABILITACAO"].ToString().Contains("NOVOTEC") == true) && (idade <= 13))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        else if ((item["HABILITACAO"].ToString() != "ENSINO MÉDIO" && item["HABILITACAO"].ToString() != "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" && item["HABILITACAO"].ToString().Contains("NOVOTEC") == false) && (idade <= 14))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        
                        if (item["matriculado"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.GreenYellow;
                        }
                        else
                        {
                            sob = sob + 1;
                        }
                        if (item["ausente"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.ForeColor = Color.Red;
                            dataGridView1.Rows[n].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Strikeout);
                        }
                        progressBar1.Value++;
                    }
                    vg.Sobra = sob;
                    co.GravarSobras(vg);

                    if (dataGridView1.Rows.Count < 40)
                    {
                        co.SalvarObs(gr);
                        string msg = "COMPLETE A TURMA COM A LISTA DA 2º OPÇÃO E/OU COM O LISTÃO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                    }
                }
                else if (comboBox3.Text == "2ª chamada")
                {
                    sob = 0;
                    int vag = Convert.ToInt32(cv.SelecionaVagas(vg).Rows[0]["VAGAS"].ToString()) + 1;
                    vg.Periodo = comboBox4.Text;
                    vg.Vaga = Convert.ToString(vag);
                    int cont = co.SomaMatricula(vg).Rows[0][0].GetHashCode();
                    vg.Banco = Convert.ToString(Convert.ToInt32((comboBox3.Text).Remove((comboBox3.Text).Length - 9)) - 1) + "ª Chamada";
                    int sobra = co.VerificaSobra(vg).Rows[0][0].GetHashCode() - 1;
                    vg.Sobra = vag + sobra;
                    vg.Escola = comboBox2.Text.Substring(comboBox2.Text.Length - 7);
                    progressBar1.Visible = true;
                    progressBar1.Maximum = co.SegundaChamada(vg).Rows.Count;
                    foreach (DataRow item in co.SegundaChamada(vg).Rows)
                    {
                        int n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = item["CLAS"].ToString();
                        dataGridView1.Rows[n].Cells[1].Value = item["NOME"].ToString();
                        dataGridView1.Rows[n].Cells[2].Value = item["TELEFONE"].ToString();
                        dataGridView1.Rows[n].Cells[3].Value = item["CELULAR"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = item["EMAIL"].ToString();

                        int idade = Idade(Convert.ToDateTime(item["dtNasc"].ToString()));

                        if ((item["HABILITACAO"].ToString() == "ENSINO MÉDIO" || item["HABILITACAO"].ToString() == "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" || item["HABILITACAO"].ToString().Contains("NOVOTEC") == true) && (idade <= 13))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        else if ((item["HABILITACAO"].ToString() != "ENSINO MÉDIO" && item["HABILITACAO"].ToString() != "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" && item["HABILITACAO"].ToString().Contains("NOVOTEC") == false) && (idade <= 14))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        if (item["matriculado"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.GreenYellow;
                        }
                        else
                        {
                            sob = sob + 1;
                        }
                        if (item["ausente"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.ForeColor = Color.Red;
                            dataGridView1.Rows[n].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Strikeout);
                        }
                        vg.Ultimo = dataGridView1.Rows[n].Cells[0].Value.ToString();
                        progressBar1.Value++;
                    }
                    vg.Banco = comboBox3.Text;
                    vg.Sobra = sob;
                    co.GravarSobras(vg);

                    vg.Banco = comboBox3.Text;
                    co.GravarChamadas(vg);

                    if (dataGridView1.Rows.Count < sob)
                    {
                        co.SalvarObs(gr);
                        string msg = "COMPLETE A TURMA COM A LISTA DA 2º OPÇÃO E/OU COM O LISTÃO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                    }
                }
                else if (comboBox3.Text == "Pós Chamadas")
                {
                    int cont = co.UltimaChamada().Rows.Count;
                    for (int i = 0; i < cont; i++)
                    {
                        chamada = co.UltimaChamada().Rows[i][0].ToString();
                    }
                    vg.Banco = chamada;
                    vg.Periodo = comboBox4.Text;
                    string ult = co.VerificaChamada(vg).Rows[0][0].ToString();
                    vg.Ultimo = ult;
                    vg.Escola = comboBox2.Text.Substring(comboBox2.Text.Length - 7);
                    progressBar1.Visible = true;
                    progressBar1.Maximum = co.PosChamadas(vg).Rows.Count;
                    foreach (DataRow item in co.PosChamadas(vg).Rows)
                    {
                        vg.Banco = comboBox3.Text;
                        int n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = item["CLAS"].ToString();
                        dataGridView1.Rows[n].Cells[1].Value = item["NOME"].ToString();
                        dataGridView1.Rows[n].Cells[2].Value = item["TELEFONE"].ToString();
                        dataGridView1.Rows[n].Cells[3].Value = item["CELULAR"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = item["EMAIL"].ToString();

                        int idade = Idade(Convert.ToDateTime(item["dtNasc"].ToString()));

                        if ((item["HABILITACAO"].ToString() == "ENSINO MÉDIO" || item["HABILITACAO"].ToString() == "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" || item["HABILITACAO"].ToString().Contains("NOVOTEC") == true) && (idade <= 13))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        else if ((item["HABILITACAO"].ToString() != "ENSINO MÉDIO" && item["HABILITACAO"].ToString() != "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" && item["HABILITACAO"].ToString().Contains("NOVOTEC") == false) && (idade <= 14))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        if (item["matriculado"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.GreenYellow;
                        }
                        if (item["ausente"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.ForeColor = Color.Red;
                            dataGridView1.Rows[n].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Strikeout);
                        }
                        progressBar1.Value++;
                    }
                }
                else
                {
                    sob = 0;
                    vg.Banco = Convert.ToInt32((comboBox3.Text).Remove((comboBox3.Text).Length - 9)) - 1 + "ª chamada";
                    vg.Periodo = comboBox4.Text;
                    int sobra = co.VerificaSobra(vg).Rows[0][0].GetHashCode();
                    vg.Banco = Convert.ToInt32((comboBox3.Text).Remove((comboBox3.Text).Length - 9)) - 1 + "ª chamada";
                    string ult = co.VerificaChamada(vg).Rows[0][0].ToString();
                    vg.Ultimo = Convert.ToString(Convert.ToInt32(ult) + 1);
                    vg.Sobra = Convert.ToInt32(ult) + sobra;
                    vg.Escola = comboBox2.Text.Substring(comboBox2.Text.Length - 7);
                    progressBar1.Visible = true;
                    progressBar1.Maximum = co.DemaisOpcao(vg).Rows.Count;
                    foreach (DataRow item in co.DemaisOpcao(vg).Rows)
                    {
                        vg.Banco = comboBox3.Text;
                        int n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = item["CLAS"].ToString();
                        dataGridView1.Rows[n].Cells[1].Value = item["NOME"].ToString();
                        dataGridView1.Rows[n].Cells[2].Value = item["TELEFONE"].ToString();
                        dataGridView1.Rows[n].Cells[3].Value = item["CELULAR"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = item["EMAIL"].ToString();
                        int idade = Idade(Convert.ToDateTime(item["dtNasc"].ToString()));

                        if ((item["HABILITACAO"].ToString() == "ENSINO MÉDIO" || item["HABILITACAO"].ToString() == "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" || item["HABILITACAO"].ToString().Contains("NOVOTEC") == true) && (idade <= 13))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        else if ((item["HABILITACAO"].ToString() != "ENSINO MÉDIO" && item["HABILITACAO"].ToString() != "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" && item["HABILITACAO"].ToString().Contains("NOVOTEC") == false) && (idade <= 14))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        if (item["matriculado"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.GreenYellow;
                        }
                        else
                        {
                            sob = sob + 1;
                        }
                        if (item["ausente"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.ForeColor = Color.Red;
                            dataGridView1.Rows[n].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Strikeout);
                        }
                        vg.Ultimo = item["CLAS"].ToString();

                        progressBar1.Value++;
                    }
                    vg.Banco = comboBox3.Text;
                    vg.Sobra = sob;
                    co.GravarSobras(vg);
                    vg.Banco = comboBox3.Text;
                    co.GravarChamadas(vg);
                    if (dataGridView1.Rows.Count < sob)
                    {
                        co.SalvarObs(gr);
                        string msg = "COMPLETE A TURMA COM A LISTA DA 2º OPÇÃO E/OU COM O LISTÃO";
                        frm_Mensagem mg = new frm_Mensagem(msg);
                        mg.ShowDialog();
                    }
                }              
                button4.Enabled = true;
            }
            progressBar1.Value = 0;
            progressBar1.Visible = false;
            if (Convert.ToInt32(cv.SelecionaVagas(vg).Rows[0]["VAGAS"].ToString()) <= Convert.ToInt32(co.TotalMatriculado(vg).Rows[0][0].ToString()))
            {
                string msg = "TURMA COMPLETA";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int cont = dataGridView1.Rows.Count;
            if (cont == 0)
            {
                string msg = "NÃO EXISTE DADOS PARA GERAR O EXCEL";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
            }
            else
            {
                int l = 3;
                salvarArquivo.FileName = "Lista de Classificados" + " - " + comboBox2.Text.Remove(comboBox2.Text.Length - 10) + " - " + comboBox3.Text;
                salvarArquivo.DefaultExt = "*.xls";
                salvarArquivo.Filter = "Todos os Aquivos do Excel (*.xls)|*.xls| Todos os arquivos (*.*)|*.*";

                try
                {
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    xlWorkSheet.PageSetup.TopMargin = 2;
                    xlWorkSheet.PageSetup.BottomMargin = 1;
                    xlWorkSheet.PageSetup.LeftMargin = 3;
                    xlWorkSheet.PageSetup.RightMargin = 2;
                    xlWorkSheet.PageSetup.PrintTitleRows = "$A$2:$E$2";
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 5]].Merge();
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[1, 1] = "Lista de Classificados" + " - " + comboBox2.Text.Remove(comboBox2.Text.Length - 10) + " - " + comboBox3.Text;
                    xlWorkSheet.Cells[1, 1].ColumnWidth = 7;
                    xlWorkSheet.Cells[1, 2].ColumnWidth = 39;
                    xlWorkSheet.Cells[1, 3].ColumnWidth = 15;
                    xlWorkSheet.Cells[1, 4].ColumnWidth = 15;
                    xlWorkSheet.Cells[1, 5].ColumnWidth = 40;
                    xlWorkSheet.Cells[1, 1].Font.Size = 16;
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 5]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 1] = "CLASS";
                    xlWorkSheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 2] = "NOME";
                    xlWorkSheet.Cells[2, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 3] = "TELEFONE";
                    xlWorkSheet.Cells[2, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 4] = "CELULAR";
                    xlWorkSheet.Cells[2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 5] = "EMAIL";
                    xlWorkSheet.Cells[2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 5]].Font.Size = 12;
                    int quant = dataGridView1.Rows.Count;

                    progressBar1.Visible = true;
                    progressBar1.Maximum = quant;
                    for (int i = 0; i < quant; i++)
                    {
                        if (dataGridView1.Rows[i].DefaultCellStyle.BackColor == Color.Violet)
                        {
                            xlWorkSheet.Range[xlWorkSheet.Cells[l, 1], xlWorkSheet.Cells[l, 5]].Interior.Color = ColorTranslator.ToWin32(Color.Yellow);
                        }
                        else
                        {
                            xlWorkSheet.Range[xlWorkSheet.Cells[l, 1], xlWorkSheet.Cells[l, 5]].Interior.Color = ColorTranslator.ToWin32(Color.White);
                        }
                        xlWorkSheet.Cells[l, 1] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        xlWorkSheet.Cells[l, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 2] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        xlWorkSheet.Cells[l, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 3] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        xlWorkSheet.Cells[l, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 4] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        xlWorkSheet.Cells[l, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 5] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        xlWorkSheet.Cells[l, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        l = l + 1;
                        progressBar1.Value++;
                    }
                    xlWorkSheet.Application.Columns[2].ShrinkToFit = true;
                    xlWorkSheet.Application.Columns[5].ShrinkToFit = true;
                    progressBar1.Value = 0;
                    progressBar1.Visible = false;

                    new System.Threading.Thread(delegate()
                    {
                        Export();
                    }).Start();
                }
                catch (Exception ex)
                {
                    string msg = "Erro : " + ex.Message;
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                }
            }
        }

        private void Export()
        {
            System.Threading.Thread arquivo = new System.Threading.Thread(new System.Threading.ThreadStart(() =>
            {
                if (salvarArquivo.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    xlWorkBook.SaveAs(salvarArquivo.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                    Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    liberarObjetos(xlWorkSheet);
                    liberarObjetos(xlWorkBook);
                    liberarObjetos(xlApp);
                }
            }));
            arquivo.SetApartmentState(System.Threading.ApartmentState.STA);
            arquivo.IsBackground = false;
            arquivo.Start();
        }


        private void liberarObjetos(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                string msg = "Ocorreu um erro durante a liberação do objeto " + ex.ToString();
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
