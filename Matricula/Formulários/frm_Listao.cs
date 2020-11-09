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
    public partial class frm_Listao : Form
    {
        ConectaVagas cv = new ConectaVagas();
        ConectaChamada co = new ConectaChamada();
        ConectaGeral cg = new ConectaGeral();
        Geral gr = new Geral();
        Vagas vg = new Vagas();
        int[] pos = new int[10];
        int[] mat = new int[100];
        int cod = 0;
        SaveFileDialog salvarArquivo = new SaveFileDialog(); // novo
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        public frm_Listao()
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
            else
            {
                if (textBox4.Text == "" || textBox5.Text == "")
                {
                    string msg = "INFORME UM INTERVALO VÁLIDO PARA PESQUISA";
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                    textBox4.Focus();
                }
                else
                {
                    string curso;
                    int cs = cv.SelecionaCurso(vg).Rows.Count;
                    vg.Vestibulinho = comboBox1.Text;
                    curso = comboBox2.Text;
                    vg.Curso = curso.Remove(curso.Length - 10);
                    string vaga = cv.SelecionaVagas(vg).Rows[0]["VAGAS"].ToString();
                    vg.Periodo = comboBox4.Text;
                    vg.Vaga = vaga;
                    vg.Primeiro = textBox4.Text;
                    vg.Ultimo = textBox5.Text;
                    vg.Escola = comboBox2.Text.Substring(comboBox2.Text.Length - 7);
                    //dataGridView1.Rows.Clear();
                    foreach (DataRow item in co.Listao(vg).Rows)
                    {
                        int n = dataGridView1.Rows.Add();
                        //dataGridView1.Rows[n].Cells[0].Value = item["COD"].GetHashCode();
                        dataGridView1.Rows[n].Cells[1].Value = item["COD"].GetHashCode();
                        dataGridView1.Rows[n].Cells[2].Value = item["CLAS"].ToString();
                        dataGridView1.Rows[n].Cells[3].Value = item["NOTA"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = item["NOME"].ToString();
                        dataGridView1.Rows[n].Cells[5].Value = item["ENDERECO"].ToString();
                        dataGridView1.Rows[n].Cells[6].Value = item["TELEFONE"].ToString();
                        dataGridView1.Rows[n].Cells[7].Value = item["CELULAR"].ToString();
                        dataGridView1.Rows[n].Cells[8].Value = item["HABILITACAO"].ToString();
                        //DateTime dta = Convert.ToDateTime(item["dtNasc"].ToString());
                        int idade = Idade(Convert.ToDateTime(item["dtNasc"].ToString()));
                        //MessageBox.Show(""+idade);
                        if ((item["HABILITACAO"].ToString() == "ENSINO MÉDIO" || item["HABILITACAO"].ToString() == "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" || item["HABILITACAO"].ToString().Contains("NOVOTEC") == true) && (idade <= 13))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        else if ((item["HABILITACAO"].ToString() != "ENSINO MÉDIO" && item["HABILITACAO"].ToString() != "ADMINISTRAÇÃO - INTEGRADO AO ENSINO MÉDIO" && item["HABILITACAO"].ToString().Contains("NOVOTEC") == false) && (idade <= 14))
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Violet;
                        }
                        dataGridView1.Rows[n].Cells[9].Value = item["ESCOL"].ToString();
                        if (item["matriculado"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.GreenYellow;
                            //dataGridView1.Rows[n].Cells[0].Value = true;
                        }

                        if (item["ausente"].ToString() == "Sim")
                        {
                            dataGridView1.Rows[n].DefaultCellStyle.ForeColor = Color.Red;
                            dataGridView1.Rows[n].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Strikeout);
                            dataGridView1.Rows[n].Cells[0].Value = true;
                        }

                        DataGridViewComboBoxCell aux = (DataGridViewComboBoxCell)dataGridView1.Rows[n].Cells[10];
                        DataGridViewComboBoxCell aux1 = (DataGridViewComboBoxCell)dataGridView1.Rows[n].Cells[11];
                        aux.Items.Add("");
                        for (int i = 0; i < cs; i++)
                        {
                            aux.Items.Add(cv.SelecionaCurso(vg).Rows[i]["curso"].ToString() + " - " + cv.SelecionaCurso(vg).Rows[i]["escola"].ToString());
                        }

                        aux1.Items.Add("");
                        aux1.Items.Add("INTEGRAL");
                        aux1.Items.Add("MANHÃ");
                        aux1.Items.Add("NOITE");
                        aux1.Items.Add("TARDE");

                        vg.Codigo = item["COD"].GetHashCode();
                        foreach (DataRow combo in cg.comboListao(vg).Rows)
                        {
                            if (combo["chamada"].ToString() == "2ª Opção")
                            {
                                aux.Value = combo["habilitacao2"].ToString() + " - " + combo["escola"].ToString();
                                aux1.Value = combo["periodo2"].ToString();
                            }
                            else if (combo["chamada"].ToString() == "Listão")
                            {
                                aux.Value = combo["listao"].ToString() + " - " + combo["escola"].ToString();
                                aux1.Value = combo["perListao"].ToString();
                            }
                            else
                            {
                                aux.Value = combo["habilitacao"].ToString()+" - "+combo["escola"].ToString();
                                aux1.Value = combo["periodo"].ToString();
                            }
                        }
                    }
                    button3.Enabled = true;
                    button4.Enabled = true;
                    textBox3.Enabled = true;
                    maskedTextBox1.Enabled = true;
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            vg.Vestibulinho = comboBox1.Text;
            int curso = cv.SelecionaCursoSegOp(vg).Rows.Count;
            int cham = cv.SelecionaChamada(vg).Rows.Count;
            comboBox2.Items.Clear();
            for (int i = 0; i < curso; i++)
            {
                if (cv.SelecionaCursoSegOp(vg).Rows[i]["curso"].ToString() != "-")
                {
                    comboBox2.Items.Add(cv.SelecionaCursoSegOp(vg).Rows[i]["curso"].ToString() + " - " + cv.SelecionaCursoSegOp(vg).Rows[i]["escola"].ToString());
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button4.Enabled = false;
            textBox3.Enabled = false;
            maskedTextBox1.Enabled = false;

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

        private void frm_Listao_Load(object sender, EventArgs e)
        {
            button3.Enabled = false;
            button4.Enabled = false;
            textBox3.Enabled = false;
            maskedTextBox1.Enabled = false;
            progressBar1.Visible = false;

            int vest = cv.Vestibulinho().Rows.Count;
            for (int i = 0; i < vest; i++)
            {
                comboBox1.Items.Add(cv.Vestibulinho().Rows[i]["vestibulinho"].ToString());
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            comboBox2.Enabled = false;
            comboBox4.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (maskedTextBox1.MaskCompleted && textBox3.Text != "")
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
                    salvarArquivo.FileName = "Lista de APM" + " - " + comboBox2.Text.Remove(comboBox2.Text.Length - 10);
                    salvarArquivo.DefaultExt = "*.xls";
                    salvarArquivo.Filter = "Todos os Aquivos do Excel (*.xls)|*.xls| Todos os arquivos (*.*)|*.*";

                    try
                    {
                        xlApp = new Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Add(misValue);

                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                        xlWorkSheet.PageSetup.TopMargin = 1;
                        xlWorkSheet.PageSetup.BottomMargin = 1;
                        xlWorkSheet.PageSetup.LeftMargin = 1;
                        xlWorkSheet.PageSetup.RightMargin = 1;
                        xlWorkSheet.PageSetup.PrintTitleRows = "$A$2:$G$2";
                        xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 7]].Merge();
                        xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 7]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[1, 1] = "Lista de APM" + " - " + comboBox2.Text.Remove(comboBox2.Text.Length - 10);
                        xlWorkSheet.Cells[1, 1].ColumnWidth = 4.86;
                        xlWorkSheet.Cells[1, 2].ColumnWidth = 35;
                        xlWorkSheet.Cells[1, 3].ColumnWidth = 4.71;
                        xlWorkSheet.Cells[1, 4].ColumnWidth = 5.29;
                        xlWorkSheet.Cells[1, 5].ColumnWidth = 18;
                        xlWorkSheet.Cells[1, 6].ColumnWidth = 20;
                        xlWorkSheet.Cells[1, 7].ColumnWidth = 6;
                        xlWorkSheet.Cells[1, 1].Font.Size = 16;
                        xlWorkSheet.Cells[2, 1] = "MATR.";
                        xlWorkSheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[2, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[2, 1].Font.Size = 10;
                        xlWorkSheet.Cells[2, 2] = "NOME";
                        xlWorkSheet.Cells[2, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[2, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[2, 2].Font.Size = 10;
                        xlWorkSheet.Cells[2, 3] = "CLASS.";
                        xlWorkSheet.Cells[2, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[2, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[2, 3].Font.Size = 10;
                        xlWorkSheet.Cells[2, 4] = "ESC.";
                        xlWorkSheet.Cells[2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[2, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[2, 4].Font.Size = 10;
                        xlWorkSheet.Range[xlWorkSheet.Cells[2, 5], xlWorkSheet.Cells[2, 7]].Merge();
                        xlWorkSheet.Range[xlWorkSheet.Cells[2, 5], xlWorkSheet.Cells[2, 7]] = "APM";
                        xlWorkSheet.Range[xlWorkSheet.Cells[2, 5], xlWorkSheet.Cells[2, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Range[xlWorkSheet.Cells[2, 5], xlWorkSheet.Cells[2, 7]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 7]].Font.Size = 10;
                        int quant = dataGridView1.Rows.Count;

                        progressBar1.Visible = true;
                        progressBar1.Maximum = quant;
                        for (int i = 0; i < quant; i++)
                        {
                            xlWorkSheet.Cells[l, 1] = "";
                            xlWorkSheet.Cells[l, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 2] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                            xlWorkSheet.Cells[l, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 3] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                            xlWorkSheet.Cells[l, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 4] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                            xlWorkSheet.Cells[l, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 5] = "(  ) PG R$ " + textBox3.Text + label10.Text + "_______";
                            xlWorkSheet.Cells[l, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 5].Font.Size = 9;
                            xlWorkSheet.Cells[l, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 6] = "(  ) Pagará até " + maskedTextBox1.Text;
                            xlWorkSheet.Cells[l, 6].Font.Size = 9;
                            xlWorkSheet.Cells[l, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 7] = "(  ) NÃO";
                            xlWorkSheet.Cells[l, 7].Font.Size = 9;
                            xlWorkSheet.Cells[l, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            l = l + 1;
                            progressBar1.Value++;
                        }
                        xlWorkSheet.Application.Columns[2].ShrinkToFit = true;

                        int linhas = 0;
                        if (quant < 55)
                        {
                            linhas = 55;
                        }
                        else if (quant > 55 && quant < 110)
                        {
                            linhas = 110;
                        }
                        else if (quant > 110 && quant < 165)
                        {
                            linhas = 165;
                        }
                        progressBar1.Visible = true;
                        progressBar1.Value = 0;
                        progressBar1.Maximum = (linhas - l) + 1;
                        //MessageBox.Show("" + progressBar1.Maximum);
                        for (int i = l; i <= linhas; i++)
                        {
                            xlWorkSheet.Cells[l, 1] = "";
                            xlWorkSheet.Cells[l, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 2] = "";
                            xlWorkSheet.Cells[l, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 3] = "";
                            xlWorkSheet.Cells[l, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 4] = "";
                            xlWorkSheet.Cells[l, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 5] = "(  ) PG R$ " + textBox3.Text + label10.Text + "_______";
                            xlWorkSheet.Cells[l, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 5].Font.Size = 9;
                            xlWorkSheet.Cells[l, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 6] = "(  ) Pagará até " + maskedTextBox1.Text;
                            xlWorkSheet.Cells[l, 6].Font.Size = 9;
                            xlWorkSheet.Cells[l, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[l, 7] = "(  ) NÃO";
                            xlWorkSheet.Cells[l, 7].Font.Size = 9;
                            xlWorkSheet.Cells[l, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            xlWorkSheet.Cells[l, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            l = l + 1;
                            progressBar1.Value++;
                        }
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
            else
            {
                string msg = "INFORMAR O VALOR E ATÉ QUANDO PODERÁ PAGAR A APM";
                frm_Mensagem mg = new frm_Mensagem(msg);
                mg.ShowDialog();
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
                salvarArquivo.FileName = "Listão de Telefones";
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
                    xlWorkSheet.PageSetup.LeftMargin = 0;
                    xlWorkSheet.PageSetup.RightMargin = 0;
                    xlWorkSheet.PageSetup.PrintTitleRows = "$A$2:$I$2";
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 9]].Merge();
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[1, 1] = "Listão de Telefones";
                    xlWorkSheet.Cells[1, 1].ColumnWidth = 7;
                    xlWorkSheet.Cells[1, 2].ColumnWidth = 7;
                    xlWorkSheet.Cells[1, 3].ColumnWidth = 35;
                    xlWorkSheet.Cells[1, 4].ColumnWidth = 35;
                    xlWorkSheet.Cells[1, 5].ColumnWidth = 13;
                    xlWorkSheet.Cells[1, 6].ColumnWidth = 14;
                    xlWorkSheet.Cells[1, 7].ColumnWidth = 7;
                    xlWorkSheet.Cells[1, 8].ColumnWidth = 8.43;
                    xlWorkSheet.Cells[1, 9].ColumnWidth = 8.43;
                    xlWorkSheet.Cells[1, 1].Font.Size = 16;
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 8]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 1] = "CLASS";
                    xlWorkSheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 2] = "NOTA";
                    xlWorkSheet.Cells[2, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 3] = "NOME";
                    xlWorkSheet.Cells[2, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 4] = "ENDEREÇO";
                    xlWorkSheet.Cells[2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 5] = "TELEFONE";
                    xlWorkSheet.Cells[2, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 6] = "CELULAR";
                    xlWorkSheet.Cells[2, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 7] = "ESC.";
                    xlWorkSheet.Cells[2, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 8] = "HAB.";
                    xlWorkSheet.Cells[2, 8].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[2, 9] = "OBS.";
                    xlWorkSheet.Cells[2, 9].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[2, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 9]].Font.Size = 12;
                    int quant = dataGridView1.Rows.Count;

                    progressBar1.Visible = true;
                    progressBar1.Maximum = quant;
                    for (int i = 0; i < quant; i++)
                    {
                        if (dataGridView1.Rows[i].DefaultCellStyle.BackColor == Color.Violet)
                        {
                            xlWorkSheet.Range[xlWorkSheet.Cells[l, 1], xlWorkSheet.Cells[l, 9]].Interior.Color = ColorTranslator.ToWin32(Color.Yellow);
                        }
                        else
                        {
                            xlWorkSheet.Range[xlWorkSheet.Cells[l, 1], xlWorkSheet.Cells[l, 9]].Interior.Color = ColorTranslator.ToWin32(Color.White);
                        }
                        xlWorkSheet.Cells[l, 1] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        xlWorkSheet.Cells[l, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 2] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        xlWorkSheet.Cells[l, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 3] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        xlWorkSheet.Cells[l, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 4] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                        xlWorkSheet.Cells[l, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 5] = dataGridView1.Rows[i].Cells[6].Value.ToString();
                        xlWorkSheet.Cells[l, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 6] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                        xlWorkSheet.Cells[l, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 7] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                        xlWorkSheet.Cells[l, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 8] = dataGridView1.Rows[i].Cells[8].Value.ToString().Substring(0,3);
                        xlWorkSheet.Cells[l, 8].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        l = l + 1;
                        progressBar1.Value++;
                    }
                    xlWorkSheet.Application.Columns[3].ShrinkToFit = true;
                    xlWorkSheet.Application.Columns[4].ShrinkToFit = true;
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

        private void button2_Click(object sender, EventArgs e)
        {
            gr.Observacao = textBox2.Text;
            co.SalvarObs(gr);
            string msg = "OBSERVAÇÃO SALVA";
            frm_Mensagem mg = new frm_Mensagem(msg);
            mg.ShowDialog();
            textBox2.Text = "";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            cod = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value);
            gr.Codigo = cod;
            int obs = co.Observacao(gr).Rows.Count;
            if (obs != 0)
            {
                textBox2.Text = co.Observacao(gr).Rows[0][0].ToString();
            }

            if (e.ColumnIndex == dataGridView1.Columns[0].Index)
            {
                dataGridView1.EndEdit();  //Stop editing of cell.
                cod = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value);
                if ((bool)dataGridView1.Rows[e.RowIndex].Cells[0].Value)
                {
                    gr.Codigo = cod;
                    gr.Ausente = "Sim";
                    gr.Chamada = "Listão";
                    cg.Ausente(gr);
                }
                else
                {
                    gr.Codigo = cod;
                    gr.Ausente = "Não";
                    gr.Chamada = "Listão";
                    cg.Ausente(gr);
                }                
            }

            if (e.ColumnIndex == dataGridView1.Columns[0].Index)
            {
                dataGridView1.EndEdit();  //Stop editing of cell.
                if (Convert.ToBoolean(dataGridView1.CurrentCell.Value) == true)
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Strikeout);
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Regular);
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cod = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value);
            gr.Codigo = cod;
            
            int obs = co.Observacao(gr).Rows.Count;
            if (obs != 0)
            {
                textBox2.Text = co.Observacao(gr).Rows[0][0].ToString();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int cont = dataGridView1.Rows.Count;

            for (int i = 0; i < cont; i++)
            {
                if (String.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[10].Value) && String.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[11].Value))
                {
                    vg.Codigo = dataGridView1.Rows[i].Cells[1].Value.GetHashCode();
                    vg.Curso = "";
                    vg.Periodo = "";
                    vg.Chamada = "";
                    vg.Matriculado = null;
                    cg.Listao(vg);
                }
                else
                {
                    vg.Codigo = dataGridView1.Rows[i].Cells[1].Value.GetHashCode();
                    vg.Curso = dataGridView1.Rows[i].Cells[10].Value.ToString().Remove(dataGridView1.Rows[i].Cells[10].Value.ToString().Length - 10);
                    vg.Periodo = dataGridView1.Rows[i].Cells[11].Value.ToString();
                    vg.Chamada = "Listão";
                    vg.Matriculado = "Sim";
                    cg.Listao(vg);
                    string msg = "ALUNO(A) MATRICULADO(A) EM " + dataGridView1.Rows[i].Cells[10].Value.ToString().Remove(dataGridView1.Rows[i].Cells[10].Value.ToString().Length - 10) + " - " + dataGridView1.Rows[i].Cells[11].Value.ToString();
                    frm_Mensagem mg = new frm_Mensagem(msg);
                    mg.ShowDialog();
                }
            }
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            cod = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value);
            gr.Codigo = cod;
            int obs = co.Observacao(gr).Rows.Count;
            if (obs != 0)
            {
                textBox2.Text = co.Observacao(gr).Rows[0][0].ToString();
            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if ((e.Exception) is ConstraintException)
            {
                DataGridView view = (DataGridView)sender;
                view.Rows[e.RowIndex].ErrorText = "an error";
                view.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "an error";

                e.ThrowException = false;
            }
        }    
      }
    }

