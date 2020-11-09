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
    public partial class frm_NaoMatriculado : Form
    {
        ConectaVagas cv = new ConectaVagas();
        ConectaChamada co = new ConectaChamada();
        ConectaGeral cg = new ConectaGeral();
        Geral gr = new Geral();
        Vagas vg = new Vagas();
        SaveFileDialog salvarArquivo = new SaveFileDialog(); // novo
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        public frm_NaoMatriculado()
        {
            InitializeComponent();
        }

        private void frm_NaoMatriculado_Load(object sender, EventArgs e)
        {
            int vest = cv.Vestibulinho().Rows.Count;
            for (int i = 0; i < vest; i++)
            {
                comboBox1.Items.Add(cv.Vestibulinho().Rows[i]["vestibulinho"].ToString());
            }
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
            else
            {
                vg.Vestibulinho = comboBox1.Text;
                dataGridView1.Rows.Clear();
                foreach (DataRow item in co.NaoMatriculados(vg).Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = item["CLAS"].ToString();
                    dataGridView1.Rows[n].Cells[1].Value = item["NOTA"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["NOME"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["ENDERECO"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["TELEFONE"].ToString();
                    dataGridView1.Rows[n].Cells[5].Value = item["CELULAR"].ToString();
                    dataGridView1.Rows[n].Cells[6].Value = item["HABILITACAO"].ToString();
                }
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
                salvarArquivo.FileName = "Lista de Telefones dos que não foram convocados";
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
                    xlWorkSheet.PageSetup.PrintTitleRows = "$A$2:$G$2";
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 7]].Merge();
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[1, 1] = "Lista de Telefones";
                    xlWorkSheet.Cells[1, 1].ColumnWidth = 7;
                    xlWorkSheet.Cells[1, 2].ColumnWidth = 7;
                    xlWorkSheet.Cells[1, 3].ColumnWidth = 39;
                    xlWorkSheet.Cells[1, 4].ColumnWidth = 39;
                    xlWorkSheet.Cells[1, 5].ColumnWidth = 13;
                    xlWorkSheet.Cells[1, 6].ColumnWidth = 14;
                    xlWorkSheet.Cells[1, 7].ColumnWidth = 20;
                    xlWorkSheet.Cells[1, 1].Font.Size = 16;
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 7]].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
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
                    xlWorkSheet.Cells[2, 7] = "CURSO.";
                    xlWorkSheet.Cells[2, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 7]].Font.Size = 12;
                    int quant = dataGridView1.Rows.Count;

                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = quant;
                    for (int i = 0; i < quant; i++)
                    {
                        xlWorkSheet.Cells[l, 1] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        xlWorkSheet.Cells[l, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 2] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        xlWorkSheet.Cells[l, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 3] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        xlWorkSheet.Cells[l, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 4] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        xlWorkSheet.Cells[l, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 5] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        xlWorkSheet.Cells[l, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 6] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                        xlWorkSheet.Cells[l, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[l, 7] = dataGridView1.Rows[i].Cells[6].Value.ToString() + " - " + comboBox1.Text;
                        xlWorkSheet.Cells[l, 7].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        xlWorkSheet.Cells[l, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        l = l + 1;
                        progressBar1.Value++;
                    }
                    xlWorkSheet.Application.Columns[3].ShrinkToFit = true;
                    xlWorkSheet.Application.Columns[4].ShrinkToFit = true;

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
