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
    public partial class frm_AvisoSobra : Form
    {
        ConectaChamada co = new ConectaChamada();
        Vagas vg = new Vagas();
        int sobra;
        string banco, curso, vestibulinho, periodo;
        public frm_AvisoSobra(string bd, int sob, string cs, string vest, string per)
        {
            InitializeComponent();
            banco = bd;
            sobra = sob;
            curso = cs;
            vestibulinho = vest;
            periodo = per;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            vg.Banco = banco;
            vg.Sobra = sobra;
            vg.Curso = curso;
            vg.Vestibulinho = vestibulinho;
            vg.Periodo = periodo;
            co.GravarSobras(vg);
            string msg = "SOBRAS ATUALIZADAS";
            frm_Mensagem mg = new frm_Mensagem(msg);
            mg.ShowDialog();
            this.Close();
        }

        private void frm_AvisoSobra_Load(object sender, EventArgs e)
        {
            button1.Visible = false;
            button2.Focus();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                button1.Visible = true;
            }
            else if (checkBox1.Checked == false)
            {
                button1.Visible = false;
            }
        }

    }
}
