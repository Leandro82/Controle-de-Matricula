using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using System.Drawing.Imaging;
using System.Data.OleDb;
using System.Timers;
using System.Reflection;

namespace Matricula
{
    class ConectaVagas
    {
        public MySqlConnection conexao;
        string caminho = "Persist Security Info=false;SERVER=10.66.121.42;DATABASE=vestibulinho;UID=secac;pwd=secac";

        public void cadastroVagas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string inserir = "INSERT INTO VAGAS(curso, escola, periodo, vagas, vestibulinho)VALUES('" + vg.Curso + "','" + vg.Escola + "','" + vg.Periodo + "','" + vg.Vaga + "','" + vg.Vestibulinho + "')";
                MySqlCommand comandos = new MySqlCommand(inserir, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void cadastroDatas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string inserir = "INSERT INTO DATAS(chamada, dtInicial, dtFinal, vestibulinho)VALUES('" + vg.Chamada + "','" + Convert.ToDateTime(vg.DtInicio).ToString("yyyy-MM-dd") + "','" + Convert.ToDateTime(vg.DtFim).ToString("yyyy-MM-dd") + "','" + vg.Vestibulinho + "')";
                MySqlCommand comandos = new MySqlCommand(inserir, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void atualizarVagas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE vagas SET vagas= '" + vg.Vaga + "'WHERE cod= '" + vg.Codigo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void atualizarChamadas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE datas SET dtInicial='" + Convert.ToDateTime(vg.DtInicio).ToString("yyyy-MM-dd") + "', dtFinal= '" + Convert.ToDateTime(vg.DtFim).ToString("yyyy-MM-dd") + "'WHERE cod= '" + vg.Codigo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Vestibulinho()
        {
            conexao = new MySqlConnection(caminho);
            conexao.Open();
            string vSQL = "Select distinct vestibulinho from geral where (ocultar!= 'Ok' or ocultar is null)";
            MySqlDataAdapter vDataAdapter = new MySqlDataAdapter(vSQL, conexao);
            DataTable vTable = new DataTable();
            vDataAdapter.Fill(vTable);
            conexao.Close();
            return vTable;
        }

        public DataTable VestibulinhoTodos()
        {
            conexao = new MySqlConnection(caminho);
            conexao.Open();
            string vSQL = "Select distinct vestibulinho, ocultar from geral";
            MySqlDataAdapter vDataAdapter = new MySqlDataAdapter(vSQL, conexao);
            DataTable vTable = new DataTable();
            vDataAdapter.Fill(vTable);
            conexao.Close();
            return vTable;
        }

        public void ocultarVestibulinho(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE geral SET ocultar='" + vg.Ocultar + "'WHERE vestibulinho= '" + vg.Vestibulinho + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }


        public DataTable CursoCad(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT DISTINCT habilitacao, periodo, escola,(SELECT DISTINCT escola) FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'order by habilitacao";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable CursoVag(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, curso, periodo, escola, vagas FROM vagas WHERE vestibulinho= '" + vg.Vestibulinho + "'";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Chamada(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, chamada, dtInicial, dtFinal FROM datas WHERE vestibulinho= '" + vg.Vestibulinho + "'";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaCurso(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT DISTINCT curso, escola, periodo, (SELECT DISTINCT escola) FROM vagas WHERE vestibulinho= '" + vg.Vestibulinho + "'";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaCursoSegOp(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT curso, escola, periodo FROM vagas WHERE vestibulinho= '" + vg.Vestibulinho + "'ORDER BY curso";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaChamada(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT chamada, dtInicial, dtFinal FROM datas WHERE vestibulinho= '" + vg.Vestibulinho + "'";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaDatas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT dtInicial, dtFinal FROM datas WHERE vestibulinho= '" + vg.Vestibulinho + "' AND chamada= '" + vg.Chamada + "'";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaPorChamada(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, email, habilitacao, escol, matriculado, ausente, chamada FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "' AND (chamada!= '2ª Opção' or chamada is null) AND CAST(clas as unsigned integer) >= 1 AND CAST(clas as unsigned integer) <= '" + vg.Vaga + "'ORDER BY CAST(clas as unsigned integer)";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaVagas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT DISTINCT periodo, vagas FROM vagas WHERE vestibulinho= '" + vg.Vestibulinho + "' AND curso= '" + vg.Curso + "'";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaFaltantes(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, habilitacao, escol, matriculado FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND (matriculado is null OR matriculado='Não' OR matriculado='') AND CAST(clas as unsigned integer) >= 1 AND CAST(clas as unsigned integer) <= '" + vg.Vaga + "'ORDER BY CAST(clas as unsigned integer)";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SegundaChamadaFaltantes(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, habilitacao, escol, matriculado FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND (matriculado is null OR matriculado='Não') AND CAST(clas as unsigned integer) >= '"+vg.Vaga+"' AND CAST(clas as unsigned integer) <= '" + vg.Sobra + "'ORDER BY CAST(clas as unsigned integer)";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable DemaisOpcaoFaltantes(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, habilitacao, escol, matriculado FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND (matriculado is null OR matriculado='Não') AND CAST(clas as unsigned integer) BETWEEN '" + vg.Ultimo + "' AND '" + vg.Sobra + "'ORDER BY CAST(clas as unsigned integer)";
                MySqlDataAdapter comand = new MySqlDataAdapter(receber, conexao);
                DataTable dt = new System.Data.DataTable();
                comand.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }
    }
}
