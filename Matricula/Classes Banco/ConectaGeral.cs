using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Data.OleDb;

namespace Matricula
{
    class ConectaGeral
    {
        public MySqlConnection conexao;
        string caminho = "Persist Security Info=false;SERVER=10.66.121.42;DATABASE=vestibulinho;UID=secac;pwd=secac";

        public void cadastro(Geral gr)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string inserir = "INSERT INTO geral(clas, nota, nome, sexo, habilitacao, endereco, telefone, celular, cidade, dtNasc, email, afro, escol, periodo, situacao, vestibulinho, escola, habilitacao2, periodo2, clas2)VALUES('" + gr.Classificacao + "','" + gr.Nota + "','" + gr.Nome + "','" + gr.Sexo + "','" + gr.Habilitacao + "','" + gr.Endereco + "','" + gr.Telefone + "','" + gr.Celular + "','" + gr.Cidade + "','" + gr.DtNascimento + "','" + gr.Email + "','" + gr.Afrodescendente + "','" + gr.Escolaridade + "','" + gr.Periodo + "','" + gr.Situacao + "','" + gr.Vestibulinho + "','" + gr.Escola + "', '" + gr.Habilitacao2 + "','" + gr.Periodo2 + "','" + gr.Classificacao2 + "')";
                MySqlCommand comandos = new MySqlCommand(inserir, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void Matricular(Geral gr)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE geral SET matriculado='" + gr.Matriculado + "', chamada='" + gr.Chamada + "'WHERE cod = '" + gr.Codigo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void Ausente(Geral gr)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE geral SET ausente='" + gr.Ausente + "', chamada='" + gr.Chamada + "'WHERE cod = '" + gr.Codigo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void AusenteSegOp(Geral gr)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE geral SET ausSegOp='" + gr.Ausente + "', chamada='" + gr.Chamada + "'WHERE cod = '" + gr.Codigo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void Listao(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE geral SET listao='" + vg.Curso + "', perListao='" + vg.Periodo + "', matriculado='" + vg.Matriculado + "', chamada='" + vg.Chamada + "' WHERE cod = '" + vg.Codigo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Verificar(Geral gr)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string selecionar = "SELECT vestibulinho, escola FROM geral";
                MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                DataTable dt = new System.Data.DataTable();
                comandos.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable SelecionaMatriculados(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT clas, nome, sexo FROM (SELECT clas, nome, sexo FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao= '" + vg.Curso + "' AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND matriculado='Sim' AND (chamada= '1ª Chamada' OR chamada='2ª Chamada' OR chamada='Pós chamadas') UNION SELECT clas, nome, sexo FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND listao='" + vg.Curso + "' AND perListao='" + vg.Periodo + "' AND matriculado='Sim' AND chamada='Listão'  UNION SELECT clas, nome, sexo FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao2='" + vg.Curso + "' AND chamada='2ª Opção') geral ORDER BY nome";
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

        public DataTable comboListao(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, chamada, escola, matriculado, habilitacao, periodo, habilitacao2, periodo2, listao, perListao FROM geral WHERE cod= '" + vg.Codigo + "'AND matriculado='sim'";
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
