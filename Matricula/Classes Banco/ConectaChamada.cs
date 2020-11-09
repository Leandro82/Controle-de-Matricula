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
    class ConectaChamada
    {
        public MySqlConnection conexao;
        string caminho = "Persist Security Info=false;SERVER=10.66.121.42;DATABASE=vestibulinho;UID=secac;pwd=secac";

        public DataTable SomaMatricula(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT COUNT(matriculado) FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao= '" + vg.Curso + "'AND chamada='" + vg.Chamada + "'AND periodo='" + vg.Periodo + "'";
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

        public DataTable ContaMatricula(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT COUNT(matriculado) FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'";
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

        public DataTable TotalMatriculado(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT SUM(total) AS vagas FROM (SELECT COUNT(nome) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao= '" + vg.Curso + "' AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND matriculado='Sim' AND (chamada= '1ª Chamada' OR chamada='2ª Chamada' OR chamada='Pós chamadas') UNION SELECT COUNT(nome) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND listao='" + vg.Curso + "' AND perListao='" + vg.Periodo + "' AND matriculado='Sim' AND chamada='Listão'  UNION SELECT COUNT(nome) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao2='" + vg.Curso + "' AND chamada='2ª Opção') geral";
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

        public DataTable TotalMasculino(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT SUM(total) AS vagas FROM (SELECT COUNT(sexo) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao= '" + vg.Curso + "' AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND matriculado='Sim' AND sexo='MASCULINO' AND (chamada= '1ª Chamada' OR chamada='2ª Chamada' OR chamada='Pós chamadas') UNION SELECT COUNT(sexo) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND listao='" + vg.Curso + "' AND perListao='" + vg.Periodo + "' AND matriculado='Sim' AND chamada='Listão' AND sexo='MASCULINO' UNION SELECT COUNT(sexo) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao2='" + vg.Curso + "' AND chamada='2ª Opção' AND sexo='MASCULINO') geral";
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

        public DataTable TotalFeminino(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT SUM(total) AS vagas FROM (SELECT COUNT(sexo) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao= '" + vg.Curso + "' AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND matriculado='Sim' AND sexo='FEMININO' AND (chamada= '1ª Chamada' OR chamada='2ª Chamada' OR chamada='Pós chamadas') UNION SELECT COUNT(sexo) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND listao='" + vg.Curso + "' AND perListao='" + vg.Periodo + "' AND matriculado='Sim' AND chamada='Listão' AND sexo='FEMININO' UNION SELECT COUNT(sexo) AS total FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "' AND habilitacao2='" + vg.Curso + "' AND chamada='2ª Opção' AND sexo='FEMININO') geral";
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

        public String BuscaDataServidor()
        {
            string data;
            using (MySqlConnection cn = new MySqlConnection())
            {
                cn.ConnectionString = caminho;
                try
                {
                    MySqlDataAdapter sda = new MySqlDataAdapter("SELECT NOW()", cn);
                    cn.Open();
                    DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    int cont = dt.Rows.Count;
                    if (cont > 0)
                    {
                        data = dt.Rows[0][0].ToString();
                    }
                    else
                    {
                        data = "";
                    }
                }
                catch (MySqlException e)
                {
                    throw new Exception(e.Message);
                }
                finally
                {
                    cn.Close();
                }
                return data;
            }
        }

        public DataTable datasChamadas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, chamada, dtFinal, vestibulinho FROM datas WHERE vestibulinho='"+ vg.Vestibulinho +"'";
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

        public void GravarChamadas(Vagas vg)
        {
            if (vg.Banco.Remove(vg.Banco.Length - 9) == "2")
            {
                vg.Banco = "dois";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "3")
            {
                vg.Banco = "tres";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "4")
            {
                vg.Banco = "quatro";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "5")
            {
                vg.Banco = "cinco";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "6")
            {
                vg.Banco = "seis";
            }
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE vagas SET "+vg.Banco+"='" +vg.Ultimo +"'WHERE curso= '" + vg.Curso + "'AND vestibulinho='"+vg.Vestibulinho+"'AND periodo='"+vg.Periodo+"'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void GravarSobras(Vagas vg)
        {
            if (vg.Banco.Remove(vg.Banco.Length - 9) == "1")
            {
                vg.Banco = "sobUm";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "2")
            {
                vg.Banco = "sobDois";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "3")
            {
                vg.Banco = "sobTres";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "4")
            {
                vg.Banco = "sobQuatro";
            }
            else if (vg.Banco.Remove(vg.Banco.Length - 9) == "5")
            {
                vg.Banco = "sobCinco";
            }
            else if (vg.Banco == "Pós Chamadas")
            {
                vg.Banco = "pos";
            }
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE vagas SET " + vg.Banco + "='" + vg.Sobra + "'WHERE curso= '" + vg.Curso + "'AND vestibulinho='" + vg.Vestibulinho + "'AND periodo='" + vg.Periodo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable VerificaChamada(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                if (vg.Banco.Remove(vg.Banco.Length - 9) == "2")
                {
                    vg.Banco = "dois";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "3")
                {
                    vg.Banco = "tres";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "4")
                {
                    vg.Banco = "quatro";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "5")
                {
                    vg.Banco = "cinco";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "6")
                {
                    vg.Banco = "seis";
                }
                string receber = "SELECT "+vg.Banco+" FROM vagas WHERE vestibulinho= '" + vg.Vestibulinho + "' AND curso= '" + vg.Curso + "'AND periodo= '"+vg.Periodo+"'";
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

        public DataTable VerificaSobra(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                if (vg.Banco.Remove(vg.Banco.Length - 9) == "1")
                {
                    vg.Banco = "sobUm";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "2")
                {
                    vg.Banco = "sobDois";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "3")
                {
                    vg.Banco = "sobTres";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "4")
                {
                    vg.Banco = "sobQuatro";
                }
                else if (vg.Banco.Remove(vg.Banco.Length - 9) == "5")
                {
                    vg.Banco = "sobCinco";
                }
                string receber = "SELECT " + vg.Banco + " FROM vagas WHERE vestibulinho= '" + vg.Vestibulinho + "' AND curso= '" + vg.Curso + "'AND periodo= '" + vg.Periodo + "'";
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

        public DataTable SegundaOpcao(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, habilitacao, periodo, clas2, nota, habilitacao2, periodo2, nome, dtNasc, endereco, telefone, celular, escol, matriculado, chamada, ausente, ausSegOp FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao2= '" + vg.Curso + "'AND periodo2='" + vg.Periodo + "' ORDER BY CAST(clas2 as unsigned integer)";
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

        public DataTable SegundaChamada(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, email, habilitacao, escol, matriculado, ausente FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND (chamada!= '2ª Opção' or chamada is NULL) AND CAST(clas as unsigned integer) >= '" + vg.Vaga + "' AND CAST(clas as unsigned integer) <= '" + vg.Sobra + "'ORDER BY CAST(clas as unsigned integer)";
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

        public DataTable DemaisOpcao(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, email, habilitacao, escol, matriculado, ausente FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola= '" + vg.Escola + "'AND(chamada!= '2ª Opção' or chamada is null) AND CAST(clas as unsigned integer) BETWEEN '" + vg.Ultimo + "' AND '" + vg.Sobra + "'ORDER BY CAST(clas as unsigned integer)";
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

        public DataTable PosChamadas(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, email, habilitacao, chamada, escol, matriculado, ausente FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND(chamada!= '2ª Opção' or chamada is NULL) AND CAST(clas as unsigned integer) > '" + vg.Ultimo + "' ORDER BY CAST(clas as unsigned integer)";
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

        public DataTable UltimaChamada()
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT chamada FROM datas";
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

        public void SalvarObs(Geral gr)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string alterar = "UPDATE geral SET observacao='" + gr.Observacao + "'WHERE cod = '" + gr.Codigo + "'";
                MySqlCommand comandos = new MySqlCommand(alterar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Observacao(Geral gr)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT observacao FROM geral WHERE cod= '" + gr.Codigo + "'";
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

        public DataTable Listao(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, habilitacao, escol, matriculado, ausente, listao, perListao FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao= '" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'AND CAST(clas as unsigned integer) BETWEEN '" + vg.Primeiro + "' AND '" + vg.Ultimo + "'ORDER BY CAST(clas as unsigned integer)";
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

        public DataTable ListaoPorNome(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT cod, clas, nota, nome, dtNasc, endereco, telefone, celular, habilitacao, escol, matriculado, chamada FROM geral WHERE nome LIKE '%" + vg.Nome + "%' AND vestibulinho= '" + vg.Vestibulinho + "'ORDER BY nome";
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

        public DataTable Pesquisa(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT clas, nota, nome, dtNasc, endereco, telefone, celular, habilitacao, escol, matriculado, chamada FROM geral WHERE vestibulinho= '" + vg.Vestibulinho + "'AND habilitacao='" + vg.Curso + "'AND periodo='" + vg.Periodo + "'AND escola='" + vg.Escola + "'ORDER BY CAST(clas as unsigned integer)";
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

        public DataTable NaoMatriculados(Vagas vg)
        {
            try
            {
                conexao = new MySqlConnection(caminho);
                conexao.Open();
                string receber = "SELECT clas, nota, nome, endereco, telefone, celular, habilitacao FROM geral WHERE vestibulinho='" + vg.Vestibulinho + "' AND (matriculado IS NULL OR matriculado='Não') AND ausente IS NULL AND ausSegOP IS NULL ORDER BY habilitacao, nome";
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
