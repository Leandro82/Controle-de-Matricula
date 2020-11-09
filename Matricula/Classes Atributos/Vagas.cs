using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Matricula
{
    class Vagas
    {
        private int nCod;
        private string nCurso;
        private string nNome;
        private string nPer;
        private string nEsc;
        private string nVagas;
        private string nChamada;
        private string nBanco;
        private DateTime nDtIn;
        private DateTime nDtFin;
        private string nVest;
        private string nUlt;
        private int nSob;
        private string nPri;
        private string nMat;
        private string nOc;

        public int Codigo
        {
            get { return nCod; }
            set { nCod = value; }
        }

        public string Curso
        {
            get { return nCurso; }
            set { nCurso = value; }
        }

        public string Periodo
        {
            get { return nPer; }
            set { nPer = value; }
        }

        public string Escola
        {
            get { return nEsc; }
            set { nEsc = value; }
        }

        public string Vaga
        {
            get { return nVagas; }
            set { nVagas = value; }
        }

        public string Chamada
        {
            get { return nChamada; }
            set { nChamada = value; }
        }

        public string Banco
        {
            get { return nBanco; }
            set { nBanco = value; }
        }

        public DateTime DtInicio
        {
            get { return nDtIn; }
            set { nDtIn = value; }
        }

        public DateTime DtFim
        {
            get { return nDtFin; }
            set { nDtFin = value; }
        }

        public string Vestibulinho
        {
            get { return nVest; }
            set { nVest = value; }
        }

        public int Sobra
        {
            get { return nSob; }
            set { nSob = value; }
        }

        public string Ultimo
        {
            get { return nUlt; }
            set { nUlt = value; }
        }

        public string Primeiro
        {
            get { return nPri; }
            set { nPri = value; }
        }

        public string Nome
        {
            get { return nNome; }
            set { nNome = value; }
        }

        public string Matriculado
        {
            get { return nMat; }
            set { nMat = value; }
        }

        public string Ocultar
        {
            get { return nOc; }
            set { nOc = value; }
        }
    }
}
