using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotasTerminaisDownload
{
    class CNPJValido
    {
        public bool valido;
        public string CNPJ;
        public string Nota;

        public CNPJValido(bool val, string cnpj, string nota)
        {

            valido = val;
            CNPJ = cnpj;
            Nota = nota;
        }
    }
}