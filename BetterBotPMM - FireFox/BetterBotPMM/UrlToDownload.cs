using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BetterBotPMM
{
    class UrlToDownload
    {
        public Uri URI;
        public string CNPJ;

        public UrlToDownload(Uri uri, string cnpj)
        {
            URI = uri;
            CNPJ = cnpj;

        }
    }
}
