using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace AlgorithmTestConsole
{
    class SuperTerminais
    {
        private OleDbConnection connection = new OleDbConnection();

        string armazenagem1periodo = "";
        int armazenagem1periodoqnt = 0;

        string pesagem = "";
        int pesagemqnt = 0;

        string invoiceprice = "";
        string invoice = "";
        int invoiceqnt = 0;

        string handling = "";
        int handlingqnt = 0;

        string insnevasivacontainers = "";
        int insnevasivacontainersqnt = 0;

        string DIDOC = "";

        string utilizacaoservico = "";
        int utilizacaoservicoqnt = 0;

        string BL = "";

        string colocacaodelacre = "";
        int colocacaodelacreqnt = 0;

        string quebradolacre = "";
        int quebradolacreqnt = 0;

        string TransInt = "";
        int TransIntqnt = 0;

        string ovacao = "";
        int ovacaoqnt = 0;
        int ovacao20 = 0;
        int ovacao40 = 0;

        string desovacao = "";
        int desovacaoqnt = 0;
        int desovacao20 = 0;
        int desovacao40 = 0;

        string armazenagem1a2periodo = "";
        int armazenagem1a2periodoqnt = 0;

        string armazenagem1a3periodo = "";
        int armazenagem1a3periodoqnt = 0;

        string segurosulamerica = "";
        int segurosulamericaqnt = 0;

        string ispscode = "";
        int ispscodeqnt = 0;

        

        string query;

       

        public void StartAnalysis()
        {
            
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
Persist Security Info=False;";

            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "select NFe, DiscriminacaodoServico from Notas";
            command.CommandText = query;
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                string nfe = reader["NFe"].ToString();
                string discriminacao = reader["DiscriminacaodoServico"].ToString();
                Console.WriteLine(discriminacao);
                findArmazenagem1periodo(discriminacao);
                findPesagem(discriminacao);
                findInvoice(discriminacao);
                findHandling(discriminacao);
                findInsnevasivacontainers(discriminacao);
                findDIDOC(discriminacao);
                findUtilizacaoservico(discriminacao);
                findBL(discriminacao);
                findColocacaoDeLacre(discriminacao);
                findQuebraDoLacre(discriminacao);
                findTransIntContainer(discriminacao);
                findOvacao(discriminacao);
                findDesovacao(discriminacao);
                findArmazenagem1a2periodo(discriminacao);
                findArmazenagem1a3periodo(discriminacao);
                findSeguroSulAmerica(discriminacao);
                findISPSCode(discriminacao);
                inserirNoBancoSuperTerminais(nfe);
            }
            connection.Close();


        }

        private void inserirNoBancoSuperTerminais(string nfe)
        {
            // connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "insert into SuperTerminais (NFe , ARMAZENAGEM1 , ARMAZENAGEM1QNT , PESAGEM , PESAGEMQNT , INVOICE ,INVOICEPRICE, INVOICEQNT, HANDLING , HANDLINGQNT, INSNEVASIVACONTAINERS , INSNEVASIVACONTAINERSQNT, DIDOC , UTILIZACAODESERVICOS, UTILIZACAODESERVICOSQNT, BL , COLOCACAODOLACRE , COLOCACAODOLACREQNT , QUEBRADOLACRE , QUEBRADOLACREQNT, TRANSINTERNO , TRANSINTERNOQNT , OVACAO , OVACAO20 , OVACAO40 , OVACAOQNT , DESOVACAO, DESOVACAO20, DESOVACAO40, DESOVACAOQNT , ARMAZENAGEM1A2 , ARMAZENAGEM1A2QNT, ARMAZENAGEM1A3 , ARMAZENAGEM1A3QNT , SEGUROSULAMERICA , SEGUROSULAMERICAQNT , ISPSCODE , ISPSCODEQNT ) values ('" + nfe + "','" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + invoice + "','" + invoiceprice + "'," + invoiceqnt + ",'" + handling + "'," + handlingqnt + ",'" + insnevasivacontainers + "'," + insnevasivacontainersqnt + ",'" + DIDOC + "','" + utilizacaoservico + "'," + utilizacaoservicoqnt + ",'" + BL + "','" + colocacaodelacre + "'," + colocacaodelacreqnt + ",'" + quebradolacre + "'," + quebradolacreqnt + ",'" + TransInt + "'," + TransIntqnt + ",'" + ovacao + "'," + ovacao20 + "," + ovacao40 + "," + ovacaoqnt + ",'" + desovacao + "'," + desovacao20 + "," + desovacao40 + "," + desovacaoqnt + ",'" + armazenagem1a2periodo + "'," + armazenagem1a2periodoqnt + ",'" + armazenagem1a3periodo + "'," + armazenagem1a3periodoqnt + ",'" + segurosulamerica + "'," + segurosulamericaqnt + ",'" + ispscode + "'," + ispscodeqnt + ")";
            command.CommandText = query;
            command.ExecuteNonQuery();
            //  connection.Close();

        }

        private void findArmazenagem1periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM REFERENTE AO 1 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem1periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem1periodo);
                Console.WriteLine(armazenagem1periodoqnt);
            }
            else
            {
                armazenagem1periodo = "";
                armazenagem1periodoqnt = 0;
            }
        }

        private void findPesagem(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("PESAGEM");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                pesagem = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                pesagemqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(pesagem);
                Console.WriteLine(pesagemqnt);
            }
            else
            {
                pesagem = "";
                pesagemqnt = 0;
            }

        }

        private void findInvoice(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("INVOICE");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                int indexpara = discriminacao.Substring(indexbegin).IndexOf("(") - 1;
                indexpara += indexbegin;
                invoice = discriminacao.Substring(indexbegin + 10, indexpara - (indexbegin + 10));
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                invoiceprice = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                invoiceqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(invoiceprice);
                Console.WriteLine(invoiceqnt);
            }
            else
            {
                invoice = "";
                invoiceprice = "";
                invoiceqnt = 0;
            }
        }

        private void findHandling(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("HANDLING");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                handling = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                handlingqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(handling);
                Console.WriteLine(handlingqnt);
            }
            else
            {
                handling = "";
                handlingqnt = 0;
            }
        }

        private void findInsnevasivacontainers(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("INSPECAO NAO INVASIVA DE CONTAINERS");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                insnevasivacontainers = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                insnevasivacontainersqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(insnevasivacontainers);
                Console.WriteLine(insnevasivacontainersqnt);
            }
            else
            {
                insnevasivacontainers = "";
                insnevasivacontainersqnt = 0;
            }
        }

        private void findDIDOC(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("DI/DOC.:");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 9;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                DIDOC = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine(DIDOC);

            }
            else
            {
                DIDOC = "";

            }
        }

        private void findUtilizacaoservico(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("UTILIZACAO DE SERVICOS");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                utilizacaoservico = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                utilizacaoservicoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(utilizacaoservico);
                Console.WriteLine(utilizacaoservicoqnt);
            }
            else
            {
                utilizacaoservico = "";
                utilizacaoservicoqnt = 0;
            }
        }

        private void findBL(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("BL.:");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 5;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                BL = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine(BL);
            }
            else
            {
                BL = "";
            }
        }

        private void findColocacaoDeLacre(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("COLOCACAO DO LACRE");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                colocacaodelacre = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                colocacaodelacreqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(utilizacaoservico);
                Console.WriteLine(utilizacaoservicoqnt);
            }
            else
            {
                colocacaodelacre = "";
                colocacaodelacreqnt = 0;
            }
        }

        private void findQuebraDoLacre(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("QUEBRA DO LACRE");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                quebradolacre = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                quebradolacreqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(quebradolacre);
                Console.WriteLine(quebradolacreqnt);
            }
            else
            {
                quebradolacre = "";
                quebradolacreqnt = 0;
            }
        }

        private void findTransIntContainer(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("TRANSPORTE INTERNO DE CONTAINER");
            if (indexbegin == -1)
            {
                indexbegin = discriminacao.IndexOf("TRANSPORTE INTERNO (M.A./FUMIGACAO)");
            }
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                TransInt = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                TransIntqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(TransInt);
                Console.WriteLine(TransIntqnt);
            }
            else
            {
                TransInt = "";
                TransIntqnt = 0;
            }
        }

        private void findOvacao(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("OVACAO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                ovacao = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                ovacaoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                ovacao20 = int.Parse(discriminacao.Substring(indexbegin - 30, 1));
                ovacao40 = int.Parse(discriminacao.Substring(indexbegin - 9, 1));
                Console.WriteLine(ovacao);
                Console.WriteLine(ovacao20);
                Console.WriteLine(ovacao40);
                Console.WriteLine(ovacaoqnt);
            }
            else
            {
                ovacao = "";
                ovacaoqnt = 0;
                ovacao20 = 0;
                ovacao40 = 0;
            }
        }

        private void findDesovacao(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("DESOVA");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                desovacao = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                desovacaoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                desovacao20 = int.Parse(discriminacao.Substring(indexbegin - 30, 1));
                desovacao40 = int.Parse(discriminacao.Substring(indexbegin - 9, 1));
                Console.WriteLine(desovacao);
                Console.WriteLine(desovacao20);
                Console.WriteLine(desovacao40);
                Console.WriteLine(desovacaoqnt);
            }
            else
            {
                desovacao = "";
                desovacaoqnt = 0;
                desovacao20 = 0;
                desovacao40 = 0;
            }
        }

        private void findArmazenagem1a2periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM REFERENTE AOS PERIODOS DE 1 A 2");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem1a2periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1a2periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem1a2periodo);
                Console.WriteLine(armazenagem1a2periodoqnt);
            }
            else
            {
                armazenagem1a2periodo = "";
                armazenagem1a2periodoqnt = 0;
            }
        }

        private void findArmazenagem1a3periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM REFERENTE AOS PERIODOS DE 1 A 3");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem1a3periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1a3periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem1a3periodo);
                Console.WriteLine(armazenagem1a3periodoqnt);
            }
            else
            {
                armazenagem1a3periodo = "";
                armazenagem1a3periodoqnt = 0;
            }
        }

        private void findSeguroSulAmerica(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("COBERTURA DE SEGURO - CIA SULAMERICA");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                segurosulamerica = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                segurosulamericaqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(segurosulamerica);
                Console.WriteLine(segurosulamericaqnt);
            }
            else
            {
                segurosulamerica = "";
                segurosulamericaqnt = 0;
            }
        }

        private void findISPSCode(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("TARIFA ISPS CODE");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                ispscode = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                ispscodeqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(ispscode);
                Console.WriteLine(ispscodeqnt);
            }
            else
            {
                ispscode = "";
                ispscodeqnt = 0;
            }
        }



    }
}

