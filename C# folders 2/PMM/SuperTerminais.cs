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

namespace BotPMM
{
    class SuperTerminais
    {
        private OleDbConnection connection = new OleDbConnection();

        private string dis;
        private string nfe;

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

        string query;

        public SuperTerminais(string discriminacao, string notafiscale)
        {
            dis = discriminacao;
            nfe = notafiscale;
        }

        public Boolean BeginAnalysis()
        {
            //Console.WriteLine(dis);
            if (findDIDOC())
            {
                findArmazenagem1periodo();
                findPesagem();
                findInvoice();
                findHandling();
                findInsnevasivacontainers();
                findUtilizacaoservico();
                findBL();
                inserirNoBancoSuperTerminais();
                return true;
            }
            else {
                return false;
            }
            
            
            
        }

        private void inserirNoBancoSuperTerminais()
        {
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\migue\OneDrive\Documentos\Notas.accdb;
Persist Security Info=False;";
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "insert into SuperTerminais (NFe , ARMAZENAGEM1 , ARMAZENAGEM1QNT , PESAGEM , PESAGEMQNT , INVOICE ,INVOICEPRICE, INVOICEQNT, HANDLING , HANDLINGQNT, INSNEVASIVACONTAINERS , INSNEVASIVACONTAINERSQNT, DIDOC , UTILIZACAODESERVICOS, UTILIZACAODESERVICOSQNT, BL) values ('" + nfe + "','" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + invoice + "','" + invoiceprice + "'," + invoiceqnt + ",'" + handling + "'," + handlingqnt + ",'" + insnevasivacontainers + "'," + insnevasivacontainersqnt + ",'" + DIDOC + "','" + utilizacaoservico + "'," + utilizacaoservicoqnt + ",'" + BL + "')";
            command.CommandText = query;
            command.ExecuteNonQuery();
            connection.Close();

        }

        private void findArmazenagem1periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM REFERENTE AO 1 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem1periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1periodoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Armazenagem 1 periodo: " + armazenagem1periodo);
                //Console.WriteLine(armazenagem1periodoqnt);
            }
            else
            {
                armazenagem1periodo = "";
                armazenagem1periodoqnt = 0;
            }
        }

        private void findPesagem()
        {
            int indexbegin = dis.IndexOf("PESAGEM");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                pesagem = dis.Substring(indexbegin, indexEnd - indexbegin);
                pesagemqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Pesagem: " + pesagem);
                //Console.WriteLine(pesagemqnt);
            }
            else
            {
                pesagem = "";
                pesagemqnt = 0;
            }

        }

        private void findInvoice()
        {
            int indexbegin = dis.IndexOf("INVOICE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                int indexpara = dis.Substring(indexbegin).IndexOf("(") - 1;
                indexpara += indexbegin;
                invoice = dis.Substring(indexbegin + 10, indexpara - (indexbegin + 10));
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                invoiceprice = dis.Substring(indexbegin, indexEnd - indexbegin);
                invoiceqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Invoice: " + invoiceprice);
                //Console.WriteLine(invoiceqnt);
            }
            else
            {
                invoice = "";
                invoiceprice = "";
                invoiceqnt = 0;
            }
        }
        private void findHandling()
        {
            int indexbegin = dis.IndexOf("HANDLING");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                handling = dis.Substring(indexbegin, indexEnd - indexbegin);
                handlingqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Handling: " + handling);
                //Console.WriteLine(handlingqnt);
            }
            else
            {
                handling = "";
                handlingqnt = 0;
            }
        }

        private void findInsnevasivacontainers()
        {
            int indexbegin = dis.IndexOf("INSPECAO NAO INVASIVA DE CONTAINERS");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                insnevasivacontainers = dis.Substring(indexbegin, indexEnd - indexbegin);
                insnevasivacontainersqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Inspecao: " + insnevasivacontainers);
                //Console.WriteLine(insnevasivacontainersqnt);
            }
            else
            {
                insnevasivacontainers = "";
                insnevasivacontainersqnt = 0;
            }
        }

        private bool findDIDOC()
        {
            int indexbegin = dis.IndexOf("DI/DOC.:");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 9;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                DIDOC = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("DIDOC: " + DIDOC);

                if (DIDOC.IndexOf("/") != -1)
                {
                    return true;
                }
                else {
                    return false;
                }

            }
            else
            {
                DIDOC = "";
                return false;

            }
            
        }

        private void findUtilizacaoservico()
        {
            int indexbegin = dis.IndexOf("UTILIZACAO DE SERVICOS");
           // Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
               // Console.WriteLine(indexbegin);
               // Console.WriteLine(indexEnd);
                utilizacaoservico = dis.Substring(indexbegin, indexEnd - indexbegin);
                utilizacaoservicoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Servico: " + utilizacaoservico);
               // Console.WriteLine(utilizacaoservicoqnt);
            }
            else
            {
                utilizacaoservico = "";
                utilizacaoservicoqnt = 0;
            }
        }

        private void findBL()
        {
            int indexbegin = dis.IndexOf("BL.:");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 5;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                BL = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("BL: "+ BL);
            }
            else
            {
                BL = "";
            }
        }

    }
}

