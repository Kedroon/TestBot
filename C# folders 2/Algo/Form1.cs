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

namespace AlgorithmTest
{
    public partial class Form1 : Form
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

        string query;

        public Form1()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
Persist Security Info=False;";
        }

        private void button1_Click(object sender, EventArgs e)
        {


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
                inserirNoBancoSuperTerminais(nfe);
            }
            connection.Close();


        }

        private void inserirNoBancoSuperTerminais(string nfe)
        {
            // connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "insert into SuperTerminais (NFe , ARMAZENAGEM1 , ARMAZENAGEM1QNT , PESAGEM , PESAGEMQNT , INVOICE ,INVOICEPRICE, INVOICEQNT, HANDLING , HANDLINGQNT, INSNEVASIVACONTAINERS , INSNEVASIVACONTAINERSQNT, DIDOC , UTILIZACAODESERVICOS, UTILIZACAODESERVICOSQNT, BL) values ('" + nfe + "','" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + invoice + "','" + invoiceprice + "'," + invoiceqnt + ",'" + handling + "'," + handlingqnt + ",'" + insnevasivacontainers + "'," + insnevasivacontainersqnt + ",'" + DIDOC + "','" + utilizacaoservico + "'," + utilizacaoservicoqnt + ",'" + BL + "')";
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

    }
}
