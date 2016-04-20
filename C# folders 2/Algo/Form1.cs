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
                inserirNoBancoSuperTerminais(nfe);
            }
            connection.Close();


        }

        private void inserirNoBancoSuperTerminais(string nfe)
        {
            // connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "insert into SuperTerminais (NFe , ARMAZENAGEM1 , ARMAZENAGEM1QNT , PESAGEM , PESAGEMQNT , INVOICE , INVOICEQNT, HANDLING , HANDLINGQNT) values ('" + nfe + "','" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + invoice + "'," + invoiceqnt + ",'" + handling + "'," + handlingqnt + ")";
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
                int indexinvoice = discriminacao.Substring(indexbegin).IndexOf(":");
                int indexpara = discriminacao.Substring(indexbegin).IndexOf("(");
                invoice = discriminacao.Substring(indexinvoice + 1, indexpara - indexinvoice);
                MessageBox.Show(invoice);
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
            Console.WriteLine(indexbegin + "HANDLING");
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

    }
}
