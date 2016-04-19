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
        string query;
        
        public Form1()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
Persist Security Info=False;";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string armazenagem1periodo = "";
            int armazenagem1periodoqnt = 0;
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "select DiscriminacaodoServico from Notas";
            command.CommandText = query;
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            reader.Read();
            string discriminacao = reader["DiscriminacaodoServico"].ToString();
            Console.WriteLine(discriminacao);
            int indexArmazenagem1 = discriminacao.LastIndexOf("ARMAZENAGEM REFERENTE AO 1 PERIODO");
            Console.WriteLine(indexArmazenagem1);
            if (indexArmazenagem1 != -1)
            {
                indexArmazenagem1 += 78;
                Console.WriteLine(indexArmazenagem1);
                int indexArmazenagem1End = discriminacao.IndexOf(";");
                indexArmazenagem1End -= 1;
                armazenagem1periodo = discriminacao.Substring(indexArmazenagem1, indexArmazenagem1End - indexArmazenagem1);
                armazenagem1periodoqnt = int.Parse(discriminacao.Substring(indexArmazenagem1 - 6, 1));
                Console.WriteLine(armazenagem1periodo);
                Console.WriteLine(armazenagem1periodoqnt);
            }
            
            connection.Close();
            inserirNoBancoSuperTerminais(armazenagem1periodo,armazenagem1periodoqnt);

        }

        private void inserirNoBancoSuperTerminais(string armazenagem1 , int armazenagemqnt) {
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "insert into SuperTerminais (ARMAZENAGEM1 , ARMAZENAGEM1QNT) values ('"+armazenagem1+"',"+armazenagemqnt+")";
            command.CommandText = query;
            command.ExecuteNonQuery();
            connection.Close();

        }

    }
}
