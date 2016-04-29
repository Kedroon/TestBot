using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace AlgorithmTestConsole
{
    class Chibatao
    {
        private OleDbConnection connection = new OleDbConnection();

        private string dis;
        private string nfe;

        string DIDOC = "";

        string query = "";

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
                findDIDOC(discriminacao);
                inserirNoBancoSuperTerminais(nfe);

            }
            connection.Close();


        }

        private void inserirNoBancoSuperTerminais(string nfe)
        {
            // connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "insert into Chibatao (NFe , DIDOC) values ('"+nfe+"','"+DIDOC+"')";
            command.CommandText = query;
            command.ExecuteNonQuery();
            //  connection.Close();

        }

        private bool findDIDOC(string dis)
        {
            int indexbegin = dis.IndexOf("DI:");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf("|");
                indexEnd += indexbegin;
                indexbegin += 3;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                DIDOC = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("DIDOC: " + DIDOC);

                if (DIDOC.IndexOf("/") != -1)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            else
            {
                DIDOC = "";
                return false;

            }

        }

    }
}
