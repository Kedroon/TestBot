using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BetterBotPMM
{
    class KN
    {
        private SqlConnection connection = new SqlConnection();
        string nfe;
        string desc;
        string dis;
        string CNPJ;
        string Agente;

        string query;

        string tipo;

        string house;

        public KN(string notafiscale, string desconsolidacao, string discriminacao, string cnpj, string agente)
        {
            nfe = notafiscale;
            desc = desconsolidacao;
            dis = discriminacao;
            CNPJ = cnpj;
            Agente = agente;
        }

        public void BeginAnalysis()
        {
            if (findIfImp() == true)
            {
                tipo = "IMP";
                InserirNoBancoKN();
            }
            else
            {
                tipo = "EXP";
                InserirNoBancoKN();
            }

        }

        public void InserirNoBancoKN()
        {
            //connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";
            InsertKN:
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                desc = desc.Replace(".", "");
                desc = desc.Replace(",", ".");
                query = "insert into Agentes (Desconsolidacao , NFe , HOUSE , Tipo , Data , Chave , Agente) values (" + desc + ",'" + nfe + "','" + house + "','" + tipo + "', CURRENT_TIMESTAMP,'" + nfe + CNPJ + "','" + Agente + "')";
                command.CommandText = query;
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (SqlException exception)
            {
                if (exception.Number == 2627) // Cannot insert duplicate key row in object error
                {
                    connection.Close();
                    return;
                }
                else
                {
                    connection.Close();
                    goto InsertKN;
                }
            }

            catch (Exception err)
            {
                connection.Close();
                Console.WriteLine(err.Message);
                goto InsertKN;
            }
        }

        private bool findIfImp()
        {

            int indexbegin = dis.IndexOf("HAWB.");
            int mais = 6;
            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("BL");
                mais = 3;
            }
            if (indexbegin != -1)
            {
                indexbegin += mais;
                int indexEnd = dis.Substring(indexbegin).IndexOf(" ");
                indexEnd += indexbegin;
                house = dis.Substring(indexbegin, indexEnd - indexbegin);
                house = house.Replace(" ", "").Replace("-", "").Replace(".", "");
                Console.WriteLine(house);

            }
            OracleConnection oracleConnection = new OracleConnection();
            bool encontrou = false;
            QueryOracle:
            try
            {
                oracleConnection.ConnectionString = "Data Source=(DESCRIPTION =" +
                "(ADDRESS = (PROTOCOL = TCP)(HOST = hda01132)(PORT = 1521))" +
                "(CONNECT_DATA =" +
                "(SERVER = DEDICATED)" +
                "(SERVICE_NAME = SFWPRD)));" +
                "User Id=SB022613;Password=SB022613;";
                oracleConnection.Open();
                OracleCommand oracleCommand = new OracleCommand();
                oracleCommand.Connection = oracleConnection;
                string oraclequery = "select HOUSE from sfwishmm.processos_importacao where HOUSE = '" + house + "'";
                oracleCommand.CommandText = oraclequery;
                OracleDataReader oracleReader = oracleCommand.ExecuteReader();
                while (oracleReader.Read())
                {
                    encontrou = true;
                    oracleConnection.Close();
                    return true;

                }
                if (!encontrou)
                {
                    oracleConnection.Close();
                    return false;
                }

            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                oracleConnection.Close();
                Thread.Sleep(2000);
                goto QueryOracle;
            }
            return false;
        }

    }
}
