using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BetterBotPMM
{
    class Agility
    {
        private SqlConnection connection = new SqlConnection();
        string nfe;
        string desc;
        string dis;
        string CNPJ;
        string Agente;
        string tipo = "IMP";

        string query;

        string house;

        public Agility(string notafiscale, string desconsolidacao, string discriminacao, string cnpj, string agente)
        {
            nfe = notafiscale;
            desc = desconsolidacao;
            dis = discriminacao;
            CNPJ = cnpj;
            Agente = agente;
        }

        public void BeginAnalysis()
        {
            findHAWB();
            InserirNoBancoAgility();
            

        }

        private void findHAWB()
        {

            int indexbegin = dis.IndexOf("Nr.Docto: ");
            if (indexbegin != -1)
            {
                indexbegin += 10;
                int indexEnd = dis.Substring(indexbegin).IndexOf(" ");
                indexEnd += indexbegin;
                house = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine(house);
            }
        }

        public void InserirNoBancoAgility()
        {
            //connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";
            InsertAgility:
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
                    goto InsertAgility;
                }
            }

            catch (Exception err)
            {
                connection.Close();
                Console.WriteLine(err.Message);
                goto InsertAgility;
            }
        }
    }
}
