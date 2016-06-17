using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BetterBotPMM
{
    class Expeditors
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

        public Expeditors(string notafiscale, string desconsolidacao,string discriminacao , string cnpj , string agente)
        {
            nfe = notafiscale;
            desc = desconsolidacao;
            dis = discriminacao;
            CNPJ = cnpj;
            Agente = agente;
        }

        public void BeginAnalysis()
        {
            if (findIfImp()==true)
            {
                tipo = "IMP";
                InserirNoBancoExpeditors();
            }
            else
            {
                tipo = "EXP";
                InserirNoBancoExpeditors();
            }
            
        }

        public void InserirNoBancoExpeditors()
        {
            //connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";
            InsertExpeditors:
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                desc = desc.Replace(".", "");
                desc = desc.Replace(",", ".");
                query = "insert into Agentes (Desconsolidacao , NFe , HOUSE , Tipo , Data , Chave , Agente) values (" + desc + ",'" + nfe + "','" + house + "','" + tipo + "', CURRENT_TIMESTAMP,'"+nfe+CNPJ+ "','" + Agente + "')";
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
                    goto InsertExpeditors;
                }
            }

            catch (Exception err)
            {
                connection.Close();
                Console.WriteLine(err.Message);
                goto InsertExpeditors;
            }
        }

        private bool findIfImp()
        {
            int indexbegin = dis.IndexOf("4401 -");
            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("4007 -");
            }
            

            if (indexbegin != -1)
            {
                int indexHouse = dis.IndexOf("HAWB:");
                indexHouse += 5;
                int indexEnd = dis.Substring(indexHouse).IndexOf(" ");
                indexEnd += indexHouse;
                house = dis.Substring(indexHouse, indexEnd - indexHouse);
                Console.WriteLine(house);
                return true;

            }
            else
            {
                int indexHouse = dis.IndexOf("LOG:");
                indexHouse += 5;
                int indexEnd = dis.Substring(indexHouse).IndexOf(" ");
                indexEnd += indexHouse;
                house = dis.Substring(indexHouse, indexEnd - indexHouse);
                Console.WriteLine(house);
                return false;

            }
        }

    }
}
