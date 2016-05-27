using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace NotasTerminaisDownload
{
    class AuroraEadi
    {
        private SqlConnection connection = new SqlConnection();

        private string dis;
        private string nfe;
        private string razaosocialnome;
        private Uri urlofnote;
        private string CNPJ;

        string DIDOC = "";

        string armazenagem1periodo;
        string armazenagemXperiodo;

        string periodo;
        string tipo;

        static string mes = DateTime.Now.ToString("MM");
        static string ano = DateTime.Now.ToString("yyyy");

        string query;

        public AuroraEadi(string discriminacao, string notafiscale , string nome , Uri url, string cnpj)
        {
            dis = discriminacao;
            nfe = notafiscale;
            razaosocialnome = nome;
            urlofnote = url;
            CNPJ = cnpj;
        }

        public void BeginAnalysis()
        {
            
            findDIDOC();
            if (findArmazenagem1periodo()==-1)
            {
                if (findArmazenagemXperiodo()==-1)
                {
                    System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\migue\Desktop\Discriminacao.txt", true);
                    file.WriteLine(razaosocialnome +" "+ dis);
                    file.Close();
                    return;
                }
                else
                {
                    inserirNoBanco(armazenagemXperiodo);
                }
            }
            else
            {
                inserirNoBanco(armazenagem1periodo);
            }

            




        }

        private void inserirNoBanco(string valor)
        {
            connection.ConnectionString = @"Data Source=192.168.0.110,59160;Initial Catalog=NotaTerminais;Persist Security Info=True;User ID=sa;Password=ca94404llc;Pooling=False";
            valor = valor.Replace(",", ".");
        InsertAurora:
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                query = "insert into Notas (Valor , Periodo , Terminal , Mes , Ano , NFe , Discriminacao , URL , CNPJPrestador , Tipo) values (" + valor + ",'" + periodo + "','" + razaosocialnome + "','" + mes + "','" + ano + "','" + nfe + "','" + dis + "','" + urlofnote + "','" + CNPJ + "','" + tipo + "')";
                command.CommandText = query;
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception err)
            {
                connection.Close();
                Console.WriteLine(err.Message);
                goto InsertAurora;
            }


        }

        private void findDIDOC()
        {
            int indexbegin = dis.IndexOf("DI:");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 4;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                DIDOC = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("DIDOC: " + DIDOC);

                if (DIDOC.IndexOf("/") != -1)
                {
                    tipo = "Importação";
                }
                else
                {
                    tipo = "Exportação";
                }

            }
            else
            {
                DIDOC = "";
                

            }

        }

        

        

        private int findArmazenagem1periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 1 PERIODO");
            int encontrou = indexbegin;
            if (indexbegin != -1)
            {
                periodo = "1";
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                
                armazenagem1periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                
                Console.WriteLine("Armazenagem 1: " + armazenagem1periodo);
                
            }
            else
            {
                armazenagem1periodo = "";
                
            }
            return encontrou;
        }

 

        private int findArmazenagemXperiodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 2 PERIODO");
            periodo = "2";

            if (indexbegin==-1)
            {
                indexbegin = dis.IndexOf("ARMAZENAGEM 3 PERIODO");
                periodo = "3";
            }
            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("ARMAZENAGEM 4 PERIODO");
                periodo = "4";
            }
            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("ARMAZENAGEM 5 PERIODO");
                periodo = "5";
            }
            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("ARMAZENAGEM 6 PERIODO");
                periodo = "6";
            }
            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("ARMAZENAGEM 7 PERIODO");
                periodo = "7";
            }
            if (indexbegin == -1)
            {
                periodo = "";
                return indexbegin;
            }

            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                armazenagemXperiodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                
                Console.WriteLine("Armazenagem X: " + armazenagemXperiodo);
                
            }
            else
            {
                armazenagemXperiodo = "";
                
            }
            return indexbegin;
        }


        
    }
}