using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NotasTerminaisDownload
{
    class Chibatao
    {
        private SqlConnection connection = new SqlConnection();

        private string dis;
        private string nfe;
        private string razaosocialnome;
        private Uri urlofnote;
        private string CNPJ;

        string armazenagem1;
        string armazenagemX;
        string armazenagemexportacao;

        string DIDOC = "";

        string query = "";

        string periodo;

        string tipo;

        static string mes = DateTime.Now.ToString("MM");
        static string ano = DateTime.Now.ToString("yyyy");

        public Chibatao(string discriminacao, string notafiscale , string nome , Uri url, string cnpj)
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

            if (findArmazenagem1() == -1)
            {
                if (findArmazenagemX() == -1)
                {
                    if (findArmazenagemExportacao() == -1)
                    {
                        System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\migue\Desktop\Discriminacao.txt", true);
                        file.WriteLine(razaosocialnome + " " + dis);
                        file.Close();
                        return;
                    }
                    else
                    {
                        inserirNoBanco(armazenagemexportacao);
                    }
                    
                }
                else
                {
                    inserirNoBanco(armazenagemX);
                }
            }
            else
            {
                inserirNoBanco(armazenagem1);
            } 
             
            



        }

        private void inserirNoBanco(string valor)
        {
            connection.ConnectionString = @"Data Source=192.168.0.110,59160;Initial Catalog=NotaTerminais;Persist Security Info=True;User ID=sa;Password=ca94404llc;Pooling=False";
            valor = valor.Replace(",", ".");
        InsertChibatao:
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
                Console.WriteLine(err.Message);
                goto InsertChibatao;
            }


        }


        private void findDIDOC()
        {
            int indexbegin = dis.IndexOf("DI:");
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf("|");
                indexEnd += indexbegin;
                indexbegin += 3;
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

        private int findArmazenagem1()
        {
            int indexbegin = dis.IndexOf("REF. 1");
            int encontrou = indexbegin;

            if (indexbegin != -1)
            {
                periodo = "1";
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                armazenagem1 = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("Armazenagem 1: " + armazenagem1);
            }
            else
            {
                armazenagem1 = "";
            }
            return encontrou;
        }

        private int findArmazenagemX()
        {
            int indexbegin = dis.IndexOf("REF. 2");
            periodo = "2";

            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("REF. 3");
                periodo = "3";
            }

            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("REF. 4");
                periodo = "4";
            }

            if (indexbegin == -1)
            {
                periodo = "";
                return indexbegin;
            }

            if (indexbegin != -1)
            {
               
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                armazenagemX = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("Armazenagem X: " + armazenagemX);
            }
            else
            {
                armazenagemX = "";
            }
            return indexbegin;
        }

        private int findArmazenagemExportacao()
        {

            int indexbegin = dis.IndexOf("ARMAZENAGEM DE EXPORTACAO");
            int encontrou = indexbegin;

            if (indexbegin != -1)
            {
                periodo = "1";
                tipo = "Exportação";
                int indexEnd = dis.Substring(indexbegin).IndexOf("|");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                armazenagemexportacao = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("Armazenagem Exportacao: " + armazenagemexportacao);

            }
            else
            {
                armazenagemexportacao = "";
            }
            return encontrou;
        }


    }
}