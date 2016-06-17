using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BetterBotPMM
{
    class Chibatao
    {
        private SqlConnection connection = new SqlConnection();

        private string dis;
        private string nfe;

        string DIDOC = "";

        string armazenagem1;
        string armazenagem2;
        string armazenagem3;
        string armazenagem4;

        string armazenagemexportacao;
        int armazenagemexportacaoqnt;

        string pesagem;
        int pesagemqnt;

        string seguro;
        int seguroqnt;

        string inspnevasiva;
        int inspnevasivaqnt;

        string administrativo;
        int administrativoqnt;

        string colocacaodelacre;
        int colocacaodelacreqnt;

        string quebradolacre;
        int quebradolacreqnt;

        string TransInt;
        int TransIntqnt;

        string arquaviaria;
        int arquaviariaqnt;

        string ISPSCode;
        int ISPSCodeqnt;

        string query = "";

        public Chibatao(string discriminacao, string notafiscale)
        {
            dis = discriminacao;
            nfe = notafiscale;
        }

        public void BeginAnalysis()
        {
            if (findDIDOC())
            {
                findArmazenagem1();
                findPesagem();
                findSeguro();
                findInspnevasiva();
                findAdministrativo();
                findColocacaoLacre();
                findQuebraLacre();
                findTransporte();
                findArquaviaria();
                findISPSCode();
                findArmazenagem2();
                findArmazenagem3();
                findArmazenagem4();
                findArmazenagemExportacao();
                // findUnkown(discriminacao, nfe);
                inserirNoBancoChibatao();
            }

            

        }

        private void inserirNoBancoChibatao()
        {
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";
            InsertChibatao:
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                query = "insert into Chibatao (NFe , DIDOC , ARMAZENAGEM1 , PESAGEM , PESAGEMQNT , SEGURO , SEGUROQNT , INSNEVASIVACONTAINERS , INSNEVASIVACONTAINERSQNT , ADMINISTRATIVO , ADMINISTRATIVOQNT , COLOCACAODOLACRE , COLOCACAODOLACREQNT , QUEBRADOLACRE , QUEBRADOLACREQNT , TRANSINTERNO , TRANSINTERNOQNT , INFRAESTRUTURAAQUAVIARIA , INFRAESTRUTURAAQUAVIARIAQNT , ISPSCODE , ISPSCODEQNT , ARMAZENAGEM2 , ARMAZENAGEM3 , ARMAZENAGEM4 , ARMAZENAGEMEXPORTACAO , ARMAZENAGEMEXPORTACAOQNT) values ('" + nfe + "','" + DIDOC + "','" + armazenagem1 + "','" + pesagem + "'," + pesagemqnt + ",'" + seguro + "'," + seguroqnt + ",'" + inspnevasiva + "'," + inspnevasivaqnt + ",'" + administrativo + "'," + administrativoqnt + ",'" + colocacaodelacre + "'," + colocacaodelacreqnt + ",'" + quebradolacre + "'," + quebradolacreqnt + ",'" + TransInt + "'," + TransIntqnt + ",'" + arquaviaria + "'," + arquaviariaqnt + ",'" + ISPSCode + "'," + ISPSCodeqnt + ",'" + armazenagem2 + "','" + armazenagem3 + "','" + armazenagem4 + "','" + armazenagemexportacao + "'," + armazenagemexportacaoqnt + ")";
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
                    goto InsertChibatao;
                }
            }

            catch (Exception err)
            {
                connection.Close();
                Console.WriteLine(err.Message);
                goto InsertChibatao;
            }


        }


        private bool findDIDOC()
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

        private void findArmazenagem1()
        {
            int indexbegin = dis.IndexOf("REF. 1");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                // MessageBox.Show("hi");
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem1 = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("Armazenagem 1: "+ armazenagem1);
            }
            else
            {
                armazenagem1 = "";
            }
        }

        private void findArmazenagem2()
        {
            int indexbegin = dis.IndexOf("REF. 2");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                // MessageBox.Show("hi");
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem2 = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("Armazenagem 2: " + armazenagem2);
            }
            else
            {
                armazenagem2 = "";
            }
        }

        private void findArmazenagem3()
        {
            int indexbegin = dis.IndexOf("REF. 3");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                // MessageBox.Show("hi");
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem3 = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("Armazenagem 3: " + armazenagem3);
            }
            else
            {
                armazenagem3 = "";
            }
        }

        private void findArmazenagem4()
        {
            int indexbegin = dis.IndexOf("REF. 4");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                // MessageBox.Show("hi");
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem4 = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("Armazenagem 4: " + armazenagem4);
            }
            else
            {
                armazenagem4 = "";
            }
        }

        private void findPesagem()
        {
            int indexbegin = dis.IndexOf("PESAGEM DE CONTAINER");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                pesagem = dis.Substring(indexbegin, indexEnd - indexbegin);
                pesagemqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Pesagem: " + pesagem);
                //Console.WriteLine("Pesagem qnt: " + pesagemqnt);
            }
            else
            {
                pesagem = "";
                pesagemqnt = 0;
            }
        }

        private void findSeguro()
        {
            int indexbegin = dis.IndexOf("SEGURO DE CARGA");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                seguro = dis.Substring(indexbegin, indexEnd - indexbegin);
                seguroqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Seguro: " + seguro);
                //Console.WriteLine("Seguro qnt: " + seguroqnt);
            }
            else
            {
                seguro = "";
                seguroqnt = 0;
            }
        }

        private void findInspnevasiva()
        {
            int indexbegin = dis.IndexOf("VISTORIA NAO INVASIVA");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                inspnevasiva = dis.Substring(indexbegin, indexEnd - indexbegin);
                inspnevasivaqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Inspecao: " + inspnevasiva);
                //Console.WriteLine("Inspecao qnt: " + inspnevasivaqnt);
            }
            else
            {
                inspnevasiva = "";
                inspnevasivaqnt = 0;
            }
        }

        private void findAdministrativo()
        {
            int indexbegin = dis.IndexOf("SERVICOS ADMINISTRATIVOS");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                administrativo = dis.Substring(indexbegin, indexEnd - indexbegin);
                administrativoqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Administrativo: " + administrativo);
                //Console.WriteLine("Administrativo qnt: " + administrativoqnt);
            }
            else
            {
                administrativo = "";
                administrativoqnt = 0;
            }
        }

        private void findColocacaoLacre()
        {
            int indexbegin = dis.IndexOf("COLOCACAO DE LACRE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                colocacaodelacre = dis.Substring(indexbegin, indexEnd - indexbegin);
                colocacaodelacreqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Colocacao lacre: " + colocacaodelacre);
                //Console.WriteLine("Colocacao lacre qnt: " + colocacaodelacreqnt);
            }
            else
            {
                colocacaodelacre = "";
                colocacaodelacreqnt = 0;
            }
        }

        private void findQuebraLacre()
        {
            int indexbegin = dis.IndexOf("QUEBRE DO LACRE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                quebradolacre = dis.Substring(indexbegin, indexEnd - indexbegin);
                quebradolacreqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Quebra do lacre: " + quebradolacre);
                //Console.WriteLine("Quebra do lacre qnt: " + quebradolacreqnt);
            }
            else
            {
                quebradolacre = "";
                quebradolacreqnt = 0;
            }
        }

        private void findTransporte()
        {
            int indexbegin = dis.IndexOf("TRANSPORTE INTERNO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                TransInt = dis.Substring(indexbegin, indexEnd - indexbegin);
                TransIntqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Transporte: " + TransInt);
                //Console.WriteLine("Transporte qnt: " + TransIntqnt);
            }
            else
            {
                TransInt = "";
                TransIntqnt = 0;
            }
        }

        private void findArquaviaria()
        {
            int indexbegin = dis.IndexOf("INFRAESTRUTURA AQUAVIARIA");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                arquaviaria = dis.Substring(indexbegin, indexEnd - indexbegin);
                arquaviariaqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Aquaviaria: " + arquaviaria);
                //Console.WriteLine("Aquaviaria qnt: " + arquaviariaqnt);
            }
            else
            {
                arquaviaria = "";
                arquaviariaqnt = 0;
            }
        }

        private void findISPSCode()
        {
            int indexbegin = dis.IndexOf("MANUTENCAO DO ISPS-CODE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                ISPSCode = dis.Substring(indexbegin, indexEnd - indexbegin);
                ISPSCodeqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("ISPS: " + ISPSCode);
                //Console.WriteLine("ISPS qnt: " + ISPSCodeqnt);
            }
            else
            {
                ISPSCode = "";
                ISPSCodeqnt = 0;
            }
        }

        private void findArmazenagemExportacao()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM DE EXPORTACAO");
            if (indexbegin != -1)
            {
                int indexqnt = indexbegin - 2;
                int indexEnd = dis.Substring(indexbegin).IndexOf("|") - 1;
                indexEnd += indexbegin;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                armazenagemexportacao = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagemexportacaoqnt = int.Parse(dis.Substring(indexqnt, 1));
                Console.WriteLine("Armazenagem Exportacao: " + armazenagemexportacao);
            }
            else
            {
                armazenagemexportacao = "";
                armazenagemexportacaoqnt = 0;
            }
        }
    }
}