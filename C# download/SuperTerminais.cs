using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotasTerminaisDownload
{
    class SuperTerminais
    {
        //OleDbCommand connection = new OleDbCommand();
        private OleDbConnection connection = new OleDbConnection();

        private string dis;
        private string nfe;

        string armazenagem1periodo = "";
        int armazenagem1periodoqnt = 0;

        string armazenagem1a2periodo = "";
        int armazenagem1a2periodoqnt = 0;

        string armazenagem1a3periodo = "";
        int armazenagem1a3periodoqnt = 0;

        string query;



        public SuperTerminais(string discriminacao, string notafiscale)
        {
            dis = discriminacao;
            nfe = notafiscale;
        }

        public void BeginAnalysis()
        {

            if (findDIDOC())
            {
                findArmazenagem1periodo();
                findPesagem();
                findInvoice();
                findHandling();
                findInsnevasivacontainers();
                findUtilizacaoservico();
                findBL();
                findColocacaoDeLacre();
                findQuebraDoLacre();
                findTransIntContainer();
                findOvacao();
                findDesovacao();
                findArmazenagem1a2periodo();
                findArmazenagem1a3periodo();
                findSeguroSulAmerica();
                findISPSCode();
                inserirNoBancoSuperTerminais();

            }
        }

        private void inserirNoBancoSuperTerminais()
        {
            //connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
            Persist Security Info=False;";
            InsertSuperTerminais:
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                query = "insert into SuperTerminais (NFe , ARMAZENAGEM1 , ARMAZENAGEM1QNT , PESAGEM , PESAGEMQNT , INVOICE ,INVOICEPRICE, INVOICEQNT, HANDLING , HANDLINGQNT, INSNEVASIVACONTAINERS , INSNEVASIVACONTAINERSQNT, DIDOC , UTILIZACAODESERVICOS, UTILIZACAODESERVICOSQNT, BL , COLOCACAODOLACRE , COLOCACAODOLACREQNT , QUEBRADOLACRE , QUEBRADOLACREQNT, TRANSINTERNO , TRANSINTERNOQNT , OVACAO , OVACAO20 , OVACAO40 , OVACAOQNT , DESOVACAO, DESOVACAO20, DESOVACAO40, DESOVACAOQNT , ARMAZENAGEM1A2 , ARMAZENAGEM1A2QNT, ARMAZENAGEM1A3 , ARMAZENAGEM1A3QNT , SEGUROSULAMERICA , SEGUROSULAMERICAQNT , ISPSCODE , ISPSCODEQNT ) values ('" + nfe + "','" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + invoice + "','" + invoiceprice + "'," + invoiceqnt + ",'" + handling + "'," + handlingqnt + ",'" + insnevasivacontainers + "'," + insnevasivacontainersqnt + ",'" + DIDOC + "','" + utilizacaoservico + "'," + utilizacaoservicoqnt + ",'" + BL + "','" + colocacaodelacre + "'," + colocacaodelacreqnt + ",'" + quebradolacre + "'," + quebradolacreqnt + ",'" + TransInt + "'," + TransIntqnt + ",'" + ovacao + "'," + ovacao20 + "," + ovacao40 + "," + ovacaoqnt + ",'" + desovacao + "'," + desovacao20 + "," + desovacao40 + "," + desovacaoqnt + ",'" + armazenagem1a2periodo + "'," + armazenagem1a2periodoqnt + ",'" + armazenagem1a3periodo + "'," + armazenagem1a3periodoqnt + ",'" + segurosulamerica + "'," + segurosulamericaqnt + ",'" + ispscode + "'," + ispscodeqnt + ")";
                command.CommandText = query;
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                goto InsertSuperTerminais;
            }


        }

        private void findArmazenagem1periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM REFERENTE AO 1 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem1periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1periodoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Armazenagem 1 periodo: " + armazenagem1periodo);
                //Console.WriteLine(armazenagem1periodoqnt);
            }
            else
            {
                armazenagem1periodo = "";
                armazenagem1periodoqnt = 0;
            }
        }

        

       
       

       

      

      

      

        

        private void findQuebraDoLacre()
        {
            int indexbegin = dis.IndexOf("QUEBRA DO LACRE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                quebradolacre = dis.Substring(indexbegin, indexEnd - indexbegin);
                quebradolacreqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Quebra do lacre: " + quebradolacre);
                //Console.WriteLine(quebradolacreqnt);
            }
            else
            {
                quebradolacre = "";
                quebradolacreqnt = 0;
            }
        }

        private void findTransIntContainer()
        {
            int indexbegin = dis.IndexOf("TRANSPORTE INTERNO DE CONTAINER");
            if (indexbegin == -1)
            {
                indexbegin = dis.IndexOf("TRANSPORTE INTERNO (M.A./FUMIGACAO)");
            }
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                TransInt = dis.Substring(indexbegin, indexEnd - indexbegin);
                TransIntqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Transporte: " + TransInt);
                //Console.WriteLine(TransIntqnt);
            }
            else
            {
                TransInt = "";
                TransIntqnt = 0;
            }
        }

        private void findOvacao()
        {
            int indexbegin = dis.IndexOf("OVACAO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                ovacao = dis.Substring(indexbegin, indexEnd - indexbegin);
                ovacaoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                ovacao20 = int.Parse(dis.Substring(indexbegin - 30, 1));
                ovacao40 = int.Parse(dis.Substring(indexbegin - 9, 1));
                Console.WriteLine("Ovacao: " + ovacao);
                Console.WriteLine("Ovacao 20: " + ovacao20);
                Console.WriteLine("Ovacao 40: " + ovacao40);
                //Console.WriteLine(ovacaoqnt);
            }
            else
            {
                ovacao = "";
                ovacaoqnt = 0;
                ovacao20 = 0;
                ovacao40 = 0;
            }
        }

        private void findDesovacao()
        {
            int indexbegin = dis.IndexOf("DESOVA");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                desovacao = dis.Substring(indexbegin, indexEnd - indexbegin);
                desovacaoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                desovacao20 = int.Parse(dis.Substring(indexbegin - 30, 1));
                desovacao40 = int.Parse(dis.Substring(indexbegin - 9, 1));
                Console.WriteLine("Desovacao: " + desovacao);
                Console.WriteLine("Desovacao 20: " + desovacao20);
                Console.WriteLine("Desovacao 40: " + desovacao40);
                //Console.WriteLine(desovacaoqnt);
            }
            else
            {
                desovacao = "";
                desovacaoqnt = 0;
                desovacao20 = 0;
                desovacao40 = 0;
            }
        }

        private void findArmazenagem1a2periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM REFERENTE AOS PERIODOS DE 1 A 2");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem1a2periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1a2periodoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Armazenagem 1 a 2: " + armazenagem1a2periodo);
                //Console.WriteLine("Armazenagem 1 a 2: " + armazenagem1a2periodoqnt);
            }
            else
            {
                armazenagem1a2periodo = "";
                armazenagem1a2periodoqnt = 0;
            }
        }

        private void findArmazenagem1a3periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM REFERENTE AOS PERIODOS DE 1 A 3");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem1a3periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1a3periodoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Armazenagem 1 a 3: " + armazenagem1a3periodo);
                //Console.WriteLine(armazenagem1a3periodoqnt);
            }
            else
            {
                armazenagem1a3periodo = "";
                armazenagem1a3periodoqnt = 0;
            }
        }

        private void findSeguroSulAmerica()
        {
            int indexbegin = dis.IndexOf("COBERTURA DE SEGURO - CIA SULAMERICA");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                segurosulamerica = dis.Substring(indexbegin, indexEnd - indexbegin);
                segurosulamericaqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Seguro sul america: " + segurosulamerica);
                //Console.WriteLine(segurosulamericaqnt);
            }
            else
            {
                segurosulamerica = "";
                segurosulamericaqnt = 0;
            }
        }

        private void findISPSCode()
        {
            int indexbegin = dis.IndexOf("TARIFA ISPS CODE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                ispscode = dis.Substring(indexbegin, indexEnd - indexbegin);
                ispscodeqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Ips: " + ispscode);
                //Console.WriteLine(ispscodeqnt);
            }
            else
            {
                ispscode = "";
                ispscodeqnt = 0;
            }
        }
    }
}
