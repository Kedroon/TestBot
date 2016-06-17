using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace BetterBotPMM
{
    class SuperTerminais
    {
        //OleDbCommand connection = new OleDbCommand();
        private SqlConnection connection = new SqlConnection();

        private string dis;
        private string nfe;

        string armazenagem1periodo = "";
        int armazenagem1periodoqnt = 0;

        string pesagem = "";
        int pesagemqnt = 0;

        string invoiceprice = "";
        string invoice = "";
        int invoiceqnt = 0;

        string handling = "";
        int handlingqnt = 0;

        string insnevasivacontainers = "";
        int insnevasivacontainersqnt = 0;

        string DIDOC = "";

        string utilizacaoservico = "";
        int utilizacaoservicoqnt = 0;

        string BL = "";

        string colocacaodelacre = "";
        int colocacaodelacreqnt = 0;

        string quebradolacre = "";
        int quebradolacreqnt = 0;

        string TransInt = "";
        int TransIntqnt = 0;

        string ovacao = "";
        int ovacaoqnt = 0;
        int ovacao20 = 0;
        int ovacao40 = 0;

        string desovacao = "";
        int desovacaoqnt = 0;
        int desovacao20 = 0;
        int desovacao40 = 0;

        string armazenagem1a2periodo = "";
        int armazenagem1a2periodoqnt = 0;

        string armazenagem1a3periodo = "";
        int armazenagem1a3periodoqnt = 0;

        string armazenagem1a4periodo = "";
        int armazenagem1a4periodoqnt = 0;

        string armazenagemexportacao = "";
        int armazenagemexportacaoqnt = 0;

        string segurosulamerica = "";
        int segurosulamericaqnt = 0;

        string ispscode = "";
        int ispscodeqnt = 0;



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
                findArmazenagem1a4periodo();
                findArmazenagemExportacao();
                inserirNoBancoSuperTerminais();

            }
        }

        private void inserirNoBancoSuperTerminais()
        {
            //connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";
            InsertSuperTerminais:
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                query = "insert into SuperTerminais (NFe , ARMAZENAGEM1 , ARMAZENAGEM1QNT , PESAGEM , PESAGEMQNT , INVOICE ,INVOICEPRICE, INVOICEQNT, HANDLING , HANDLINGQNT, INSNEVASIVACONTAINERS , INSNEVASIVACONTAINERSQNT, DIDOC , UTILIZACAODESERVICOS, UTILIZACAODESERVICOSQNT, BL , COLOCACAODOLACRE , COLOCACAODOLACREQNT , QUEBRADOLACRE , QUEBRADOLACREQNT, TRANSINTERNO , TRANSINTERNOQNT , OVACAO , OVACAO20 , OVACAO40 , OVACAOQNT , DESOVACAO, DESOVACAO20, DESOVACAO40, DESOVACAOQNT , ARMAZENAGEM1A2 , ARMAZENAGEM1A2QNT, ARMAZENAGEM1A3 , ARMAZENAGEM1A3QNT , SEGUROSULAMERICA , SEGUROSULAMERICAQNT , ISPSCODE , ISPSCODEQNT  , ARMAZENAGEM1A4 , ARMAZENAGEM1A4QNT , ARMAZENAGEMEXPORTACAO , ARMAZENAGEMEXPORTACAOQNT) values ('" + nfe + "','" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + invoice + "','" + invoiceprice + "'," + invoiceqnt + ",'" + handling + "'," + handlingqnt + ",'" + insnevasivacontainers + "'," + insnevasivacontainersqnt + ",'" + DIDOC + "','" + utilizacaoservico + "'," + utilizacaoservicoqnt + ",'" + BL + "','" + colocacaodelacre + "'," + colocacaodelacreqnt + ",'" + quebradolacre + "'," + quebradolacreqnt + ",'" + TransInt + "'," + TransIntqnt + ",'" + ovacao + "'," + ovacao20 + "," + ovacao40 + "," + ovacaoqnt + ",'" + desovacao + "'," + desovacao20 + "," + desovacao40 + "," + desovacaoqnt + ",'" + armazenagem1a2periodo + "'," + armazenagem1a2periodoqnt + ",'" + armazenagem1a3periodo + "'," + armazenagem1a3periodoqnt + ",'" + segurosulamerica + "'," + segurosulamericaqnt + ",'" + ispscode + "'," + ispscodeqnt + ",'" + armazenagem1a4periodo + "'," + armazenagem1a4periodoqnt + ",'" + armazenagemexportacao + "'," + armazenagemexportacaoqnt + ")";
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
                    goto InsertSuperTerminais;
                }
            }

            catch (Exception err)
            {
                connection.Close();
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

        private void findPesagem()
        {
            int indexbegin = dis.IndexOf("PESAGEM");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                pesagem = dis.Substring(indexbegin, indexEnd - indexbegin);
                pesagemqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Pesagem: " + pesagem);
                //Console.WriteLine(pesagemqnt);
            }
            else
            {
                pesagem = "";
                pesagemqnt = 0;
            }

        }

        private void findInvoice()
        {
            int indexbegin = dis.IndexOf("INVOICE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                int indexpara = dis.Substring(indexbegin).IndexOf("(") - 1;
                indexpara += indexbegin;
                invoice = dis.Substring(indexbegin + 10, indexpara - (indexbegin + 10));
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                invoiceprice = dis.Substring(indexbegin, indexEnd - indexbegin);
                invoiceqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Invoice: " + invoiceprice);
                //Console.WriteLine(invoiceqnt);
            }
            else
            {
                invoice = "";
                invoiceprice = "";
                invoiceqnt = 0;
            }
        }
        private void findHandling()
        {
            int indexbegin = dis.IndexOf("HANDLING");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                handling = dis.Substring(indexbegin, indexEnd - indexbegin);
                handlingqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Handling: " + handling);
                //Console.WriteLine(handlingqnt);
            }
            else
            {
                handling = "";
                handlingqnt = 0;
            }
        }

        private void findInsnevasivacontainers()
        {
            int indexbegin = dis.IndexOf("INSPECAO NAO INVASIVA DE CONTAINERS");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                insnevasivacontainers = dis.Substring(indexbegin, indexEnd - indexbegin);
                insnevasivacontainersqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Inspecao: " + insnevasivacontainers);
                //Console.WriteLine(insnevasivacontainersqnt);
            }
            else
            {
                insnevasivacontainers = "";
                insnevasivacontainersqnt = 0;
            }
        }

        private bool findDIDOC()
        {
            int indexbegin = dis.IndexOf("DI/DOC.:");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 9;
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

        private void findUtilizacaoservico()
        {
            int indexbegin = dis.IndexOf("UTILIZACAO DE SERVICOS");
            // Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                // Console.WriteLine(indexbegin);
                // Console.WriteLine(indexEnd);
                utilizacaoservico = dis.Substring(indexbegin, indexEnd - indexbegin);
                utilizacaoservicoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Servico: " + utilizacaoservico);
                // Console.WriteLine(utilizacaoservicoqnt);
            }
            else
            {
                utilizacaoservico = "";
                utilizacaoservicoqnt = 0;
            }
        }

        private void findBL()
        {
            int indexbegin = dis.IndexOf("BL.:");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 5;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                BL = dis.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine("BL: " + BL);
            }
            else
            {
                BL = "";
            }
        }

        private void findColocacaoDeLacre()
        {
            int indexbegin = dis.IndexOf("COLOCACAO DO LACRE");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                colocacaodelacre = dis.Substring(indexbegin, indexEnd - indexbegin);
                colocacaodelacreqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Colocao lacre: " + colocacaodelacre);
                //Console.WriteLine(utilizacaoservicoqnt);
            }
            else
            {
                colocacaodelacre = "";
                colocacaodelacreqnt = 0;
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
            }
            else
            {
                armazenagem1a3periodo = "";
                armazenagem1a3periodoqnt = 0;
            }
        }

        private void findArmazenagem1a4periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM REFERENTE AOS PERIODOS DE 1 A 4");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                armazenagem1a4periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1a4periodoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
                Console.WriteLine("Armazenagem 1 a 4: " + armazenagem1a4periodo);
            }
            else
            {
                armazenagem1a4periodo = "";
                armazenagem1a4periodoqnt = 0;
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

        private void findArmazenagemExportacao()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM (EXPORTACAO)");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                armazenagemexportacao = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagemexportacaoqnt = int.Parse(dis.Substring(indexbegin - 6, 1));
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