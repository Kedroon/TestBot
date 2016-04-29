using System;
using System.Data.OleDb;


namespace AlgorithmTestConsole
{
    class AuroraEadi
    {
        private OleDbConnection connection = new OleDbConnection();

        private string dis;
        private string nfe;

        string DIDOC = "";

        string movimentacaoDeCarga = "";
        int movimentacaoDeCargaqnt = 0;

        string pesagem = "";
        int pesagemqnt = 0;

        string armazenagem1periodo;
        int armazenagem1periodoqnt;

        string armazenagem2periodo;
        int armazenagem2periodoqnt;

        string armazenagem3periodo;
        int armazenagem3periodoqnt;

        string armazenagem4periodo;
        int armazenagem4periodoqnt;

        string armazenagem5periodo;
        int armazenagem5periodoqnt;

        string armazenagem6periodo;
        int armazenagem6periodoqnt;

        string armazenagem7periodo;
        int armazenagem7periodoqnt;

        string tarifanf;
        int tarifanfqnt;

        string tarifa44;
        int tarifa44qnt;

        string BL;

        string query;

        

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
                findMovimentacaoDeCarga(discriminacao);
                findPesagem(discriminacao);
                findArmazenagem1periodo(discriminacao);
                findBL(discriminacao);
                findArmazenagem2periodo(discriminacao);
                findArmazenagem3periodo(discriminacao);
                findArmazenagem4periodo(discriminacao);
                findArmazenagem5periodo(discriminacao);
                findArmazenagem6periodo(discriminacao);
                findArmazenagem7periodo(discriminacao);
                findTarifaNF(discriminacao);
                findTarifa44(discriminacao);
                inserirNoBancoSuperTerminais(nfe);

            }
            connection.Close();


        }

        private void inserirNoBancoSuperTerminais(string nfe)
        {
            // connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            query = "insert into AuroraEadi (NFe, DIDOC, MOVIMENTACAO , MOVIMENTACAOQNT, PESAGEM , PESAGEMQNT, ARMAZENAGEM1 , ARMAZENAGEM1QNT, BL , ARMAZENAGEM2 , ARMAZENAGEM2QNT, ARMAZENAGEM3 , ARMAZENAGEM3QNT , ARMAZENAGEM4 , ARMAZENAGEM4QNT , ARMAZENAGEM5 , ARMAZENAGEM5QNT , ARMAZENAGEM6 , ARMAZENAGEM6QNT, ARMAZENAGEM7 , ARMAZENAGEM7QNT , TARIFANF , TARIFANFQNT , TARIFA44 , TARIFA44QNT) values ('" + nfe+"','"+DIDOC+ "','" + movimentacaoDeCarga + "'," + movimentacaoDeCargaqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + BL + "','" + armazenagem2periodo + "'," + armazenagem2periodoqnt + ",'" + armazenagem3periodo + "'," + armazenagem3periodoqnt + ",'" + armazenagem4periodo + "'," + armazenagem4periodoqnt + ",'" + armazenagem5periodo + "'," + armazenagem5periodoqnt + ",'" + armazenagem6periodo + "'," + armazenagem6periodoqnt + ",'" + armazenagem7periodo + "'," + armazenagem7periodoqnt + ",'" + tarifanf + "'," + tarifanfqnt + ",'" + tarifa44 + "'," + tarifa44qnt + ")";
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
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 4;
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

        private void findMovimentacaoDeCarga(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("MOVIMENTACAO DE CARGA");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                movimentacaoDeCarga = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                movimentacaoDeCargaqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(movimentacaoDeCarga);
                Console.WriteLine(movimentacaoDeCargaqnt);
            }
            else
            {
                movimentacaoDeCarga = "";
                movimentacaoDeCargaqnt = 0;
            }
        }

        private void findPesagem(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("PESAGEM");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                pesagem = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                pesagemqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(pesagem);
                Console.WriteLine(pesagemqnt);
            }
            else
            {
                pesagem = "";
                pesagemqnt = 0;
            }

        }

        private void findArmazenagem1periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM 1 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem1periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem1periodo);
                Console.WriteLine(armazenagem1periodoqnt);
            }
            else
            {
                armazenagem1periodo = "";
                armazenagem1periodoqnt = 0;
            }
        }

        private void findBL(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("BL:");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 4;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                BL = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                Console.WriteLine(BL);
            }
            else
            {
                BL = "";
            }
        }

        private void findArmazenagem2periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM 2 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem2periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem2periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem2periodo);
                Console.WriteLine(armazenagem2periodoqnt);
            }
            else
            {
                armazenagem2periodo = "";
                armazenagem2periodoqnt = 0;
            }
        }

        private void findArmazenagem3periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM 3 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem3periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem3periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem3periodo);
                Console.WriteLine(armazenagem3periodoqnt);
            }
            else
            {
                armazenagem3periodo = "";
                armazenagem3periodoqnt = 0;
            }
        }

        private void findArmazenagem4periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM 4 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem4periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem4periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem4periodo);
                Console.WriteLine(armazenagem4periodoqnt);
            }
            else
            {
                armazenagem4periodo = "";
                armazenagem4periodoqnt = 0;
            }
        }

        private void findArmazenagem5periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM 5 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem5periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem5periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem5periodo);
                Console.WriteLine(armazenagem5periodoqnt);
            }
            else
            {
                armazenagem5periodo = "";
                armazenagem5periodoqnt = 0;
            }
        }

        private void findArmazenagem6periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM 6 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem6periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem6periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem6periodo);
                Console.WriteLine(armazenagem6periodoqnt);
            }
            else
            {
                armazenagem6periodo = "";
                armazenagem6periodoqnt = 0;
            }
        }

        private void findArmazenagem7periodo(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("ARMAZENAGEM 7 PERIODO");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                armazenagem7periodo = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem7periodoqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(armazenagem7periodo);
                Console.WriteLine(armazenagem7periodoqnt);
            }
            else
            {
                armazenagem7periodo = "";
                armazenagem7periodoqnt = 0;
            }
        }

        private void findTarifaNF(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("TARIFA MINIMA DE EMISSAO DE NOTA FISCAL");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                tarifanf = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                tarifanfqnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(tarifanf);
                Console.WriteLine(tarifanfqnt);
            }
            else
            {
                tarifanf = "";
                tarifanfqnt = 0;
            }
        }

        private void findTarifa44(string discriminacao)
        {
            int indexbegin = discriminacao.IndexOf("TARIFA DE SERVICOS CONFORME TABELA ITEM 4.4");
            Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = discriminacao.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = discriminacao.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 1;
                Console.WriteLine(indexbegin);
                Console.WriteLine(indexEnd);
                tarifa44 = discriminacao.Substring(indexbegin, indexEnd - indexbegin);
                tarifa44qnt = int.Parse(discriminacao.Substring(indexbegin - 6, 1));
                Console.WriteLine(tarifa44);
                Console.WriteLine(tarifa44qnt);
            }
            else
            {
                tarifa44 = "";
                tarifa44qnt = 0;
            }
        }



    }


}
