using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace BetterBotPMM
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

        public AuroraEadi(string discriminacao, string notafiscale)
        {
            dis = discriminacao;
            nfe = notafiscale;
        }

        public void BeginAnalysis()
        {
            //Console.WriteLine(dis);
            if (findDIDOC())
            {
                findMovimentacaoDeCarga();
                findPesagem();
                findArmazenagem1periodo();
                findBL();
                findArmazenagem2periodo();
                findArmazenagem3periodo();
                findArmazenagem4periodo();
                findArmazenagem5periodo();
                findArmazenagem6periodo();
                findArmazenagem7periodo();
                findTarifaNF();
                findTarifa44();
                inserirNoBancoSuperTerminais();

            }




        }

        private void inserirNoBancoSuperTerminais()
        {
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
            Persist Security Info=False;";
            InsertAurora:
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                query = "insert into AuroraEadi (NFe, DIDOC, MOVIMENTACAO , MOVIMENTACAOQNT, PESAGEM , PESAGEMQNT, ARMAZENAGEM1 , ARMAZENAGEM1QNT, BL , ARMAZENAGEM2 , ARMAZENAGEM2QNT, ARMAZENAGEM3 , ARMAZENAGEM3QNT , ARMAZENAGEM4 , ARMAZENAGEM4QNT , ARMAZENAGEM5 , ARMAZENAGEM5QNT , ARMAZENAGEM6 , ARMAZENAGEM6QNT, ARMAZENAGEM7 , ARMAZENAGEM7QNT , TARIFANF , TARIFANFQNT , TARIFA44 , TARIFA44QNT) values ('" + nfe + "','" + DIDOC + "','" + movimentacaoDeCarga + "'," + movimentacaoDeCargaqnt + ",'" + pesagem + "'," + pesagemqnt + ",'" + armazenagem1periodo + "'," + armazenagem1periodoqnt + ",'" + BL + "','" + armazenagem2periodo + "'," + armazenagem2periodoqnt + ",'" + armazenagem3periodo + "'," + armazenagem3periodoqnt + ",'" + armazenagem4periodo + "'," + armazenagem4periodoqnt + ",'" + armazenagem5periodo + "'," + armazenagem5periodoqnt + ",'" + armazenagem6periodo + "'," + armazenagem6periodoqnt + ",'" + armazenagem7periodo + "'," + armazenagem7periodoqnt + ",'" + tarifanf + "'," + tarifanfqnt + ",'" + tarifa44 + "'," + tarifa44qnt + ")";
                command.CommandText = query;
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                goto InsertAurora;
            }


        }

        private bool findDIDOC()
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

        private void findMovimentacaoDeCarga()
        {
            int indexbegin = dis.IndexOf("MOVIMENTACAO DE CARGA");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                movimentacaoDeCarga = dis.Substring(indexbegin, indexEnd - indexbegin);
                movimentacaoDeCargaqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Movimentacao carga: " + movimentacaoDeCarga);
                //Console.WriteLine(movimentacaoDeCargaqnt);
            }
            else
            {
                movimentacaoDeCarga = "";
                movimentacaoDeCargaqnt = 0;
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
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                pesagem = dis.Substring(indexbegin, indexEnd - indexbegin);
                pesagemqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Pesagem: " + pesagem);
                //Console.WriteLine(pesagemqnt);
            }
            else
            {
                pesagem = "";
                pesagemqnt = 0;
            }

        }

        private void findArmazenagem1periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 1 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem1periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem1periodoqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Armazenagem 1: " + armazenagem1periodo);
                //Console.WriteLine(armazenagem1periodoqnt);
            }
            else
            {
                armazenagem1periodo = "";
                armazenagem1periodoqnt = 0;
            }
        }

        private void findBL()
        {
            int indexbegin = dis.IndexOf("BL:");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin;
                indexbegin += 4;
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

        private void findArmazenagem2periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 2 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem2periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem2periodoqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Armazenagem 2: " + armazenagem2periodo);
                //Console.WriteLine(armazenagem2periodoqnt);
            }
            else
            {
                armazenagem2periodo = "";
                armazenagem2periodoqnt = 0;
            }
        }

        private void findArmazenagem3periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 3 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem3periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem3periodoqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Armazenagem 3: " + armazenagem3periodo);
                //Console.WriteLine(armazenagem3periodoqnt);
            }
            else
            {
                armazenagem3periodo = "";
                armazenagem3periodoqnt = 0;
            }
        }

        private void findArmazenagem4periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 4 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem4periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem4periodoqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Armazenagem 4: " + armazenagem4periodo);
                //Console.WriteLine(armazenagem4periodoqnt);
            }
            else
            {
                armazenagem4periodo = "";
                armazenagem4periodoqnt = 0;
            }
        }

        private void findArmazenagem5periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 5 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem5periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem5periodoqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Armazenagem 5: " + armazenagem5periodo);
                //Console.WriteLine(armazenagem5periodoqnt);
            }
            else
            {
                armazenagem5periodo = "";
                armazenagem5periodoqnt = 0;
            }
        }

        private void findArmazenagem6periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 6 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem6periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem6periodoqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Armazenagem 6: " + armazenagem6periodo);
                //Console.WriteLine(armazenagem6periodoqnt);
            }
            else
            {
                armazenagem6periodo = "";
                armazenagem6periodoqnt = 0;
            }
        }

        private void findArmazenagem7periodo()
        {
            int indexbegin = dis.IndexOf("ARMAZENAGEM 7 PERIODO");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                armazenagem7periodo = dis.Substring(indexbegin, indexEnd - indexbegin);
                armazenagem7periodoqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Armazenagem 7: " + armazenagem7periodo);
                //Console.WriteLine(armazenagem7periodoqnt);
            }
            else
            {
                armazenagem7periodo = "";
                armazenagem7periodoqnt = 0;
            }
        }

        private void findTarifaNF()
        {
            int indexbegin = dis.IndexOf("TARIFA MINIMA DE EMISSAO DE NOTA FISCAL");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                tarifanf = dis.Substring(indexbegin, indexEnd - indexbegin);
                tarifanfqnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Tarifa nota fiscal: " + tarifanf);
                //Console.WriteLine(tarifanfqnt);
            }
            else
            {
                tarifanf = "";
                tarifanfqnt = 0;
            }
        }

        private void findTarifa44()
        {
            int indexbegin = dis.IndexOf("TARIFA DE SERVICOS CONFORME TABELA ITEM 4.4");
            //Console.WriteLine(indexbegin);
            if (indexbegin != -1)
            {
                int indexEnd = dis.Substring(indexbegin).IndexOf(";");
                indexEnd += indexbegin - 1;
                int indexmoney = dis.Substring(indexbegin).IndexOf("$");
                indexbegin += indexmoney + 2;
                //Console.WriteLine(indexbegin);
                //Console.WriteLine(indexEnd);
                tarifa44 = dis.Substring(indexbegin, indexEnd - indexbegin);
                tarifa44qnt = int.Parse(dis.Substring(indexbegin - 7, 1));
                Console.WriteLine("Tarifa 4.4: " + tarifa44);
                //Console.WriteLine(tarifa44qnt);
            }
            else
            {
                tarifa44 = "";
                tarifa44qnt = 0;
            }
        }
    }
}