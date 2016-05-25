using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotasTerminaisDownload
{
    class ReadNote
    {
        string source;
        Uri urlofnote;
        string CNPJ;
        HtmlDocument doc;

        public ReadNote(string src, Uri url, string cnpj)
        {
            source = src;
            urlofnote = url;
            CNPJ = cnpj;
        }

        public void StartAnalysis()
        {
            //MySqlConnection connection = new MySqlConnection();
            //connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
            Persist Security Info=False;";
            string query;

            doc = new HtmlDocument();
            doc.LoadHtml(source.Replace("&nbsp;", ""));
            string CNPJPrestador = "";
            string nfe = "";
            string rps = "";
            string dis = "";
            string valorliquido = "";
            string valorservico = "";
            string ISSQNRetido = "";
            string CODServico = "";
            string NFeSub = "";
            string DataHoraEmissao = "";
            string Competencia = "";
            string CODVerificacao = "";
            string CNPJTomador = "";
            string RazaoSocialNome = "";
            string CIA = "";
            


            try //Try CNPJ/CPF Prestador
            {
                CNPJPrestador = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:129px;top:176px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("CNPJ/CPF Prestador: " + CNPJPrestador);
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                Console.WriteLine("cade o CNPJ Prestador");

            }

            try //Try nota fiscal eletronica
            {
                nfe = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:502px;top:52px;')]").SelectSingleNode(".//*").InnerHtml;

                Console.WriteLine("NFe: " + nfe);
            }
            catch (Exception)
            {
                nfe = "Não tem nota fiscal???????";
                Console.WriteLine("Não tem nota fiscal???????");

            }


            try //Try Discriminacao
            {
                dis = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:12px;top:335px;')]").SelectSingleNode(".//*").InnerHtml.Replace("<br>", " ").Replace("'", "''");
                Console.WriteLine("Discriminacao: " + dis);
            }
            catch (Exception)
            {
                dis = "Não possui Discriminação do Serviço";
                Console.WriteLine("Não possui Discriminação");
            }

       

            if (CNPJPrestador == "04.335.535/0002-55")  //Insert BD SuperTerminais Table
            {
                SuperTerminais superterminais = new SuperTerminais(dis, nfe);
                superterminais.BeginAnalysis();

                //insert no banco
                InsertSuperTerminais:
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL , CNPJPrestadorNumber) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "','" + CNPJ + "')";
                    command.CommandText = query;
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception err)
                {
                    connection.Close();
                    Console.WriteLine(nfe + " - " + err.Message);
                    goto InsertSuperTerminais;
                }

            }

            else if (CNPJPrestador == "04.694.548/0001-30")  //Insert BD Aurora Table
            {

                AuroraEadi auroraeadi = new AuroraEadi(dis, nfe);
                auroraeadi.BeginAnalysis();
                InsertAurora:
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL , CNPJPrestadorNumber) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "','" + CNPJ + "')";
                    command.CommandText = query;
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception err)
                {
                    connection.Close();
                    Console.WriteLine(nfe + " - " + err.Message);
                    goto InsertAurora;
                }


            }

            else if (CNPJPrestador == "84.098.383/0001-72")  //Insert BD Chibatao Table
            {
                Chibatao chibatao = new Chibatao(dis, nfe);
                chibatao.BeginAnalysis();

                InsertChibatao:
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL , CNPJPrestadorNumber) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "','" + CNPJ + "')";
                    command.CommandText = query;
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception err)
                {
                    connection.Close();
                    Console.WriteLine(nfe + " - " + err.Message);
                    goto InsertChibatao;
                }

            }

            else if (CNPJPrestador == "00.711.083/0005-50") //Insert EXPEDITORS
            {
                InsertExpeditors:
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    query = "insert into EXPEDITORS (Desconsolidacao, NFe) values ('" + valorservico + "','" + nfe + "')";
                    Console.WriteLine("Desconsolidacao: " + valorservico);
                    command.CommandText = query;
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception err)
                {
                    connection.Close();
                    Console.WriteLine(nfe + " - " + err.Message);
                    goto InsertExpeditors;

                }
            }



            else
            {
                InsertWhatever:
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL , CNPJPrestadorNumber) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "','" + CNPJ + "')";
                    command.CommandText = query;
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception err)
                {
                    connection.Close();
                    Console.WriteLine(nfe + " - " + dis + " - " + err.Message);
                    goto InsertWhatever;
                }
            }



        }
    }
}
