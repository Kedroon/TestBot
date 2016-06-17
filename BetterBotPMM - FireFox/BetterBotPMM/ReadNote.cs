using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BetterBotPMM
{
    public class ReadNote
    {
        string source;
        Uri urlofnote;
        string CNPJ;
        string CIA;
        HtmlDocument doc;

        public ReadNote(string src, Uri url, string cnpj, string cia)
        {
            source = src;
            urlofnote = url;
            CNPJ = cnpj;
            CIA = cia;
        }

        public void StartAnalysis()
        {
            //MySqlConnection connection = new MySqlConnection();
            //connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";
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
            string DescontoIncondicionado = "";
            string DescontoCondicionado = "";
            string RetençõesFederais = "";
            string OutrasRetenções = "";
            string PIS = "";
            string COFINS = "";
            string IR = "";
            string INSS = "";
            string CSLL = "";


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


            try //Try RPS
            {
                rps = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:124px;top:102px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("RPS: " + rps);

            }
            catch (Exception)
            {
                rps = "Não possui RPS";
                Console.WriteLine("Não possui RPS");
            }

            //Try Valor Liquido

            valorliquido = returnValueofElementDetalhamento("(=) Valor Líquido R$", "Não possui valor liquido", "Valor liquido: ");



            //Try Valor Servico

            valorservico = returnValueofElementDetalhamento("Valor do Serviço  R$", "Não possui valor do serviço", "Valor servico: ");





            //Try ISSQN Retido
            ISSQNRetido = returnValueofElementDetalhamento("(-) ISSQN Retido", "Não possui ISSQN Retido", "ISSQN Retido: ");





            //Try Codigo do Servico
            CODServico = returnValueofElemenetCodigoServiço("Código do Serviço / Atividade", "Não possui codigo de serviço", "Codigo servico: ");





            try //Try NFe Substituido
            {
                NFeSub = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:319px;top:102px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("NFe Substituido: " + NFeSub);

            }
            catch (Exception)
            {
                NFeSub = "Nao possui Nfe Substituido";
                Console.WriteLine("Não possui NFe Substituido");
            }

            try //Try Data e Hora
            {
                DataHoraEmissao = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:124px;top:82px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Data e Hora de Emissao: " + DataHoraEmissao);

            }
            catch (Exception)
            {
                DataHoraEmissao = "Nao possui Data e Hora de Emissao";
                Console.WriteLine("Não possui Data e Hora de Emissao");
            }

            try //Try Competencia
            {
                Competencia = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:319px;top:82px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Competencia: " + Competencia);

            }
            catch (Exception)
            {
                Competencia = "Nao possui Competencia";
                Console.WriteLine("Não possui Competencia");
            }

            try //Try Codigo de Verificação
            {
                CODVerificacao = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:470px;top:82px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Codigo de Verificacao: " + CODVerificacao);

            }
            catch (Exception)
            {
                CODVerificacao = "Nao possui Codigo de Verificacao";
                Console.WriteLine("Não possui Codigo de Verificacao");
            }

            try //Try CNPJ do Tomador
            {
                CNPJTomador = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:59px;top:264px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("CNPJ/CPF Tomador: " + CNPJTomador);

            }
            catch (Exception)
            {
                CNPJTomador = "Nao possui CNPJ Tomador";
                Console.WriteLine("Não possui CNPJ Tomador");
            }

            try //Try CNPJ do RazaoSocialNome
            {
                RazaoSocialNome = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:175px;top:142px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Razao Social Nome: " + RazaoSocialNome);

            }
            catch (Exception)
            {
                RazaoSocialNome = "Nao possui Razao Social Nome";
                Console.WriteLine("Não possui Razao Social Nome");
            }

            DescontoIncondicionado = returnValueofElementDetalhamento("(-) Desconto Incondicionado", "Não possui Desconto Incondicionado", "Desconto incondicionado: ");


            DescontoCondicionado = returnValueofElementDetalhamento("(-) Desconto Condicionado", "Não possui Desconto Condicionado", "Desconto condicionado: ");


            RetençõesFederais = returnValueofElementDetalhamento("(-) Retenções Federais", "Não possui Retenções Federais", "Retenções Federais: ");


            OutrasRetenções = returnValueofElementDetalhamento("(-) Outras Retenções", "Não possui Outras Retenções", "Outras Retenções: ");


            PIS = returnValueofElementTributos("PIS(R$)", "73", "Não possui PIS", "PIS: ");


            COFINS = returnValueofElementTributos("COFINS(R$)", "185", "Não possui COFINS", "COFINS: ");


            IR = returnValueofElementTributos("IR(R$)", "297", "Não possui IR", "IR: ");


            INSS = returnValueofElementTributos("INSS(R$)", "409", "Não possui INSS", "INSS: ");


            CSLL = returnValueofElementTributos("CSLL(R$)", "528", "Não possui CSLL", "CSLL: ");



            if (CNPJPrestador == "04.335.535/0002-55")  //Insert BD SuperTerminais Table
            {
                SuperTerminais superterminais = new SuperTerminais(dis, nfe);
                superterminais.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);

            }

            else if (CNPJPrestador == "04.694.548/0001-30")  //Insert BD Aurora Table
            {

                AuroraEadi auroraeadi = new AuroraEadi(dis, nfe);
                auroraeadi.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);


            }

            else if (CNPJPrestador == "84.098.383/0001-72")  //Insert BD Chibatao Table
            {
                Chibatao chibatao = new Chibatao(dis, nfe);
                chibatao.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);

            }

            else if (CNPJPrestador == "00.711.083/0005-50") //Insert EXPEDITORS
            {
                Expeditors expeditors = new Expeditors(nfe, valorservico, dis, CNPJ, RazaoSocialNome);
                expeditors.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);
            }

            else if (CNPJPrestador == "57.067.928/0007-04") //Insert Agility
            {
                Agility agility = new Agility(nfe, valorservico, dis, CNPJ, RazaoSocialNome);
                agility.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);
            }

            else if (CNPJPrestador == "11.233.216/0002-02") //Insert Capital
            {
                Capital capital = new Capital(nfe, valorservico, dis, CNPJ, RazaoSocialNome);
                capital.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);
            }

            else if (CNPJPrestador == "51.595.908/0004-79") //Insert Nippon
            {
                Nippon nippon = new Nippon(nfe, valorservico, dis, CNPJ, RazaoSocialNome);
                nippon.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);
            }

            else if (CNPJPrestador == "53.284.634/0005-03") //Insert UPS
            {
                UPS ups = new UPS(nfe, valorservico, dis, CNPJ, RazaoSocialNome);
                ups.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);
            }

            else if (CNPJPrestador == "02.886.427/0043-13") //Insert KN
            {
                KN kn = new KN(nfe, valorservico, dis, CNPJ, RazaoSocialNome);
                kn.BeginAnalysis();
                InsertBanco(nfe, CNPJ, rps, dis, valorliquido, valorservico, ISSQNRetido, CODServico, NFeSub, DataHoraEmissao, Competencia, CODVerificacao, CNPJPrestador, CNPJTomador, RazaoSocialNome, DescontoIncondicionado, DescontoCondicionado, RetençõesFederais, OutrasRetenções, PIS, COFINS, IR, INSS, CSLL, urlofnote, CIA);

            }


            /*else
            {
                InsertWhatever:
                try
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand();
                    command.Connection = connection;
                    query = "insert into Despesas (Chave , NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL , CNPJPrestadorNumber , CIA , Data) values ('" + nfe + CNPJ + "','" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "','" + CNPJ + "','" + CIA + "', CURRENT_TIMESTAMP )";
                    command.CommandText = query;
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (SqlException exception)
                {
                    if (exception.Number == 2627) // Cannot insert duplicate key row in object error
                    {
                        // handle duplicate key error
                        return;
                    }
                    else
                        return; // throw exception if this exception is unexpected
                }
                catch (Exception err)
                {
                    connection.Close();
                    Console.WriteLine(nfe + " - " + dis + " - " + err.Message);
                    goto InsertWhatever;
                }
            }*/



        }

        string returnValueofElementDetalhamento(string searchValue, string failure, string descricao)
        {
            string temp;
            try
            {
                temp = doc.DocumentNode.SelectSingleNode("//span[.='" + searchValue + "']").OuterHtml;
                int indextop = temp.IndexOf("top") + 4;
                temp = temp.Substring(indextop, temp.Substring(indextop).IndexOf("px;"));
                temp = doc.DocumentNode.SelectSingleNode(@"//span[starts-with(@style,'position:absolute;left:135px;top:" + temp + "px;')]").InnerText;
                Console.WriteLine(descricao + temp);
                //Console.WriteLine(" '" + temp + "' ");
                //Console.Read();

            }
            catch (Exception)
            {
                temp = failure;
            }
            if (temp == "")
            {
                return "0,00";
            }
            return temp;
        }

        string returnValueofElementTributos(string searchValue, string left, string failure, string descricao)
        {
            string temp;
            try
            {
                temp = doc.DocumentNode.SelectSingleNode("//span[.='" + searchValue + "']").OuterHtml;
                int indextop = temp.IndexOf("top") + 4;
                temp = temp.Substring(indextop, temp.Substring(indextop).IndexOf("px;"));
                temp = doc.DocumentNode.SelectSingleNode(@"//span[starts-with(@style,'position:absolute;left:" + left + "px;top:" + temp + "px;')]").InnerText;
                Console.WriteLine(descricao + temp);
                //Console.WriteLine(" '" + temp + "' ");
                //Console.Read();
            }
            catch (Exception)
            {

                temp = failure;
            }
            if (temp == "")
            {
                return "0,00";
            }
            return temp;
        }

        string returnValueofElemenetCodigoServiço(string searchValue, string failure, string descricao)
        {
            string temp;
            try
            {
                temp = doc.DocumentNode.SelectSingleNode("//span[.='" + searchValue + "']").OuterHtml;
                int indextop = temp.IndexOf("top") + 4;
                temp = temp.Substring(indextop, temp.Substring(indextop).IndexOf("px;"));
                int i = int.Parse(temp) + 20;
                temp = doc.DocumentNode.SelectSingleNode(@"//span[starts-with(@style,'position:absolute;left:12px;top:" + i + "px;')]").InnerText;
                int indexEnd = temp.IndexOf("-");
                temp = temp.Substring(0, indexEnd - 1);
                Console.WriteLine(descricao + temp);
            }
            catch (Exception)
            {

                temp = failure;
            }
            return temp;
        }

        void InsertBanco(string nfe, string CNPJ, string rps, string dis, string valorliquido, string valorservico, string ISSQNRetido, string CODServico, string NFeSub, string DataHoraEmissao, string Competencia, string CODVerificacao, string CNPJPrestador, string CNPJTomador, string RazaoSocialNome, string DescontoIncondicionado, string DescontoCondicionado, string RetençõesFederais, string OutrasRetenções, string PIS, string COFINS, string IR, string INSS, string CSLL, Uri urlofnote, string CIA)
        {
            SqlConnection connection = new SqlConnection();
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";
            string query;
            InsertBanco:
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                query = "insert into Despesas (Chave , NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL , CNPJPrestadorNumber , CIA , Data) values ('" + nfe + CNPJ + "','" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "','" + CNPJ + "','" + CIA + "', CURRENT_TIMESTAMP )";
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
                    goto InsertBanco;
                }
            }

            catch (Exception err)
            {
                connection.Close();
                Console.WriteLine(err.Message);
                goto InsertBanco;
            }
        }

    }

}