using HtmlAgilityPack;
using System;
using System.Collections.Generic;

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

            doc = new HtmlDocument();
            doc.LoadHtml(source.Replace("&nbsp;", ""));
            string CNPJPrestador = "";
            string nfe = "";
            string dis = "";
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


            if (CNPJPrestador == "04.335.535/0002-55")  //Insert BD SuperTerminais Table
            {
                SuperTerminais superterminais = new SuperTerminais(dis, nfe , RazaoSocialNome, urlofnote , CNPJ);
                superterminais.BeginAnalysis();

            

            }

            else if (CNPJPrestador == "04.694.548/0001-30")  //Insert BD Aurora Table
            {

                AuroraEadi auroraeadi = new AuroraEadi(dis, nfe , RazaoSocialNome , urlofnote , CNPJ);
                auroraeadi.BeginAnalysis();        


            }

            else if (CNPJPrestador == "84.098.383/0001-72")  //Insert BD Chibatao Table
            {
                Chibatao chibatao = new Chibatao(dis, nfe , RazaoSocialNome , urlofnote , CNPJ);
                chibatao.BeginAnalysis();
           

            }


        }
    }
}