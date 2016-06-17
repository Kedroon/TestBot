using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.IO;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OpenQA.Selenium.Remote;
using System.Drawing.Imaging;
using HtmlAgilityPack;
using System.Net;
using System.Diagnostics;
using System.Drawing;
using System.Collections.Concurrent;
using System.Data.SqlClient;

namespace BetterBotPMM
{
    class Automation
    {
        FirefoxDriver js;
        SqlConnection connection;
        string query;
        List<CNPJValido> cnpjcpfValidos = new List<CNPJValido>();
        //string excelpath = @"C:\TempExcel2\rel_notas_aceite_05-2016.xls";
        string excelpath = @"C:\TempExcel2\rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy") + ".xls";
        int superterminaiscount = 0;
        int auroraeadicount = 0;
        int chibataocount = 0;
        int notasvalidas = 0;
        List<UrlToDownload> urls_to_download;
        string MOBI;
        string CIA;
        int countIfZero = 0;


        public Automation(string mobi , string cia)
        {
            MOBI = mobi;
            CIA = cia;
            connection = new SqlConnection();
            connection.ConnectionString = @"Data Source=TERMDT0174,80;Initial Catalog=NotasTerminais;Persist Security Info=True;User ID=sa;Password=301d05150063";


        }

        public void startautomation()
        {
            string pathToCurrentUserProfiles = Environment.ExpandEnvironmentVariables("%APPDATA%") + @"\Mozilla\Firefox\Profiles"; // Path to profile
            string[] pathsToProfiles = Directory.GetDirectories(pathToCurrentUserProfiles, "*.default*", SearchOption.TopDirectoryOnly);
            FirefoxProfile profile = null;
            if (pathsToProfiles.Length != 0)
            {
                profile = new FirefoxProfile();
                profile.SetPreference("browser.download.dir", @"C:\TempExcel2");
                profile.SetPreference("browser.download.folderList", 2);
                profile.SetPreference("browser.tabs.loadInBackground", false); // set preferences you need
                profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream;application/csv;text/csv;application/vnd.ms-excel;");
                Console.WriteLine("Profile carregado");
            }
            OpenQA.Selenium.Cookie cookie1 = new OpenQA.Selenium.Cookie("PID", "2524");
            OpenQA.Selenium.Cookie cookie2 = new OpenQA.Selenium.Cookie("MOBI", MOBI);
            Console.WriteLine(MOBI);
            OpenQA.Selenium.Cookie cookie3 = new OpenQA.Selenium.Cookie("TIPO", "0");
            OpenQA.Selenium.Cookie cookie4 = new OpenQA.Selenium.Cookie("SUBTIPO", "");
            OpenQA.Selenium.Cookie cookie5 = new OpenQA.Selenium.Cookie("mes", DateTime.Now.ToString("MM"));
            OpenQA.Selenium.Cookie cookie6 = new OpenQA.Selenium.Cookie("ano", DateTime.Now.Year.ToString());
            First:
            

                //options.AddArgument("--no-startup-window");


                Console.WriteLine("Iniciando Firefox");

                js = new FirefoxDriver(new FirefoxBinary(), profile);
                js.Manage().Window.Position = new System.Drawing.Point(-2000, 0);

                /*OpenQA.Selenium.Cookie cookie5 = new OpenQA.Selenium.Cookie("mes", "05");
                OpenQA.Selenium.Cookie cookie6 = new OpenQA.Selenium.Cookie("ano", "2016");*/

                Page0:
                try
                {

                    Console.WriteLine("Pagina 1");
                    js.Navigate().GoToUrl("https://www3.gissonline.com.br/interna/default.cfms");
                    js.Manage().Cookies.AddCookie(cookie1);
                    js.Manage().Cookies.AddCookie(cookie2);
                    js.Manage().Cookies.AddCookie(cookie3);
                    js.Manage().Cookies.AddCookie(cookie4);
                    js.Manage().Cookies.AddCookie(cookie5);
                    js.Manage().Cookies.AddCookie(cookie6);
                }
                catch (Exception err)
                {
                    Console.WriteLine(err.Message);
                    goto Page0;


                }

                Page1:
            try
            {
                js.Navigate().GoToUrl("https://www3.gissonline.com.br/recebidas/listaNotas.cfm?modalidade=T");
                //Console.Read();
                new SelectElement(js.FindElement(By.Name("maxrow"))).SelectByText("500");

                if (File.Exists(excelpath))
                {
                    File.Delete(excelpath);
                }

                js.FindElementByXPath("//a[contains(text(),'GERAR ARQUIVO EXCEL')]").Click();
            }



            catch (Exception err)
            {
                countIfZero++;
                if (countIfZero < 3)
                {
                    Console.WriteLine(err.Message);
                    goto Page1;
                }
                else
                {
                    js.Quit();
                    return;

                }


            }
            //Esperar o termino do download da planilha
            Console.WriteLine("Downloading Planilha");
            for (var i = 0; i < 30; i++)
            {
                if (File.Exists(excelpath))
                {
                    break;
                }
                Thread.Sleep(1000);
                if (i == 29)
                {
                    goto Page1;
                }
            }
            long length;
            FileLength:
            try
            {
                length = new FileInfo(excelpath).Length;
            }
            catch (Exception)
            {

                goto FileLength;
            }

            for (var i = 0; i < 30; i++)
            {
                Thread.Sleep(1000);
                var newLength = new FileInfo(excelpath).Length;
                if (newLength == length && length != 0) { break; }
                length = newLength;
            }
            Console.WriteLine("Download concluido");
            Console.WriteLine("Analisando planilha");
            ListOfCNPJCPF(); //Analisar planilha
            Thread.Sleep(3000);
            Console.WriteLine("Analise concluida");

            Console.WriteLine(cnpjcpfValidos.Count);

            HtmlDocument page = new HtmlDocument();
            page.LoadHtml(js.PageSource);
            List<HtmlNode> urlOfNotas = new List<HtmlNode>();
            try
            {
                foreach (var item in page.DocumentNode.SelectNodes("//a[starts-with(@onclick,'janela')]"))
                {
                    urlOfNotas.Add(item);
                }
                Console.WriteLine(urlOfNotas.Count);

                int loopPaginas = 500;
                while (cnpjcpfValidos.Count > loopPaginas)
                {
                    int tempvalue = loopPaginas + 1;
                    js.FindElementByXPath("//a[contains(@onclick,'document.formPag.startrow.value=" + tempvalue + ";document.formPag.submit();')]").Click();
                    HtmlDocument nextpage = new HtmlDocument();
                    nextpage.LoadHtml(js.PageSource);

                    nextpage.DocumentNode.SelectNodes("//a[starts-with(@onclick,'janela')]");


                    foreach (var item in nextpage.DocumentNode.SelectNodes("//a[starts-with(@onclick,'janela')]"))
                    {
                        urlOfNotas.Add(item);

                    }

                    loopPaginas += 500;

                }
            }
            catch (Exception)
            {
                cnpjcpfValidos.Clear();
                urlOfNotas.Clear();
                goto Page1;
            }



            string temp;

            urls_to_download = new List<UrlToDownload>();
            int count = 0;
            if (urlOfNotas.Count != cnpjcpfValidos.Count)
            {
                js.Quit();
                return;
            }
            foreach (var item in urlOfNotas)
            {




                if (cnpjcpfValidos[count].valido == true)
                {

                    temp = item.OuterHtml;
                    temp = temp.Replace(@"<a onclick=""janela('..", "").Replace(@"',430,260)"" href=""javascript:;""><img border=""0"" title=""Dados da nota fiscal"" src=""../biblioteca/images/PL_FindResults_R.png""></a>", "");
                    temp = "https://www3.gissonline.com.br" + temp;
                    temp = temp.Replace("amp;", "");
                    Console.WriteLine(temp);
                    try
                    {
                        urls_to_download.Add(new UrlToDownload(new Uri(temp), cnpjcpfValidos[count].CNPJ));
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("URL invalida");
                        return;
                    }


                }
                count++;

            }

            Console.WriteLine(count);

            var watch = System.Diagnostics.Stopwatch.StartNew();
            Client clienttest = new Client();
            clienttest.Headers.Add(HttpRequestHeader.Cookie,
              "PID=2524;" +
              "MOBI=560801"
              );

            foreach (var item in urls_to_download)
            {
                string source;
                int countTryAgain = 0;
                Tryagain:
                try
                {
                    source = clienttest.DownloadString(item.URI);
                }
                catch (Exception)
                {
                    countTryAgain++;
                    Console.WriteLine("Tentando de novo " +countTryAgain);
                    Thread.Sleep(2000);
                    goto Tryagain;
                }

                ReadNote note = new ReadNote(source, item.URI, item.CNPJ , CIA);
                note.StartAnalysis();
            }

            //loop para abrir as notas

            watch.Stop();
            Console.WriteLine("Execution Time: " + (watch.ElapsedMilliseconds / 1000) + "Seconds");
            js.Manage().Cookies.DeleteAllCookies();
            js.Quit();

        }

        private void ListOfCNPJCPF()
        {

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            //string workbookPath = @"C:\TempExcel\rel_notas_aceite_05-2016.xls";
            string workbookPath = @"C:\TempExcel2\rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy");
            var workbooks = excelApp.Workbooks;
            Excel.Workbook excelWorkbook = workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            //string currentSheet = "rel_notas_aceite_05-2016";
            string currentSheet = "rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy");
            //MessageBox.Show(currentSheet);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            int i = 4;
            while (excelWorksheet.Cells[i, 11].Value != null)
            {
                string cnpjcpf = excelWorksheet.Cells[i, 11].Value2.ToString();
                if (cnpjcpf == "84098383000172") //Chibatao
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    chibataocount++;
                    notasvalidas++;
                    //Console.WriteLine(excelWorksheet.Cells[i, 11].Value2);
                }
                else if (cnpjcpf == "4694548000130") //Aurora Eadi
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    auroraeadicount++;
                    notasvalidas++;
                }
                else if (cnpjcpf == "4335535000255") //Super Terminais
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    superterminaiscount++;
                    notasvalidas++;
                }

                else if (cnpjcpf == "711083000550") //EXPEDITORS INTERNATIONAL
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }

                else if (cnpjcpf == "53284634000503") //UPS SCS TRANSPORTES
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }

                else if (cnpjcpf == "57067928000704") //AGILLITY (ITATRANS)
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }

                else if (cnpjcpf == "51595908000479") //NIPPON EXPRESS
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }

                else if (cnpjcpf == "2886427004313") //KUEHNE+NAGEL
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }

                else if (cnpjcpf == "11233216000202") //CAPITALLOG LOGISTICA
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }
                else if (cnpjcpf == "84098383001063") //CHIBATAO NAVEGACAO 2 ??
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }

                else
                {
                    cnpjcpfValidos.Add(new CNPJValido(false, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                }

                i++;
            }

            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelSheets);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(workbooks);
            excelApp.Quit();

            FilterNotes:
            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                query = "select Chave from Despesas";
                command.CommandText = query;
                SqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {

                    string nota = reader["Chave"].ToString();
                    foreach (var item in cnpjcpfValidos)
                    {
                        if (nota == item.Nota + item.CNPJ)
                        {
                            //Console.WriteLine(nota + " " + item.Nota + item.CNPJ);
                            item.valido = false;
                        }
                        else
                        {
                            //Console.WriteLine("oi");
                        }
                    }

                }
                connection.Close();
            }
            catch (Exception)
            {
                connection.Close();
                goto FilterNotes;
            }
        }





    }





}