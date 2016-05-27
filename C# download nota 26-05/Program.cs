using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;
using HtmlAgilityPack;
using System.Threading;
using System.Net;
using System.Runtime.InteropServices;

namespace NotasTerminaisDownload
{
    class Program
    {
        static ChromeDriver js;
        static SqlConnection connection;
        static string query;
        static List<CNPJValido> cnpjcpfValidos = new List<CNPJValido>();
        static string excelpath = @"C:\TempExcel\rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy") + ".xls";
        static int superterminaiscount = 0;
        static int auroraeadicount = 0;
        static int chibataocount = 0;
        static int notasvalidas = 0;
        static int countMultiPage = 0;
        static List<UrlToDownload> urls_to_download;

        static void Main(string[] args)
        {
            connection = new SqlConnection();
            connection.ConnectionString = @"Data Source=192.168.0.110,59160;Initial Catalog=NotaTerminais;Persist Security Info=True;User ID=sa;Password=ca94404llc;Pooling=False";
            automation();


        }
        static void automation()
        {
            
            Environment.SetEnvironmentVariable("webdriver.chrome.driver", "chromedriver.exe");
            ChromeOptions options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", @"C:\TempExcel");
            //options.AddArgument("--no-startup-window");

            Console.WriteLine("Iniciando Chrome");

            js = new ChromeDriver(options);
            OpenQA.Selenium.Cookie cookie1 = new OpenQA.Selenium.Cookie("PID", "2524");
            OpenQA.Selenium.Cookie cookie2 = new OpenQA.Selenium.Cookie("MOBI", "560801");
            OpenQA.Selenium.Cookie cookie3 = new OpenQA.Selenium.Cookie("TIPO", "0");
            OpenQA.Selenium.Cookie cookie4 = new OpenQA.Selenium.Cookie("SUBTIPO", "");

            //Console.WriteLine("Profile do firefox nao encontrado");





            Console.WriteLine("Pagina 1");
            js.Navigate().GoToUrl("https://www3.gissonline.com.br/interna/default.cfms");
            js.Manage().Cookies.AddCookie(cookie1);
            js.Manage().Cookies.AddCookie(cookie2);
            js.Manage().Cookies.AddCookie(cookie3);
            js.Manage().Cookies.AddCookie(cookie4);
            bool page2 = false;
        Page1:
            try
            {
                js.Navigate().GoToUrl("https://www3.gissonline.com.br/interna/default.cfm");
                js.SwitchTo().Frame(0);
                js.FindElementByXPath("//img[contains(@src,'images/bt_menu__06_off.jpg')]").Click();
            }

            catch (Exception err)
            {


                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1;

            }

        Page2:
            Console.WriteLine("Pagina 2");
            try
            {
                js.SwitchTo().DefaultContent();
                js.SwitchTo().Frame(2);
                DateTime time = DateTime.Now;
                if (!page2)
                {
                    js.FindElement(By.Name("mes")).SendKeys(time.ToString("MM"));
                    js.FindElement(By.Name("ano")).SendKeys(time.Year.ToString());
                    page2 = true;
                }
                else
                {
                    js.FindElement(By.Name("ano")).SendKeys(" ");
                    js.FindElement(By.Name("ano")).Click();
                    Console.WriteLine("oi");
                }
                js.FindElement(By.LinkText("Notas Recebidas")).Click();
            }
            catch (UnhandledAlertException)
            {
                js.SwitchTo().Alert().Accept();
                js.SwitchTo().DefaultContent();
                goto Page1;
            }

            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1;
            }
            int count = 0;
            bool first = true;
        Page3:
            Console.WriteLine("Pagina 3");
            try
            {
                js.SwitchTo().DefaultContent();
                js.SwitchTo().Frame(2);
                new SelectElement(js.FindElement(By.Name("maxrow"))).SelectByText("500");

                if (File.Exists(excelpath))
                {
                    File.Delete(excelpath);
                }

                js.FindElementByXPath("//a[contains(text(),'GERAR ARQUIVO EXCEL')]").Click();
                

            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1;
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
                countMultiPage++;
                HtmlDocument nextpage = new HtmlDocument();
                nextpage.LoadHtml(js.PageSource);
                try
                {
                    nextpage.DocumentNode.SelectNodes("//a[starts-with(@onclick,'janela')]");
                }
                catch (Exception)
                {

                    reloadmultipage();
                }
                foreach (var item in nextpage.DocumentNode.SelectNodes("//a[starts-with(@onclick,'janela')]"))
                {
                    urlOfNotas.Add(item);

                }

                loopPaginas += 500;

            }


            string temp;

            urls_to_download = new List<UrlToDownload>();
            foreach (var item in urlOfNotas)
            {




                if (cnpjcpfValidos[count].valido == true)
                {
                    
                    temp = item.OuterHtml;
                    temp = temp.Replace(@"<a href=""javascript:;"" onclick=""janela('..", "").Replace(@"',430,260)""><img src=""../biblioteca/images/PL_FindResults_R.png"" title=""Dados da nota fiscal"" border=""0""></a>", "");
                    temp = "https://www3.gissonline.com.br" + temp;
                    temp = temp.Replace("amp;", "");
                    Console.WriteLine(temp);
                    urls_to_download.Add(new UrlToDownload(new Uri(temp), cnpjcpfValidos[count].CNPJ));

                }
                count++;

            }

            Console.WriteLine(count);

            var watch = System.Diagnostics.Stopwatch.StartNew();
            int urlatual = 0;
            Client clienttest = new Client();
            clienttest.Headers.Add(HttpRequestHeader.Cookie,
              "PID=2524;" +
              "MOBI=560801"
              );

            foreach (var item in urls_to_download)
            {
                string source;
            Tryagain:
                try
                {
                    source = clienttest.DownloadString(item.URI);
                }
                catch (Exception)
                {
                    Console.WriteLine("Tentando de novo");
                    goto Tryagain;
                }

                ReadNote note = new ReadNote(source, item.URI, item.CNPJ);
                note.StartAnalysis();
            }
          




            //loop para abrir as notas

            watch.Stop();
            Console.WriteLine("Execution Time: " + (watch.ElapsedMilliseconds / 1000) + "Seconds");
            Console.ReadLine();
            Thread.Sleep(10000);
            

        }

        static private void ListOfCNPJCPF()
        {

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            string workbookPath = @"C:\TempExcel\rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy");
            var workbooks = excelApp.Workbooks;
            Excel.Workbook excelWorkbook = workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
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
                else if (cnpjcpf == "84098383001063") //CHIBATAO NAVEGACAO 2 ??
                {
                    cnpjcpfValidos.Add(new CNPJValido(true, cnpjcpf, excelWorksheet.Cells[i, 1].Text.ToString()));
                    notasvalidas++;
                }

                else
                {
                    cnpjcpfValidos.Add(new CNPJValido(false, "", ""));
                }

                i++;
            }
            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelSheets);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(workbooks);
            excelApp.Quit();


        }

        private static void reloadmultipage()
        {
        Page1Reload:
            try
            {
                js.Navigate().GoToUrl("https://www3.gissonline.com.br/interna/default.cfm");
                js.SwitchTo().Frame(0);
                js.FindElementByXPath("//img[contains(@src,'images/bt_menu__06_off.jpg')]").Click();
            }

            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1Reload;
            }

        Page2Reload:
            Console.WriteLine("Pagina 2");
            try
            {
                js.SwitchTo().DefaultContent();
                js.SwitchTo().Frame(2);
                DateTime time = DateTime.Now;
                js.FindElement(By.Name("ano")).SendKeys(" ");
                js.FindElement(By.Name("ano")).Click();
                Console.WriteLine("oi");
                js.FindElement(By.LinkText("Notas Recebidas")).Click();
            }
            catch (UnhandledAlertException)
            {
                js.SwitchTo().Alert().Accept();
                js.SwitchTo().DefaultContent();
                goto Page1Reload;
            }

            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1Reload;
            }

            try
            {
                js.SwitchTo().DefaultContent();
                js.SwitchTo().Frame(2);
                new SelectElement(js.FindElement(By.Name("maxrow"))).SelectByText("500");
            }
            catch (Exception err)
            {


                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1Reload;

            }
            int loopPaginas = 500;
            for (int i = 0; i < countMultiPage; i++)
            {

                int tempvalue = loopPaginas + 1;
                try
                {
                    js.FindElementByXPath("//a[contains(@onclick,'document.formPag.startrow.value=" + tempvalue + ";document.formPag.submit();')]").Click();
                }
                catch (Exception)
                {

                    reloadmultipage();
                    break;
                }

                loopPaginas += 500;
            }

        }
    }
}