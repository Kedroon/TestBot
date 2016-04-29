using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.IO;
using OpenQA.Selenium.Support.UI;
using System.Data.OleDb;
using System.Threading;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace BotPMM
{
    class Program
    {
        static FirefoxDriver fox;
        static OleDbConnection connection;
        static string query;
        static List<bool> cnpjcpfValidos = new List<bool>();
        static string excelpath = @"C:\TempExcel\rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy") + ".xls";


        static void Main(string[] args)
        {
            connection = new OleDbConnection();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
Persist Security Info=False;";
            automation();

        }
        static private void clickVirtualButton(string num, FirefoxDriver fox)
        {
            fox.FindElementByXPath("//img[contains(@src,'/images/teclado/tec_" + num + ".gif')]").Click();

        }

        static void automation()
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();

            Console.WriteLine("Iniciando Firefox");
            string pathToCurrentUserProfiles = Environment.ExpandEnvironmentVariables("%APPDATA%") + @"\Mozilla\Firefox\Profiles"; // Path to profile
            string[] pathsToProfiles = Directory.GetDirectories(pathToCurrentUserProfiles, "*.default*", SearchOption.TopDirectoryOnly);
            if (pathsToProfiles.Length != 0)
            {

                FirefoxProfile profile = new FirefoxProfile(pathsToProfiles[0]);
                profile.SetPreference("browser.tabs.loadInBackground", false); // set preferences you need
                profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream;application/csv;text/csv;application/vnd.ms-excel;");
                profile.SetPreference("browser.helperApps.alwaysAsk.force", false);
                profile.SetPreference("browser.download.folderList", 2);
                profile.SetPreference("browser.download.dir", @"C:\TempExcel");
                profile.SetPreference("capability.policy.policynames", "strict");
                profile.SetPreference("capability.policy.strict.sites", "http://www.netdating.dk");
                profile.SetPreference("capability.policy.strict.Window.alert", "noAccess");
                profile.SetPreference("capability.policy.strict.Window.confirm", "noAccess");
                profile.SetPreference("capability.policy.strict.Window.prompt", "noAccess");
                fox = new FirefoxDriver(new FirefoxBinary(), profile);
                Console.WriteLine("Profile do firefox carregado com sucesso");
            }
            else
            {
                fox = new FirefoxDriver();
                Console.WriteLine("Profile do firefox nao encontrado");
            }

            Page1:
            Console.WriteLine("Pagina 1");
            try
            {

                fox.Navigate().GoToUrl("https://acessoseguro.gissonline.com.br/index.cfm?m=portal");
                fox.FindElementByName("TxtIdent").SendKeys("560801");
                fox.FindElementByName("TxtSenha").SendKeys("honda2011");
                SendKeys.SendWait("{TAB}");
                SendKeys.SendWait("{TAB}");
                fox.SwitchTo().Frame(0);
                string num1 = fox.FindElementByXPath(@"/html/body/table/tbody/tr/td[1]/img").GetAttribute("value");
                string num2 = fox.FindElementByXPath(@"/html/body/table/tbody/tr/td[2]/img").GetAttribute("value");
                string num3 = fox.FindElementByXPath(@"/html/body/table/tbody/tr/td[3]/img").GetAttribute("value");
                string num4 = fox.FindElementByXPath(@"/html/body/table/tbody/tr/td[4]/img").GetAttribute("value");
                fox.SwitchTo().DefaultContent();
                clickVirtualButton(num1, fox);
                clickVirtualButton(num2, fox);
                clickVirtualButton(num3, fox);
                clickVirtualButton(num4, fox);
                fox.FindElementById("imgLogin").Click();
                Thread.Sleep(5000);
                try
                {
                    fox.SwitchTo().Alert().Accept();
                }
                catch (Exception err)
                {
                    Console.WriteLine(err.Message);

                }

                Thread.Sleep(15000);
                try
                {
                    fox.SwitchTo().Alert().Accept();
                }

                catch (Exception err)
                {

                    Console.WriteLine(err.Message);

                }

                Thread.Sleep(8000);
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                goto Page1;
            }
            Page2:
            Console.WriteLine("Pagina 2");
            try
            {
                fox.SwitchTo().Frame(0);
                fox.FindElement(By.Id("6")).Click();
            }
            

            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                fox.Navigate().Refresh();
                goto Page2;

            }
    
            Page3:
            Console.WriteLine("Pagina 3");
            try
            {
                fox.SwitchTo().DefaultContent();
                fox.SwitchTo().Frame(2);
                DateTime time = DateTime.Now;
                fox.FindElement(By.Name("mes")).SendKeys(time.ToString("MM"));
                fox.FindElement(By.Name("ano")).SendKeys(time.Year.ToString());
                fox.FindElement(By.LinkText("Notas Recebidas")).Click();
            }
            

            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                fox.Navigate().Refresh(); 
                goto Page3;
            }
            int count = 0;
            ReadOnlyCollection<IWebElement> element;
            string mwh;
            bool first = true;
            Page4:
            Console.WriteLine("Pagina 4");
            try
            {
                fox.SwitchTo().DefaultContent();
                fox.SwitchTo().Frame(2);
                new SelectElement(fox.FindElement(By.Name("maxrow"))).SelectByText("500");
                element = fox.FindElementsByXPath("//img[contains(@title,'Dados da nota fiscal')]");
                if (File.Exists(excelpath))
                {
                    File.Delete(excelpath);
                }

                fox.FindElementByXPath("//a[contains(text(),'GERAR ARQUIVO EXCEL')]").Click();
                mwh = fox.CurrentWindowHandle;

            }

            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                fox.Navigate().Refresh();
                goto Page4;
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
            }
            var length = new FileInfo(excelpath).Length;
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

            foreach (var item in cnpjcpfValidos)
            {
                // Console.WriteLine(item.ToString());

            }
            Console.WriteLine(cnpjcpfValidos.Count);

            //loop para abrir as notas
            foreach (var item in element)
            {
                if (cnpjcpfValidos[count] == true)
                {
                    LineCNPJ:
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
                    string CIA = "";

                    if (!first)
                    {
                        fox.SwitchTo().Frame(2);
                    }
                    item.Click();
                    ReadOnlyCollection<string> popups = fox.WindowHandles;
                    fox.SwitchTo().Window(popups[1]);
                    Console.WriteLine(fox.Url);
                    try //Try CNPJ/CPF Prestador
                    {
                        FirefoxWebElement parentcnpj = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:129px;top:176px;')]");
                        CNPJPrestador = parentcnpj.FindElementByXPath(".//*").Text;
                        Console.WriteLine("CNPJ/CPF Prestador: " + CNPJPrestador);
                    }
                    catch (Exception err)
                    {
                        Console.WriteLine(err.Message);
                        Console.WriteLine("cade o CNPJ Prestador");
                        fox.Close();
                        fox.SwitchTo().Window(mwh);
                        goto LineCNPJ;
                    }
                    if (CNPJPrestador == "84.098.383/0001-72")
                    {
                        try //Try nota fiscal eletronica
                        {
                            FirefoxWebElement parentnfe = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:502px;top:52px;')]");
                            nfe = parentnfe.FindElementByXPath(".//*").Text;
                            Console.WriteLine("NFe: " + nfe);
                        }
                        catch (Exception)
                        {
                            nfe = "Não tem nota fiscal???????";
                            Console.WriteLine("Não tem nota fiscal???????");

                        }


                        try //Try Discriminacao
                        {
                            FirefoxWebElement parentdis = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:12px;top:335px;')]");

                            dis = parentdis.FindElementByXPath(".//*").Text;

                            Console.WriteLine("Discriminacao: " + dis);
                        }
                        catch (Exception)
                        {
                            dis = "Não possui Discriminação do Serviço";
                            Console.WriteLine("Não possui Discriminação");
                        }


                        try //Try RPS
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:124px;top:102px;')]");
                            rps = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("RPS: " + rps);

                        }
                        catch (Exception)
                        {
                            rps = "Não possui RPS";
                            Console.WriteLine("Não possui RPS");
                        }

                        try //Try Valor Liquido
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:135px;top:690px;')]");
                            valorliquido = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("Valor liquido: " + valorliquido);

                        }
                        catch (Exception)
                        {
                            valorliquido = "Não possui valor liquido";
                            Console.WriteLine("Não possui valor liquido");
                        }

                        try //Try Valor Servico
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:135px;top:570px;')]");
                            valorservico = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("Valor Servico: " + valorservico);

                        }
                        catch (Exception)
                        {
                            valorservico = "Não possui valor do servico";
                            Console.WriteLine("Não possui valor do servico");
                        }

                        try //Try ISSQN Retido
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:135px;top:670px;')]");
                            ISSQNRetido = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("ISSQN Retido: " + ISSQNRetido);

                        }
                        catch (Exception)
                        {
                            ISSQNRetido = "Nao possui ISSQN Retido";
                            Console.WriteLine("Não possui ISSQN Retido");
                        }

                        try //Try Codigo do Servico
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:12px;top:450px;')]");
                            CODServico = parentrps.FindElementByXPath(".//*").Text;
                            int indexEnd = CODServico.IndexOf("-");
                            CODServico = CODServico.Substring(0, indexEnd - 1);
                            Console.WriteLine("Codigo do Servico: " + CODServico);

                        }
                        catch (Exception)
                        {
                            CODServico = "Nao possui Codigo do Servico";
                            Console.WriteLine("Não possui Codigo do Servico");
                        }

                        try //Try NFe Substituido
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:319px;top:102px;')]");
                            NFeSub = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("NFe Substituido: " + NFeSub);

                        }
                        catch (Exception)
                        {
                            NFeSub = "Nao possui Nfe Substituido";
                            Console.WriteLine("Não possui NFe Substituido");
                        }

                        try //Try Data e Hora
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:124px;top:82px;')]");
                            DataHoraEmissao = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("Data e Hora de Emissao: " + DataHoraEmissao);

                        }
                        catch (Exception)
                        {
                            DataHoraEmissao = "Nao possui Data e Hora de Emissao";
                            Console.WriteLine("Não possui Data e Hora de Emissao");
                        }

                        try //Try Competencia
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:319px;top:82px;')]");
                            Competencia = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("Competencia: " + Competencia);

                        }
                        catch (Exception)
                        {
                            Competencia = "Nao possui Competencia";
                            Console.WriteLine("Não possui Competencia");
                        }

                        try //Try Codigo de Verificação
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:470px;top:82px;')]");
                            CODVerificacao = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("Codigo de Verificacao: " + CODVerificacao);

                        }
                        catch (Exception)
                        {
                            CODVerificacao = "Nao possui Codigo de Verificacao";
                            Console.WriteLine("Não possui Codigo de Verificacao");
                        }

                        try //Try CNPJ do Tomador
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:59px;top:264px;')]");
                            CNPJTomador = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine("CNPJ/CPF Tomador: " + CNPJTomador);

                        }
                        catch (Exception)
                        {
                            CNPJTomador = "Nao possui CNPJ Tomador";
                            Console.WriteLine("Não possui CNPJ Tomador");
                        }

                        if (CNPJPrestador == "04.335.535/0002-55")  //Insert BD SuperTerminais Table
                        {
                            SuperTerminais superterminais = new SuperTerminais(dis, nfe);
                            if (superterminais.BeginAnalysis())
                            {
                                //insert no banco
                                connection.Open();
                                OleDbCommand command = new OleDbCommand();
                                command.Connection = connection;
                                query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , Numeracao , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador) values ('" + nfe + "','" + rps + "','" + dis + "','" + count.ToString() + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','"+NFeSub +"','"+DataHoraEmissao+"','"+ Competencia + "','"+ CODVerificacao + "','"+CNPJPrestador+"','"+CNPJTomador+"')";
                                command.CommandText = query;
                                command.ExecuteNonQuery();
                                connection.Close();
                            }
                            else
                            {
                                Console.WriteLine("DI ZOADA");
                            }
                        }

                        if (CNPJPrestador == "04.694.548/0001-30")  //Insert BD Aurora Table
                        {

                            AuroraEadi auroraeadi = new AuroraEadi(dis, nfe);
                            if (auroraeadi.BeginAnalysis())
                            {
                                connection.Open();
                                OleDbCommand command = new OleDbCommand();
                                command.Connection = connection;
                                query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , Numeracao , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador) values ('" + nfe + "','" + rps + "','" + dis + "','" + count.ToString() + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "')";
                                command.CommandText = query;
                                command.ExecuteNonQuery();
                                connection.Close();
                            }
                            else
                            {
                                Console.WriteLine("Colocaram um navio no meio da avenida!");
                            }


                        }

                        if (CNPJPrestador == "84.098.383/0001-72")  //Insert BD Chibatao Table
                        {
                            Chibatao chibatao = new Chibatao(dis, nfe);
                            if (chibatao.BeginAnalysis())
                            {
                                connection.Open();
                                OleDbCommand command = new OleDbCommand();
                                command.Connection = connection;
                                query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , Numeracao , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador) values ('" + nfe + "','" + rps + "','" + dis + "','" + count.ToString() + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "')";
                                command.CommandText = query;
                                command.ExecuteNonQuery();
                                connection.Close();
                            }
                            else
                            {
                                Console.WriteLine("Colocaram um navio no meio da avenida!");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("WTF");
                    }

                    fox.Close();
                    fox.SwitchTo().Window(mwh);
                    first = false;
                }

                count++;
                Console.WriteLine(count);
            }
            watch.Stop();
            Console.WriteLine("Execution Time: " + (watch.ElapsedMilliseconds / 1000) / 60 + "Minutes");
            Console.ReadLine();

        }

        static private void ListOfCNPJCPF()
        {

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            string workbookPath = @"C:\TempExcel\rel_notas_aceite_04-2016.xls";
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
                if (cnpjcpf == "84098383000172")
                {
                    cnpjcpfValidos.Add(true);

                    //Console.WriteLine(excelWorksheet.Cells[i, 11].Value2);
                }
                else
                {
                    cnpjcpfValidos.Add(false);
                }

                i++;
            }
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelSheets);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(workbooks);
            excelApp.Quit();

        }
    }

}