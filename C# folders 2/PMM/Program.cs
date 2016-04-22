using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.IO;
using OpenQA.Selenium.Support.UI;
using System.Data.OleDb;
using System.Threading;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;

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

            string pathToCurrentUserProfiles = Environment.ExpandEnvironmentVariables("%APPDATA%") + @"\Mozilla\Firefox\Profiles"; // Path to profile
            string[] pathsToProfiles = Directory.GetDirectories(pathToCurrentUserProfiles, "*.default*", SearchOption.TopDirectoryOnly);
            if (pathsToProfiles.Length != 0)
            {
                Console.WriteLine("hi");
                FirefoxProfile profile = new FirefoxProfile(pathsToProfiles[0]);
                profile.SetPreference("browser.tabs.loadInBackground", false); // set preferences you need
                profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream;application/csv;text/csv;application/vnd.ms-excel;");
                profile.SetPreference("browser.helperApps.alwaysAsk.force", false);
                profile.SetPreference("browser.download.folderList", 2);
                profile.SetPreference("browser.download.dir", @"C:\TempExcel");
                fox = new FirefoxDriver(new FirefoxBinary(), profile);

            }
            else
            {
                fox = new FirefoxDriver();
            }

            Page1:

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
                
                Thread.Sleep(5000);
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
                fox.Navigate().Refresh(); ;
                goto Page3;
            }
            int count = 0;
            ReadOnlyCollection<IWebElement> element;
            string mwh;
            bool first = true;
            Page4:

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

            ListOfCNPJCPF(); //Analisar planilha
            Thread.Sleep(2000);


            foreach (var item in cnpjcpfValidos)
            {
                Console.WriteLine(item.ToString());

            }
            Console.WriteLine(cnpjcpfValidos.Count);

            //loop para abrir as notas
            foreach (var item in element)
            {
                if (cnpjcpfValidos[count] == true)
                {
                    LineCNPJ:
                    string cnpj = "";
                    string nfe = "";
                    string rps = "";
                    string dis = "";
                    if (!first)
                    {
                        fox.SwitchTo().Frame(2);
                    }
                    item.Click();
                    ReadOnlyCollection<string> popups = fox.WindowHandles;
                    fox.SwitchTo().Window(popups[1]);
                    try //Try CNPJ/CPF
                    {
                        FirefoxWebElement parentcnpj = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:129px;top:176px;')]");
                        cnpj = parentcnpj.FindElementByXPath(".//*").Text;
                        Console.WriteLine(cnpj);
                    }
                    catch (Exception err)
                    {
                        Console.WriteLine(err.Message);
                        Console.WriteLine("cade o CNPJ");
                        fox.Close();
                        fox.SwitchTo().Window(mwh);
                        goto LineCNPJ;
                    }
                    if (cnpj == "04.335.535/0002-55")
                    {
                        try //Try nota fiscal eletronica
                        {
                            FirefoxWebElement parentnfe = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:502px;top:52px;')]");
                            nfe = parentnfe.FindElementByXPath(".//*").Text;
                            Console.WriteLine(nfe);
                        }
                        catch (Exception)
                        {
                            nfe = "Não tem nota fiscal???????";
                            Console.WriteLine("Não tem nota fiscal???????");

                        }

                        try //Try RPS
                        {
                            FirefoxWebElement parentrps = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:124px;top:102px;')]");
                            rps = parentrps.FindElementByXPath(".//*").Text;
                            Console.WriteLine(rps);

                        }
                        catch (Exception)
                        {
                            rps = "Não possui RPS";
                            Console.WriteLine("Não possui RPS");
                        }


                        try //Try Discriminacao
                        {
                            FirefoxWebElement parentdis = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:12px;top:335px;')]");

                            dis = parentdis.FindElementByXPath(".//*").Text;
                            Console.WriteLine(dis);
                        }
                        catch (Exception)
                        {
                            dis = "Não possui Discriminação do Serviço";
                            Console.WriteLine("Não possui Discriminação");
                        }

                        //insert no banco
                        connection.Open();
                        OleDbCommand command = new OleDbCommand();
                        command.Connection = connection;
                        query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , Numeracao) values ('" + nfe + "','" + rps + "','" + dis + "','" + count.ToString() + "')";
                        command.CommandText = query;
                        command.ExecuteNonQuery();
                        connection.Close();


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


        }

        static private void ListOfCNPJCPF()
        {

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            string workbookPath = @"C:\TempExcel\rel_notas_aceite_04-2016.xls";
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath,
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
                if (cnpjcpf == "4335535000255")
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
            excelApp.Quit();
        }
    }

}

