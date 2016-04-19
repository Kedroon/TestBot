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

namespace TestBot
{
    public partial class Form1 : Form
    {
        FirefoxDriver fox;
        private OleDbConnection connection = new OleDbConnection();
        string query;

        public Form1()
        {

            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Usuários\sb042182\Desktop\Notas.accdb;
Persist Security Info=False;";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            automation();

        }

        private void clickVirtualButton(string num, FirefoxDriver fox)
        {
            fox.FindElementByXPath("//img[contains(@src,'/images/teclado/tec_" + num + ".gif')]").Click();

        }

        private void automation() {

            string pathToCurrentUserProfiles = Environment.ExpandEnvironmentVariables("%APPDATA%") + @"\Mozilla\Firefox\Profiles"; // Path to profile
            string[] pathsToProfiles = Directory.GetDirectories(pathToCurrentUserProfiles, "*.default", SearchOption.TopDirectoryOnly);
            if (pathsToProfiles.Length != 0)
            {
                FirefoxProfile profile = new FirefoxProfile(pathsToProfiles[0]);
                profile.SetPreference("browser.tabs.loadInBackground", false); // set preferences you need
                fox = new FirefoxDriver(new FirefoxBinary(), profile);
            }
            else
            {
                fox = new FirefoxDriver();
            }
            try
            {
                fox.Navigate().GoToUrl("https://acessoseguro.gissonline.com.br/index.cfm?m=portal");
                //fox.FindElementByName("TxtIdent").Clear();
                // SendKeys.SendWait("{TAB}");
                fox.FindElementByName("TxtIdent").SendKeys("560801");
                //fox.FindElementByName("TxtSenha").Clear();
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
                fox.SwitchTo().Alert().Accept();
                //SendKeys.SendWait("{ENTER}");
                Thread.Sleep(10000);
                fox.SwitchTo().Alert().Accept();
                Thread.Sleep(8000);
                fox.SwitchTo().Frame(0);
                fox.FindElement(By.Id("6")).Click();
                fox.SwitchTo().DefaultContent();
                fox.SwitchTo().Frame(2);
                fox.FindElement(By.Name("mes")).SendKeys("04");
                fox.FindElement(By.Name("ano")).SendKeys("2016");
                fox.FindElement(By.LinkText("Notas Recebidas")).Click();
                fox.SwitchTo().DefaultContent();
                fox.SwitchTo().Frame(2);
                new SelectElement(fox.FindElement(By.Name("maxrow"))).SelectByText("500");
                ReadOnlyCollection<IWebElement> element = fox.FindElementsByXPath("//img[contains(@title,'Dados da nota fiscal')]");
                // MessageBox.Show(element.Count.ToString());
                string mwh = fox.CurrentWindowHandle;
                int count = 0;
                foreach (var item in element)
                {
                    string cnpj = "";
                    string nfe = "";
                    string rps = "";
                    string dis = "";
                    if (count != 0)
                    {
                        fox.SwitchTo().Frame(2);
                    }
                    item.Click();
                    ReadOnlyCollection<string> popups = fox.WindowHandles;
                    fox.SwitchTo().Window(popups[1]);
                    try
                    {
                        Thread.Sleep(1000);
                        FirefoxWebElement parentcnpj = (FirefoxWebElement)fox.FindElementByXPath("//span[starts-with(@style,'position:absolute;left:129px;top:176px;')]");
                        cnpj = parentcnpj.FindElementByXPath(".//*").Text;
                        Console.WriteLine(cnpj);
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
                            finally
                            {
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
                                finally
                                {
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
                                    finally //Finally para inserir no banco e repetir o loop
                                    {
                                        dis.Replace("'", "");
                                        connection.Open();
                                        OleDbCommand command = new OleDbCommand();
                                        command.Connection = connection;
                                        query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , Numeracao) values ('" + nfe + "','" + rps + "','" + dis + "','" + count.ToString() + "')";
                                        command.CommandText = query;
                                        command.ExecuteNonQuery();
                                        connection.Close();


                                    }
                                }

                            }
                        }
                    }
                    catch (Exception err)
                    {
                        Console.WriteLine(err.Message);
                        Console.WriteLine("cade o CNPJ");
                    }
                    finally
                    {
                        fox.Close();
                        fox.SwitchTo().Window(mwh);
                        count++;
                        Console.WriteLine(count);
                    }





                }



            }
            catch (Exception err)
            {
               // MessageBox.Show(err.Message);
                fox.Close();
                fox.Dispose();
                automation();

            }

        }
    }
}