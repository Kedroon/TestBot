using System;
using System.Threading;


namespace BetterBotPMM
{
    class Program
    {


        static void Main(string[] args)
        {
            while (true)
            {

                Automation automationhda = new Automation("560801","HDA");
                Console.WriteLine("Acessando HDA");
                automationhda.startautomation();
                Console.WriteLine("HDA Concluído");
                Automation automationhca = new Automation("4244701","HCA");
                Console.WriteLine("Acessando HCA");
                automationhca.startautomation();
                Console.WriteLine("HCA Concluído");
                Automation automationhta = new Automation("5951701","HTA");
                Console.WriteLine("Acessando HTA");
                automationhta.startautomation();
                Console.WriteLine("HTA Concluído");
                Thread.Sleep(900000);
            }

        }


    }
}