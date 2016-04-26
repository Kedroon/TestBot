using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AlgorithmTestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            SuperTerminais superterminais = new SuperTerminais();
            var watch = System.Diagnostics.Stopwatch.StartNew();
            superterminais.StartAnalysis();
            watch.Stop();
            Console.WriteLine("Execution Time: " + (watch.ElapsedMilliseconds) + "Ms");
            Console.Read();
        }
    }
}
