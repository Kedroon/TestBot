using System;


namespace AlgorithmTestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //SuperTerminais superterminais = new SuperTerminais();
            //AuroraEadi auroraeadi = new AuroraEadi();
            Chibatao chibatao = new Chibatao();
            var watch = System.Diagnostics.Stopwatch.StartNew();
            chibatao.StartAnalysis();
            //superterminais.StartAnalysis();
            //auroraeadi.StartAnalysis();
            watch.Stop();
            Console.WriteLine("Execution Time: " + (watch.ElapsedMilliseconds) + "Ms");
            Console.Read();
        }
    }
}
