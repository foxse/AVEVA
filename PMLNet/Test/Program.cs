using Aveva.Core.InstLoader;
using System;

namespace Test
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var loader = new InstLoader();

            loader.Start();

            Console.ReadLine();
        }
    }
}
