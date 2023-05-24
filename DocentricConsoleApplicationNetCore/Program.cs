// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using DocentricConsoleApplicationNetCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DocentricConsoleApplication
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ImageGeneration imageGeneration = new ImageGeneration();
            var file = imageGeneration.GenerateReport();
            Console.WriteLine(file);

        }


    }
}

