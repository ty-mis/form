using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using tomlib;
namespace cons
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime dt = DateTime.Now;

            Console.WriteLine(TomLib.getWeek(dt));

            Console.ReadKey();
        }
    }
}
