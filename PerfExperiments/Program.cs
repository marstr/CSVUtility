using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PerfExperiments
{
    class Program
    {
        static void Main(string[] args)
        {
            var rows = 100000;
            var cols = 10;

            var totalCells = rows * cols;
            var tempFile = Path.GetTempFileName();

            var experiments = new[] {
                new WriteToStreamExperiment(new MemoryStream(), string.Format("Write {0} Cells to Memory", totalCells)),
                new WriteToStreamExperiment(new FileStream(tempFile,FileMode.Create), string.Format("Write {0} Cells to File", totalCells))};
                    

            foreach (var runner in experiments)
            {
                Console.Write("Starting {0}...", runner.Title);
                var result = runner.RunExperiment(rows, cols);
                Console.WriteLine("Finished in {0} seconds.", result.TotalSeconds);                               
            }            

            Console.Write("Press any key to continue...");
            Console.ReadKey();
        }
    }
}
