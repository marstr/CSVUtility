using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utilities.CSV;

namespace PerfExperiments
{
    public class WriteToStreamExperiment : IPerfExperiment
    {
        protected CSVWriter writer;

        public string Title
        {
            get
            {
                return _Title;
            }
        }
        private readonly string _Title;

        public const string DEFAULT_TITLE = "Stream Writer Experiment";

        public WriteToStreamExperiment(CSVWriter writer, string title = DEFAULT_TITLE)
        {
            this.writer = writer;
            _Title = title;
        }

        public WriteToStreamExperiment(Stream output, string title = DEFAULT_TITLE)
        {
            writer = new CSVWriter(output);
            _Title = title;
        }

        public TimeSpan RunExperiment()
        {
            return RunExperiment(10000, 10);
        }

        public TimeSpan RunExperiment(int rows, int columns)
        {
            var writeData = new int[columns];

            var r = new Random();

            for(int i = 0; i < columns; i++)
            {
                writeData[i] = r.Next();
            }

            var clock = new Stopwatch();
            clock.Start();
            for(int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    writer.WriteTuple(writeData);
                }
            }
            clock.Stop();
            return clock.Elapsed;
        }
    }
}
