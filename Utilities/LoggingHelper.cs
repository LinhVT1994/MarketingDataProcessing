using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketingDataProcessing.Utilities
{
    class LoggingHelper
    {
        internal static void WriteDown(string message)
        {
            Debug.WriteLine(message);
        }
        internal static void LogForSuccession(string message)
        {
            Debug.WriteLine(message);
            using (StreamWriter stream = new StreamWriter(MainWindow.RootDir + @"\data.json",append:true))
            {
                stream.WriteLine(message);
            }
        }
    }
}
