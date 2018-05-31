using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvToKml
{
    class Program
    {
        static void Main(string[] args)
        {
            Map m;
            try
            {
                switch (args.Length)
                {
                    case 0:
                        Console.WriteLine("No file(s) specified.");
                        return;
                    case 1:
                        m = new Map(args[0]);
                        break;
                    case 2:
                        m = new Map(args[0], args[1]);
                        break;
                    default:
                        m = new Map(args[0], args[1]);
                        return;
                }
                m.ProduceKml(args[0] + ".kml");
                m.Labels = LabelMode.Some;
                m.ProduceKml(args[0] + "_somelabels.kml");
                m.Labels = LabelMode.None;
                m.ProduceKml(args[0] + "_nolabels.kml");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }            
        }
    }
}
