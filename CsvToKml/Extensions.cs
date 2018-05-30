using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvToKml
{
    public static class Extensions
    {
        public static string Sanitize(this string me)
        {
            return me.Replace(' ', '_');
        }
    }
}
