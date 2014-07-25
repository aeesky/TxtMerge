using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TextMerge.Helpers
{
    public class DiretoryHelper
    {
        public static string Prex = "bsFrequencies.txt";

        public static IEnumerable<string> GetFiles(string dpath)
        {
            return Directory.GetFiles(dpath, "*.txt", SearchOption.AllDirectories)
                .Where(p => p.EndsWith(Prex))
                .OrderBy<string, int>(
                    file =>
                    {
                        try
                        {
                            return Convert.ToInt16(Directory.GetParent(file).Name);
                        }
                        catch (System.Exception)
                        {
                            return 0;
                        }
                    });
        }
    }
}
