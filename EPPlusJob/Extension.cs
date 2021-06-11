using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusJob
{
    public static class Extension
    {
        public static string GetKeyValue(this string key)
        {
            return ConfigurationManager.AppSettings[key];
        }
    }
}
