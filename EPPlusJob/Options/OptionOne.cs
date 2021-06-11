using ConsoleCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusJob.Options
{
    /// <summary>
    /// Collection to Excel with formatting option
    /// </summary>
    public class OptionOne : ParamsObject
    {
        public OptionOne(string[] args) 
            : base(args)
        {
                
        }

        [Switch("CS",true)]
        public string ConnectionString { get; set; }
        [Switch("Q", true)]
        public string Query { get; set; }
        [Switch("FP", true)]
        public string  FilePath { get; set; }
        [Switch("SN", true)]
        public string SheetName { get; set; }
        [Switch("NF", false)]
        public string NumberFormat { get; set; }
        [Switch("DTC", false)]
        public string DateTimeColumns { get; set; }
        [Switch("DIF", false)]
        public int DeleteIfExist { get; set; }
        [Switch("RS", true)]
        public int RowStart { get; set; }
        //public int CheckDate { get; set; }
        [Switch("OT", true)]
        public outputType OutputType { get; set; }

        public bool DeleteExistingFile
        {
            get
            {
                if (DeleteIfExist == 1)
                    return true;
                else
                    return false;
            }
        }
    }

    public enum outputType
    {
        Excel = 1,
        CSV = 2
    }
}
