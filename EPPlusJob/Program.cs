using EPPlusJob.Actions;
using EPPlusJob.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusJob
{
    class Program
    {
        static Logger logger = new Logger();
        static void Main(string[] args)
        {
            logger.Info("EPPlus Job Started");
            try
            {
                OptionOne opt = new OptionOne(args);
                DatasetToExcelWithFormattingAndRowStart collectionToExcel = new DatasetToExcelWithFormattingAndRowStart(opt, logger);
                collectionToExcel.Run();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                logger.Info("An error occurred");
               // throw;
            }
            
        }
    }
}
