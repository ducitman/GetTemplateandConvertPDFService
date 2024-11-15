using GetTemplateandConvertPDFService.Models;
using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetTemplateandConvertPDFService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DBHelperOracle.SetConnection("(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 10.118.11.26)(PORT = 1521))(CONNECT_DATA =(SERVER = SHARED)(SERVICE_NAME = BOSSDB)))", "SPEC", "SPEC");
            DBLOGHelper.SetConnection("10.118.11.111", "BTMVAppsLog", "sa", "@dmin15!!");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
