using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace SEI
{

    class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        static void Main(string[] args)
        {
            if (System.Environment.UserInteractive)
            {
                var sw = new StreamWriter(Console.OpenStandardOutput())
                {
                    AutoFlush = true
                };
                Console.SetOut(sw);
                Console.WriteLine("sdfsdfs");

                if (args[0].ToUpper() == "/I")
                {
                    try
                    {
                        System.Configuration.Install.ManagedInstallerClass.InstallHelper(new string[] { System.Reflection.Assembly.GetExecutingAssembly().Location });
                    }
                    catch
                    {
                    }
                    return;
                }
                if (args[0].ToUpper() == "/D")
                {
                    try
                    {
                        System.Configuration.Install.ManagedInstallerClass.InstallHelper(new string[] { "/u", System.Reflection.Assembly.GetExecutingAssembly().Location });
                    }
                    catch
                    {
                    }
                return;
                }
            }
            else {               
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                new SEIS()
                };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
