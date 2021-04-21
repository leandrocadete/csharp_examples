using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService_example {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main() {
#if DEBUG
            ServiceExample service = new ServiceExample();
            service.DebugMode();
            
            
            return;
#endif
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] { new ServiceExample() };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
