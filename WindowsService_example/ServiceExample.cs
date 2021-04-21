using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService_example
{
    public partial class ServiceExample : ServiceBase
    {
        public ServiceExample()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
        }

        public void DebugMode() {
            Debug.WriteLine("Debug mode!");
        }

        protected override void OnStop()
        {
        }
    }
}
