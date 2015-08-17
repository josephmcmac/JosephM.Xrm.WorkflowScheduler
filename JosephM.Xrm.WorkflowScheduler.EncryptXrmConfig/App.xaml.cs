using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using JosephM.ObjectEncryption;
using JosephM.Prism.Infrastructure.Prism;
using JosephM.Xrm.WorkflowScheduler.EncryptXrmConfig.Module;
using JosephM.Xrm.Settings.Test;

namespace JosephM.Xrm.WorkflowScheduler.EncryptXrmConfig
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var prism = new PrismApplication("Test Prism Application");
            prism.AddModule<ObjectEncryptModule<EncryptXrmConfigurationDialog, TestXrmConfiguration>>();
            prism.Run();
        }
    }
}
