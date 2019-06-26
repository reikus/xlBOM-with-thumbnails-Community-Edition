using Prism.DryIoc;
using Prism.Ioc;
using Prism.Regions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Navigation;

namespace ThumbnailedBOM
{
    public class Application : PrismApplication
    {

        protected override Window CreateShell()
        {
            return this.Container.Resolve<Shell>();
        }
        protected override void InitializeShell(Window shell)
        {
            this.MainWindow = shell;
            IntPtr windowHandle = new WindowInteropHelper(shell).Handle;
            AddInContext.ApplicationWindow = new AppWindow() { Handle = windowHandle };

        }
        public void Restart()
        {
            this.Initialize();
            this.MainWindow.ShowDialog();
        }
        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterForNavigation<Views.Main>();
        }


        
    }
}
