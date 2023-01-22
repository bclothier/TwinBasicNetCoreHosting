using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace WpfServer
{

    [
        ComVisible(true),
        Guid("19B1E709-EBFD-4561-859D-CD38A8D575AF"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("Demo.WpfHost")
    ]
    public class WpfHost : HwndSource, IWpfHost
    {
        internal WpfHost(HwndSourceParameters parameters) : base(parameters)
        {
            Dispatcher.UnhandledException += Dispatcher_UnhandledException;
            RootVisual = new Simple();
            ((Simple)RootVisual).Visibility = Visibility.Visible;
        }

        private void Dispatcher_UnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString());
        }
    }
}