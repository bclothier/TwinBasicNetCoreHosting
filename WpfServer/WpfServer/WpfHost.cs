using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace WpfServer
{

    [
        ComVisible(true),
        Guid("19B1E709-EBFD-4561-859D-CD38A8D575AF"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("Demo.WpfHost"),
        ComSourceInterfaces(typeof(IWpfHostEvents))
    ]
    public class WpfHost : HwndSource, IWpfHost
    {
        private readonly Simple _simple;

        internal WpfHost(HwndSourceParameters parameters) : base(parameters)
        {
            AppDomain.CurrentDomain.UnhandledException += Domain_UnhandledException;
            Dispatcher.UnhandledException += Dispatcher_UnhandledException;
            _simple = new Simple();
            RootVisual = _simple;
            _simple.Visibility = Visibility.Visible;
            _simple.MouseDown += Host_MouseDown;
        }

        private delegate void MouseDownDelegate();
        private event MouseDownDelegate? MouseDown;
        private void Host_MouseDown(object sender, MouseButtonEventArgs e)
        {
            MouseDown?.Invoke();
        }

        private void Host_Resize(int width, int height)
        {
            MessageBox.Show($"New size: {width}, {height}");
        }

        private void Host_Unload(ref bool cancel)
        {
            cancel = MessageBox.Show("Close?", default, MessageBoxButton.YesNo) != MessageBoxResult.Yes;
        }

        private void Dispatcher_UnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString());
        }

        private void Domain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show(e.ExceptionObject.ToString());
        }

        public void Advise([MarshalAs(UnmanagedType.IUnknown)] ref object client)
        {
            ComEventsHelper.Combine(client, new Guid("C1B1A5B2-979E-409C-B573-7943971A4CCA"), 1, Host_Resize);
            ComEventsHelper.Combine(client, new Guid("C1B1A5B2-979E-409C-B573-7943971A4CCA"), 2, Host_Unload);
        }
    }
}