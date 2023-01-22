using System.Runtime.InteropServices;

namespace WpfServer
{
    [
        ComVisible(true),
        Guid("19B1E709-EBFD-4561-859D-CD38A8D575AF"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IWpfHost
    {
        IntPtr Handle { get; }
        void Dispose();
    }

    /*
    public static class WpfHostFactory
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        public delegate object CreateParametersDelegate();
        public static object CreateParameters()
        {
            return new WpfHostCreateParameters();
        }

        [return: MarshalAs(UnmanagedType.IUnknown)]
        public delegate object CreateHostDelegate([MarshalAs(UnmanagedType.IUnknown)] IWpfHostCreateParameters parameters);
        public static object CreateHost(IWpfHostCreateParameters parameters)
        {
            var _parameters = parameters as WpfHostCreateParameters;
            return new WpfHost(_parameters.HwndSourceParameters());
        }
    }
    */
}