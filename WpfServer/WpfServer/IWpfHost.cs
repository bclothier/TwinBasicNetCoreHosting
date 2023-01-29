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
        void Advise([MarshalAs(UnmanagedType.IUnknown)] ref object client);
        IntPtr Handle { get; }
        void Dispose();
    }
}