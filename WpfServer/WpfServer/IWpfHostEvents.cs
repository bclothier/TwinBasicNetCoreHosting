using System.Runtime.InteropServices;

namespace WpfServer
{
    [
        ComVisible(true),
        Guid("7F18FA39-1108-4E7F-A8C1-90BB30E96E71"),
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)
    ]
    public interface IWpfHostEvents
    {
        [DispId(1)]
        void MouseDown();
    }
}
