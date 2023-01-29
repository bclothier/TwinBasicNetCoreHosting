using System.Runtime.InteropServices;

namespace WpfServer
{
    [
        ComVisible(true),
        Guid("1E43F347-4E45-42E7-88BF-088E53F428CB"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IWpfHostFactory
    {
        IWpfHostCreateParameters CreateParameters();
        IWpfHost CreateHost(IWpfHostCreateParameters parameters);
    }
}