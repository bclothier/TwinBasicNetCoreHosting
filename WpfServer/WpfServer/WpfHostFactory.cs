using System.Runtime.InteropServices;

namespace WpfServer
{
    [
        ComVisible(true),
        Guid("65FCC31B-65E8-43F8-A3E0-832C6C56AA7E"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("Demo.WpfHostFactory")
    ]
    public class WpfHostFactory : IWpfHostFactory
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        public delegate object CreateWpfHostFactoryDelegate();
        public static object CreateWpfHostFactory()
        {
            return new WpfHostFactory();
        }

        public IWpfHostCreateParameters CreateParameters()
        {
            return new WpfHostCreateParameters();
        }

        public IWpfHost CreateHost(IWpfHostCreateParameters parameters)
        {
            if (parameters is not WpfHostCreateParameters param) throw new ArgumentOutOfRangeException(nameof(parameters));

            return new WpfHost(param.HwndSourceParameters());
        }
    }
}