using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Interop;

namespace WpfServer
{
    [
        ComVisible(true),
        Guid("A903DA0A-86EA-443B-BE97-B770F849E5C8"),
        ClassInterface(ClassInterfaceType.None),
        ProgId("Demo.WpfHostCreateParameters")
    ]
    public class WpfHostCreateParameters : IWpfHostCreateParameters
    {
        public int WindowClassStyle { get; set; }

        private int _width = 1;
        public int Width
        {
            get => _width;
            set
            {
                _width = value;
                _hasAssignedSize = true;
            }
        }
        public bool UsesPerPixelTransparency { get; set; }
        public bool UsesPerPixelOpacity { get; set; }
        public bool TreatAsInputRoot { get; set; }
        public bool TreatAncestorsAsNonClientArea { get; set; }

        private RestoreFocusMode _restoreFocusMode = System.Windows.Input.RestoreFocusMode.Auto;
        public int RestoreFocusMode
        {
            get => (int)_restoreFocusMode;
            set
            {
                if (value == 0)
                {
                    _restoreFocusMode = System.Windows.Input.RestoreFocusMode.Auto;
                }
                else
                {
                    _restoreFocusMode = System.Windows.Input.RestoreFocusMode.None;
                }
            }
        }
        public int PositionY { get; set; }
        public int PositionX { get; set; }
        public IntPtr ParentWindow { get; set; }
        public IntPtr HwndSourceHook { get; set; }

        private int _height = 1;
        public int Height
        {
            get => _height;
            set
            {
                _height = value;
                _hasAssignedSize = true;
            }
        }

        private bool _hasAssignedSize = false;
        public bool HasAssignedSize => _hasAssignedSize;

        public int ExtendedWindowStyle { get; set; }
        public bool AdjustSizingForNonClientArea { get; set; }
        public bool AcquireHwndFocusInMenuMode { get; set; }
        public string WindowName { get; set; } = string.Empty;
        public int WindowStyle { get; set; }

        public void SetPosition(int x, int y)
        {
            PositionX = x;
            PositionY = y;
        }

        public void SetSize(int width, int height)
        {
            Width = width;
            Height = height;
        }

        internal HwndSourceParameters HwndSourceParameters() => new()
        {
            AcquireHwndFocusInMenuMode = AcquireHwndFocusInMenuMode,
            AdjustSizingForNonClientArea = AdjustSizingForNonClientArea,
            ExtendedWindowStyle = ExtendedWindowStyle,
            Height = Height,
            HwndSourceHook = HwndSourceHook == IntPtr.Zero ? null : Marshal.GetDelegateForFunctionPointer<HwndSourceHook>(HwndSourceHook),
            ParentWindow = ParentWindow,
            PositionX = PositionX,
            PositionY = PositionY,
            RestoreFocusMode = _restoreFocusMode,
            TreatAncestorsAsNonClientArea = TreatAncestorsAsNonClientArea,
            TreatAsInputRoot = TreatAsInputRoot,
            UsesPerPixelOpacity = UsesPerPixelOpacity,
            UsesPerPixelTransparency = UsesPerPixelTransparency,
            Width = Width,
            WindowClassStyle = WindowClassStyle,
            WindowName = WindowName,
            WindowStyle = WindowStyle
        };
    }
}