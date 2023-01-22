using System.Runtime.InteropServices;

namespace WpfServer
{
    [
        ComVisible(true),
        Guid("A4C1E9FC-BE79-4546-BD3D-AFA1093C4A27"),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IWpfHostCreateParameters
    {
        // Summary:
        //     Gets or sets the Microsoft Windows class style for the window.
        //
        // Returns:
        //     The window class style. See Window Class Styles.
        public int WindowClassStyle { get; set; }
        //
        // Summary:
        //     Gets or sets a value that indicates the width of the window.
        //
        // Returns:
        //     The window width, in device pixels. The default value is 1.
        public int Width { get; set; }
        //
        // Summary:
        //     Gets a value that declares whether the per-pixel transparency of the source window
        //     content is respected.
        //
        // Returns:
        //     true if using per-pixel transparency; otherwise, false.
        public bool UsesPerPixelTransparency { get; set; }
        //
        // Summary:
        //     Gets a value that declares whether the per-pixel opacity of the source window
        //     content is respected.
        //
        // Returns:
        //     true if using per-pixel opacity; otherwise, false.
        public bool UsesPerPixelOpacity { get; set; }
        //
        // Summary:
        //     Gets or sets a value that indicates whether the System.Windows.Interop.HwndSource
        //     should receive window messages raised by the message pump via the System.Windows.Interop.ComponentDispatcher.
        //
        // Returns:
        //     true if the System.Windows.Interop.HwndSource should receive window messages
        //     raised by the message pump via the System.Windows.Interop.ComponentDispatcher;
        //     otherwise, false. The default is true if the System.Windows.Interop.HwndSource
        //     corresponds to a top-level window; otherwise, the default is false.
        public bool TreatAsInputRoot { get; set; }
        //
        // Summary:
        //     Gets or sets a value that indicates whether the parent windows of the System.Windows.Interop.HwndSource
        //     should be considered the non-client area of the window during layout passes.
        //
        // Returns:
        //     true if parent windows of the System.Windows.Interop.HwndSource should be considered
        //     the non-client area of the window during layout passes.; otherwise, false. The
        //     default is false.
        public bool TreatAncestorsAsNonClientArea { get; set; }
        //
        // Summary:
        //     Gets or sets how WPF handles restoring focus to the window.
        //
        // Returns:
        //     One of the enumeration values that specifies how WPF handles restoring focus
        //     for the window. The default is System.Windows.Input.RestoreFocusMode.Auto.
        public int RestoreFocusMode { get; set; }
        //
        // Summary:
        //     Gets or sets the upper-edge position of the window.
        //
        // Returns:
        //     The upper-edge position of the window. The default is CW_USEDEFAULT, as processed
        //     by CreateWindow.
        public int PositionY { get; set; }
        //
        // Summary:
        //     Gets or sets the left-edge position of the window.
        //
        // Returns:
        //     The left-edge position of the window. The default is CW_USEDEFAULT, as processed
        //     by CreateWindow.
        public int PositionX { get; set; }
        //
        // Summary:
        //     Gets or sets the window handle (HWND) of the parent for the created window.
        //
        // Returns:
        //     The HWND of the parent window.
        public IntPtr ParentWindow { get; set; }
        //
        // Summary:
        //     Gets or sets the message hook for the window.
        //
        // Returns:
        //     The message hook for the window.
        public IntPtr HwndSourceHook { get; set; }
        //
        // Summary:
        //     Gets or sets a value that indicates the height of the window.
        //
        // Returns:
        //     The height of the window, in device pixels. The default value is 1.
        public int Height { get; set; }
        //
        // Summary:
        //     Gets a value that indicates whether a size was assigned.
        //
        // Returns:
        //     true if the window size was assigned; otherwise, false. The default is false,
        //     unless the structure was created with provided height and width, in which case
        //     the value is true.
        public bool HasAssignedSize { get; }
        //
        // Summary:
        //     Gets or sets the extended Microsoft Windows styles for the window.
        //
        // Returns:
        //     The extended window styles. See CreateWindowEx.
        public int ExtendedWindowStyle { get; set; }
        //
        // Summary:
        //     Gets or sets a value that indicates whether to include the nonclient area for
        //     sizing.
        //
        // Returns:
        //     true if the layout manager sizing logic should include the nonclient area; otherwise,
        //     false. The default is false.
        public bool AdjustSizingForNonClientArea { get; set; }
        //
        // Summary:
        //     Gets or sets the value that determines whether to acquire Win32 focus for the
        //     WPF containing window when an System.Windows.Interop.HwndSource is created.
        //
        // Returns:
        //     true to acquire Win32 focus for the WPF containing window when the user interacts
        //     with menus; otherwise, false. null to use the value of System.Windows.Interop.HwndSource.DefaultAcquireHwndFocusInMenuMode.
        public bool AcquireHwndFocusInMenuMode { get; set; }
        //
        // Summary:
        //     Gets or sets the name of the window.
        //
        // Returns:
        //     The window name.
        public string WindowName { get; set; }
        //
        // Summary:
        //     Gets or sets the style for the window.
        //
        // Returns:
        //     The window style. See the CreateWindowEx function for a complete list of style
        //     bits. Defaults: WS_VISIBLE, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX,
        //     WS_MAXIMIZEBOX, WS_CLIPCHILDREN.
        public int WindowStyle { get; set; }

        //
        // Summary:
        //     Sets the values that are used for the screen position of the window for the System.Windows.Interop.HwndSource.
        //
        // Parameters:
        //   x:
        //     The position of the left edge of the window.
        //
        //   y:
        //     The position of the upper edge of the window.
        public void SetPosition(int x, int y);
        //
        // Summary:
        //     Sets the values that are used for the window size of the System.Windows.Interop.HwndSource.
        //
        // Parameters:
        //   width:
        //     The width of the window, in device pixels.
        //
        //   height:
        //     The height of the window, in device pixels.
        public void SetSize(int width, int height);
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