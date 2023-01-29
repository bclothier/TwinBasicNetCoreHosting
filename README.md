# TwinBasicNetCoreHosting

This is a demonstration project showing how twinBASIC can host a .NET 7. Though .NET does support COM activation where the hosting is automatically handled by the internal implementation of the generated `.comhost.dll` file, this simply show how the hosting can be done manually and controlled by twinBASIC without any COM registration required. Furthermore, the project loads a WPF Page UI control into the window provided by the twinBASIC, demonstrating how to host WPF directly by twinBASIC.

# Setup

This requires Visual Studio 2022, .NET 7 SDK, and twinBASIC (beta 235). The output should be placed in a folder named `Build` at the root of directory which is not included in the git repository. The Visual Studio project will attempt to copy the output of its build into the folder which is the same as the output used by twinBASIC's project. The post build event sues a simple copy command that assuems the path exists. 

Furthermore, the `nethost.dll` should be copied from the `Resources\nethost.dll` into the `Build`. After a successful build of both projects, the `Build` folder should contain the minimum of the following files:

* `nethost.dll`
* `WpfClient_win64.exe`
* `WpfServer.deps.json`
* `WpfServer.dll`
* `WpfServer.runtimeconfig.json`

The output is tested for 64-bit only and most likely will not work with 32-bit. Because the .NET resources are used in-process by an unmanaged host, the `Any CPU` is not strictly proper. Therefore, the `x64` is the default built in both Visual Studio and twinBASIC project. 

# Mechanics 

In order for the .NET Core library to be dynamically loaded at runtime, it must have the following property in the `.csproj`:

```
<EnableDynamicLoading>true</EnableDynamicLoading>
```

The demonstation does not use [the usual approach to activate COM objects via .NET host](https://github.com/dotnet/samples/tree/main/core/extensions/COMServerDemo). Rather, the project uses a function as an entry point to provide the COM object. This can be found in the `WpfHostFactory` class:
```
public class WpfHostFactory : IWpfHostFactory
{
	[return: MarshalAs(UnmanagedType.IUnknown)]
	public delegate object CreateWpfHostFactoryDelegate();
	public static object CreateWpfHostFactory()
	{
		return new WpfHostFactory();
	}
	
	...
}
```

On the twinBASIC side, the function is loaded and called in the `MyForm.twin`:
```
If This.CreateFactoryDelegate = 0 Then
	Set This.Manager = New UniversalCaller(This.HostFrxLibraryHandle, STRINGPARAMS_ENUM.STR_NONE, CALLINGCONVENTION_ENUM.CC_CDECL)
	Result = This.Manager.CallDllFunction( _
		AssemblyDelegate, _
		CALLRETURNTUYPE_ENUM.CR_LONGLONG, _
		StrPtr(AssemblyPath), _ 
		StrPtr(TypeName), _ 
		StrPtr(MethodName), _
		StrPtr(DelegateName), _ 
		CLngPtr(0), _ 
		This.CreateFactoryDelegate)
	...
End If

...

Dim Factory As IWpfHostFactory
Set Factory = This.Manager.CallDllFunction( _
	This.CreateFactoryDelegate, _
	CALLRETURNTUYPE_ENUM.CR_OBJECT)
```

Note furthermore that twinBASIC and .NET project has their own definitions of the interfacess. Compare .NET host's `IWpfHost` with twinBASIC's `IWpfHost` for an example. This allows both projects to pass around COM objects without having to register or use regfree approach which would require a separate reference. In a production system, it's probably better to have a separate type library that both projects can reference to avoid having to define the same thing twice in respective language. Nonetheless, this demonstates how a COM object can be passed around without having a type library available. 

# Event Handling 

The demonstration now includes examples of raising events from either side. 

## From .NET Core

From .NET Core, we can raise the MouseDown event from `WpfHost` by defining a source interface `IWpfHostEvents`:
```
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
```

and then attach it to the `WpfHost`:
```
[
	...
	ComSourceInterfaces(typeof(IWpfHostEvents))
]
public class WpfHost : HwndSource, IWpfHost
{
	...

	internal WpfHost(HwndSourceParameters parameters) : base(parameters)
	{
		...
		_simple.MouseDown += Host_MouseDown;
	}

	private delegate void MouseDownDelegate();
	private event MouseDownDelegate? MouseDown;
	private void Host_MouseDown(object sender, MouseButtonEventArgs e)
	{
		MouseDown?.Invoke();
	}

	...
}
```

With the corresponding handler in twinBASIC's `MyForm.twin`:
```
Private WithEvents Host As WpfHost

...

Private Sub Host_MouseDown() Handles Host.MouseDown
	MsgBox "Mouse down"
End Sub
```

Using the `WpfHost` as a coclass:
```
...

[ InterfaceId ("7F18FA39-1108-4E7F-A8C1-90BB30E96E71") ]
Interface IWpfHostEvents
    [ DispId (1) ]
    Sub MouseDown()
End Interface

[ CoClassId ("19B1E709-EBFD-4561-859D-CD38A8D575AF") ]
CoClass WpfHost
    [ Default ] Interface IWpfHost
    [ Default, Source ] Interface IWpfHostEvents
End CoClass
```

## From twinBASIC side

twinBASIC does not yet support defining a source interface and assigning it to an implementation. However, we still can approximate this by defining an `EventInterface` and defning the `Event`s the usual way:
```
[ EventInterfaceId ("C1B1A5B2-979E-409C-B573-7943971A4CCA") ]
Public Class MyForm

...

    Public Event Resize(ByVal Width As Long, ByVal Height As Long)
    Private Sub Form_Resize() Handles Form.Resize
        RaiseEvent Resize(Me.PixelsWidth, Me.PixelsHeight)
    End Sub
    
    Public Event Unload(ByRef Cancel As Boolean)
    Private Sub Form_Unload(Cancel As Integer) Handles Form.Unload
        Dim _Cancel As Boolean = Cancel
        RaiseEvent Unload(_Cancel)
        Cancel = _Cancel
    End Sub
End Class
```

Then on .NET side, we use `ComEventsHelper` to listen and respond to the events from twinBASIC. However, we need to pass back the twinBASIC object. We do this via the `IWpfHost::Advise`:
```
[ InterfaceId ("19B1E709-EBFD-4561-859D-CD38A8D575AF") ]
Interface IWpfHost
    Sub Advise(Client As Object)
    ...
End Interface
```

The `IWpfHost` interface is also defined in the .NET project. We respond to the `Advise` method and call `ComEventsHelper.Combine()`:

```
public class WpfHost : HwndSource, IWpfHost
{
    ...
	
	private void Host_Resize(int width, int height)
	{
		MessageBox.Show($"New size: {width}, {height}");
	}

	private void Host_Unload(ref bool cancel)
	{
		cancel = MessageBox.Show("Close?", default, MessageBoxButton.YesNo) != MessageBoxResult.Yes;
	}

	public void Advise([MarshalAs(UnmanagedType.IUnknown)] ref object client)
	{
		ComEventsHelper.Combine(client, new Guid("C1B1A5B2-979E-409C-B573-7943971A4CCA"), 1, Host_Resize);
		ComEventsHelper.Combine(client, new Guid("C1B1A5B2-979E-409C-B573-7943971A4CCA"), 2, Host_Unload);
	}
}
```

# References

[.NET Core Hosting](https://learn.microsoft.com/en-us/dotnet/core/tutorials/netcore-hosting)
