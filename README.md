# TwinBasicNetCoreHosting

This is a demonstration project showing how twinBASIC can host a .NET 7. Though .NET does support COM activation where the hosting is automatically handled by the internal implementation of the generated `.comhost.dll` file, this simply show how the hosting can be done manually and controlled by twinBASIC without any COM registration required. Furthermore, the project loads a WPF Page UI control into the window provided by the twinBASIC, demonstrating how to host WPF directly by twinBASIC.

# Setup

This requires Visual Studio 2022, .NET 7 SDK, and twinBASIC (beta 235). The output should be placed in a folder named `Build` at the root of directory which is not included in the git repository. The Visual Studio project will attempt to copy the output of its build into the folder which is the same as the output used by twinBASIC's project. 

Furthermore, the `nethost.dll` should be copied from the `Resources\nethost.dll` into the `Build`. After a successful build of both projects, the `Build` folder should contain the minimum of the following files:

* `nethost.dll`
* `WpfClient_win64.exe`
* `WpfServer.deps.json`
* `WpfServer.dll`
* `WpfServer.runtimeconfig.json`

The output is tested for 64-bit only and most likely will not work with 32-bit. Because the .NET resources are used in-process by an unmanaged host, the `Any CPU` is not strictly proper. Therefore, the `x64` is the default built in both Visual Studio and twinBASIC project. 

# References

[.NET Core Hosting](https://learn.microsoft.com/en-us/dotnet/core/tutorials/netcore-hosting)
