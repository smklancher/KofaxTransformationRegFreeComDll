# Kofax Transformation Registration-Free COM DLL Example

There is a help topic that explains that you can use a registration-free COM dll in a Transformation project:

[Dynamically load DLLs in script](https://docshield.kofax.com/KTA/en_US/7.11.0-h49vd5omev/help/ScriptDocumentation/ScriptDocumentation/c_DynamicallyLoadDLLsinScript.html)

However it is not easy to go from that bit of documenation to a working configuration.  It is easier to copy a working example, which is provided here.

The dll just has an example class to return a bit of text, and the Transformation project uses the dll to populate locators/fields.

## Testing

To test in Transformation Designer, import the package and open RegFreeDllExample shared project.  Testing extraction or the locators shows values returned from the dll.  Note that if you update the dll, Transformation Designer must be closed and reopened for the updated dll to be loaded.

To test at runtime and verify the functionality, import the package, open RegFreeTransformationDllTest_Scan.form in the workspace, import any page and create a job.  After the Transformation activity processes, take the Validation activity to see field results populuated from the dll.

## Details

The [Create a registration-free COM DLL and manifest file](https://docshield.kofax.com/KTA/en_US/7.11.0-h49vd5omev/help/ScriptDocumentation/ScriptDocumentation/c_CreateaRegistrationFreeCOMDLLandManifestFile.html) topic mentions this external application manifest [(example)](https://github.com/smklancher/KofaxTransformationRegFreeComDll/blob/main/TransformationScript/application.manifest), but does not discuss the need for the embedded component/assembly manifest file [(example)](https://github.com/smklancher/KofaxTransformationRegFreeComDll/blob/main/DllSource/app.manifest).

As discussed in [Place the DLL](https://docshield.kofax.com/KTA/en_US/7.11.0-h49vd5omev/help/ScriptDocumentation/ScriptDocumentation/t_PlacetheDLL.html), the dll and external application manifest go under the Custom folder within the project folder (which is in a subfolder of %temp% when open in TD).  

The [Invoke the assembly](https://docshield.kofax.com/KTA/en_US/7.11.0-h49vd5omev/help/ScriptDocumentation/ScriptDocumentation/c_InvoketheAssembly.html) shows code to create an object from the reg free dll, however in the example [Transformation script](https://github.com/smklancher/KofaxTransformationRegFreeComDll/blob/main/TransformationScript/_ProjectClass.vb) here, the similar code is organized into a few different functions to try to make it a bit more understandable and reusable.  Note also the two required references that are not mentioned in the help topic:

* C:\Program Files (x86)\Common Files\Kofax\Components\Kofax.Transformation.Extensions.tlb
* Required reference: C:\Windows\SysWOW64\scrrun.dll

## Troubleshooting

Since this example project is open and available it can be compared to your project to determine whatever differences there are.  The code and manifest files have a lot of comments specifically to help with this.

The only element required in the [csproj file](https://github.com/smklancher/KofaxTransformationRegFreeComDll/blob/main/DllSource/RegFreeTransformationDll.csproj) itself is that it is told to actually use the included manifest file:

```xml
    <ApplicationManifest>app.manifest</ApplicationManifest>
```

One notable element of the example csproj file itself is that it uses [SDK-style](https://learn.microsoft.com/en-us/dotnet/core/project-sdk/overview) format.  This is not required, but it makes the file much more terse and more easy to read and show as an example.  The SDK-style format is not used by default for new .NET Framework projects.  However new .NET Framework projects and can be created this way with a few minor extra steps:

[How to create an SDK-style .NET Framework project in VS?](https://stackoverflow.com/a/63065708/221018)
