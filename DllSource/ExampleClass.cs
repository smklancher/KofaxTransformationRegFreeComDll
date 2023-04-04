using System.Collections;
using System.Runtime.InteropServices;

namespace RegFreeTransformationDll
{
    // ComVisible is already true by default, so you really only need to specify the attribute
    // if setting to false to hide a class from COM
    //[ComVisible(true)]

    // ProgId isn't explictly required as it will use Namespace.Classname by default.
    //[ProgId("RegFreeTransformationDll.ExampleClass")]

    // For a class to get created by the ManifestActivator in Transformation script, the GUID
    // must be defined explictly so it can be specified in the embedded component/assembly manifest.
    [Guid("efb13163-f3f9-4e56-8f90-2af836401c60")]

    // ClassInterfaceType: COM best practice to set this to None and explictly implement an interface
    //  (https://stackoverflow.com/a/1436814/221018), though this may not matter as much when used in
    //  a registration-free manner.
    [ClassInterface(ClassInterfaceType.None)]

    // ComDefaultInterface: in case a class implements multiple interfaces (which can also happen implicitly if
    //  ClassInterfaceType isn't set to None), this specifies which is the default interface to COM
    [ComDefaultInterface(typeof(IExampleClass))]
    public class ExampleClass : IExampleClass
    {
        public string HelloWorld(string msg)
        {
            return $"Hello from registration-free COM DLL (msg: {msg})";
        }

        public ArrayList ReturnComplexType()
        {
            ArrayList ret = new ArrayList
            {
                new ComplexType() { Confidence = 0, Text = "Low confidence alternative" },
                new ComplexType() { Confidence = .5, Text = "Med confidence alternative" },
                new ComplexType() { Confidence = 1, Text = "High confidence alternative" }
            };

            return ret;
        }
    }
}