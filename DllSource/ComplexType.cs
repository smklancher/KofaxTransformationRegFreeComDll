using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RegFreeTransformationDll
{
    // Unlike classes that get returned directly by the ManifestActivator in Transformation script,
    // other classes and interfaces don't need a GUID explictly defined.
    // Usually this isn't needed since you're supposed to change the GUID any time the interface (implicit or not) changes
    // and .NET already does this automatically when the attribute isn't explicitly specified.
    //
    // See various references:
    //  https://stackoverflow.com/a/8719479/221018
    //  "You only use the [Guid] attribute if you want to override the auto-generated guid.
    //  Which is only useful in COM interop scenarios to get a declaration to match an
    //  existing COM interface or coclass."

    //  https://stackoverflow.com/a/26614463/221018
    //  "Explicitly specifying like you do is quite risky, there is a rock-hard rule in COM that you must change
    //  the guid when you change the declaration. Not doing so causes very nasty DLL Hell problems."
    //
    //  https://stackoverflow.com/a/13705404/221018
    //  "the better approach is to not use the [Guid] attribute but leave it up to the CLR to auto-generate one.
    //  It has a very good algorithm that keeps the guid stable as long as the interface or class members don't change."
    //[Guid("ca075c9d-387b-4a46-98d5-0eda0bb09621")]

    // ClassInterfaceType: Though None + an explicit interface is a COM best practice,
    //  setting to AutoDispatch allows late bound access without explicitly implementing an interface
    //  (https://stackoverflow.com/a/1436814/221018)
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class ComplexType
    {
        public double Confidence { get; set; }

        public string Text { get; set; } = string.Empty;
    }
}