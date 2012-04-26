using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.EnterpriseServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Component Services")]
[assembly: AssemblyDescription("Media Center Network Controller Server for Microsoft Windows Vista")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("NrgUp")]
[assembly: AssemblyProduct("Media Center Network Controller")]
[assembly: AssemblyCopyright("Copyright ©2007 Jonathan Bradshaw")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// General information about the serviced component
[assembly: ApplicationName("Media Center Network Controller")]
[assembly: ApplicationID("38A99532-3787-11DC-8AA9-656856D89593")]
[assembly: ApplicationActivation(ActivationOption.Server)]
[assembly: ApplicationAccessControl(false,
                    AccessChecksLevel = AccessChecksLevelOption.Application,
                    Authentication = AuthenticationOption.Packet,
                    ImpersonationLevel = ImpersonationLevelOption.Identify)]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("6b442151-06f2-433c-a392-c3a39b9ea09c")]
