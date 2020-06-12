using SharePointPnP.PowerShell.CmdletHelpAttributes;
using System.Management.Automation;
using System.Reflection;

namespace SharePointPnP.PowerShell.Commands.Base
{
    [Cmdlet(VerbsCommon.Get, "PnPVersion")]
    [CmdletHelp("Returns the PnP PowerShell version currently loaded",
        Category = CmdletHelpCategory.Base,
        OutputType = typeof(GetVersionResult))]
    [CmdletExample(
        Code = "PS:> Get-PnPVersion",
        Remarks = "Returns the PnP PowerShell version currently loaded",
        SortOrder = 1)]
    public class GetVersion : BasePSCmdlet
    {
        protected override void ExecuteCmdlet()
        {
            var result = new GetVersionResult();

            var assembly = Assembly.GetExecutingAssembly();
            result.Version = ((AssemblyFileVersionAttribute)assembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version;

#if ONPREMISES
#if SP2013
            result.Platform = "SP2013";
#elif SP2016
            result.Platform = "SP2016";
#elif SP2019
            result.Platform = "SP2019";
#else
            // Unknown version (foolproof for next ONPREMISES versions)
            result.Platform = "ONPREMISES";
#endif
#else
            result.Platform = "SPO";
#endif

            WriteObject(result);
        }
    }

    class GetVersionResult
    {
        public string Version { get; set; }

        public string Platform { get; set; }
    }
}
