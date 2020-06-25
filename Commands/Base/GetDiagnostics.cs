using SharePointPnP.PowerShell.CmdletHelpAttributes;
using SharePointPnP.PowerShell.Commands.Model;
using System;
using System.Collections;
using System.Management.Automation;
using System.Reflection;

namespace SharePointPnP.PowerShell.Commands.Base
{
    [Cmdlet(VerbsCommon.Get, "PnPDiagnostics")]
    [CmdletHelp("Returns diagnostic information",
        Category = CmdletHelpCategory.Base,
        OutputType = typeof(GetDiagnosticsResult))]
    [CmdletExample(
        Code = "PS:> Get-PnPDiagnostics",
        Remarks = "Returns diagnostic information",
        SortOrder = 1)]
    public class GetDiagnostics : BasePSCmdlet
    {
        protected override void ExecuteCmdlet()
        {
            var result = new GetDiagnosticsResult();

            FillVersion(result);
            FillPlatform(result);
            result.ModuleName = NotImplemented<string>();
            FillOperatingSystem(result);
            FillConnectionMethod(result);
            FillCurrentSite(result);
            result.AccessTokenExpirationTime = NotImplemented<DateTime?>();
            result.ModulePath = NotImplemented<string>();
            FillNewerVersionAvailable(result);
            FillLastException(result);

            WriteObject(result, true);
        }

        void FillVersion(GetDiagnosticsResult result)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = ((AssemblyFileVersionAttribute)assembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version;
            result.Version = version;
        }

        void FillPlatform(GetDiagnosticsResult result)
        {
            string platform;
#if ONPREMISES
#if SP2013
            platform = "SP2013";
#elif SP2016
            platform = "SP2016";
#elif SP2019
            platform = "SP2019";
#else
            // Unknown version (foolproof for next ONPREMISES versions)
            platform = "ONPREMISES";
#endif
#else
            platform = "SPO";
#endif

            result.Platform = platform;
        }

        void FillOperatingSystem(GetDiagnosticsResult result)
        {
            result.OperatingSystem = Environment.OSVersion.VersionString;
        }

        void FillConnectionMethod(GetDiagnosticsResult result)
        {
            result.ConnectionMethod = PnPConnection.CurrentConnection?.ConnectionMethod;
        }

        void FillCurrentSite(GetDiagnosticsResult result)
        {
            result.CurrentSite = PnPConnection.CurrentConnection?.Url;
        }

        void FillNewerVersionAvailable(GetDiagnosticsResult result)
        {
            result.NewerVersionAvailable = !string.IsNullOrEmpty(PnPConnectionHelper.GetLatestVersion());
        }

        void FillLastException(GetDiagnosticsResult result)
        {
            // Most of this code has been copied from GetException cmdlet
            var exceptions = (ArrayList)this.SessionState.PSVariable.Get("error").Value;
            if (exceptions.Count > 0)
            {
                var exception = (ErrorRecord)exceptions[0];
                var correlationId = string.Empty;
                if (exception.Exception.Data.Contains("CorrelationId"))
                {
                    correlationId = exception.Exception.Data["CorrelationId"].ToString();
                }
                var timeStampUtc = DateTime.MinValue;
                if (exception.Exception.Data.Contains("TimeStampUtc"))
                {
                    timeStampUtc = (DateTime)exception.Exception.Data["TimeStampUtc"];
                }
                var pnpException = new PnPException() { CorrelationId = correlationId, TimeStampUtc = timeStampUtc, Message = exception.Exception.Message, Stacktrace = exception.Exception.StackTrace, ScriptLineNumber = exception.InvocationInfo.ScriptLineNumber, InvocationInfo = exception.InvocationInfo, Exception = exception.Exception };

                result.LastCorrelationId = pnpException.CorrelationId;
                result.LastExceptionTimeStampUtc = pnpException.TimeStampUtc;
                result.LastExceptionMessage = pnpException.Message;
                result.LastExceptionStacktrace = pnpException.Stacktrace;
                result.LastExceptionScriptLineNumber = pnpException.ScriptLineNumber;
            }
        }

        T NotImplemented<T>()
        {
            return default;
        }
    }

    class GetDiagnosticsResult
    {
        public string Version { get; set; }
        public string Platform { get; set; }
        public string ModuleName { get; set; }
        public string OperatingSystem { get; set; }
        // Was this actually ConnectionMethod ?
        //public string AuthenticationMode { get; set; }
        public ConnectionMethod? ConnectionMethod { get; set; }
        public string CurrentSite { get; set; }
        public DateTime? AccessTokenExpirationTime { get; set; }
        public string ModulePath { get; set; }
        public bool NewerVersionAvailable { get; set; }
        public string LastCorrelationId { get; set; }
        public DateTime? LastExceptionTimeStampUtc { get; set; }
        public string LastExceptionMessage { get; set; }
        public string LastExceptionStacktrace { get; set; }
        public int? LastExceptionScriptLineNumber { get; set; }
    }
}
