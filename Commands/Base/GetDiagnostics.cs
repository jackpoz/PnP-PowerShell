using SharePointPnP.PowerShell.CmdletHelpAttributes;
using SharePointPnP.PowerShell.Commands.Enums;
using SharePointPnP.PowerShell.Commands.Model;
using System;
using System.Collections;
using System.IO;
using System.Management.Automation;
using System.Reflection;

namespace SharePointPnP.PowerShell.Commands.Base
{
    [Cmdlet(VerbsCommon.Get, "PnPDiagnostics")]
    [CmdletHelp("Returns diagnostic information",
        Category = CmdletHelpCategory.Base)]
    [CmdletExample(
        Code = "PS:> Get-PnPDiagnostics",
        Remarks = "Returns diagnostic information",
        SortOrder = 1)]
    public class GetDiagnostics : BasePSCmdlet
    {
        protected override void ExecuteCmdlet()
        {
            var result = new PSObject();

            FillVersion(result);
            FillPlatform(result);
            FillModuleInfo(result);
            FillOperatingSystem(result);
            FillConnectionMethod(result);
            FillCurrentSite(result);
            FillAccessTokenExpirationTimes(result);
            FillNewerVersionAvailable(result);
            FillLastException(result);

            WriteObject(result, true);
        }

        void FillVersion(PSObject result)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = ((AssemblyFileVersionAttribute)assembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version;
            AddProperty(result, "Version", version);
        }

        void FillPlatform(PSObject result)
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

            AddProperty(result, "Platform", platform);
        }

        void FillModuleInfo(PSObject result)
        {
            var modulePath = AssemblyDirectoryFromLocation;
            DirectoryInfo dirInfo = new DirectoryInfo(modulePath);
            if (!dirInfo.Exists)
            {
                modulePath = AssemblyDirectoryFromCodeBase;
                dirInfo = new DirectoryInfo(modulePath);
                if (!dirInfo.Exists)
                {
                    modulePath = "Could not retrieve the information";
                    dirInfo = null;
                }
            }

            string moduleName = dirInfo?.Name ?? "Could not retrieve the information";

            AddProperty(result, "ModulePath", modulePath);
            AddProperty(result, "ModuleName", moduleName);
        }

        void FillOperatingSystem(PSObject result)
        {
            AddProperty(result, "OperatingSystem", Environment.OSVersion.VersionString);
        }

        void FillConnectionMethod(PSObject result)
        {
            AddProperty(result, "ConnectionMethod", PnPConnection.CurrentConnection?.ConnectionMethod);
        }

        void FillCurrentSite(PSObject result)
        {
            AddProperty(result, "CurrentSite", PnPConnection.CurrentConnection?.Url);
        }

        void FillAccessTokenExpirationTimes(PSObject result)
        {
            var connection = PnPConnection.CurrentConnection;
            if (connection == null)
                return;

            foreach(TokenAudience audience in Enum.GetValues(typeof(TokenAudience)))
            {
                var expirationTime = connection.TryGetTokenExpirationTime(audience);
                if (expirationTime != null)
                    AddProperty(result, $"AccessTokenExpirationTime{audience}", expirationTime);
            }
        }

        void FillNewerVersionAvailable(PSObject result)
        {
            AddProperty(result, "NewerVersionAvailable", !string.IsNullOrEmpty(PnPConnectionHelper.GetLatestVersion()));
        }

        void FillLastException(PSObject result)
        {
            // Most of this code has been copied from GetException cmdlet
            PnPException pnpException = null;
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
                pnpException = new PnPException() { CorrelationId = correlationId, TimeStampUtc = timeStampUtc, Message = exception.Exception.Message, Stacktrace = exception.Exception.StackTrace, ScriptLineNumber = exception.InvocationInfo.ScriptLineNumber, InvocationInfo = exception.InvocationInfo, Exception = exception.Exception };

            }

            AddProperty(result, "LastCorrelationId", pnpException?.CorrelationId);
            AddProperty(result, "LastExceptionTimeStampUtc", pnpException?.TimeStampUtc);
            AddProperty(result, "LastExceptionMessage", pnpException?.Message);
            AddProperty(result, "LastExceptionStacktrace", pnpException?.Stacktrace);
            AddProperty(result, "LastExceptionScriptLineNumber", pnpException?.ScriptLineNumber);
        }

        void AddProperty(PSObject pso, string name, object value)
        {
            pso.Properties.Add(new PSVariableProperty(new PSVariable(name, value)));
        }
    }
}
