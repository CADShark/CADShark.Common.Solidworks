using CADBooster.SolidDna;
using CADShark.Common.Analytics;
using CADShark.Common.Logging;
using CADShark.Common.SolidWorks.Errors;

namespace CADShark.Common.SolidWorks
{
    public static class SoliDOperations
    {
        static SoliDOperations()
        {
            if (AddInIntegration.ConnectToActiveSolidWorksForStandAlone()) return;

            var message = "Failed to connect to SolidWorks";
            CadLogger.Error(message);
            throw new ConnectionException(message);
        }

        private static void SwVersion() => Telemetry.LogVersions(Version(), "");

        public static int Version()
        {
            var sw = SolidWorksEnvironment.Application;
            var solidWorksVersion = sw.SolidWorksVersion;
            var version = solidWorksVersion.Version;
            var servicePackMajor = solidWorksVersion.ServicePackMajor;
            var servicePackMinor = solidWorksVersion.ServicePackMinor;
            return version;
        }

        public static void SetVisibilityDocument(bool visible, ComponentTypes swDocType)
        {
            SolidWorksEnvironment.Application.UnsafeObject.DocumentVisible(visible, (int)swDocType);
        }

        public static string GetProperties(string path, string propName)
        {
            SetVisibilityDocument(false, ComponentTypes.Part);
            SolidWorksEnvironment.Application.OpenFile(path);
            propName = SolidWorksEnvironment.Application.ActiveModel.GetCustomProperty(propName, null, true);

            SetVisibilityDocument(true, ComponentTypes.Part);
            return propName;
        }
    }
}