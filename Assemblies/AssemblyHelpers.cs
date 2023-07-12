using System;
using System.Collections.Generic;
using System.Linq;
using CADBooster.SolidDna;
using CADShark.Common.Logging;
using CADShark.Common.SolidWorks.Errors;
using SolidWorks.Interop.sldworks;

namespace CADShark.Common.SolidWorks.Assemblies
{
    internal class AssemblyHelpers
    {
        internal AssemblyHelpers()
        {
            if (AddInIntegration.ConnectToActiveSolidWorksForStandAlone()) return;

            var message = "Failed to connect to SolidWorks";
            CadLogger.Error(message);
            throw new ConnectionException(message);
        }

        internal static List<CADBooster.SolidDna.Component> GetComponents(AssemblyDoc assembly, bool includeSubComponents)
        {
            if (assembly == null)
                throw new ArgumentNullException(nameof(assembly));
            var ToplevelOnly = !includeSubComponents;
            var components = assembly.GetComponents(ToplevelOnly);
            return components != null ? ((object[])components).Cast<Component2>().Select(x => new CADBooster.SolidDna.Component(x)).ToList() : new List<CADBooster.SolidDna.Component>();
        }
    }
}
