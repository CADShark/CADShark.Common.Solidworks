using CADBooster.SolidDna;
using SolidWorks.Interop.swconst;

namespace CADShark.Common.Solidworks
{
    internal class SetMaterial
    {
        private const string Database = @"C:\Program Files\CAD Shark\";
        private const string MaterialProp = "Материал";

        public static void Apply(Materials mi)
        {

            var swApp = AddInIntegration.SolidWorks;

            if (!swApp.ActiveModel.IsPart) return;

            swApp.SetUserPreferencesString(swUserPreferenceStringValue_e.swFileLocationsMaterialDatabases, Database);
            var configName = swApp.ActiveModel.ActiveConfiguration.Name;

            var material = new Material
            {
                Name = mi.SWProperty
            };

            swApp.ActiveModel.SetMaterial(material, configName);
            swApp.ActiveModel.SetCustomProperty(MaterialProp, mi.SWProperty, configName);
        }
    }
}
