using CADShark.Common.Logging;
using CADShark.Common.SolidWorks.Documents;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CADShark.Common.SolidWorks
{
    public class AssemblyDocument : IAssemblyDocument
    {
        private readonly SldWorks _swApp;
        private ModelDoc2 _swModel;
        private AssemblyDoc _swAssy;
        private int _errors;
        private int _warnings;

        private static readonly CadLogger Logger = CadLogger.GetLogger(className: nameof(AssemblyDocument));

        public AssemblyDocument()
        {
        }

        public AssemblyDocument(SldWorks swApp)
        {
            _swApp = swApp;
        }

        public ModelDoc2 ActivateDoc(string filePath)
        {
            _swModel = (ModelDoc2)_swApp.ActivateDoc3(filePath, false, (int)swRebuildOnActivation_e.swUserDecision,
                ref _errors);
            return _swModel;
        }

        public ModelDoc2 OpenFile(string filePath, OpenDocumentOptions options = OpenDocumentOptions.None,
            string configuration = null)
        {
            // Get file type
            var fileType =
                filePath.ToLower().EndsWith(".sldprt") ? DocumentType.Part :
                filePath.ToLower().EndsWith(".sldasm") ? DocumentType.Assembly :
                filePath.ToLower().EndsWith(".slddrw") ? DocumentType.Drawing :
                throw new ArgumentException("Unknown file type");

            _swModel = _swApp.OpenDoc6(filePath, (int)fileType, (int)options, configuration, ref _errors,
                ref _warnings);

            if (_swModel != null)
            {
                Logger.Info($"Open document: {filePath}");
                if (_warnings != 0)
                {
                    Logger.Warning($"Warning to open document. Warning code: {_warnings}");
                }

                return _swModel;
            }

            Logger.Error($"Error to open document {filePath}. Error code: {_errors}");
            return null;
        }

        public Dictionary<string, Component2> GetDistinctPartComponents(ref object[] vComponents)
        {
            var swModel = (ModelDoc2)_swApp.ActiveDoc;

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return null;

            var groupedComponents = new Dictionary<string, Component2>();

            _swAssy = (AssemblyDoc)swModel;
            _swAssy.ResolveAllLightWeightComponents(true);

            vComponents = (object[])_swAssy.GetComponents(false);

            foreach (Component2 component in vComponents)
            {
                if (component.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed) continue;
                swModel = (ModelDoc2)component.GetModelDoc2();
                if (swModel?.GetType() != (int)swDocumentTypes_e.swDocPART) continue;

                var pathName = component.GetPathName();

                if (!groupedComponents.ContainsKey(pathName))
                {
                    groupedComponents[pathName] = component;
                }
            }

            return groupedComponents;

            //foreach (Component2 component in vComponents)
            //{
            //    var path = component.GetPathName();

            //    if (component.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed) continue;

            //    swModel = (ModelDoc2)component.GetModelDoc2();

            //    if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART) continue;
            //    listPath.Add(path);
            //}

            //var array = listPath.GroupBy(x => x.ToString()).Select(x => x.Key).ToArray();

            //return array;
        }

        public string[] GetDistinctComponents()
        {
            var swModel = (ModelDoc2)_swApp.ActiveDoc;
            var listPath = new List<string>();
            string drawPath = null;

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return null;
            if (CheckExistDrawingFile(swModel.GetPathName(), ref drawPath))
                listPath.Add(drawPath);

            _swAssy = (AssemblyDoc)swModel;
            _swAssy.ResolveAllLightWeightComponents(true);

            var vComponents = (object[])_swAssy.GetComponents(false);

            foreach (Component2 component in vComponents)
            {
                //Logger.Info($"GetPathName {component.GetPathName()}");
                var path = component.GetPathName();
                if (CheckExistDrawingFile(path, ref drawPath))
                    listPath.Add(drawPath);
            }

            var array = listPath.GroupBy(x => x.ToString()).Select(x => x.Key).ToArray();

            return array;
        }

        public bool CheckExistDrawingFile(string path, ref string drawPath)
        {
            drawPath = Regex.Replace(path.ToUpper(), "SLDASM|SLDPRT", "SLDDRW");
            return File.Exists(drawPath);
        }

        public string[] GetDerivedConfig()
        {
            var swModel = (ModelDoc2)_swApp.ActiveDoc;
            var configList = new List<string>();

            var configNames = (string[])swModel.GetConfigurationNames();

            for (var index = 0; index < configNames.Length; index++)
            {
                var configName = configNames[index];
                var swConfig = (Configuration)swModel.GetConfigurationByName(configName);
                if (swConfig == null)
                {
                    Logger.Error($"Error to get configuration by name: {configName}");
                }
                else
                {
                    if (swConfig.IsDerived()) continue;
                    configList.Add(swConfig.Name);
                    Logger.Trace($"Config is not Derived: {swConfig.Name}");
                }
            }

            return configList.ToArray();
        }

        public void SuppressUpdates(bool enable)
        {
            _swModel = (ModelDoc2)_swApp.ActiveDoc;

            var swView = (ModelView)_swModel.ActiveView;
            swView.EnableGraphicsUpdate = enable;
            _swModel.FeatureManager.EnableFeatureTree = enable;
            _swModel.FeatureManager.EnableFeatureTreeWindow = enable;
        }
    }
}