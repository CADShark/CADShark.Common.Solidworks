using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
//using CADShark.Common.Logging;
using CADShark.Common.SolidWorks.Documents;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace CADShark.Common.SolidWorks;

public class AssemblyDocument(ISldWorks swApp) : IAssemblyDocument
{
    private ModelDoc2 _swModel;
    private AssemblyDoc _swAssy;
    private int _errors;
    private int _warnings;

    //private static readonly CadLogger Logger = CadLogger.GetLogger<AssemblyDocument>();

    public ModelDoc2 ActivateDoc(string filePath)
    {
        _swModel = (ModelDoc2)swApp.ActivateDoc3(filePath, false, (int)swRebuildOnActivation_e.swUserDecision,
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

        _swModel = swApp.OpenDoc6(filePath, (int)fileType, (int)options, configuration, ref _errors,
            ref _warnings);

        if (_swModel != null)
        {
            //Logger.Info($"Open document: {filePath}");
            //if (_warnings != 0) Logger.Warning($"Warning to open document. Warning code: {_warnings}");

            return _swModel;
        }

        //Logger.Error($"Error to open document {filePath}. Error code: {_errors}");
        return null;
    }

    public Dictionary<string, Component2> GetDistinctPartComponents(ref object[] vComponents)
    {
        var swModel = (ModelDoc2)swApp.ActiveDoc;

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

            if (!groupedComponents.ContainsKey(pathName)) groupedComponents[pathName] = component;
        }

        return groupedComponents;
    }

    public string[] GetDistinctComponents()
    {
        var swModel = (ModelDoc2)swApp.ActiveDoc;
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
            var path = component.GetPathName();
            if (CheckExistDrawingFile(path, ref drawPath))
                listPath.Add(drawPath);
        }

        var array = listPath.GroupBy(x => x.ToString()).Select(x => x.Key).ToArray();

        return array;
    }

    public List<(string Path, string Config, int Count)> GetDistinctComponentsDxf()
    {
        var swModel = (ModelDoc2)swApp.ActiveDoc;
        if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return [];

        _swAssy = (AssemblyDoc)swModel;
        _swAssy.ResolveAllLightWeightComponents(true);

        var vComponents = (object[])_swAssy.GetComponents(false);
        //if (vComponents == null || vComponents.Length == 0)
        //{
        //    Logger.Info("Компоненты в сборке отсутствуют.");
        //    return [];
        //}

        var allComponents = new List<(string Path, string Config)>();

        foreach (Component2 swComp in vComponents)
        {
            if (swComp.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed) continue;

            var path = swComp.GetPathName();
            var config = swComp.ReferencedConfiguration;

            allComponents.Add((path, config));
        }

        var grouped = allComponents
            .GroupBy(c => new { c.Path, c.Config })
            .Select(g => (g.Key.Path, g.Key.Config, g.Count()))
            .OrderBy(g => g.Path)
            .ToList();
        return grouped;
    }


    public bool CheckExistDrawingFile(string path, ref string drawPath)
    {
        drawPath = Path.ChangeExtension(path.ToUpper(), "SLDDRW");
        return File.Exists(drawPath);
    }

    public string[] GetDerivedConfig(ModelDoc2 swModel)
    {
        var configList = new List<string>();

        var configNames = (string[])swModel.GetConfigurationNames();

        for (var index = 0; index < configNames.Length; index++)
        {
            var configName = configNames[index];
            var swConfig = (Configuration)swModel.GetConfigurationByName(configName);
            if (swConfig == null)
            {
                //Logger.Error($"Error to get configuration by name: {configName}");
            }
            else
            {
                if (swConfig.IsDerived()) continue;
                configList.Add(swConfig.Name);
                //Logger.Trace($"Config is not Derived: {swConfig.Name}");
            }
        }

        return configList.ToArray();
    }

    public void SuppressUpdates(bool enable, ModelDoc2 swModel)
    {
        var swView = (ModelView)swModel.ActiveView;

        var enableGraphicsUpdate = swView.EnableGraphicsUpdate = enable;
        //Logger.Debug($"EnableFeatureTree {enableGraphicsUpdate}");

        var enableFeatureTree = swModel.FeatureManager.EnableFeatureTree = enable;
        //Logger.Debug($"EnableFeatureTree {enableFeatureTree}");

        var enableFeatureTreeWindow = swModel.FeatureManager.EnableFeatureTreeWindow = enable;
        //Logger.Debug($"EnableFeatureTreeWindow {enableFeatureTreeWindow}");
    }
}