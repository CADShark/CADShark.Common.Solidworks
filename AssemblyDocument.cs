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

    //private static readonly CadLogger Logger = CadLogger.GetLogger<AssemblyDocument>();

    public ModelDoc2 ActivateDoc(string filePath, ref int errors)
    {
        _swModel = (ModelDoc2)swApp.ActivateDoc3(
            filePath, false,
            (int)swRebuildOnActivation_e.swUserDecision,
            ref errors);
        return _swModel;
    }
    public ModelDoc2 OpenFile(string filePath, ref int errors, ref int warnings, OpenDocumentOptions options = OpenDocumentOptions.None, string configuration = null)
    {
        var fileType =
            filePath.ToLower().EndsWith(".sldprt") ? DocumentType.Part :
            filePath.ToLower().EndsWith(".sldasm") ? DocumentType.Assembly :
            filePath.ToLower().EndsWith(".slddrw") ? DocumentType.Drawing :
            throw new ArgumentException($"Unknown file type: {filePath}");

        var model = swApp.OpenDoc6(
            filePath, (int)fileType, (int)options, configuration,
            ref errors, ref warnings);

        _swModel = model;

        //if (model != null)
        //    Logger.Info($"Opened document: {filePath}");
        //else
        //    Logger.Error($"Failed to open document: {filePath}. Error code: {_errors}");

        //if (_warnings != 0)
        //    Logger.Warning($"Warning opening document: {filePath}. Warning code: {_warnings}");
        return _swModel; 
        //return new OpenFileResult(model, _errors, _warnings);
    }

    public Dictionary<string, Component2> GetDistinctPartComponents(ref object[] vComponents)
    {
        var swModel = (ModelDoc2)swApp.ActiveDoc;

        if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return null;

        var groupedComponents = new Dictionary<string, Component2>();

        _swAssy = swModel as AssemblyDoc;
        _swAssy.ResolveAllLightWeightComponents(true);

        vComponents = (object[])_swAssy.GetComponents(false);

        foreach (Component2 component in vComponents)
        {
            if (component.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed) continue;

            swModel = (ModelDoc2)component.GetModelDoc2();
            if (swModel?.GetType() != (int)swDocumentTypes_e.swDocPART) continue;

            var pathName = component.GetPathName();
            if (!groupedComponents.ContainsKey(pathName))
                groupedComponents[pathName] = component;
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

        return listPath.GroupBy(x => x).Select(x => x.Key).ToArray();
    }

    public List<(string Path, string Config, int Count)> GetDistinctComponentsDxf()
    {
        var swModel = (ModelDoc2)swApp.ActiveDoc;
        if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return [];

        _swAssy = (AssemblyDoc)swModel;
        _swAssy.ResolveAllLightWeightComponents(true);

        var vComponents = (object[])_swAssy.GetComponents(false);


        var allComponents = new List<(string Path, string Config)>();

        foreach (Component2 swComp in vComponents)
        {
            if (swComp.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed) continue;

            allComponents.Add((swComp.GetPathName(), swComp.ReferencedConfiguration));
        }

        return allComponents
            .GroupBy(c => new { c.Path, c.Config })
            .Select(g => (g.Key.Path, g.Key.Config, g.Count()))
            .OrderBy(g => g.Path)
            .ToList();
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

        foreach (var configName in configNames)
        {
            var swConfig = (Configuration)swModel.GetConfigurationByName(configName);

            if (swConfig == null)
                //Logger.Error($"Error to get configuration by name: {configName}");
                continue;

            if (swConfig.IsDerived()) continue;

            configList.Add(swConfig.Name);
            //Logger.Trace($"Config is not derived: {swConfig.Name}");
        }

        return configList.ToArray();
    }

    public void SuppressUpdates(bool enable, ModelDoc2 swModel)
    {
        var swView = (ModelView)swModel.ActiveView;

        swView.EnableGraphicsUpdate = enable;
        //Logger.Debug($"EnableGraphicsUpdate: {enable}");

        swModel.FeatureManager.EnableFeatureTree = enable;
        //Logger.Debug($"EnableFeatureTree: {enable}");

        swModel.FeatureManager.EnableFeatureTreeWindow = enable;
        //Logger.Debug($"EnableFeatureTreeWindow: {enable}");
    }


}