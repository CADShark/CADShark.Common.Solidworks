using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace CADShark.Common.SolidWorks;

public enum SwPropertyWriteStatus
{
    Success = 0,
    DocumentNotOpen = 1,
    AddFailed = 2
}

public sealed class SwPropertyWriteResult
{
    private SwPropertyWriteResult(SwPropertyWriteStatus status, int nativeResultCode)
    {
        Status = status;
        NativeResultCode = nativeResultCode;
    }

    public SwPropertyWriteStatus Status { get; }
    public int NativeResultCode { get; }
    public bool Success => Status == SwPropertyWriteStatus.Success;

    public static SwPropertyWriteResult Succeeded(int nativeResultCode) =>
        new(SwPropertyWriteStatus.Success, nativeResultCode);

    public static SwPropertyWriteResult DocumentNotOpen() =>
        new(SwPropertyWriteStatus.DocumentNotOpen, 0);

    public static SwPropertyWriteResult AddFailed(int nativeResultCode) =>
        new(SwPropertyWriteStatus.AddFailed, nativeResultCode);
}

public class SwPropertyManager
{
    public static string GetProperty(ModelDoc2 model, string configName, string propName)
    {
        var propMgr = model.Extension.CustomPropertyManager[configName];

        var res = propMgr.Get6(propName, false, out _, out var resolvedVal, out _, out _);

        if (res != (int)swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue) return "";
        return resolvedVal;

    }

    public static string GetProperty(SldWorks swApp, string modelPath, string configName, string propName)
    {
        var model = (ModelDoc2)swApp.GetOpenDocumentByName(modelPath);

        var propMgr = model.Extension.CustomPropertyManager[configName];

        var res = propMgr.Get6(propName, false, out _, out var resolvedVal, out _, out _);
        Debug.WriteLine(
            $"Get property: {propName}");

        Debug.WriteLine(
            $"Get property status {res}");
        if (res == (int)swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue)
            return resolvedVal;

        return "";
    }

    public static SwPropertyWriteResult TrySetProperty(
        SldWorks swApp,
        string modelPath,
        string propName,
        string newValue,
        string configName = "")
    {
        var model = (ModelDoc2)swApp.GetOpenDocumentByName(modelPath);
        if (model == null)
            return SwPropertyWriteResult.DocumentNotOpen();

        return TrySetProperty(model, propName, newValue, configName);
    }

    public static SwPropertyWriteResult TrySetProperty(
        ModelDoc2 model,
        string propName,
        string newValue,
        string configName = "")
    {
        if (model == null)
            return SwPropertyWriteResult.DocumentNotOpen();

        var propMgr = model.Extension.CustomPropertyManager[configName];
        var result = TrySetProperty(propMgr, propName, newValue, swCustomInfoType_e.swCustomInfoText);

        // Legacy SetProperty(ModelDoc2, ...) marked the model dirty after every attempted
        // write, including a native Add3 failure. Preserve that observable behavior while
        // exposing a non-UI result API to new consumers.
        model.SetSaveFlag();

        return result;
    }

    public static SwPropertyWriteResult TrySetProperty(
        CustomPropertyManager propMgr,
        string propName,
        string newValue,
        swCustomInfoType_e infoType)
    {
        var res = propMgr.Add3(
            propName,
            (int)infoType,
            newValue,
            (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

        return res == (int)swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged
            ? SwPropertyWriteResult.Succeeded(res)
            : SwPropertyWriteResult.AddFailed(res);
    }

    /// <summary>
    /// Sets a custom property in a SolidWorks document.
    /// </summary>
    /// <param name="swApp"></param>
    /// <param name="modelPath">
    /// The full path to the open SolidWorks document (part, assembly, or drawing).
    /// </param>
    /// <param name="propName">
    /// The name of the custom property to set.
    /// </param>
    /// <param name="newValue">
    /// The new value of the property.
    /// </param>
    /// <param name="configName">
    /// The name of the configuration for which the property is set.
    /// If an empty string is provided, the property is applied to the general document level.
    /// </param>
    /// <remarks>
    /// If the property already exists, it will be overwritten.
    /// The method uses <see cref="ICustomPropertyManager.Add3"/> to apply the change.
    /// </remarks>
    /// <example>
    /// Example:
    /// <code>
    /// SetProperty(@"C:\Models\Part1.SLDPRT", "Material", "Steel", "Default");
    /// </code>
    /// </example>
    public static void SetProperty(SldWorks swApp, string modelPath, string propName, string newValue,
        string configName = "")
    {
        var result = TrySetProperty(swApp, modelPath, propName, newValue, configName);

        if (!result.Success)
            MessageBox.Show($@"Не удалось сохранить свойство '{propName}' = '{newValue}'");
    }

    public static void SetProperty(ModelDoc2 model, string propName, string newValue, string configName = "")
    {
        var result = TrySetProperty(model, propName, newValue, configName);

        if (!result.Success)
            MessageBox.Show($@"Не удалось сохранить свойство '{propName}' = '{newValue}'");
    }

    public static void SetProperty(CustomPropertyManager propMgr, string propName, string newValue,
        swCustomInfoType_e infoType)
    {
        var result = TrySetProperty(propMgr, propName, newValue, infoType);

        if (!result.Success)
            MessageBox.Show($@"Не удалось сохранить свойство '{propName}' = '{newValue}'");
    }
}