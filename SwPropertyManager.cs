using System.Diagnostics;
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
    public static SwPropertyWriteResult Succeeded(int nativeResultCode) => new(SwPropertyWriteStatus.Success, nativeResultCode);
    public static SwPropertyWriteResult DocumentNotOpen() => new(SwPropertyWriteStatus.DocumentNotOpen, 0);
    public static SwPropertyWriteResult AddFailed(int nativeResultCode) => new(SwPropertyWriteStatus.AddFailed, nativeResultCode);
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
        Debug.WriteLine($"Get property: {propName}");
        Debug.WriteLine($"Get property status {res}");
        return res == (int)swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue ? resolvedVal : "";
    }

    public static SwPropertyWriteResult TrySetProperty(SldWorks swApp, string modelPath, string propName, string newValue, string configName = "")
    {
        var model = (ModelDoc2)swApp.GetOpenDocumentByName(modelPath);
        return model == null ? SwPropertyWriteResult.DocumentNotOpen() : TrySetProperty(model, propName, newValue, configName);
    }

    public static SwPropertyWriteResult TrySetProperty(ModelDoc2 model, string propName, string newValue, string configName = "")
    {
        if (model == null) return SwPropertyWriteResult.DocumentNotOpen();
        var propMgr = model.Extension.CustomPropertyManager[configName];
        var result = TrySetProperty(propMgr, propName, newValue, swCustomInfoType_e.swCustomInfoText);
        model.SetSaveFlag();
        return result;
    }

    public static SwPropertyWriteResult TrySetProperty(CustomPropertyManager propMgr, string propName, string newValue, swCustomInfoType_e infoType)
    {
        var res = propMgr.Add3(propName, (int)infoType, newValue, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
        return res == (int)swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged
            ? SwPropertyWriteResult.Succeeded(res)
            : SwPropertyWriteResult.AddFailed(res);
    }
}
