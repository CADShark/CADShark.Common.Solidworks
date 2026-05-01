using System.Collections.Generic;
using CADShark.Common.SolidWorks.Documents;
using SolidWorks.Interop.sldworks;

namespace CADShark.Common.SolidWorks;

public interface IAssemblyDocument
{
    /// <summary>
    /// Opens an existing document and returns a pointer to the document object. 
    /// </summary>
    /// <param name="filePath">File path to document</param>
    /// <param name="options">How to open documents</param>
    /// <param name="configuration">Model configuration in which to open this document</param>
    /// <returns>Model document</returns>
    ModelDoc2 OpenFile(string filePath, OpenDocumentOptions options = OpenDocumentOptions.None,
        string configuration = null);

    /// <summary>
    /// Activates a document that has already been loaded. This file becomes the active document, and this method returns a pointer to that document object. 
    /// </summary>
    /// <param name="filePath">Name of document to activate</param>
    /// <returns>Model document</returns>
    ModelDoc2 ActivateDoc(string filePath);

    /// <summary>
    /// Gets all of the components in the active configuration of this assembly. 
    /// </summary>
    /// <param name="vComponents">Array of IComponent2s</param>
    /// <returns></returns>
    Dictionary<string, Component2> GetDistinctPartComponents(ref object[] vComponents);

    string[] GetDistinctComponents();
    bool CheckExistDrawingFile(string path, ref string drawPath);
    string[] GetDerivedConfig(ModelDoc2 swModel);

    void SuppressUpdates(bool enable, ModelDoc2 model);

    public List<(string Path, string Config, int Count)> GetDistinctComponentsDxf();
}