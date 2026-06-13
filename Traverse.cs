using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CADShark.Common.SolidWorks;

public class Traverse(ModelDoc2 swModelDoc)
{
    private int _nextId = 1;

    private int GetNextId()
    {
        return _nextId++;
    }

    public List<SwNode> BuildFlatList()
    {
        var result = new List<SwNode>();

        if (swModelDoc == null)
            return result;

        var root = CreateNode(swModelDoc, null);
        result.Add(root);

        if (swModelDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
        {
            var asm = (AssemblyDoc)swModelDoc;
            asm.ResolveAllLightweight();

            var components = (object[])asm.GetComponents(true);

            if (components != null)
                foreach (Component2 comp in components)
                    TraverseFlat(comp, root.Id, result);
        }

        return result;
    }

    private void TraverseFlat(Component2 comp, int parentId, List<SwNode> result)
    {
        if (comp == null) return;

        var model = (ModelDoc2)comp.GetModelDoc2();
        if (model == null) return;

        var node = CreateNode(model, parentId);

        result.Add(node);

        var children = (object[])comp.GetChildren();
        if (children == null) return;

        foreach (Component2 child in children) TraverseFlat(child, node.Id, result);
    }

    private SwNode CreateNode(ModelDoc2 doc, int? parentId)
    {
        var configName = doc.ConfigurationManager.ActiveConfiguration.Name;

        var node = new SwNode
        {
            Id = GetNextId(),
            ParentId = parentId,

            CompName = Path.GetFileNameWithoutExtension(doc.GetPathName()),
            ModelPath = doc.GetPathName(),
            ConfigName = configName,
            IsOpenedReadOnly = doc.IsOpenedReadOnly(),
            Quantity = 1,
            Properties = new Dictionary<string, string>()
        };

        return node;
    }

    public List<SwNode> GroupNodes(List<SwNode> nodes)
    {
        if (nodes == null || nodes.Count == 0) return new List<SwNode>();

        var groupedNodes = nodes
            .GroupBy(n => new { n.ModelPath, n.ConfigName })
            .Select(g => new SwNode
            {
                CompName = g.First().CompName,
                ModelPath = g.Key.ModelPath,
                IsOpenedReadOnly = g.First().IsOpenedReadOnly,
                ConfigName = g.Key.ConfigName,
                Quantity = g.Sum(n => n.Quantity),
                Properties = g.First().Properties
            })
            .ToList();

        return groupedNodes;
    }
}