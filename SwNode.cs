using System.Collections.Generic;

namespace CADShark.Common.SolidWorks
{
    public class SwNode
    {
        public int Id { get; set; }
        public int? ParentId { get; set; }
        public string PathKey { get; set; }
        public Dictionary<string, string> Properties { get; set; } = new();
        public string CompName { get; set; }
        public string ModelPath { get; set; }
        public string ConfigName { get; set; }
        public int Quantity { get; set; } = 1;
        public bool IsOpenedReadOnly { get; set; }
    }
}
