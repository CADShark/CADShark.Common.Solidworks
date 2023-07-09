namespace CADShark.SolidWorks.AddIn.Models
{
    public class PropertyModel
    {
        public string Number { get; set; }
        public string Description { get; set; }
        public ObjectTypes ObjectType { get; set; }
        public string ConfigName { get; set; }
        public int Count { get; set; }
        public int F_ID { get; set; }
        public string FilePath { get; set; }
        public string ObjectGuid { get; set; }
    }
}
