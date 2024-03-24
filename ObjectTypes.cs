using System.ComponentModel;

namespace CADShark.Common.SolidWorks
{
    public enum ObjectTypes
    {
        /// <summary>
        /// Component is a part
        /// </summary>
        [Description("Деталі")] Part = 1052,

        /// <summary>
        /// Component is a sub-assembly
        /// </summary>
        [Description("Складальні одиниці")] Assembly = 1074,

        /// <summary>
        /// Component is a StandardProd
        /// </summary>
        [Description("Стандартные изделия")] StandardProd = 1105,

        /// <summary>
        /// Component is a Other products
        /// </summary>
        [Description("Прочие изделия")] OtherProducts = 1138,

        /// <summary>
        /// wasn't defined
        /// </summary>
        [Description("None")] None = 0
    }
}