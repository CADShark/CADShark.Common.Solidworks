using CADBooster.SolidDna;
using System.Collections.Generic;
using System;
using SolidWorks.Interop.sldworks;

namespace CADShark.Common.Solidworks
{
    public class Items
    {
        #region Default constructor
        static Items()
        {
            if (!AddInIntegration.ConnectToActiveSolidWorksForStandAlone())
            {
                const string message = "Failed to connect to SolidWorks";
                //CadLogger.Error(message);
                //throw new ConnectionException(message);
            }
            SwApp = SolidWorksEnvironment.Application;
        }
        #endregion

        #region Private properties
        private static readonly SolidWorksApplication SwApp;
        private static string typeForDoc = null;
        private static string codeForDoc = null;
        private static string objectGuid = null;
        #endregion

        #region Public Prpoperties
        public static string AddDocument => GetDocumentType();
        #endregion

        #region Public Methods
        public static Dictionary<string, string> DocumentProps()
        {
            var swApp = SwApp.ActiveModel;
            string valueDesignation;
            //string filePath = swApp.FilePath;
            string fileName = swApp.UnsafeObject.GetTitle();

            var propertyModelList = new Dictionary<string, string>();

            if (swApp.IsDrawing) swApp = DrawingItem();

            var activConfigurationName = swApp.ActiveConfiguration.Name;

            var editor = swApp.Extension.CustomPropertyEditor(activConfigurationName);
            var status = editor.CustomPropertyExists(ConstatnProps.attributeDesignation);
            if (status)
            {
                valueDesignation = editor.GetCustomProperty(ConstatnProps.attributeDesignation, true);
            }
            else
            {
                valueDesignation = fileName;
                editor.SetCustomProperty(ConstatnProps.attributeDesignation, valueDesignation);
            }
            propertyModelList.Add(SystemGUIDs.attributeDesignation, valueDesignation);


            foreach (var item in ConfiguretionPropsDic())
            {
                var prop = editor.GetCustomProperty(item.Value, true);
                propertyModelList.Add(item.Key, prop);
            }

            return propertyModelList;


        }
        #endregion

        #region Private Methods
        private static string GetDocumentType()
        {
            var model = SwApp.ActiveModel;
            switch (model.ModelType)
            {
                case ModelType.Assembly:
                    typeForDoc = @"Сборки SOLIDWORKS";
                    codeForDoc = "ЭСБ";
                    objectGuid = SystemGUIDs.objtypeSwASM;
                    break;
                case ModelType.Part:
                    typeForDoc = @"Чертежи SolidWorks";
                    codeForDoc = "МД";
                    objectGuid = SystemGUIDs.objtypeSwPart;
                    break;
                case ModelType.Drawing:
                    objectGuid = GetDocumentTypeForDrawing();
                    break;
            }

            Console.WriteLine($"{typeForDoc}\t{codeForDoc}");
            return objectGuid;
        }
        private static string GetDocumentTypeForDrawing()
        {
            switch (DrawingItem().ModelType)
            {
                case ModelType.Assembly:
                    typeForDoc = @"Сборочные чертежи SolidWorks";
                    codeForDoc = "CK";
                    objectGuid = SystemGUIDs.objtypeSwASMDrawing;
                    break;
                case ModelType.Part:
                    typeForDoc = @"Чертежи SolidWorks";
                    objectGuid = SystemGUIDs.objtypeSwPartDrawing;
                    break;
                case ModelType.None:
                    break;
            }
            return objectGuid;

        }

        //private static ObjectTypes CheckWhatTypeisModel(string type)
        //{
        //    switch (type)
        //    {
        //        case @"Детали":
        //        case @"Деталі":
        //            return ObjectTypes.Part;

        //        case @"Сборочные единицы":
        //        case @"Складальні одиниці":
        //            return ObjectTypes.Assembly;

        //        case @"Стандартные изделия":
        //            return ObjectTypes.StandardProd;

        //        case @"Прочие изделия":
        //            return ObjectTypes.OtherProducts;
        //        default:
        //            return ObjectTypes.None;
        //    }
        //}

        private static Dictionary<string, string> ConfiguretionPropsDic()
        {
            var configuretionPropsDic = new Dictionary<string, string>
            {
                { SystemGUIDs.attributeCADMECH_Density, "CADMECH_Density" },
                { SystemGUIDs.attributeWeight,  "Масса" },
                { SystemGUIDs.attrMaterial, "Материал" },
                { SystemGUIDs.attributeName, "Наименование" },
                //{ SystemGUIDs.attributeDesignation, "MaterialID" },
                { SystemGUIDs.attributeSectionNum, "Раздел" },
                { SystemGUIDs.attributeForeignkeyObject, "Внешний ключ изделия" },
            };
            return configuretionPropsDic;
        }
        private static Model DrawingItem()
        {
            var swDraw = SwApp.ActiveModel.AsDrawing();

            var sheets = (object[])swDraw.GetSheetNames();

            var firstSheet = sheets[0].ToString();

            swDraw.ActivateSheet(firstSheet);

            var swView = (View)swDraw.GetFirstView();

            swView = (View)swView.GetNextView();

            if (swView == null) return null;

            var swModel = new Model(swView.ReferencedDocument);

            return swModel;
        }
        #endregion
    }
}
