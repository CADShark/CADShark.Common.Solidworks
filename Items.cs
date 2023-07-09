using System;
using System.Collections.Generic;
using System.Linq;
using CADBooster.SolidDna;
using CADShark.Common.Logging;
using CADShark.Common.SolidWorks.Errors;
using CADShark.Kernel;
using CADShark.SolidWorks.AddIn.Assemblies;
using CADShark.SolidWorks.AddIn.Models;
using SolidWorks.Interop.sldworks;

namespace CADShark.SolidWorks.AddIn
{
    public static class Items
    {
        #region Private properties
        private static readonly SolidWorksApplication SwApp;
        private static string typeForDoc = null;
        private static string codeForDoc = null;
        private static string objectGuid = null;
        #endregion

        #region Default constructor
        static Items()
        {
            if (!AddInIntegration.ConnectToActiveSolidWorksForStandAlone())
            {
                const string message = "Failed to connect to SolidWorks";
                CadLogger.Error(message);
                throw new ConnectionException(message);
            }
            SwApp = SolidWorksEnvironment.Application;
        }
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
        private static ObjectTypes CheckWhatTypeisModel(string type)
        {
            switch (type)
            {
                case @"Детали":
                case @"Деталі":
                    return ObjectTypes.Part;

                case @"Сборочные единицы":
                case @"Складальні одиниці":
                    return ObjectTypes.Assembly;

                case @"Стандартные изделия":
                    return ObjectTypes.StandardProd;

                case @"Прочие изделия":
                    return ObjectTypes.OtherProducts;
                default:
                    return ObjectTypes.None;
            }
        }
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

        #region Tresh

        private static List<PropertyModel> CreateItem()
        {
            var propertyModelList = new List<PropertyModel>();

            var configNames = SwApp.ActiveModel.ConfigurationNames;

            foreach (var configName in configNames)
            {
                var number = SwApp.ActiveModel.GetCustomProperty(@"Обозначение", configName, true);
                var description = SwApp.ActiveModel.GetCustomProperty(@"Наименование", configName, true);
                var section = SwApp.ActiveModel.GetCustomProperty(@"Раздел", configName, true);

                var propertyData = new PropertyModel
                {
                    Number = number,
                    Description = description,
                    ObjectType = CheckWhatTypeisModel(section),
                    ConfigName = configName
                };
                propertyModelList.Add(propertyData);
            }

            return propertyModelList;
        }







        private static Dictionary<string, string> CustomPropsDic()
        {
            var customPropsDic = new Dictionary<string, string>
            {
                //{SystemGUIDs.attributeDesignation, "Разработал"},
                //{SystemGUIDs.attributeDesignation, "Проверил" },
                //{SystemGUIDs.attributeDesignation, "Техконтроль"},
                //{SystemGUIDs.attributeDesignation, "Н.контр."},
                //{SystemGUIDs.attributeDesignation, "Нач.отд."},
                //{SystemGUIDs.attributeDesignation, "Утвердил"},
                //{ "Перв.примен", SystemGUIDs.attributeDesignation, },
                //{ "Справ.N", SystemGUIDs.attributeDesignation, },
                //{ "Инв.N подл", SystemGUIDs.attributeDesignation, },
                //{ "Подп. и дата 1", SystemGUIDs.attributeDesignation, },
                //{ "Взам. инв.N", SystemGUIDs.attributeDesignation, },
                //{ "Инв. N дубл.", SystemGUIDs.attributeDesignation, },
                //{ "Подп. и дата 2", SystemGUIDs.attributeDesignation, },
                //{ "Подразделение", SystemGUIDs.attributeDesignation, },
                //{ "MassaFormat", SystemGUIDs.attributeDesignation, },
                //{ "ФОРМАТ", SystemGUIDs.attributeDesignation, },
                //{ "ЛИСТОВ", SystemGUIDs.attributeDesignation, },
                //{ "КОЛИЧЕСТВО ФОРМАТОВ A3", SystemGUIDs.attributeDesignation, },
                //{SystemGUIDs.attributeDesignation, "Код документа" },
                //{ SystemGUIDs.attributeDesignation, "Тип документа"}
            };

            return customPropsDic;
        }

        private static Dictionary<string, string> DocumentProps1()
        {
            string valueDesignation;
            var swapp = SwApp.ActiveModel;

            if (swapp.IsDrawing) swapp = DrawingItem();


            string filePath = swapp.FilePath;
            string fileName = swapp.UnsafeObject.GetTitle();

            var activConfigurationName = swapp.ActiveConfiguration.Name;

            var editor = swapp.Extension.CustomPropertyEditor(activConfigurationName);
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

            var valueName = swapp.GetCustomProperty(ConstatnProps.attributeName, null, true);


            var propertyModelList = new Dictionary<string, string>
            {
                {SystemGUIDs.attributeDesignation, valueDesignation},
                {SystemGUIDs.attributeName, valueName},
                {SystemGUIDs.attributeFile, filePath}
            };


            return propertyModelList;
        }








        private static void SuppressUpdates()
        {
            ModelView myModelView = null;

            var swDraw = SwApp.ActiveModel.AsDrawing();
        }
        private static List<PropertyModel> GetComponents()
        {
            var componentsList = new List<PropertyModel>();

            var model = SwApp.ActiveModel;
            var assembly = model.AsAssembly();

            foreach (var component in AssemblyHelpers.GetComponents(assembly, false))
            {
                var comp = component.AsModel;
                var config = component.ConfigurationName;

                if (component.IsSuppressed)
                    continue;

                var number = comp.GetCustomProperty(@"Обозначение", config, true);
                var description = comp.GetCustomProperty(@"Наименование", config, true);
                if (string.IsNullOrEmpty(description))
                {
                    description = comp.GetCustomProperty(@"Наименование", null, true);
                }

                var section = comp.GetCustomProperty(@"Раздел", config, true);

                var componentData = new PropertyModel
                {
                    Number = number,
                    Description = description,
                    ObjectType = CheckWhatTypeisModel(section),
                    ConfigName = config,
                    FilePath = component.FilePath
                };
                componentsList.Add(componentData);
            }
            var groupedComponents = componentsList
                .GroupBy(c => new { c.FilePath, c.ConfigName })
                .Select(g => new PropertyModel
                {
                    Number = g.First().Number,
                    Description = g.First().Description,
                    ObjectType = g.First().ObjectType,
                    ConfigName = g.Key.ConfigName,
                    FilePath = g.Key.FilePath,
                    Count = g.Count()
                })
                .ToList();

            return groupedComponents;
        }

        #endregion



    }
}
