using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace CADShark.Common.SolidWorks
{
    public class Materials
    {
        public bool IsDirty { get; set; }

        public string ID { get; set; }
        public string ParentID { get; set; }


        #region
        private string materialName;
        private string density;
        private string swProperty;
        private string xxhatch;
        private string aangle;
        private string sscale;
        private string ppwshader2;
        private string ppath;
        private string rrgb;
        #endregion

        public string MaterialName
        {
            get { return materialName; }
            set { materialName = value; IsDirty = true; }
        }
        public string Density
        {
            get { return density; }
            set { density = value; IsDirty = true; }
        }
        public string SWProperty
        {
            get { return swProperty; }
            set { swProperty = value; IsDirty = true; }
        }
        public string xhatch
        {
            get { return xxhatch; }
            set { xxhatch = value; IsDirty = true; }
        }
        public string angle
        {
            get { return aangle; }
            set { aangle = value; IsDirty = true; }
        }
        public string scale
        {
            get { return sscale; }
            set { sscale = value; IsDirty = true; }
        }
        public string pwshader2
        {
            get { return ppwshader2; }
            set { ppwshader2 = value; IsDirty = true; }
        }
        public string path
        {
            get { return ppath; }
            set { ppath = value; IsDirty = true; }
        }

        public string rgb
        {
            get { return rrgb; }
            set { rrgb = value; IsDirty = true; }
        }


        public static List<Materials> allofmaterials = null;
        /// <summary>
        /// Используеться только для сравнения, НИЧЕМУ НЕ ПРИСВАИВАТЬ
        /// </summary>
        public static List<Materials> allMaterialsUnchangeble = null;


        public Materials()
        {

        }
        public Materials(DataTable allProps)
        {
            allofmaterials = new List<Materials>(allProps.Rows.Count);
            allMaterialsUnchangeble = new List<Materials>();

            foreach (DataRow item in allProps.Rows)
            {
                Materials m = new Materials()
                {
                    angle = item.Field<string>("angle"),
                    Density = item.Field<string>("Density"),
                    ID = item.Field<string>("ID"),
                    ParentID = item.Field<string>("ParentID"),
                    MaterialName = item.Field<string>("MaterialName"),
                    path = item.Field<string>("path"),
                    pwshader2 = item.Field<string>("pwshader2"),
                    rgb = item.Field<string>("rgb"),
                    scale = item.Field<string>("scale"),
                    SWProperty = item.Field<string>("SWProperty"),
                    xhatch = item.Field<string>("xhatch"),
                    IsDirty = false
                };

                allofmaterials.Add(m);
                allMaterialsUnchangeble.Add(m.MemberwiseClone() as Materials);  //клонирование копирует все свойства, но обьекты 

                //остаются разными. Изменения в одном не затрвгивают другой, его копию

            }
        }

        public override bool Equals(object obj)
        {

            Materials mt = obj as Materials;


            if (this.angle == mt.angle && this.Density == mt.Density && this.ID == mt.ID && this.MaterialName == mt.MaterialName
                && this.ParentID == mt.ParentID && this.path == mt.path && this.pwshader2 == mt.pwshader2 && this.rgb == mt.rgb
                && this.scale == mt.scale && this.SWProperty == mt.SWProperty && this.xhatch == mt.xhatch)
                return true;


            return false;
        }

        public List<Materials> SortByParent(string type)
        {

            return allofmaterials.Where(x => x.ParentID == type).ToList();
        }

        public static DataTable TransleteIntoTable(string type)
        {
            var mo = allofmaterials.Where(x => x.ParentID == type).ToList();

            DataTable tabMat = new DataTable();

            tabMat.Columns.Add("ID", typeof(long));
            tabMat.Columns.Add("Name", typeof(string));

            foreach (Materials item in mo)
            {
                tabMat.Rows.Add(item.ID, item.MaterialName);
            }



            return tabMat;
        }

        public static bool CompareTwoItems(Materials m1, Materials m2)
        {

            if (m1.Equals(m2))
                return true;

            return false;

        }

        public static string WhichParametrsHasChanged(Materials m1, Materials m2)
        {
            string set = "";

            if (m1.MaterialName != m2.MaterialName)
            {
                set += " MaterialName = '" + m1.MaterialName + "' ,";
            }
            if (m1.Density != m2.Density)
            {
                set += " Density = '" + m1.Density + "' ,";
            }
            if (m1.SWProperty != m2.SWProperty)
            {
                set += " SWProperty = '" + m1.SWProperty + "' ,";
            }
            if (m1.xhatch != m2.xhatch)
            {
                set += " xhatch = '" + m1.xhatch + "' ,";
            }
            if (m1.angle != m2.angle)
            {
                set += " angle = '" + m1.angle + "' ,";
            }
            if (m1.scale != m2.scale)
            {
                set += " scale = '" + m1.scale + "' ,";
            }
            if (m1.pwshader2 != m2.pwshader2)
            {
                set += " pwshader2 = '" + m1.pwshader2 + "' ,";
            }
            if (m1.path != m2.path)
            {
                set += " path = '" + m1.path + "' ,";
            }
            if (m1.rgb != m2.rgb)
            {
                set += " rgb = '" + m1.rgb + "' ,";
            }

            if (set.Length > 2)
            {
                set = set.Remove(set.Length - 2, 2);

                set += " where ID = " + m1.ID + ";";
            }

            return set;
        }

        public static DataTable Transponse(Materials m)
        {
            DataTable tt = new DataTable();

            tt.Columns.Add("Свойство", typeof(string));
            tt.Columns.Add("Значение");


            PropertyInfo[] allProps = m.GetType().GetProperties();

            foreach (PropertyInfo item in allProps)
            {
                if (item.Name != "IsDirty")
                {
                    object t = item.GetValue(m);
                    if (t != null)
                        t = t?.ToString();

                    tt.Rows.Add(item.Name, t);

                }
            }

            return tt;
        }

        public static Materials ConvertTransponedTableToMateraial(DataTable tv)
        {

            Materials m = new Materials();


            foreach (DataRow item in tv.Rows)
            {
                if (item[0].ToString() == "angle")
                {
                    m.angle = item[1].ToString();
                }
                else if (item[0].ToString() == "Density")
                {
                    m.Density = item[1].ToString();
                }
                else if (item[0].ToString() == "ID")
                {
                    m.ID = item[1].ToString();
                }
                else if (item[0].ToString() == "MaterialName")
                {
                    m.MaterialName = item[1].ToString();
                }
                else if (item[0].ToString() == "ParentID")
                {
                    m.ParentID = item[1].ToString();
                }
                else if (item[0].ToString() == "path")
                {
                    m.path = item[1].ToString();
                }
                else if (item[0].ToString() == "pwshader2")
                {
                    m.pwshader2 = item[1].ToString();
                }
                else if (item[0].ToString() == "rgb")
                {
                    m.rgb = item[1].ToString();
                }
                else if (item[0].ToString() == "scale")
                {
                    m.scale = item[1].ToString();
                }
                else if (item[0].ToString() == "SWProperty")
                {
                    m.SWProperty = item[1].ToString();
                }
                else if (item[0].ToString() == "xhatch")
                {
                    m.xhatch = item[1].ToString();
                }
            }


            return m;
        }
    }
}
