using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

using OfficeOpenXml;
using Evs.Measurement.Domain.Entities;
using Q.Shared.CommonTypes;
using System.IO;

namespace DomainXLSXTemplateGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var domainList = GetDomainList();

            using (ExcelPackage excel = new ExcelPackage())
            {
                int i = 0;
                foreach(Type t in domainList)
                {
                    ++i;

                    excel.Workbook.Worksheets.Add(t.Name);
                    ExcelWorksheet ws = excel.Workbook.Worksheets[i];

                    List<string> listPropNames = new List<string>();

                    TraverseType(t, (prefix,prop) => listPropNames.Add($"{prefix}{prop.Name}"));

                    List<string[]> headerRow = new List<string[]>()
                    {
                        listPropNames.ToArray()
                    };

                    ws.Cells[1, 1, 1, listPropNames.Count].LoadFromArrays(headerRow);
                }

                string filePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + @"\domain.xlsx";

                FileInfo excelFile = new FileInfo(filePath);
                excel.SaveAs(excelFile);
            }
        }
        
        private static List<Type> GetDomainList()
        {
            var list = new List<Type>();

            list.Add(typeof(ProductQuality));
            list.Add(typeof(ProductQuantity));

            return list;
        }

        private static List<MolecularComposition> GetMolecularCompositions()
        {
            var list = new List<MolecularComposition>();

            MolecularComposition mc0 = new MolecularComposition();
            mc0.MassPercent = 94;
            list.Add(mc0);

            MolecularComposition mc1 = new MolecularComposition();
            mc0.MassPercent = 94;
            list.Add(mc1);

            MolecularComposition mc2 = new MolecularComposition();
            mc0.MassPercent = 94;
            list.Add(mc2);

            MolecularComposition mc3 = new MolecularComposition();
            mc0.MassPercent = 94;
            list.Add(mc3);

            MolecularComposition mc4 = new MolecularComposition();
            mc0.MassPercent = 94;
            list.Add(mc4);

            return list;
        }

        static void TraverseType(Type t, Action<string, PropertyInfo> action, int nestLevel = 0, string prefix = null)
        {
            if(String.IsNullOrEmpty(prefix))
            {
                prefix = "";
            }

            foreach (PropertyInfo p in t.GetProperties())
            {
                //Handle nullables
                Type propType;
                
                propType = Nullable.GetUnderlyingType(p.PropertyType);
                if (propType == null)
                {
                    propType = p.PropertyType;
                }

                if (nestLevel > 3 ||
                     propType.IsPrimitive || propType.IsEnum || propType == typeof(string) || propType == typeof(DateTime) || propType == typeof(Decimal))
                {
                    action(prefix, p);
                }
                else if (IsIEnumerable(propType))
                {
                    Type[] typeArguments = propType.GetGenericArguments();

                    if (typeArguments.Length > 0)
                    {
                        TraverseType(typeArguments[0], action, nestLevel + 1, $"{prefix}{p.Name}/0/");
                    }
                }
                else
                {
                    if(p.PropertyType == typeof(Component) ||
                       p.PropertyType == typeof(QuantityValues) ||
                       p.PropertyType == typeof(BalanceQuantityValues) ||
                       p.PropertyType == typeof(DateSpan) ||
                       IsSubclassOfRawGeneric(typeof(SystemKeyValuePair<>), p.PropertyType)
                        )
                    {
                        TraverseType(propType, action, nestLevel + 1, $"{prefix}{p.Name}/");
                    }
                    
                }
            }
        }

        static bool IsIEnumerable(Type t)
        {
            bool isIEnumerable = t.IsInterface && t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>);

            foreach (var i in t.GetInterfaces())
            {
                isIEnumerable |= i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>);
            }

            return isIEnumerable;
        }
        static bool IsSubclassOfRawGeneric(Type generic, Type toCheck)
        {
            while (toCheck != null && toCheck != typeof(object))
            {
                var cur = toCheck.IsGenericType ? toCheck.GetGenericTypeDefinition() : toCheck;
                if (generic == cur)
                {
                    return true;
                }
                toCheck = toCheck.BaseType;
            }
            return false;
        }
    }
}
