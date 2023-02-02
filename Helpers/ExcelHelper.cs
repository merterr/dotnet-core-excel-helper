using ClosedXML.Excel;
using ExcelHelperProject.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelHelperProject.Helpers
{
    public static class ExcelHelper<T> where T : class
    {
        public static byte[] CreateExcelFile(List<T> list)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var cols = GetColumnList(typeof(T), out List<PropertyInfo> propertyInfoList);

            for (var c = 0; c < cols.Count; c++)
                ws.Cell(1, c + 1).Value = cols[c]; // header
            

            for (var r = 0; r < list.Count; r++)
            {
                for (int i = 0; i < propertyInfoList.Count; i++)
                {
                    object item = propertyInfoList[i].GetValue(list[r], null) ?? "";

                    switch (item)
                    {
                        case decimal d:
                            ws.Cell(r + 2, i + 1).SetValue(d).Style.NumberFormat.Format = "#,##0.00";
                            break;
                        case DateTime date:
                            ws.Cell(r + 2, i + 1).SetValue(date).Style.DateFormat.Format = "dd.MM.yyyy";
                            break;
                        case bool b:
                            ws.Cell(r + 2, i + 1).Value = b ? "Yes" : "No";
                            break;
                        case int integer:
                            ws.Cell(r + 2, i + 1).SetValue(integer).Style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Integer);
                            break;
                        default:
                            ws.Cell(r + 2, i + 1).SetValue(item.ToString()).Style.NumberFormat.Format = "@";
                            break;
                    }
                }
            }

            using MemoryStream stream = new ();
            wb.SaveAs(stream);
            return stream.ToArray();
        }

        private static List<string> GetColumnList(Type type, out List<PropertyInfo> propertyInfoList)
        {
            propertyInfoList = new List<PropertyInfo>();
            var propertList = type.GetProperties();
            var resultList = new List<string>();
            foreach (var property in propertList)
            {
                var attr = property.GetCustomAttributes<ExcelColumnNameAttribute>().FirstOrDefault();
                if (attr != null)
                {
                    resultList.Add(attr.GetName());
                    propertyInfoList.Add(property);
                }
            }
            return resultList;
        }
    }
}
