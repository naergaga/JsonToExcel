using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcel
{
    public class ExportService
    {
        public bool FormatDecimal { get; set; } = true;

        public byte[] Export(ExportModel item2)
        {
            using (var p = new ExcelPackage())
            {
                CreateSheet(item2.propAndheader, item2.list, item2.sheetName, p);
                return p.GetAsByteArray();
            }
        }

        public byte[] Export<T1>(ExportModel<T1> item2)
        {
            using (var p = new ExcelPackage())
            {
                CreateSheet(item2.propAndheader, item2.list, item2.sheetName, p);
                return p.GetAsByteArray();
            }
        }

        private void CreateSheet<T>(Dictionary<string, string> propAndheader, IEnumerable<T> list, string sheetName, ExcelPackage p)
        {
            var props = list.First().GetType().GetProperties();
            List<PropertyInfo> propsList = new List<PropertyInfo>();

            //查找并排序
            foreach (var item in propAndheader)
            {
                var prop = props.FirstOrDefault(t => t.Name == item.Key);
                if (prop != null)
                {
                    propsList.Add(prop);
                }
            }

            var ws = p.Workbook.Worksheets.Add(sheetName);

            var row = 1;
            var col = 1;

            foreach (var item in propAndheader)
            {
                ws.Cells[row, col++].Value = item.Value;
            }

            row += 1;


            foreach (var item in list)
            {
                col = 1;
                foreach (var prop in propsList)
                {
                    var value = prop.GetValue(item);
                    //ws.Cells[row, col++].Value = handleValue(value);
                    HandleCellAndValue(ws.Cells[row, col++], value);
                }
                row += 1;
            }
        }

        private void CreateSheet(Dictionary<string, string> propAndheader, IEnumerable<Dictionary<string, object>> list, string sheetName, ExcelPackage p)
        {
            var props = list.First().Keys;
            List<string> propsList = new List<string>();

            //查找并排序
            foreach (var item in propAndheader)
            {
                var prop = props.FirstOrDefault(t => t == item.Key);
                if (prop != null)
                {
                    propsList.Add(prop);
                }
            }

            var ws = p.Workbook.Worksheets.Add(sheetName);

            var row = 1;
            var col = 1;

            foreach (var item in propAndheader)
            {
                ws.Cells[row, col++].Value = item.Value;
            }

            row += 1;


            foreach (var item in list)
            {
                col = 1;

                foreach (var prop in propsList)
                {
                    object value;
                    item.TryGetValue(prop, out value);
                    HandleCellAndValue(ws.Cells[row, col++], value);
                }
                row += 1;
            }
        }


        private void HandleCellAndValue(ExcelRange cell, object value)
        {
            if (value == null) return;
            var type = value.GetType();
            if (type == typeof(DateTime))
            {
                var t1 = (DateTime)value;
                cell.Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
            }
            else if (type == typeof(decimal))
            {
                if (!FormatDecimal)
                {

                    cell.Style.Numberformat.Format = "0.00###";
                }
                else
                {
                    cell.Style.Numberformat.Format = "0.00";
                }
            }
            else if (type == typeof(JObject))
            {
                value = value.ToString();
            }

            cell.Value = value;
        }
    }

    public class ExportModel
    {
        public Dictionary<string, string> propAndheader { get; set; }
        public IEnumerable<Dictionary<string, object>> list { get; set; }
        public string sheetName { get; set; }
    }

    public class ExportModel<T>
    {
        public Dictionary<string, string> propAndheader { get; set; }
        public IEnumerable<T> list { get; set; }
        public string sheetName { get; set; }
    }
}
