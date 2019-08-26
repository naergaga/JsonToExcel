using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcel
{
    public class JsonConverter
    {
        private static Dictionary<string, string> headerMap1 = new Dictionary<string, string> {
                        { "Key",Properties.Resources.Key },
                        { "Value",Properties.Resources.Value }
                    };

        private ConvertOptions options;

        public JsonConverter(ConvertOptions options)
        {
            this.options = options;
        }

        public void ToExcelFile(string path, string savePath)
        {
            File.WriteAllBytes(savePath, ToExcel(path));
        }

        public byte[] ToExcel(string path)
        {
            var jsonObj = JsonConvert.DeserializeObject(File.ReadAllText(path));
            var jsonObjType = jsonObj.GetType();

            if (jsonObjType == typeof(JObject))
            {

                var obj = (JObject)jsonObj;
                if (options.ListMode)
                {
                    JArray array = FindJArray(obj);
                    if (array == null)
                    {
                        throw new Exception(Properties.Resources.Ex_UnsupportedFormat);
                    }

                    return ExportJArray(array);
                }
                else
                {
                    var jList = (IEnumerable<KeyValuePair<string, JToken>>)jsonObj;

                    //获取list<KeyValueModel>
                    var list = jList.Select(t =>
                    {
                        var item = t.Value.ToObject<dynamic>();
                        var model = new KeyValueModel(t.Key, item);
                        return model;
                    }).ToList();

                    var exportModel = new ExportModel<KeyValueModel>
                    {
                        list = list,
                        propAndheader = headerMap1,
                        sheetName = "Sheet1"
                    };
                    var service = new ExportService();
                    return service.Export(exportModel);
                }
            }
            else if (jsonObjType == typeof(JArray))
            {
                return ExportJArray((JArray)jsonObj);
            }

            throw new Exception(Properties.Resources.Ex_UnsupportedFormat);
        }


        private byte[] ExportJArray(JArray array)
        {
            var list = array.Children();
            var first = list.First();
            var headerMap = first.Children().Select(t => (JProperty)t).ToDictionary(t => t.Name, t => t.Name);
            var exportList = list.Select(t =>
            {
                Dictionary<string, object> map1 = new Dictionary<string, object>();
                foreach (var item in t.Children())
                {
                    var p1 = (JProperty)item;
                    map1.Add(p1.Name, p1.Value.ToObject<dynamic>());
                }
                return map1;
            });

            var exportModel = new ExportModel
            {
                list = exportList,
                propAndheader = headerMap,
                sheetName = "Sheet1"
            };

            var service = new ExportService();
            return service.Export(exportModel);
        }

       

        private JArray FindJArray(JToken jToken)
        {
            switch (jToken.Type)
            {
                case JTokenType.Object:
                    var jList = (IEnumerable<KeyValuePair<string, JToken>>)jToken;
                    foreach (var item in jList)
                    {
                        var array1 = FindJArray(item.Value);
                        if (array1 != null)
                        {
                            return array1;
                        }
                    }
                    break;
                case JTokenType.Array:
                    var array2 = (JArray)jToken;
                    if (array2.Count < options.ArrayMinCount)
                    {
                        foreach (var item1 in array2)
                        {
                            var array3 = FindJArray(item1);
                            if (array3 != null)
                            {
                                return array3;
                            }
                        }
                        return null;
                    }
                    return (JArray)jToken;
                default:
                    break;
            }
            return null;
        }
    }
}
