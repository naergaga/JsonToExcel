using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcel
{
    public class MainDataContext : ObservableObject
    {
        private string jsonPath;
        private string outputPath;
        private ConvertOptions options= new ConvertOptions {ArrayMinCount=2,ListMode=true };

        public string JsonPath { get => jsonPath; set { Set(ref jsonPath, value); } }
        public string OutputPath { get => outputPath; set { Set(ref outputPath, value); } }
        public ConvertOptions Options { get => options; set { Set(ref options, value); } }

        public void Export(string path)
        {
            outputPath = Path.GetDirectoryName(path);
            var converter = new JsonConverter(options);
            converter.ToExcelFile(jsonPath, path);
        }
        
    }
}
