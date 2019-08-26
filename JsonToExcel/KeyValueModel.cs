using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcel
{
    public class KeyValueModel
    {

        public KeyValueModel(string key, object value)
        {
            this.Key = key;
            this.Value = value;
        }

        public string Key { get; set; }
        public object Value { get; set; }
    }
}
