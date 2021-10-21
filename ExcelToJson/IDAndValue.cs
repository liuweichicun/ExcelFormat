using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormat
{
    [Serializable]
    public class IDAndValue
    {
        public int id;
        public int value;


        public override string ToString()
        {
            return "{\"id\"" + ":" + id + ",\"value\"" + ":" + value + "}";
        }
    }
}
