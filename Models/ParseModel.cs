using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevocationParser.Models
{
    class ParseModel
    {
        public int ID { get; set; }
        public int id { get; set; }
        public string link { get; set; }
        public string date { get; set; }
        public string author { get; set; }
        public int rating { get; set; }
        public string description { get; set; }
        public string pros { get; set; }
        public string cons { get; set; }
    }
}
