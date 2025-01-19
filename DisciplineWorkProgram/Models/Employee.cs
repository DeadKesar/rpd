using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace DisciplineWorkProgram.Models
{
    internal class Employee : HierarchicalCheckableElement
    {
        protected override IEnumerable<HierarchicalCheckableElement> GetNodes() => Enumerable.Empty<HierarchicalCheckableElement>();

        [JsonIgnore]
        public IDictionary<string, string> Props { get; } = new Dictionary<string, string>();

        public String name {  
            get => Props["Employee"]; 
            set => Props["Employee"] = value;
        }
        public String nameFordoc
        {
            get => Props["nameFordoc"];
            set => Props["nameFordoc"] = value;
        }
        public String position
        {
            get => Props["position"];
            set => Props["position"] = value;
        }
        public String FIO
        {
            get => Props["FIO"];
            set => Props["FIO"] = value;
        }
        public String institut
        {
            get => Props["institut"];
            set => Props["institut"] = value;
        }
    }
}
