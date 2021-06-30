using System.Collections.Generic;

namespace ExcelReader.Excelservice
{
    public class ModelList
    {
        public string NameModel { get; set; }

        public List<Model> Models { get; set; }

        public ModelList()
        {
            Models = new List<Model>();
        }
    }
}
