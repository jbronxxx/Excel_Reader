using System.Collections.Generic;

namespace ExcelReader.Excelservice
{
    public class Model
    {
        public string Name { get; set; }

        public List<Price> Prices { get; set; }

        public Model()
        {
            Prices = new List<Price>();
        }
    }
}
