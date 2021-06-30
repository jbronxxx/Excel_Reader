using System.Collections.Generic;

namespace ExcelReader.Excelservice
{
    public class ModelNameView
    {
        public string Title { get; set; }

        public List<ModelList> ModelNames { get; set; }

        public ModelNameView()
        {
            ModelNames = new List<ModelList>(); 
        }
    }
}
