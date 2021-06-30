using ClosedXML.Excel;
using ExcelReader.Excelservice;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Linq;

namespace Excel_Reader.Controllers
{
    public class ExcelController : Controller
    {
        // GET: ExcelController1
        public ActionResult Index()
        {
            return View();
        }

        // GET: ExcelController1/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: ExcelController1/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ExcelController1/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: ExcelController1/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: ExcelController1/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: ExcelController1/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: ExcelController1/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        [HttpGet]
        public ActionResult Import()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(IFormFile fileExcel)
        {
            if (ModelState.IsValid)
            {
                // Список моделей 
                ModelList modelList = new ModelList();

                // Открывает Excel документ
                using (XLWorkbook workbook = new XLWorkbook(fileExcel.FileName, XLEventTracking.Disabled))
                {
                    // Проходит по листам открытого документа
                    foreach (IXLWorksheet worksheet in workbook.Worksheets)
                    {
                        // Проходит по колонкам, пропуская первыю колонку
                        foreach (IXLColumn column in worksheet.ColumnsUsed().Skip(1))
                        {
                            Model model = new Model();
                            // Записывает в список моделей значения первых ячеек каждой колонки
                            model.Name = column.Cell(1).Value.ToString();

                            // Проходит по строкам, пропуская первую
                            foreach (IXLRow row in worksheet.RowsUsed().Skip(1))
                            {
                                // Цены
                                Price price = new Price();
                                // Записывает значения первых ячеек каждой строки в качестве наименования цены
                                price.PriceName = row.Cell(1).Value.ToString();

                                // Записывает значения ячеек остальных строк в качестве цены на услугу
                                price.ServicePrice = row.Cell(column.ColumnNumber()).Value.ToString();

                                // Записывает цену в список цен у каждой модели
                                model.Prices.Add(price);
                            }

                            // Записывает модель в список моделей
                            modelList.Models.Add(model);
                        }
                    }

                    return View(modelList);
                }
            }

            return RedirectToAction("Index");
            
        }
        
    }
}
