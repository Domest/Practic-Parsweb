using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using ParsWeb.Models;
using Microsoft.EntityFrameworkCore;
using ClosedXML.Excel;


namespace ParsWeb.Controllers
{
    public class HomeController : Controller
    {
        private ApplicationContext db;
        public HomeController(ApplicationContext context)
        {
            db = context;
        }

        public async Task<IActionResult> Index()
        {
            return View(await db.Wagons.ToListAsync());
        }

        [HttpPost]
        public async Task<IActionResult> Create()
        {
            var enumexcel = EnumerateExcel();
            LocalList ll = new LocalList();
            foreach (var e in enumexcel)
            {
                ll.LocalWagonList.Add(new ParserModel() { TCNumber = e.TCNumber, WaybillNumber = e.WaybillNumber });
            }

            for (int i = 0; i < ll.LocalWagonList.Count; i++)
            {
                var wagon = new ParserModel()
                {
                    TCNumber = ll.LocalWagonList[i].TCNumber,
                    WaybillNumber = ll.LocalWagonList[i].WaybillNumber
                };
                db.Wagons.Add(wagon);
                
            }

            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        static IEnumerable<ParserModel> EnumerateExcel()
        {
            string xlsxpath = "WAGON_LIST 150520.xlsx";
            using (var workbook = new XLWorkbook(xlsxpath))
            {
                var worksheet = workbook.Worksheets.Worksheet(1);
                int rows = worksheet.RowsUsed().Count();
                for (int row = 2; row <= rows; ++row)
                {
                    var metric = new ParserModel
                    {
                        TCNumber = worksheet.Cell(row, 2).GetValue<int>(),
                        WaybillNumber = worksheet.Cell(row, 3).GetValue<string>()
                    };
                    yield return metric;
                }
            }
        }
    }
}
