using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Karban.Models;
using System.IO;
using OfficeOpenXml;
using System.Text;
using Newtonsoft.Json;
using System.Data;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace Karban.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            // path to your excel file
            string path = "D:\\data.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            StringBuilder sb = new StringBuilder();
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows; // 20
            int columns = worksheet.Dimension.Columns; // 7

            List<KanbanModel> lstkanbanModels = new List<KanbanModel>();

            // loop through the worksheet rows and columns
            for (int row = 2; row <= rows; row++)
            {
                KanbanModel kanban = new KanbanModel();
                kanban.CodeId = worksheet.Cells[row, 1].Value.ToString();
                kanban.Subject = worksheet.Cells[row, 2].Value.ToString();
                kanban.DeveloperName = worksheet.Cells[row, 3].Value.ToString();
                kanban.AssignedOn = worksheet.Cells[row, 4].Value.ToString();
                kanban.Priority = worksheet.Cells[row, 5].Value.ToString();
                kanban.StatusCode = worksheet.Cells[row, 6].Value.ToString();
                lstkanbanModels.Add(kanban);

            }
            TempData["data"] = lstkanbanModels;

            return View(lstkanbanModels);
        }

       public IActionResult CriticalPartners()
        {
            // path to your excel file
            string path = "D:\\data.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            StringBuilder sb = new StringBuilder();
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows; // 20
            int columns = worksheet.Dimension.Columns; // 7

            List<KanbanModel> lstkanbanModels = new List<KanbanModel>();

            // loop through the worksheet rows and columns
            for (int row = 2; row <= rows; row++)
            {
                KanbanModel kanban = new KanbanModel();
                kanban.CodeId = worksheet.Cells[row, 1].Value.ToString();
                kanban.Subject = worksheet.Cells[row, 2].Value.ToString();
                kanban.DeveloperName = worksheet.Cells[row, 3].Value.ToString();
                kanban.AssignedOn = worksheet.Cells[row, 4].Value.ToString();
                kanban.Priority = worksheet.Cells[row, 5].Value.ToString();
                kanban.StatusCode = worksheet.Cells[row, 6].Value.ToString();
                lstkanbanModels.Add(kanban);

            }
            TempData["data"] = lstkanbanModels;

            return View(lstkanbanModels);
        }

        public IActionResult OnlyMyPartners()
        {
            // path to your excel file
            string path = "D:\\data.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            StringBuilder sb = new StringBuilder();
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows; // 20
            int columns = worksheet.Dimension.Columns; // 7

            List<KanbanModel> lstkanbanModels = new List<KanbanModel>();

            // loop through the worksheet rows and columns
            for (int row = 2; row <= rows; row++)
            {
                KanbanModel kanban = new KanbanModel();
                kanban.CodeId = worksheet.Cells[row, 1].Value.ToString();
                kanban.Subject = worksheet.Cells[row, 2].Value.ToString();
                kanban.DeveloperName = worksheet.Cells[row, 3].Value.ToString();
                kanban.AssignedOn = worksheet.Cells[row, 4].Value.ToString();
                kanban.Priority = worksheet.Cells[row, 5].Value.ToString();
                kanban.StatusCode = worksheet.Cells[row, 6].Value.ToString();
                lstkanbanModels.Add(kanban);

            }
            TempData["data"] = lstkanbanModels;

            return View(lstkanbanModels);
        }

        public IActionResult RecentlyUpdated()
        {
            // path to your excel file
            string path = "D:\\data.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            StringBuilder sb = new StringBuilder();
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows; // 20
            int columns = worksheet.Dimension.Columns; // 7

            List<KanbanModel> lstkanbanModels = new List<KanbanModel>();

            // loop through the worksheet rows and columns
            for (int row = 2; row <= rows; row++)
            {
                KanbanModel kanban = new KanbanModel();
                kanban.CodeId = worksheet.Cells[row, 1].Value.ToString();
                kanban.Subject = worksheet.Cells[row, 2].Value.ToString();
                kanban.DeveloperName = worksheet.Cells[row, 3].Value.ToString();
                kanban.AssignedOn = worksheet.Cells[row, 4].Value.ToString();
                kanban.Priority = worksheet.Cells[row, 5].Value.ToString();
                kanban.StatusCode = worksheet.Cells[row, 6].Value.ToString();
                lstkanbanModels.Add(kanban);

            }
            TempData["data"] = lstkanbanModels;

            return View(lstkanbanModels);
        }




        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
