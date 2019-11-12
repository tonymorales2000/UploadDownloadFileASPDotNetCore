using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace UploadDownloadFileASPDotNetCore
{
    [Route("api/[controller]")]
    public class DownloadController : Controller
    {
        public IHostingEnvironment _hostingEnvironment { get; }

        public DownloadController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        // GET: api/<controller>
        [HttpGet]
        public async Task<IActionResult> Get()
        {
            var path = "C:\\develop\\OdeToFood\\Notes\\Sql Server Object Explorer.PNG";
            var memory = new MemoryStream();
            using (var stream = new FileStream(path, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return File(memory, MimeTypes.GetFileType()[ext], Path.GetFileName(path));
        }

        // GET api/<controller>/5
        [HttpGet("GetExcel") , Route("GetExcel")]
        public async Task<IActionResult> GetExcelAsync(int id)
        {
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"demo.xlsx";
            string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            var memory = new MemoryStream();
            using (var fs = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook;
                workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Demo");
                IRow row = excelSheet.CreateRow(0);

                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("Name");
                row.CreateCell(2).SetCellValue("Age");

                row = excelSheet.CreateRow(1);
                row.CreateCell(0).SetCellValue(1);
                row.CreateCell(1).SetCellValue("Kane Williamson");
                row.CreateCell(2).SetCellValue(29);

                row = excelSheet.CreateRow(2);
                row.CreateCell(0).SetCellValue(2);
                row.CreateCell(1).SetCellValue("Martin Guptil");
                row.CreateCell(2).SetCellValue(33);

                row = excelSheet.CreateRow(3);
                row.CreateCell(0).SetCellValue(3);
                row.CreateCell(1).SetCellValue("Colin Munro");
                row.CreateCell(2).SetCellValue(23);

                workbook.Write(fs);
            }
            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;

            return File(memory, MimeTypes.GetFileType()[".xlsx"], "Report.xlsx");
        }

        [HttpGet("GetExcel2"), Route("GetExcel2")]
        public async Task<IActionResult> DownloadFile(int id)
        {
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"test1.xlsx";
            //var wb = await BuildExcelFile(id);
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("This is Crop Plan");
            wb.SaveAs(Path.Combine(sWebRootFolder, sFileName));
            var memory = new MemoryStream();
            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;

            return File(memory, MimeTypes.GetFileType()[".xlsx"], sFileName);
        }


        [HttpGet("GetExcel3"), Route("GetExcel3")]
        public async Task<IActionResult> DownloadFile3(int id)
        {
            var memory = new MemoryStream();
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"test2.xlsx";
            //var wb = await BuildExcelFile(id);
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var firstCell = ws.FirstCell().SetValue("This is Crop Plan from Decisive Farming Application");
            firstCell.Style.Font.SetBold()
                .Fill.SetBackgroundColor(XLColor.CyanProcess);
            ws.SheetView.FreezeRows(2);
            var titleCell = firstCell.CellBelow().SetValue("Crop Type");
            var totalCell = titleCell.CellBelow().SetValue("Barley")
            .CellBelow().SetValue("Canola")
            .CellBelow().SetValue("Wheat");
            ws.FirstCell().WorksheetColumn().AdjustToContents();
            
            //af.AppendChild(firstCell.CellBelow().SetValue("Crop Type"));
            //var fcu = ws.FirstCell().WorksheetRow().FirstCellUsed();
            //var a6 = ws.Cell("A6").SetValue(fcu.CurrentRegion.ToString());
            
            ws.Range(titleCell, totalCell).SetAutoFilter().Sort();
            wb.SaveAs(memory);
            
           
            memory.Position = 0;

            return File(memory, MimeTypes.GetFileType()[".xlsx"], sFileName);
        }

        [HttpGet("GetExcel4"), Route("GetExcel4")]
        public async Task<IActionResult> DownloadFile4(int id)
        {
            var memory = new MemoryStream();
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"test2.xlsx";
            //var wb = await BuildExcelFile(id);
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var firstCell = ws.FirstCell().SetValue("This is Crop Plan from Decisive Farming Application");
            firstCell.Style.Font.SetBold()
                .Fill.SetBackgroundColor(XLColor.CyanProcess);
            ws.SheetView.FreezeRows(2);
            var titleCell = firstCell.CellBelow().SetValue("Crop Type");
            var totalCell = titleCell.CellBelow().SetValue("Barley")
            .CellBelow().SetValue("Canola")
            .CellBelow().SetValue("Wheat");
            ws.FirstCell().WorksheetColumn().AdjustToContents();
            titleCell.Select();
            //af.AppendChild(firstCell.CellBelow().SetValue("Crop Type"));
            //var fcu = ws.FirstCell().WorksheetRow().FirstCellUsed();
            //var a6 = ws.Cell("A6").SetValue(fcu.CurrentRegion.ToString());

            //ws.Range(titleCell, totalCell).SetAutoFilter().Sort();
            //ws.Range(titleCell, totalCell).SetAutoFilter();
            wb.SaveAs(memory);


            memory.Position = 0;

            return File(memory, MimeTypes.GetFileType()[".xlsx"], sFileName);
        }



        private async Task<XLWorkbook> BuildExcelFile(int id)
        {
            //Creating the workbook
            var t = Task.Run(() =>
            {
                var wb = new XLWorkbook();
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().SetValue(id);

                return wb;
            });

            return await t;
        }


        // POST api/<controller>
        [HttpPost]
        public void Post([FromBody]string value)
        {
        }

        // PUT api/<controller>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/<controller>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
