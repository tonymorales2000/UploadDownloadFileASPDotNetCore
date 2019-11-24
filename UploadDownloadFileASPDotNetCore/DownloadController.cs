using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
//using NPOI.SS.UserModel;
//using NPOI.XSSF.UserModel;


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

        //// GET api/<controller>/5
        //[HttpGet("GetExcel") , Route("GetExcel")]
        //public async Task<IActionResult> GetExcelAsync(int id)
        //{
        //    string sWebRootFolder = _hostingEnvironment.WebRootPath;
        //    string sFileName = @"demo.xlsx";
        //    string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
        //    FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
        //    var memory = new MemoryStream();
        //    //using (var fs = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.Write))
        //    {
        //        IWorkbook workbook;
        //        workbook = new XSSFWorkbook();
        //        ISheet excelSheet = workbook.CreateSheet("Demo");
        //        IRow row = excelSheet.CreateRow(0);

        //        row.CreateCell(0).SetCellValue("ID");
        //        row.CreateCell(1).SetCellValue("Name");
        //        row.CreateCell(2).SetCellValue("Age");

        //        row = excelSheet.CreateRow(1);
        //        row.CreateCell(0).SetCellValue(1);
        //        row.CreateCell(1).SetCellValue("Kane Williamson");
        //        row.CreateCell(2).SetCellValue(29);

        //        row = excelSheet.CreateRow(2);
            //    row.CreateCell(0).SetCellValue(2);
            //    row.CreateCell(1).SetCellValue("Martin Guptil");
            //    row.CreateCell(2).SetCellValue(33);

            //    row = excelSheet.CreateRow(3);
            //    row.CreateCell(0).SetCellValue(3);
            //    row.CreateCell(1).SetCellValue("Colin Munro");
            //    row.CreateCell(2).SetCellValue(23);

            //    workbook.Write(fs);
            //}
        //    using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
        //    {
        //        await stream.CopyToAsync(memory);
        //    }
        //    memory.Position = 0;

        //    return File(memory, MimeTypes.GetFileType()[".xlsx"], "Report.xlsx");
        //}

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



        [HttpGet, Route("CreateTable")]
        public async Task<IActionResult> CreateTable(int id)
        {
            var memory = new MemoryStream();
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"table.xlsx";
            //var wb = await BuildExcelFile(id);
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            ws.PageSetup.AdjustTo(80);
            var farmNameCell = ws.FirstCell().SetValue("This is Farm Name");
            var totalAcresCell = farmNameCell.CellBelow().SetValue("Total Acres");
            DataTable table = GetNewTable();

            var iXLCell = totalAcresCell.CellBelow();
            var tableXl = iXLCell.InsertTable(table);
            tableXl.Theme = XLTableTheme.None;
            var tableColumns = tableXl.Worksheet.Columns();
            //foreach (var col in tableColumns)
            //{
            //    col.AdjustToContents();
            //}

            //table = GetNewTable();

            //iXLCell = totalAcresCell.CellBelow(10);
            //tableXl = iXLCell.InsertTable(table);
            //tableXl.Theme = XLTableTheme.TableStyleLight5;
            var row = 20;
             foreach (var theme in XLTableTheme.GetAllThemes())
            {
                table = GetNewTable();
                
                iXLCell = totalAcresCell.CellBelow(row);
                tableXl = iXLCell.InsertTable(table);
                tableXl.Theme = theme;
                iXLCell.CellRight(5).SetValue(theme.ToString());
                row += 10;
            }


            //table = GetNewTable();

            //iXLCell = totalAcresCell.CellBelow(20);
            //tableXl = iXLCell.InsertTable(table);
            //tableXl.Theme = XLTableTheme.TableStyleLight4;

            ws.FirstCell().WorksheetColumn().AdjustToContents();
            //ws.SheetView.FreezeRows(iXLCell.Address.RowNumber);

            //farmNameCell.Address;
            //ws.LastCellUsed().CellRight().Address;
            //var lastColumnAddress = ws.LastColumnUsed().LastCellUsed().Address;
            //var farmNameRange = ws.Range(farmNameCell.Address, ws.LastColumnUsed().Cell(farmNameCell.Address.RowNumber).Address);
            //farmNameRange.Merge();


            //var totalAcresCellRange = ws.Range(totalAcresCell.Address, ws.LastColumnUsed().Cell(totalAcresCell.Address.RowNumber).Address);
            //totalAcresCellRange.Merge();

            //var wsRange = ws.Range(ws.FirstCell().Address, ws.LastCellUsed().Address);
            //wsRange.Style.Border.InsideBorder = XLBorderStyleValues.Hair;
            //wsRange.Style.Border.OutsideBorder = XLBorderStyleValues.Hair;

            //var firstCell = ws.FirstCell().SetValue("This is Crop Plan from Decisive Farming Application");
            //firstCell.Style.Font.SetBold()
            //    .Fill.SetBackgroundColor(XLColor.CyanProcess);
            //ws.SheetView.FreezeRows(2);
            //var titleCell = firstCell.CellBelow().SetValue("Crop Type");
            //var totalCell = titleCell.CellBelow().SetValue("Barley")
            //.CellBelow().SetValue("Canola")
            //.CellBelow().SetValue("Wheat");
            //ws.FirstCell().WorksheetColumn().AdjustToContents();
            //titleCell.Select();
            //af.AppendChild(firstCell.CellBelow().SetValue("Crop Type"));
            //var fcu = ws.FirstCell().WorksheetRow().FirstCellUsed();
            //var a6 = ws.Cell("A6").SetValue(fcu.CurrentRegion.ToString());

            //ws.Range(titleCell, totalCell).SetAutoFilter().Sort();
            //ws.Range(titleCell, totalCell).SetAutoFilter();
            wb.SaveAs(memory);


            memory.Position = 0;

            return File(memory, MimeTypes.GetFileType()[".xlsx"], sFileName);
        }

        private static DataTable GetNewTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
            return table;
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
