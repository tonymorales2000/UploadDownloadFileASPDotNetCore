using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace UploadDownloadFileASPDotNetCore
{
    [Route("api/[controller]")]
    public class ReportController : Controller
    {
        // GET: api/<controller>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "ok", "doki" ,"hey"};
        }

        [HttpGet("GetExcel"), Route("GetExcel")]
        [Produces("application/octet-stream")]
        public async Task<IActionResult> DownloadFile3(int id)
        {
            var memory = new MemoryStream();
            //string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"CropPlans.xlsx";
            //var wb = await BuildExcelFile(id);
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var firstCell = ws.FirstCell().SetValue("This is Crop Plan from Decisive Farming Application");
            ////firstCell.Style.Font.SetBold()
            ////    .Fill.SetBackgroundColor(XLColor.CyanProcess);
            ////ws.SheetView.FreezeRows(2);
            var titleCell = firstCell.CellBelow().SetValue("Crop Type");
            var totalCell = titleCell.CellBelow().SetValue("Barley")
            .CellBelow().SetValue("Canola")
            .CellBelow().SetValue("Wheat");
            //ws.FirstCell().WorksheetColumn().AdjustToContents();




            //ws.Range(titleCell, totalCell).SetAutoFilter().Sort();
            wb.SaveAs(memory);


            memory.Position = 0;

            return File(memory, MimeTypes.GetFileType()[".xlsx"], sFileName);
        }



        [HttpGet, Route("GetCropPlan")]
        [Produces("application/octet-stream")]
        public async Task<IActionResult> GetCropPlan()
        {
            var workbook = ExportToExcelHelper.CreateWorkbook();
            var sheetName = "Crop Plans";
            var worksheet = ExportToExcelHelper.AddWorksheet(workbook, sheetName);
            

            var columnList = ExportToExcelHelper.GetOrderedColumnNames(typeof(CropPlanGridDto));
            //var tableGridStartCell = worksheet.FirstCell();

            var cropPlanGridDtos = new List<CropPlanGridDto>() {
                    new CropPlanGridDto() {
                        BudgetId = null,
                        CropPlanId = 86086,
                        CropTypeId = null,
                        CropTypeName = "Alfa Alfa",
                        //CropVariety = null,
                        CropVarietyName = "Corn",
                        CropYear = 2019,
                        FarmId = 2682,
                        FarmableAcres = (float)202.4,
                        FieldId = 12826,
                        FieldName = "Field 1",
                        GISDisplayValue = "W 6 28 25 W4 : W 6 28 25 W4",
                        IsGISBoundary = true,
                        MarketingPlan = null,
                        ToleranceTypeId = null,
                        ToleranceTypeName = null,
                        YieldGoal = 0,
                        YieldUnit = null
                    },
                    new CropPlanGridDto() {
                            BudgetId    =   null,
                            CropPlanId  =   86079,
                            CropTypeId  =   null,
                            CropTypeName    =   "Canola",
                            //CropVariety =   null,
                            CropVarietyName =   "Sweet Wheat",
                            CropYear    =   2019,
                            FarmId  =   2682,
                            FarmableAcres   =   (float)497.6,
                            FieldId =   12833,
                            FieldName   =   "Field 8/9",
                            GISDisplayValue =   "SC 32 27 25 W4",
                            IsGISBoundary   =   true,
                            MarketingPlan   =   null,
                            ToleranceTypeId =   null,
                            ToleranceTypeName   =   null,
                            YieldGoal   =   0,
                            YieldUnit   =   null,
                    },
                    new CropPlanGridDto() {
                            BudgetId    =   null,
                            CropPlanId  =   88099,
                            CropTypeId  =   null,
                            CropTypeName    =   null,
                            //CropVariety =   null,
                            CropVarietyName =   null,
                            CropYear    =   2019,
                            FarmId  =   2682,
                            FarmableAcres   =   (float)497.6,
                            FieldId =   12833,
                            FieldName   =   "Field 87/990",
                            GISDisplayValue =   "SC 32 27 25 W4",
                            IsGISBoundary   =   true,
                            MarketingPlan   =   null,
                            ToleranceTypeId =   null,
                            ToleranceTypeName   =   null,
                            YieldGoal   =   0,
                            YieldUnit   =   null,
                    }
            };
            var cropYear = 2019;
            var previousCropYear = cropYear - 1;
            var rotationColumnList = columnList.Where(a => a.ColumnName == "Rotation").OrderBy(c => c.ColumnOrder);
            foreach (var rotationCol in rotationColumnList)
            {
                rotationCol.ColumnName = rotationCol.ColumnName + " " + previousCropYear;
                previousCropYear--;
            }

            var cropPlanGridDtosList = cropPlanGridDtos.ToList<IExportToExcelDto>();

            var totalAcres = cropPlanGridDtos.Where(a => a.FarmableAcres != null).Sum(a => a.FarmableAcres);
            var totalAcresStr = totalAcres.HasValue ? totalAcres.Value.ToString("#,##0.00") : "";
            var totalAcresDescription = string.Format($"Crop Plans {cropYear} (Total Acres: {totalAcresStr})");
            var firstCell = worksheet.FirstCell();
            ExportToExcelHelper.MergeRows(worksheet, firstCell, 3);

            //header values
            var farmNameCell = firstCell.SetValue("Tony Test Farm");
            farmNameCell.Style.Font.SetBold(true);
            farmNameCell.Style.Font.SetFontSize(12);


            var totalAcreRow = farmNameCell.CellBelow();
            ExportToExcelHelper.MergeRows(worksheet, totalAcreRow, 3);
            var totalAcresCell = totalAcreRow.SetValue(totalAcresDescription);
            totalAcresCell.Style.Font.SetFontSize(11);
            var tableGridStartCell = totalAcresCell.CellBelow();

            ExportToExcelHelper.InsertTableGid(worksheet, cropPlanGridDtosList, tableGridStartCell, columnList);

            var tableCellCount = columnList.Count - 1;

            ExportToExcelHelper.MergeRows(worksheet, farmNameCell, tableCellCount);
            ExportToExcelHelper.MergeRows(worksheet, totalAcresCell, tableCellCount);


            var fileInByteArray = ExportToExcelHelper.GetByteArray(workbook);

            //return File(fileInByteArray, MimeTypes.GetFileType()[".xlsx"], "Crop Plans 2019.xlsx");
            return File(fileInByteArray, MimeKit.MimeTypes.GetMimeType("Crop Plans 2019.xlsx"), "Crop Plans 2019.xlsx");
        }


        //// GET api/<controller>/5
        //[HttpGet("{id}")]
        //public string Get(int id)
        //{
        //    return "value";
        //}

        //// POST api/<controller>
        //[HttpPost]
        //public void Post([FromBody]string value)
        //{
        //}

        //// PUT api/<controller>/5
        //[HttpPut("{id}")]
        //public void Put(int id, [FromBody]string value)
        //{
        //}

        //// DELETE api/<controller>/5
        //[HttpDelete("{id}")]
        //public void Delete(int id)
        //{
        //}
    }
}
