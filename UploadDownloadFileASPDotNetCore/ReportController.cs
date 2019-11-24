﻿using System;
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
            //var firstCell = ws.FirstCell().SetValue("This is Crop Plan from Decisive Farming Application");
            ////firstCell.Style.Font.SetBold()
            ////    .Fill.SetBackgroundColor(XLColor.CyanProcess);
            ////ws.SheetView.FreezeRows(2);
            //var titleCell = firstCell.CellBelow().SetValue("Crop Type");
            //var totalCell = titleCell.CellBelow().SetValue("Barley")
            //.CellBelow().SetValue("Canola")
            //.CellBelow().SetValue("Wheat");
            //ws.FirstCell().WorksheetColumn().AdjustToContents();




            //ws.Range(titleCell, totalCell).SetAutoFilter().Sort();
            wb.SaveAs(memory);


            memory.Position = 0;

            return File(memory, MimeTypes.GetFileType()[".xlsx"], sFileName);
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