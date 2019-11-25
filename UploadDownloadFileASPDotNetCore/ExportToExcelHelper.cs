using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;


namespace UploadDownloadFileASPDotNetCore
{
    public class ExportToExcelHelper
    {
        public static IXLWorksheet AddWorksheet(XLWorkbook workbook, string sheetName, XLPageOrientation pageOrientation = XLPageOrientation.Landscape, int pageAdjustment = 80, XLPaperSize paperSize = XLPaperSize.LegalPaper)
        {
            var worksheet = workbook.AddWorksheet(sheetName);

            worksheet.PageSetup.PageOrientation = pageOrientation;
            worksheet.PageSetup.AdjustTo(pageAdjustment);
            worksheet.PageSetup.PaperSize = paperSize;

            return worksheet;
        }

        public static XLWorkbook CreateWorkbook()
        {
            //create workbook
            return new XLWorkbook();
        }
        public static IList<ExcelExportColumnAttribute> GetOrderedColumnNames(Type valueType)
        {
            var properties = valueType.GetProperties();
            IList<ExcelExportColumnAttribute> excelExportColumnAttributeList = new List<ExcelExportColumnAttribute>();

            foreach (var property in properties)
            {
                var excelInfo = GetExcelExportAttribute(property);
                if (excelInfo != null)
                {
                    excelExportColumnAttributeList.Add(excelInfo);
                }
            }

            return excelExportColumnAttributeList.OrderBy(a => a.ColumnOrder).ToList();

        }

        public static ExcelExportColumnAttribute GetExcelExportAttribute(PropertyInfo property)
        {
            return property.GetCustomAttributes<ExcelExportColumnAttribute>().FirstOrDefault();
        }



        public static void MergeRows(IXLWorksheet worksheet, IXLCell cell, int cellNumber)
        {
            worksheet.Range(cell, cell.CellRight(cellNumber)).Merge();
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

        public static void InsertTableGid(IXLWorksheet worksheet, List<IExportToExcelDto> gridDto, IXLCell tableGridStartCell, IList<ExcelExportColumnAttribute> columnAttibuteList, XLTableTheme xlTableTheme = null)
        {
            var table = new DataTable();
            var columnValueLengthMap = new Dictionary<int, int>();
            //table header
            foreach (var attr in columnAttibuteList)
            {
                table.Columns.Add(attr.ColumnName, attr.ColumnType);
            }

            //table rows           
            foreach (var dto in gridDto)
            {
                var dtoType = dto.GetType();
                var props = dtoType.GetProperties();
                var values = new Dictionary<int, object>();
                
                foreach (PropertyInfo prop in props)
                {
                    var excelProp = GetExcelExportAttribute(prop);
                    if (excelProp != null)
                    {
                        var propValue = prop.GetValue(dto, null);
                        if (propValue != null)
                        {
                            values.Add(excelProp.ColumnOrder, propValue);
                        }
                        else
                        {
                            if (excelProp.ColumnType == typeof(string))
                                values.Add(excelProp.ColumnOrder, "");
                            else
                                values.Add(excelProp.ColumnOrder, 0);

                            propValue = string.Empty;
                        }

                        MapColumnWidthByContent(columnValueLengthMap, excelProp, propValue);

                    }
                }
                var rowValues = values.OrderBy(a => a.Key).Select(c => c.Value).ToArray();
                var row = table.Rows.Add(rowValues);
            }
            var columnValueLengthArray = columnValueLengthMap.OrderBy(a => a.Key).ToArray();
            var gridTable = tableGridStartCell.InsertTable(table);

            //fomatting
            gridTable.Style.Font.SetFontSize(10);
            if (xlTableTheme != null)
                gridTable.Theme = xlTableTheme;
            else
                gridTable.Theme = XLTableTheme.None;

            var tableColumns = gridTable.Worksheet.ColumnsUsed();
            foreach (var col in tableColumns)
            {
                var columnCell = col.Cell(tableGridStartCell.Address.RowNumber);
                var columnValueLength = columnValueLengthArray[columnCell.Address.ColumnNumber - 1].Value;
                
                var columnName = columnCell.Value;
                var dataType = columnAttibuteList.FirstOrDefault(a => a.ColumnName.Equals(columnName));
                if (dataType != null)
                {
                    if (dataType.ColumnType == typeof(float))
                        col.Style.NumberFormat.Format = "###,###,##0.00";
                    else if (dataType.ColumnType == typeof(decimal))
                        col.Style.NumberFormat.Format = "$ ###,###,##0.00";

                    //adjust column width
                    var defaultColumnWidth = dataType.ColumnWidth;
                    if (columnValueLength > defaultColumnWidth)
                        defaultColumnWidth = columnValueLength;
                    col.Width = defaultColumnWidth;
                } 

                var tableHeaderRowStyle = columnCell.Style;
                if (xlTableTheme == null)
                    tableHeaderRowStyle.Fill.SetBackgroundColor(XLColor.LightGray);
            }
            worksheet.SheetView.FreezeRows(tableGridStartCell.Address.RowNumber);
        }

        private static void MapColumnWidthByContent(Dictionary<int, int> columnValueLengthMap, ExcelExportColumnAttribute excelProp, object propValue)
        {
            if (columnValueLengthMap.ContainsKey(excelProp.ColumnOrder) && propValue != null)
            {
                var propStrLength = propValue.ToString().Length;
                var columnWidth = columnValueLengthMap[excelProp.ColumnOrder];
                if(propStrLength > columnWidth)
                    columnWidth = propStrLength;
                columnValueLengthMap[excelProp.ColumnOrder] = columnWidth;
            }
            else
            {
                columnValueLengthMap.Add(excelProp.ColumnOrder, propValue.ToString().Length);
            }
        }

        public static byte[] GetByteArray(XLWorkbook workbook)
        {
            using (var memoryStream = new MemoryStream())
            {
                workbook.SaveAs(memoryStream);
                memoryStream.Position = 0;
                return memoryStream.ToArray();
            }

        }

        public void CropPlanWorksheetTableTest()
        {
            //var workbook = ExportToExcelHelper.CreateWorkbook();
            //var sheetName = "Crop Plans";
            //var worksheet = ExportToExcelHelper.AddWorksheet(workbook, sheetName);
            //var columnList = ExportToExcelHelper.GetOrderedColumnNames(typeof(CropPlanGridDto));
            //var tableGridStartCell = worksheet.FirstCell();

            //var cropPlanGridDtos = new List<CropPlanGridDto>() {
            //        new CropPlanGridDto() {
            //            BudgetId = null,
            //            CropPlanId = 86086,
            //            CropTypeId = null,
            //            CropTypeName = null,
            //            CropVariety = null,
            //            CropVarietyName = null,
            //            CropYear = 2019,
            //            FarmId = 2682,
            //            FarmableAcres = (float)202.4,
            //            FieldId = 12826,
            //            FieldName = "Field 1",
            //            GISDisplayValue = "W 6 28 25 W4",
            //            IsGISBoundary = true,
            //            MarketingPlan = null,
            //            ToleranceTypeId = null,
            //            ToleranceTypeName = null,
            //            YieldGoal = 0,
            //            YieldUnit = null
            //        },
            //        new CropPlanGridDto() {
            //                BudgetId    =   null,
            //                CropPlanId  =   86079,
            //                CropTypeId  =   null,
            //                CropTypeName    =   null,
            //                CropVariety =   null,
            //                CropVarietyName =   null,
            //                CropYear    =   2019,
            //                FarmId  =   2682,
            //                FarmableAcres   =   (float)497.6,
            //                FieldId =   12833,
            //                FieldName   =   "Field 8/9",
            //                GISDisplayValue =   "SC 32 27 25 W4",
            //                IsGISBoundary   =   true,
            //                MarketingPlan   =   null,
            //                ToleranceTypeId =   null,
            //                ToleranceTypeName   =   null,
            //                YieldGoal   =   0,
            //                YieldUnit   =   null,
            //        }
            //};
            //var cropYear = 2019;
            //var previousCropYear = cropYear - 1;
            //var rotationColumnList = columnList.Where(a => a.ColumnName == "Rotation").OrderBy(c => c.ColumnOrder);
            //foreach (var rotationCol in rotationColumnList)
            //{
            //    rotationCol.ColumnName = rotationCol.ColumnName + " " + previousCropYear;
            //    previousCropYear--;
            //}

            //var cropPlanGridDtosList = cropPlanGridDtos.ToList<IExportToExcelDto>();

            //ExportToExcelHelper.InsertTableGid(worksheet, cropPlanGridDtosList, tableGridStartCell, columnList);

            //Assert.Equal(sheetName, worksheet.Name);
            //Assert.Equal(3, worksheet.LastRowUsed().RowNumber());

        }


    }
}
