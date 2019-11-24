using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace UploadDownloadFileASPDotNetCore
{
    public class CropPlanGridDto : IExportToExcelDto
    {
        public int CropPlanId { get; set; }
        [ExcelExportColumn("Field Id", 0, typeof(int))]
        public int FieldId { get; set; }
        [ExcelExportColumn("Field Name", 1, typeof(string))]
        public string FieldName { get; set; }
        public int CropYear { get; set; }
        public int? CropTypeId { get; set; }
        [ExcelExportColumn("Crop Type", 4, typeof(string))]
        public string CropTypeName { get; set; }
        [ExcelExportColumn("Variety", 5, typeof(string))]
        public string CropVarietyName { get; set; }
        [ExcelExportColumn("Farmable\r\nAcres", 3, typeof(float))]
        public float? FarmableAcres { get; set; }
        [ExcelExportColumn("Agronomic\r\nYield", 9, typeof(float))]
        public float? YieldGoal { get; set; }
        [ExcelExportColumn("Yield\r\nUnit", 10, typeof(string))]
        public string YieldUnit { get; set; }
        public int? ToleranceTypeId { get; set; }
        //[ExcelExportColumn("Budget", 7, typeof(int))]
        public int? BudgetId { get; set; } // to be decided
        //[ExcelExportColumn("Marketing Plan", 8, typeof(string))]
        public string MarketingPlan { get; set; } // to be decided as not mapped to any Crop Plans at this point.
        [ExcelExportColumn("Tolerance", 6, typeof(string))]
        public string ToleranceTypeName { get; set; }
        public IList<CropPlanGridRotationDto> Rotations { get; set; }
        public bool IsGISBoundary { get; set; }
        [ExcelExportColumn("Legal Description", 2, typeof(string))]
        public IEnumerable<char> GISDisplayValue { get; set; }
        public int FarmId { get; set; }
        //public CropVarieties CropVariety { get; set; }

        [ExcelExportColumn("Rotation", 11, typeof(string))]
        public virtual string CropRotationPreviousYear
        {
            get
            {
                return Rotations.FirstOrDefault() != null ?
                    Rotations.FirstOrDefault().ToString() : "None";
            }
        }

        [ExcelExportColumn("Rotation", 12, typeof(string))]
        public virtual string CropRotationPreviousYearSecond
        {
            get
            {
                return Rotations.LastOrDefault() != null ?
                    Rotations.LastOrDefault().ToString() : "None";
            }
        }
        public CropPlanGridDto()
        {
            Rotations = new List<CropPlanGridRotationDto>();
        }
    }
}
