using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UploadDownloadFileASPDotNetCore
{
    public class CropPlanGridRotationDto
    {
        public int CropYear { get; set; }
        public int? CropTypeId { get; set; }
        public string CropTypeName { get; set; }
        public string CropVarietyName { get; set; }
        public float? YieldGoal { get; set; }
        public string YieldUnit { get; set; }

        public override string ToString()
        {
            var returnDesc = "None";
            StringBuilder str = new StringBuilder();
            if (CropTypeId != null)
                str.Append(CropTypeName);
            if (!String.IsNullOrWhiteSpace(CropVarietyName))
            {
                str.Append("-");
                str.Append(CropVarietyName);
            }
            if (str.Length > 0 && YieldGoal != null)
            {
                str.Append("-");
                str.Append(YieldGoal);
                str.Append(" ");
                str.Append(YieldUnit);
            }

            if (!string.IsNullOrEmpty(str.ToString()))
                returnDesc = str.ToString();
            return returnDesc;
        }
    }
}
