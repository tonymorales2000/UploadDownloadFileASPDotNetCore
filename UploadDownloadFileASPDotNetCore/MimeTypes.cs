using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace UploadDownloadFileASPDotNetCore
{
    public class MimeTypes
    {
        public static Dictionary<string, string> GetFileType()
        {
            return new Dictionary<string, string>
            {
                {".csv", "text/csv" },
                {".doc", "application/msword" },
                {".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                {".pdf", "application/pdf" },
                {".png", "image/png" },
                {".txt", "text/plain" },
                {".xls", "	application/vnd.ms-excel" },
                {".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                {".zip", "application/zip" },
            };
        }
    }
}
