using System;

namespace UploadDownloadFileASPDotNetCore
{
    [AttributeUsage(AttributeTargets.All)]
    public class ExcelExportColumnAttribute : Attribute
    {
        public string ColumnName { get; set; }
        public int ColumnOrder { get; }
        public Type ColumnType { get; set; }
        public ExcelExportColumnAttribute(string columnName, int columnOrder, Type columnType)
        {
            ColumnName = columnName;
            ColumnOrder = columnOrder;
            ColumnType = columnType;
        }


    }
}