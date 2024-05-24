namespace OpakSoftware.Models
{
    public class UploadedDataViewModel
    {
        public List<string>  ColumnNames { get; set; }
        public List<List<string>> ? Rows { get; set; }
    }
}
