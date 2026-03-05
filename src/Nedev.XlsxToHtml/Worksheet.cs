namespace Nedev.XlsxToHtml
{
    public class Worksheet
    {
        public string Name { get; set; } = string.Empty;
        public int Index { get; set; }
        public List<Row> Rows { get; } = new List<Row>();
    }
}
