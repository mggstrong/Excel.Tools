using Aspose.Cells;

namespace Excel.Tools
{
    public class ExcelWriter
    {
        public byte[] Excute(string[,] data)
        {
            var workbook = new Workbook();

            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    cells[i, j].PutValue(data[i, j]);
                }
            }

            using (var s = workbook.SaveToStream())
            {
                return s.ToArray();
            }
        }

        public byte[] Excute(string templatePath, string[] data)
        {
            var workbook = new Workbook(templatePath);

            WorkbookDesigner wd = new WorkbookDesigner(workbook);
            wd.SetDataSource("x", data);
            wd.Process();
            using (var s = workbook.SaveToStream())
            {
                return s.ToArray();
            }
        }
    }
}
