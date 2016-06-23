using Aspose.Cells;

namespace Excel.Tools
{
    public class ExcelReader
    {
        public string[,] Excute(string file)
        {
            var workbook = new Workbook(file);

            Cells cells = workbook.Worksheets[0].Cells;
            string[,] result = new string[cells.MaxDataRow + 1, cells.MaxDataColumn + 1];

            for (int i = 0; i < cells.MaxDataRow + 1; i++)
            {
                for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                {
                    result[i, j] = cells[i, j].StringValue.Trim();
                }
            }

            return result;
        }
    }
}
