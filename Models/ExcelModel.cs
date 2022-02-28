using System.Collections.Generic;
using System.Linq;

public class ExcelModel
{
    public List<ExcelRow> Rows { get; set; }

    public ExcelColumn[] Columns
    {
        get
        {
            if (Rows == null || Rows.Count() == 0)
            {
                return null;
            }

            return Rows.First().Columns.ToArray();
        }
    }

    public ExcelModel()
    {
        Rows = new List<ExcelRow>();
    }
}

public class ExcelColumn
{
    public string Name { get; set; } = string.Empty;

    public string Value { get; set; } = string.Empty;
}

public class ExcelRow
{
    public List<ExcelColumn> Columns { get; set; }

    public ExcelRow()
    {
        Columns = new List<ExcelColumn>();
    }
}