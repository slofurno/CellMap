using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
using System.ComponentModel;

namespace CellMap
{
  public class ExcelReader
  {

    public ExcelReader()
    {

    }

    public IEnumerable<T> Read<T>(string fileName) where T : new()
    {

      using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
      using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
      {
        WorkbookPart workbookPart = doc.WorkbookPart;
        SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
        SharedStringTable sst = sstpart.SharedStringTable;

        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
        Worksheet sheet = worksheetPart.Worksheet;

        var cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ToArray();
        var numberingFormats = workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats;

        //var cells = sheet.Descendants<Cell>();
        var rows = sheet.Descendants<Row>();


        foreach (Row row in rows)
        {
          var cells = new Dictionary<int, string>();

          foreach (Cell c in row.Elements<Cell>())
          {

            var rowcol = c.CellReference;
            var rowchars = ParseChars(rowcol);
            var rowindex = AlphaToInt(rowchars);


            if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
            {
              int ssid = int.Parse(c.CellValue.Text);
              string str = sst.ChildElements[ssid].InnerText;

              cells.Add(rowindex, str);
            }
            else if (c.CellValue != null)
            {

              cells.Add(rowindex, c.CellValue.Text);

            }

          }

          yield return BuildModel<T>(cells);
        }
      }
    }

    private string ParseChars(string input)
    {

      var result = new List<char>();
      int A = (int)'A';
      int Z = (int)'Z';

      for (var i = 0; i < input.Length; ++i)
      {
        var next = (int)input[i];

        if (!(next > Z || next < A))
        {
          result.Add((char)next);
        }
      }

      return new string(result.ToArray());
    }

    private int AlphaToInt(string alpha)
    {

      int A = (int)'A';
      int total = 0;

      for (var i = 0; i < alpha.Length; ++i)
      {
        var next = (int)alpha[i] - (A - 1);
        total = total * 26 + next;
      }

      return total;
    }

    public T BuildModel<T>(IDictionary<int, string> args) where T : new()
    {

      Type type = typeof(T);

      var result = new T();

      var properties = type.GetProperties().OrderBy(x => x.GetCustomAttribute<ExcelOrderAttribute>().ExcelOrder).ToArray();

      for (var i = 0; i < properties.Length; ++i)
      {

        var index = properties[i].GetCustomAttribute<ExcelOrderAttribute>().ExcelOrder;
        var rawvalue = args[index];

        dynamic castvalue;

        if (properties[i].PropertyType == typeof(DateTime))
        {

          var days = double.Parse(rawvalue);

          var d = new DateTime(1899, 12, 30);

          castvalue = d.AddDays(days);

        }
        else if (properties[i].PropertyType == typeof(TimeSpan))
        {
          var days = double.Parse(rawvalue);
          var d = new DateTime(1899, 12, 30);
          var time = d.AddDays(days);

          TimeSpan span = time - d;

          castvalue = span;

        }
        else if (properties[i].PropertyType.IsEnum)
        {
          castvalue = Enum.Parse(properties[i].PropertyType, rawvalue);
        }
        else
        {
          castvalue = Convert.ChangeType(rawvalue, properties[i].PropertyType);
        }
        properties[i].SetValue(result, castvalue, null);

      }
      return result;
    }


  }


  [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field,
  Inherited = true, AllowMultiple = false)]
  [ImmutableObject(true)]
  public sealed class ExcelOrderAttribute : Attribute
  {
    private readonly int order;
    public int ExcelOrder { get { return order; } }
    public ExcelOrderAttribute(int order) { this.order = order; }
  }
}
