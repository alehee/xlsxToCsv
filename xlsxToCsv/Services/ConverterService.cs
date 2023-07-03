using ClosedXML.Excel;
using System.Diagnostics;

namespace xlsxToCsv.Services
{
    class ConverterService
    {
        public static void ConvertFiles(string[] files, string inputPath, string outputPath)
        {
            foreach (string file in files)
            {
                try
                {
                    Console.WriteLine($" File `{file}`");
                    Console.WriteLine("  Acquiring data");
                    var sheets = GetSheetsData(file);

                    Console.WriteLine("  Outputing data to file");
                    string fileBase = file.Replace(".xlsx", "").Replace($"{inputPath}\\", "");
                    if (sheets.Keys.Count > 1 )
                    {
                        foreach( var sheet in sheets.Keys )
                        {
                            SaveToFile($"{outputPath}/{fileBase}_{sheet}.csv", sheets[sheet]);
                        }
                    }
                    else
                        SaveToFile($"{outputPath}/{fileBase}.csv", sheets[sheets.Keys.ToArray()[0]]);

                } catch (Exception e)
                {
                    Console.WriteLine($" An error occured while parsing `{file}`. Ignoring this file.");
                    Debug.WriteLine(e.ToString());
                }
            }
        }

        private static Dictionary<string, List<List<object>>> GetSheetsData(string filePath)
        {
            var dict = new Dictionary<string, List<List<object>>>();

            var workbook = new XLWorkbook(filePath);
            foreach (var worksheet in workbook.Worksheets)
            {
                var data = new List<List<object>>();

                for (int r = worksheet.FirstRowUsed().RowNumber(); r <= worksheet.LastRowUsed().RowNumber(); r++)
                {
                    List<object> row = new List<object>();
                    for (int c = worksheet.FirstColumnUsed().ColumnNumber(); c <=  worksheet.LastColumnUsed().ColumnNumber(); c++)
                    {
                        row.Add(worksheet.Cell(r, c).Value);
                    }
                    data.Add(row);
                }

                dict[worksheet.Name] = data;
            }

            return dict;
        }

        private static void SaveToFile(string filePath, List<List<object>> data)
        {
            var lines = new List<string>();

            foreach (var row in data)
            {
                var line = "";
                foreach (var col in row)
                {
                    if (line != "")
                        line += ",";

                    line += col.ToString().Trim();
                }
                lines.Add(line);
            }

            File.WriteAllLines(filePath, lines);
        }
    }
}
