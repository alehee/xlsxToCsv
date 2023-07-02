namespace xlsxToCsv.Services
{
    class ConverterService
    {
        public static void ConvertFiles(string[] files, string outputPath)
        {
            foreach (string file in files)
            {
                try
                {
                    var sheets = GetSheetsData(file);
                } catch (Exception e)
                {
                    Console.WriteLine($" An error occured while parsing `{file}`. Ignoring this file.");
                }
            }
        }

        private static Dictionary<string, List<List<object>>> GetSheetsData(string filePath)
        {
            var dict = new Dictionary<string, List<List<object>>>();



            return dict;
        }
    }
}
