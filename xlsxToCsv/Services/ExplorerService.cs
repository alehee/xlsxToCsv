namespace xlsxToCsv.Services
{
    public class ExplorerService
    {
        public static bool InitializationCheck(string inputPath, string outputPath)
        {
            bool isOk = true;

            // Input folder fetching
            if (!Directory.Exists(inputPath))
            {
                isOk = false;
                try
                {
                    Directory.CreateDirectory(inputPath);
                } catch
                {
                    Console.WriteLine(" An error occured while creating directory. Check your privileges!");
                    return false;
                }
                Console.WriteLine($" Input directory was created. Drop .xlsx files to the `{inputPath}` folder!");
            }
            else if (GetFilesToConvert(inputPath).Count == 0)
            {
                isOk = false;
                Console.WriteLine($" There's no valid files for convertion! Drop the files to convert in the `{inputPath}` folder!");
            }

            // Output folder creating
            if (!Directory.Exists(outputPath))
            {
                try
                {
                    Directory.CreateDirectory(outputPath);
                }
                catch
                {
                    Console.WriteLine(" An error occured while creating directory. Check your privileges!");
                    return false;
                }
            }

            return isOk;
        }

        public static List<string> GetFilesToConvert(string inputPath)
        {
            var files = new List<string>();

            foreach (var file in Directory.GetFiles(inputPath).ToList())
            {
                if (Path.GetExtension(file) == ".xlsx")
                    files.Add(file);
            }

            return files;
        }
    }
}
