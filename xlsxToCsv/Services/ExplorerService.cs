namespace xlsxToCsv.Services
{
    public class ExplorerService
    {
        public static bool InitializationCheck(string InputPath, string OutputPath)
        {
            bool isOk = true;

            // Input folder fetching
            if (!Directory.Exists(InputPath))
            {
                isOk = false;
                try
                {
                    Directory.CreateDirectory(InputPath);
                } catch
                {
                    Console.WriteLine(" An error occured while creating directory. Check your privileges!");
                    return false;
                }
                Console.WriteLine($" Input directory was created. Drop .xlsx files to the `{InputPath}` folder!");
            }
            else if (GetFilesToConvert(InputPath).Count == 0)
            {
                isOk = false;
                Console.WriteLine($" There's no valid files for convertion! Drop the files to convert in the `{InputPath}` folder!");
            }

            // Output folder creating
            if (!Directory.Exists(OutputPath))
            {
                try
                {
                    Directory.CreateDirectory(OutputPath);
                }
                catch
                {
                    Console.WriteLine(" An error occured while creating directory. Check your privileges!");
                    return false;
                }
            }

            return isOk;
        }

        public static List<string> GetFilesToConvert(string InputPath)
        {
            var files = new List<string>();

            foreach (var file in Directory.GetFiles(InputPath).ToList())
            {
                if (Path.GetExtension(file) == "xlsx")
                    files.Add(file);
            }

            return files;
        }
    }
}
