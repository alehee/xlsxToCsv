namespace xlsxToCsv.Services
{
    public class ExplorerService
    {
        const string InputPath = "in";
        const string OutputPath = "out";

        public static bool InitializationCheck()
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
                Console.WriteLine(" Input directory was created. Drop .xlsx files to the `in` folder!");
            }
            else if (Directory.GetFiles(InputPath).Length == 0)
            {
                isOk = false;
                Console.WriteLine(" The `in` directory contains 0 files. Fill the directory with .xlsx files you want to convert!");
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
    }
}
