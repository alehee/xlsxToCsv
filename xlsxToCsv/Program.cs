using System.Reflection;
using xlsxToCsv.Services;

/*
    xlsxToCsv
    
    Program converts generic xlsx files containing table-like data
    to simple csv file which is easier for fetching

    Aleksander Heese
    07.2023
*/


const string InputPath = "in";
const string OutputPath = "out";

Console.WriteLine($" xlsxToCsv converter v. {Assembly.GetExecutingAssembly().GetName().Version}");
if (!ExplorerService.InitializationCheck(InputPath, OutputPath))
    return;

Console.WriteLine(" Starting the conversion");
ConverterService.ConvertFiles(ExplorerService.GetFilesToConvert(InputPath).ToArray(), InputPath, OutputPath);

Console.WriteLine(" See you soon! :)");