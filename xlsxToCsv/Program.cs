using System.Reflection;
using xlsxToCsv.Services;

/*
    xlsxToCsv
    
    Program converts generic xlsx files containing table-like data
    to simple csv file which is easier for fetching

    Aleksander Heese
    06.2023
*/ 


Console.WriteLine($" xlsxToCsv converter v. {Assembly.GetExecutingAssembly().GetName().Version}");
if (!ExplorerService.InitializationCheck())
    return;
