using System;
using System.IO;
using Nedev.XlsxToHtml;

if (args.Length < 1 || args.Length > 2)
{
    Console.WriteLine("Usage: dotnet run --project src/Nedev.XlsxToHtml.Cli -- <input.xlsx> [output.html]");
    return;
}

string input = args[0];
string output = args.Length == 2 ? args[1] : null;

if (!File.Exists(input))
{
    Console.Error.WriteLine($"Input file does not exist: {input}");
    Environment.Exit(1);
}

try
{
    var sw = System.Diagnostics.Stopwatch.StartNew();
    var reader = new XlsxReader();
    var workbook = reader.Read(input);
    var writer = new HtmlWriter();
    if (string.IsNullOrEmpty(output))
    {
        // write to stdout
        writer.Write(workbook, Console.Out);
    }
    else
    {
        using var fs = new StreamWriter(output, false, System.Text.Encoding.UTF8);
        writer.Write(workbook, fs);
    }
    sw.Stop();
    Console.Error.WriteLine($"Converted in {sw.ElapsedMilliseconds} ms");
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    Environment.Exit(2);
}
