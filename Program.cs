using System;
using System.Collections.Generic;

namespace Vi2Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Usage: .\Vi2Converter.exe "path_to_your_file.vi2"
            if (args.Length == 0)
            {
                Console.WriteLine("Error: Please provide a .vi2 file path.");
                return;
            }

            string filePath = args[0];

            try
            {
                using (var handler = new TestoVi2Handler())
                {
                    Console.WriteLine($"Reading: {filePath}...");
                    List<MeasurementPoint> data = handler.Read(filePath);

                    Console.WriteLine($"--- Data Extracted ({data.Count} points) ---");
                    foreach (var point in data)
                    {
                        // Outputs: [Timestamp]: Value1, Value2...
                        Console.WriteLine($"{point.Timestamp:yyyy-MM-dd HH:mm:ss}: {string.Join(", ", point.Values)}");
                    }
                    Console.WriteLine("--- End of Data ---");
                }
            }
            catch (Exception ex)
            {
                // InnerException often contains the real COM error
                Console.WriteLine("Read Test Failed: " + (ex.InnerException?.Message ?? ex.Message));
            }
        }
    }
}
