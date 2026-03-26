using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;

namespace Vi2Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: Vi2Processor.exe <start_dt> <end_dt> <file.vi2>");
                return;
            }

            string startStr = args[0]; // e.g., "2026-03-16 09:00:00"
            string endStr = args[1];
            string inputFile = args[2];
            string format = "yyyy-MM-dd HH:mm:ss";

            try
            {
                DateTime targetStart = DateTime.ParseExact(startStr, format, CultureInfo.InvariantCulture);
                DateTime targetEnd = DateTime.ParseExact(endStr, format, CultureInfo.InvariantCulture);

                using (var handler = new TestoVi2Handler())
                {
                    Console.WriteLine("[*] Reading Original File...");
                    List<MeasurementPoint> allPoints = handler.Read(inputFile);

                    // --- THE TIMEZONE FIX ---
                    // The CSV shows 7:00 when you want 9:00. 
                    // This means we need to find the point that is 2 hours BEFORE your target.
                    TimeSpan localOffset = TimeZoneInfo.Local.GetUtcOffset(targetStart);
                    DateTime utcSearchStart = targetStart.Subtract(localOffset);
                    DateTime utcSearchEnd = targetEnd.Subtract(localOffset);

                    Console.WriteLine($"[!] Applying Timezone Correction: -{localOffset.TotalHours} hours");
                    Console.WriteLine($"[!] Searching for Raw UTC Time: {utcSearchStart}");

                    // Find the index of the point that corresponds to the UTC time
                    var firstMatch = allPoints.FirstOrDefault(p => p.Timestamp >= utcSearchStart);
                    if (firstMatch == null) 
                    {
                        Console.WriteLine("[-] Could not find the requested range in the file.");
                        return;
                    }

                    int startIndex = allPoints.IndexOf(firstMatch);
                    var filteredData = allPoints
                        .Where(p => p.Timestamp >= utcSearchStart && p.Timestamp <= utcSearchEnd)
                        .ToList();

                    // --- GENERATE FILES ---
                    string filteredVi2 = "Filtered_" + Path.GetFileName(inputFile);
                    
                    Console.WriteLine($"[*] Writing Filtered Binary at Index: {startIndex}");
                    handler.WriteBinary(inputFile, filteredVi2, filteredData, startIndex);

                    // Optional Debug CSVs
                    handler.ExportToCsv("Debug_Filtered.csv", filteredData);

                    Console.WriteLine($"\n[SUCCESS] Start point (UTC {utcSearchStart}) mapped to Index {startIndex}.");
                    Console.WriteLine($"[SUCCESS] Software should now show local start as: {targetStart}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n[ERROR] " + ex.Message);
            }
        }
    }
}