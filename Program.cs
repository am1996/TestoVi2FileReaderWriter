using System;

namespace Vi2Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            try 
            {
                using (var handler = new TestoVi2Handler())
                {
                    // 1. Read existing data
                    Console.WriteLine("Reading data...");
                    var data = handler.Read("input.vi2");
                    Console.WriteLine($"Found {data.Count} points.");

                    // 2. Modify data (optional)
                    data[0].Timestamp = DateTime.Now; 

                    // 3. Write to a new file
                    Console.WriteLine("Writing new file...");
                    handler.Write("output.vi2", data, "MyCustomSession");
                    Console.WriteLine("Success!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + (ex.InnerException?.Message ?? ex.Message));
            }
        }
    }
}