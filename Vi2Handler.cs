
using System.Text;
using System.Globalization;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using OpenMcdf;

namespace Vi2Converter
{
    public class MeasurementPoint
    {
        public DateTime Timestamp { get; set; }
        public double[] Values { get; set; }
    }

    public class TestoVi2Handler : IDisposable
    {
        private object _loader;
        private Type _loaderType;
        private PropertyInfo _protocolsProp;
        private FieldInfo _protocolsField;
        private MethodInfo _loadMethod;
        private MethodInfo _saveMethod;
        private bool _disposed;
        public TestoVi2Handler(string modelDll = "TestoModellVI2.dll", string interopDll = "Interop.Tcddka.dll")
        {
            // Keep your existing reflection setup here, 
            // but wrap it in a try-catch if you want to allow 
            // the class to exist even if the DLLs are grumpy.
            try {
                Assembly.LoadFrom(interopDll);
                Assembly testoAssembly = Assembly.LoadFrom(modelDll);
                _loaderType = testoAssembly.GetType("Testo.Model.TestVI2Loader");
                _loader = Activator.CreateInstance(_loaderType);
                
                // Initialize your MethodInfos here as you did before...
                _loadMethod = _loaderType.GetMethod("Load", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                _protocolsProp = _loaderType.GetProperty("Protocols", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            } catch (Exception ex) {
                Console.WriteLine("Note: Testo DLLs failed to init. Read() may fail, but WriteBinary() might still work.");
            }
        }

        private object GetProtocol()
        {
            object mProtocols = _protocolsField.GetValue(_loader);
            if (mProtocols is IEnumerable en)
                foreach (object p in en) return p;
            throw new InvalidOperationException("No protocol found in mProtocols.");
        }

        // ── READ ──────────────────────────────────────────────────────────────

        public List<MeasurementPoint> Read(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"VI2 file not found: {filePath}");

            _loadMethod.Invoke(_loader, new object[] { filePath });

            var results   = new List<MeasurementPoint>();
            var protocols = (IEnumerable)_protocolsProp.GetValue(_loader);

            foreach (object p in protocols)
            {
                int   rowCount = Convert.ToInt32(Interaction.CallByName(p, "NumRows", CallType.Get));
                short colCount = Convert.ToInt16(Interaction.CallByName(p, "NumCols", CallType.Get));

                var dateArray = (DateTime[])Interaction.CallByName(p, "DateArray", CallType.Get);
                var valArray  = (double[])  Interaction.CallByName(p, "ValArray",  CallType.Get);

                for (int r = 0; r < rowCount; r++)
                {
                    var point = new MeasurementPoint
                    {
                        Timestamp = dateArray[r],
                        Values    = new double[colCount]
                    };
                    for (short c = 0; c < colCount; c++)
                        point.Values[c] = valArray[r * colCount + c];

                    results.Add(point);
                }
            }
            return results;
        }

        // ── WRITE ────────────────────────────────────────────────────────────

        public void Write(string filePath, List<MeasurementPoint> data)
        {
            if (data == null || data.Count == 0)
                throw new ArgumentException("No data to write.", nameof(data));

            int   rowCount = data.Count;
            short colCount = (short)data[0].Values.Length;

            // Get mVi2File — Save() reads from this, not from mProtocols
            FieldInfo fileField = _loaderType.GetField("mVi2File",
                BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            object vi2File = fileField.GetValue(_loader);

            // Get the live protocol directly from vi2File via its COM enumerator
            // mProtocols is a disconnected .NET list — writes to it are ignored by Save()
            object protocol = null;
            try
            {
                // vi2File exposes _NewEnum for COM enumeration
                var newEnumMethod = vi2File.GetType().GetMethod("get__NewEnum",
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (newEnumMethod != null)
                {
                    object rawEnum = newEnumMethod.Invoke(vi2File, null);
                    if (rawEnum is IEnumerable en)
                        foreach (object p in en) { protocol = p; break; }
                }
            }
            catch { }

            // Fallback: try Item(1) then Item(0) on vi2File
            if (protocol == null)
            {
                var itemMethod = vi2File.GetType().GetMethod("Item",
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (itemMethod != null)
                    foreach (int idx in new[] { 1, 0, 2 })
                        try
                        {
                            protocol = itemMethod.Invoke(vi2File, new object[] { idx });
                            if (protocol != null) break;
                        }
                        catch { }
            }

            // Last fallback: use mProtocols (may not affect Save output)
            if (protocol == null)
                protocol = GetProtocol();

            if (protocol == null)
                throw new InvalidOperationException("Could not retrieve a writable protocol.");

            // Build arrays
            double[] oleDates = new double[rowCount];
            double[] valFlat  = new double[rowCount * colCount];

            for (int i = 0; i < rowCount; i++)
            {
                oleDates[i] = data[i].Timestamp.ToOADate();
                for (int j = 0; j < colCount; j++)
                    valFlat[i * colCount + j] = data[i].Values[j];
            }

            // Write via SetData on the protocol
            var setDataMethod = protocol.GetType().GetMethod("SetData",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (setDataMethod != null)
            {
                setDataMethod.Invoke(protocol, new object[] { "DateArray", oleDates });
                setDataMethod.Invoke(protocol, new object[] { "ValArray",  valFlat  });
            }

            // Verify
            int newRows = Convert.ToInt32(Interaction.CallByName(protocol, "NumRows", CallType.Get));
            if (newRows != rowCount)
                throw new InvalidOperationException(
                    $"Write verification failed: NumRows={newRows}, expected {rowCount}. " +
                    $"The protocol object may be disconnected from vi2File.");

            _saveMethod.Invoke(_loader, new object[] { filePath });
        }

        public static void DumpCfbStructure(string filePath)
        {
            var cf = new CompoundFile(filePath);
            Console.WriteLine("=== CFB Structure ===");
            cf.RootStorage.VisitEntries(item =>
            {
                Console.WriteLine($"  [{(item.IsStorage ? "Storage" : "Stream")}] " +
                                $"'{item.Name}' size={item.Size}");
            }, recursive: true);
            cf.Close();
        }
        public static void DumpCfbStreams(string filePath)
        {
            var cf = new OpenMcdf.CompoundFile(filePath);

            // Dump everything we can find by visiting all streams
            cf.RootStorage.VisitEntries(item =>
            {
                if (!item.IsStorage)
                {
                    try
                    {
                        // Try to get this stream from root
                        byte[] bytes = null;
                        try { bytes = cf.RootStorage.GetStream(item.Name).GetData(); } catch { }

                        if (bytes == null) return;

                        Console.WriteLine($"\n=== Stream '{item.Name}' ({bytes.Length} bytes) ===");
                        int previewLen = Math.Min(64, bytes.Length);
                        byte[] preview = new byte[previewLen];
                        Array.Copy(bytes, preview, previewLen);
                        Console.WriteLine($"Hex: {BitConverter.ToString(preview)}");

                        // Try as Int32s
                        Console.Write("Int32s: ");
                        for (int i = 0; i + 3 < Math.Min(32, bytes.Length); i += 4)
                            Console.Write($"{BitConverter.ToInt32(bytes, i)} ");
                        Console.WriteLine();

                        // Try as doubles
                        Console.Write("Doubles: ");
                        for (int i = 0; i + 7 < Math.Min(64, bytes.Length); i += 8)
                            Console.Write($"{BitConverter.ToDouble(bytes, i):F4} ");
                        Console.WriteLine();
                    }
                    catch (Exception ex) { Console.WriteLine($"  Error reading '{item.Name}': {ex.Message}"); }
                }
            }, recursive: false);

            // Also try the numbered storage
            cf.RootStorage.VisitEntries(item =>
            {
                if (item.IsStorage)
                {
                    Console.WriteLine($"\n=== Storage '{item.Name}' ===");
                    try
                    {
                        var storage = cf.RootStorage.GetStorage(item.Name);
                        storage.VisitEntries(child =>
                        {
                            Console.WriteLine($"  [{(child.IsStorage ? "Storage" : "Stream")}] '{child.Name}' size={child.Size}");
                            if (!child.IsStorage)
                            {
                                try
                                {
                                    byte[] bytes = storage.GetStream(child.Name).GetData();
                                    int previewLen = Math.Min(64, bytes.Length);
                                    byte[] preview = new byte[previewLen];
                                    Array.Copy(bytes, preview, previewLen);
                                    Console.WriteLine($"    Hex: {BitConverter.ToString(preview)}");
                                    Console.Write("    Int32s: ");
                                    for (int i = 0; i + 3 < Math.Min(32, bytes.Length); i += 4)
                                        Console.Write($"{BitConverter.ToInt32(bytes, i)} ");
                                    Console.WriteLine();
                                    Console.Write("    Doubles: ");
                                    for (int i = 0; i + 7 < Math.Min(64, bytes.Length); i += 8)
                                        Console.Write($"{BitConverter.ToDouble(bytes, i):F4} ");
                                    Console.WriteLine();
                                }
                                catch (Exception ex) { Console.WriteLine($"    Error: {ex.Message}"); }
                            }
                        }, recursive: false);
                    }
                    catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }
                }
            }, recursive: false);

            cf.Close();
        }
        public static void DecodeValues(string filePath)
        {
            var cf = new OpenMcdf.CompoundFile(filePath);

            string numberedStorageName = null;
            cf.RootStorage.VisitEntries(item =>
            {
                if (item.IsStorage && item.Name != "channels")
                    numberedStorageName = item.Name;
            }, recursive: false);

            var root        = cf.RootStorage.GetStorage(numberedStorageName);
            var dataStorage = root.GetStorage("data");

            byte[] val = dataStorage.GetStream("values").GetData();
            byte[] sum = root.GetStream("summary").GetData();

            // Known: row0 = 3/16/2026 9:00:00 AM, interval = 300s
            DateTime knownDt   = new DateTime(2026, 3, 16, 9, 0, 0, DateTimeKind.Local);
            long     knownUnix = ((DateTimeOffset)knownDt).ToUnixTimeSeconds();

            // The 4 bytes of row0 timestamp
            byte[] tsBytes = new byte[4];
            Array.Copy(val, 0, tsBytes, 0, 4);

            // Try every possible interpretation of the 4-byte timestamp
            float  asFloat     = BitConverter.ToSingle(tsBytes, 0);
            int    asInt32     = BitConverter.ToInt32(tsBytes, 0);
            uint   asUint32    = (uint)asInt32;

            // Big endian versions
            Array.Reverse(tsBytes);
            float  asFloatBE   = BitConverter.ToSingle(tsBytes, 0);
            int    asInt32BE   = BitConverter.ToInt32(tsBytes, 0);
            uint   asUint32BE  = (uint)asInt32BE;
            Array.Reverse(tsBytes); // restore

            Console.WriteLine($"Known Unix seconds: {knownUnix}");
            Console.WriteLine($"Row0 bytes: {BitConverter.ToString(tsBytes)}");
            Console.WriteLine($"  as float LE:   {asFloat}");
            Console.WriteLine($"  as int32 LE:   {asInt32}");
            Console.WriteLine($"  as uint32 LE:  {asUint32}");
            Console.WriteLine($"  as float BE:   {asFloatBE}");
            Console.WriteLine($"  as int32 BE:   {asInt32BE}");
            Console.WriteLine($"  as uint32 BE:  {asUint32BE}");

            // Check what offset makes uint32 LE match unix
            long offsetLE = (long)asUint32 - knownUnix;
            long offsetBE = (long)asUint32BE - knownUnix;
            Console.WriteLine($"\n  Offset if uint32 LE is unix+offset: {offsetLE}");
            Console.WriteLine($"  Offset if uint32 BE is unix+offset: {offsetBE}");

            // Row1 bytes
            byte[] ts1Bytes = new byte[4];
            Array.Copy(val, 12, ts1Bytes, 0, 4);
            uint row1Uint = (uint)BitConverter.ToInt32(ts1Bytes, 0);
            Console.WriteLine($"\nRow1 uint32 LE: {row1Uint}");
            Console.WriteLine($"Row1-Row0 diff: {(long)row1Uint - asUint32}");
            Console.WriteLine($"Expected diff for 300s interval: 300");

            // Try summary[0] as unix+offset
            uint sumUint = (uint)BitConverter.ToInt32(sum, 0);
            Console.WriteLine($"\nSummary[0] uint32: {sumUint}");
            Console.WriteLine($"Summary[0] - known unix: {(long)sumUint - knownUnix}");
            Console.WriteLine($"Summary[0] - row0 uint:  {(long)sumUint - asUint32}");

            // Try: timestamp = (asFloat * scale) + base
            // If row0=knownUnix and row1=knownUnix+300
            float row1Float = BitConverter.ToSingle(val, 12);
            double scale = 300.0 / (row1Float - asFloat);
            double base_ = knownUnix - asFloat * scale;
            Console.WriteLine($"\nFloat scale approach:");
            Console.WriteLine($"  scale = {scale:F2}");
            Console.WriteLine($"  base  = {base_:F2}");

            // Verify with row2
            float row2Float = BitConverter.ToSingle(val, 24);
            double row2Unix = row2Float * scale + base_;
            Console.WriteLine($"  row2 decoded: {DateTimeOffset.FromUnixTimeSeconds((long)row2Unix).LocalDateTime}");
            Console.WriteLine($"  row2 expected: {knownDt.AddSeconds(600)}");

            cf.Close();
        }
       public static void DiagnosticDump(string filePath)
        {
            using (var cf = new OpenMcdf.CompoundFile(filePath))
            {
                string numberedStorageName = null;
                cf.RootStorage.VisitEntries(item =>
                {
                    if (item.IsStorage && item.Name != "channels")
                        numberedStorageName = item.Name;
                }, recursive: false);

                var root = cf.RootStorage.GetStorage(numberedStorageName);
                var dataStorage = root.GetStorage("data");
                byte[] val = dataStorage.GetStream("values").GetData();
                byte[] sum = root.GetStream("summary").GetData();

                Console.WriteLine($"--- FILE DIAGNOSTICS: {Path.GetFileName(filePath)} ---");
                Console.WriteLine($"Total 'values' stream size: {val.Length} bytes");
                Console.WriteLine($"Total 'summary' stream size: {sum.Length} bytes");

                // 1. ANALYZE SUMMARY (Looking for the Row Count)
                Console.WriteLine("\n[SUMMARY STREAM ANALYSIS]");
                Console.WriteLine($"Hex: {BitConverter.ToString(sum)}");
                for (int i = 0; i <= sum.Length - 4; i += 4)
                {
                    int valInt = BitConverter.ToInt32(sum, i);
                    Console.WriteLine($"  Offset {i}: (Int32) {valInt}");
                }

                // 2. ANALYZE FIRST ROW (Looking for Timestamp format)
                Console.WriteLine("\n[FIRST ROW (FIRST 16 BYTES OF VALUES)]");
                byte[] firstRow = new byte[Math.Min(16, val.Length)];
                Array.Copy(val, 0, firstRow, 0, firstRow.Length);
                Console.WriteLine($"Hex: {BitConverter.ToString(firstRow)}");
                
                // Try interpreting first 4 bytes
                uint tsUint = BitConverter.ToUInt32(firstRow, 0);
                float tsFloat = BitConverter.ToSingle(firstRow, 0);
                Console.WriteLine($"  First 4 bytes as UInt32: {tsUint}");
                Console.WriteLine($"  First 4 bytes as Float:  {tsFloat}");

                // Try interpreting next 8 bytes as a double (the first measurement)
                if (val.Length >= 12)
                {
                    double firstMeas = BitConverter.ToDouble(val, 4);
                    Console.WriteLine($"  Bytes 4-11 as Double:    {firstMeas}");
                }

                // 3. ANALYZE SECOND ROW (To find the Byte-Gap/Row-Size)
                // We look for where the next timestamp-like value appears
                Console.WriteLine("\n[ROW SPACING SEARCH]");
                for (int i = 4; i < Math.Min(64, val.Length - 4); i += 4)
                {
                    uint nextTs = BitConverter.ToUInt32(val, i);
                    // If this looks like a timestamp (close to the first one)
                    if (Math.Abs((long)nextTs - tsUint) < 10000 && nextTs != 0)
                    {
                        Console.WriteLine($"  Potential next timestamp found at offset {i} (Value: {nextTs})");
                        Console.WriteLine($"  Calculated Row Size: {i} bytes");
                    }
                }
            }
        }

       public void ExportToCsv(string outputPath, List<MeasurementPoint> points){
            using (var writer = new StreamWriter(outputPath, false, Encoding.UTF8))
            {
                // 1. Dynamic Header
                string header = "Timestamp";
                if (points.Count > 0)
                {
                    for (int i = 0; i < points[0].Values.Length; i++)
                        header += $",Channel_{i + 1}";
                }
                writer.WriteLine(header);

                // 2. Data Rows
                foreach (var p in points)
                {
                    var line = new StringBuilder();
                    
                    // --- THE FIX: Convert UTC to Local Time ---
                    DateTime displayTime = p.Timestamp.ToLocalTime();
                    line.Append(displayTime.ToString("yyyy-MM-dd HH:mm:ss"));

                    foreach (var val in p.Values)
                    {
                        line.Append(",");
                        line.Append(val.ToString(CultureInfo.InvariantCulture));
                    }
                    writer.WriteLine(line.ToString());
                }
            }
        }

        public void WriteBinary(string templatePath, string outputPath, List<MeasurementPoint> filteredData, int startIndex)
        {
            File.Copy(templatePath, outputPath, true);

            using (var cf = new OpenMcdf.CompoundFile(outputPath, CFSUpdateMode.Update, 0))
            {
                string storageName = null;
                cf.RootStorage.VisitEntries(item => {
                    if (item.IsStorage && item.Name != "channels") storageName = item.Name;
                }, false);

                var root = cf.RootStorage.GetStorage(storageName);
                var dataStream = root.GetStorage("data").GetStream("values");
                byte[] originalBytes = dataStream.GetData();
                
                int rowSize = 12;
                int startOffset = startIndex * rowSize;
                
                // 1. EXTRACT REAL TICKS
                // Instead of guessing 320, we take the EXACT ticks from the original file
                uint newStartTick = BitConverter.ToUInt32(originalBytes, startOffset);
                
                // Find the end tick by looking at the last point of your range in the original file
                int endOffset = (startIndex + filteredData.Count - 1) * rowSize;
                uint newEndTick = BitConverter.ToUInt32(originalBytes, endOffset);
                
                // Calculate the actual average gap to keep the timeline linear
                uint totalTickDiff = newEndTick - newStartTick;
                float actualGap = (float)totalTickDiff / (filteredData.Count - 1);

                // 2. BUILD BUFFER
                byte[] newBuffer = new byte[filteredData.Count * rowSize];
                using (var ms = new MemoryStream(newBuffer))
                using (var writer = new BinaryWriter(ms))
                {
                    for (int i = 0; i < filteredData.Count; i++)
                    {
                        // We recreate the timeline using the exact calculated gap
                        uint currentTick = newStartTick + (uint)Math.Round(i * actualGap);
                        writer.Write(currentTick);
                        writer.Write((float)filteredData[i].Values[0]);
                        writer.Write((float)filteredData[i].Values[1]);
                    }
                }
                dataStream.SetData(newBuffer);

                // 3. UPDATE SUMMARY
                var sumStream = root.GetStream("summary");
                byte[] sumBytes = sumStream.GetData();

                // Offset 0: Start Tick (Syncs the start time)
                Array.Copy(BitConverter.GetBytes(newStartTick), 0, sumBytes, 0, 4);
                
                // Offset 12: Point Count (Fixes the truncation)
                Array.Copy(BitConverter.GetBytes(filteredData.Count), 0, sumBytes, 12, 4);
                
                // Offset 32: End Tick (Syncs the end time)
                Array.Copy(BitConverter.GetBytes(newEndTick), 0, sumBytes, 32, 4);

                sumStream.SetData(sumBytes);
                cf.Commit();
            }
        }
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_loader is IDisposable d) d.Dispose();
            if (_loader != null && Marshal.IsComObject(_loader))
                Marshal.ReleaseComObject(_loader);
            _loader = null;
        }
    }
}