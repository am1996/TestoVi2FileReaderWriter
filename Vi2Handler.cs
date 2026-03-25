using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

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
        private Assembly _tcddkaAssembly;
        private PropertyInfo _protocolsProp;
        private MethodInfo _loadMethod;
        private MethodInfo _saveMethod;

        public TestoVi2Handler(string modelDll = "TestoModellVI2.dll", string interopDll = "Interop.Tcddka.dll")
        {
            if (!File.Exists(modelDll) || !File.Exists(interopDll))
                throw new FileNotFoundException("Ensure TestoModellVI2.dll and Interop.Tcddka.dll are in the folder.");

            Assembly testoAssembly = Assembly.LoadFrom(modelDll);
            _tcddkaAssembly = Assembly.LoadFrom(interopDll);

            _loaderType = testoAssembly.GetType("Testo.Model.TestVI2Loader");
            _loader = Activator.CreateInstance(_loaderType);

            // Cache Reflection info for performance
            _protocolsProp = _loaderType.GetProperty("Protocols", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            _loadMethod = _loaderType.GetMethod("Load", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            _saveMethod = _loaderType.GetMethod("Save", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
        }

        // --- READ LOGIC ---
        public List<MeasurementPoint> Read(string filePath)
        {
            var results = new List<MeasurementPoint>();
            _loadMethod.Invoke(_loader, new object[] { filePath });

            var protocols = (IEnumerable)_protocolsProp.GetValue(_loader);
            foreach (var p in protocols)
            {
                dynamic dp = p; // Uses Microsoft.CSharp (Interop.Tcddka.Protocol)
                for (long r = 0; r < dp.Count; r++)
                {
                    var point = new MeasurementPoint { Values = new double[dp.NumCols - 1] };
                    point.Timestamp = DateTime.FromOADate(Convert.ToDouble(dp.GetVal(r, (short)0)));

                    for (short c = 1; c < dp.NumCols; c++)
                    {
                        point.Values[c - 1] = Convert.ToDouble(dp.GetVal(r, c));
                    }
                    results.Add(point);
                }
            }
            return results;
        }

        // --- WRITE LOGIC ---
        public void Write(string filePath, List<MeasurementPoint> data, string sessionName = "ExportedSession")
        {
            // Create the COM Protocol object
            Type protocolType = _tcddkaAssembly.GetType("Tcddka.ProtocolClass");
            dynamic protocol = Activator.CreateInstance(protocolType);
            protocol.Name = sessionName;

            foreach (var point in data)
            {
                // Note: Tcddka Protocol objects usually expect: Add(OADate, ArrayOfValues) 
                // or Add(OADate, Value1, Value2...). 
                // Adjusting to the most common Interop signature:
                protocol.Add(point.Timestamp.ToOADate(), point.Values);
            }

            // Sync with Loader and Save
            IList protocolList = (IList)_protocolsProp.GetValue(_loader);
            protocolList.Clear();
            protocolList.Add(protocol);

            _saveMethod.Invoke(_loader, new object[] { filePath });
        }

        public void Dispose()
        {
            if (_loader is IDisposable d) d.Dispose();
        }
    }
}
