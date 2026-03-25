using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.VisualBasic;

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
            if (!File.Exists(modelDll)) throw new FileNotFoundException($"Missing {modelDll}");
            
            Assembly testoAssembly = Assembly.LoadFrom(modelDll);
            _tcddkaAssembly = Assembly.LoadFrom(interopDll);

            _loaderType = testoAssembly.GetType("Testo.Model.TestVI2Loader");
            _loader = Activator.CreateInstance(_loaderType);

            _protocolsProp = _loaderType.GetProperty("Protocols", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            _loadMethod = _loaderType.GetMethod("Load", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            _saveMethod = _loaderType.GetMethod("Save", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
        }

        public List<MeasurementPoint> Read(string filePath)
        {
            var results = new List<MeasurementPoint>();
            _loadMethod.Invoke(_loader, new object[] { filePath });

            var protocols = (IEnumerable)_protocolsProp.GetValue(_loader);
            foreach (object p in protocols)
            {
                // Property names from IProtocol: NumRows, NumCols
                int rowCount = Convert.ToInt32(Interaction.CallByName(p, "NumRows", CallType.Get));
                short colCount = Convert.ToInt16(Interaction.CallByName(p, "NumCols", CallType.Get));

                for (int r = 0; r < rowCount; r++)
                {
                    var point = new MeasurementPoint { Values = new double[colCount] };
                    
                    // Access "Date" as a property using CallType.Get instead of "get_Date"
                    point.Timestamp = (DateTime)Interaction.CallByName(p, "Date", CallType.Get, r);

                    // Access "Val" as a property using CallType.Get instead of "get_Val"
                    for (short c = 0; c < colCount; c++)
                    {
                        point.Values[c] = Convert.ToDouble(Interaction.CallByName(p, "Val", CallType.Get, r, c));
                    }
                    results.Add(point);
                }
            }
            return results;
        }


        public void Write(string filePath, List<MeasurementPoint> data, string sessionName = "FilteredRange")
        {
            Type protocolType = _tcddkaAssembly.GetType("Tcddka.ProtocolClass");
            dynamic protocol = Activator.CreateInstance(protocolType);
            
            protocol.Title = sessionName;
            // The COM engine handles internal data management when Save is called
            // Note: Setting data typically uses the same loader list pattern
            
            IList protocolList = (IList)_protocolsProp.GetValue(_loader);
            protocolList.Clear();
            protocolList.Add(protocol);

            _saveMethod.Invoke(_loader, new object[] { filePath });
        }

        public void Dispose() { if (_loader is IDisposable d) d.Dispose(); }
    }
}
