# Vi2Converter

A .NET utility for reading and writing Testo `.vi2` measurement files using the `TestoModellVI2` COM interop library via reflection.

## Requirements

- .NET Framework 4.8
- x86 platform (required by the Testo COM libraries)
- The following DLLs must be present in the project root:
  - `TestoModellVI2.dll`
  - `Interop.Tcddka.dll`

## Project Structure

```
Vi2Converter/
├── Program.cs          # Entry point — demonstrates read/modify/write workflow
├── Vi2Handler.cs       # Core handler: reflection-based wrapper around TestVI2Loader
├── TestoModellVI2.dll  # Testo model library (internal class accessed via reflection)
├── Interop.Tcddka.dll  # COM interop for Testo Protocol objects
└── Vi2Converter.csproj
```

## How It Works

`TestVI2Loader` inside `TestoModellVI2.dll` is an `internal` class, so it cannot be referenced directly. `TestoVi2Handler` uses .NET reflection to load and invoke it at runtime.

### Key Classes

- `TestoVi2Handler` — wraps `TestVI2Loader` via reflection; implements `IDisposable`
- `MeasurementPoint` — plain data model with a `Timestamp` and a `double[]` of channel values

## Usage

```csharp
using (var handler = new TestoVi2Handler())
{
    // Read
    var data = handler.Read("input.vi2");

    // Optionally modify
    data[0].Timestamp = DateTime.Now;

    // Write
    handler.Write("output.vi2", data, "MySession");
}
```

### Read

`Read(string filePath)` loads a `.vi2` file and returns a `List<MeasurementPoint>`.  
Column 0 is treated as an OA date timestamp; remaining columns become `Values`.

### Write

`Write(string filePath, List<MeasurementPoint> data, string sessionName)` creates a new `.vi2` file from the provided data points under the given session name.

## Build

Open `Vi2Converter.sln` in Visual Studio and build for **x86 | Debug** (or Release).  
The project targets **net48** and requires **C# 10** language version.

> The DLLs are not redistributed with this repo. Copy them from your Testo software installation into the project root before building.
