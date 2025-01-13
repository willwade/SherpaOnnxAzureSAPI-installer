using System;
using System.IO;
using System.Reflection;

class AssemblyInspector
{
    static void Main(string[] args)
    {
        try
        {
            // Load sherpa-onnx assembly
            string dllPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "lib", "sherpa-onnx.dll");
            var assembly = Assembly.LoadFrom(dllPath);

            Console.WriteLine($"Assembly: {assembly.FullName}");
            Console.WriteLine("\nTypes:");
            foreach (var type in assembly.GetTypes())
            {
                Console.WriteLine($"\nType: {type.FullName}");
                Console.WriteLine("Methods:");
                foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
                {
                    Console.WriteLine($"  {method.Name}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
