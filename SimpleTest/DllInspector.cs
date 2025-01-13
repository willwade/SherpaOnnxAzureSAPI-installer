using System;
using System.Reflection;

namespace DllInspector
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string dllPath = @"..\OpenSpeechTTS\lib\sherpa-onnx.dll";
                Assembly assembly = Assembly.LoadFrom(dllPath);

                Console.WriteLine($"Assembly: {assembly.FullName}\n");

                foreach (Type type in assembly.GetTypes())
                {
                    Console.WriteLine($"Type: {type.FullName}");
                    
                    foreach (PropertyInfo prop in type.GetProperties())
                    {
                        Console.WriteLine($"  Property: {prop.Name} ({prop.PropertyType})");
                    }
                    
                    foreach (MethodInfo method in type.GetMethods())
                    {
                        if (method.DeclaringType == type) // Only show methods declared in this type
                        {
                            Console.WriteLine($"  Method: {method.Name}");
                            foreach (ParameterInfo param in method.GetParameters())
                            {
                                Console.WriteLine($"    Parameter: {param.Name} ({param.ParameterType})");
                            }
                        }
                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
