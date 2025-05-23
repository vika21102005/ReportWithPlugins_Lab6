using System;
using System.IO;
using System.Linq;
using System.Reflection;
using PluginBase;

namespace ReportApp
{
    class Program
    {
        static void Main()
        {
            var pluginDir = Path.Combine(AppContext.BaseDirectory, "Plugins");
            var dlls = Directory.GetFiles(pluginDir, "*.dll");
            var generators = dlls
                .SelectMany(file =>
                {
                    var asm = Assembly.LoadFrom(file);
                    return asm.GetTypes()
                              .Where(t => typeof(IReportGenerator).IsAssignableFrom(t) && !t.IsInterface)
                              .Select(t => (IReportGenerator)Activator.CreateInstance(t)!);
                })
                .ToList();

            Console.WriteLine("Оберіть формат звіту:");
            for (int i = 0; i < generators.Count; i++)
                Console.WriteLine($"{i + 1}. {generators[i].Name}");
            var idx = int.Parse(Console.ReadLine()!) - 1;
            var gen = generators[idx];

            var data = new[] { "Пункт A", "Пункт B", "Пункт C" };
            var fileName = $"{gen.Name}_Report_{DateTime.Now:yyyyMMdd_HHmmss}" +
                           (gen.Name == "Word" ? ".docx" : ".xlsx");
            var output = Path.Combine(AppContext.BaseDirectory, fileName);

            gen.GenerateReport(data, output);
            Console.WriteLine($"Звіт створено: {output}");
        }
    }
}
