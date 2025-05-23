namespace PluginBase
{
    public interface IReportGenerator
    {
        string Name { get; }
        void GenerateReport(string[] items, string outputPath);
    }
}
