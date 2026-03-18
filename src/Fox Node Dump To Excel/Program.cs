using Fox_Node_Dump_Parser.Logic;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Threading.Tasks;
    
namespace Fox_Node_Dump_Parser
{
    internal class Program  
    {
        static async Task Main(string[] args)
        {
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            using IHost host = Host.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config.SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                          .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                          .AddEnvironmentVariables();
                })
                .ConfigureServices((context, services) =>
                {
                    // Register application services
                    services.AddTransient<AppWorker>();
                })
                .Build();

            // Create a scope and run the AppWorker resolved from DI.
            // The IConfiguration (from appsettings.json) will be provided to AppWorker by the DI container.
            using var scope = host.Services.CreateScope();
            var provider = scope.ServiceProvider;

            var appWorker = provider.GetRequiredService<AppWorker>();
            await appWorker.DoWorkAsync();
        }
    }
}
