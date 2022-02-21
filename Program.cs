using DesignerEyesService.FileFtpService;
using DesignerEyesService.Interfaces;
using DesignerEyesService.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using System.Threading.Tasks;

namespace DesignerEyesService
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Log.Logger = new LoggerConfiguration()
          .WriteTo.File("DesignerEyes.log", rollingInterval: RollingInterval.Day)
          .CreateLogger();

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);

            var serviceProvider = serviceCollection.BuildServiceProvider();

            await serviceProvider.GetService<MainProcess>().StartService(args); 

            var logger = serviceProvider.GetService<ILogger<Program>>();

            logger.LogInformation("Completed process");
        }

        private static void ConfigureServices(IServiceCollection services)
        {
            services.AddLogging(configure => configure.AddSerilog())
                    .AddTransient<MainProcess>()
                    .AddTransient<IFileFtpService, FtpFilesServ>()
                    .AddTransient<IReadInventory, ReadInventory>()
                    .AddTransient<IReadOrder, ReadOrder>()
                    .AddTransient<IReadShipConfirm, ReadShipConfirm>();
        }
    }
}
