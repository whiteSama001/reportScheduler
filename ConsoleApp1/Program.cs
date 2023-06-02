/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace mailScheduler
{
    internal class Program
    {
        static void Main(string[] args)
        {
            GetExcelFile GF = new GetExcelFile();
            GF.GetDataAndExportExcel();

            SendMail Sm = new SendMail();
            Sm.SendEmail("joshua.komolafe@optimusbank.com", "Report on Bill Payments");
        }
    }


}*/

using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace mailScheduler
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    GetExcelFile GF = new GetExcelFile();
                    GF.GetDataAndExportExcel();

                    SendMail Sm = new SendMail();
                    Sm.SendEmail("joshua.komolafe@gmail.com", "DAILY BET9JA REPORT");

                    await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken); 
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "An error occurred while executing the worker task.");
                    await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken); 
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddHostedService<Worker>();
                });
    }
}

