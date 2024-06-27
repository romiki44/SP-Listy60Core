using ListyApp;
using ListyApp.Models;
using ListyApp.Models.Exports;
using ListyApp.Repositories;
using ListyApp.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Core;
using System;
using System.Configuration;
using System.Reflection;
using System.Runtime.Serialization;

internal class Program2
{
    private static async Task Main2(string[] args)
    {               
        var enviroment = Environment.GetEnvironmentVariable("NETCORE_ENVIROMENT");

        string logfile = "loggs\\myapp.txt";
        Logger seriLogger = new LoggerConfiguration()
                            .WriteTo.Console()
                            .WriteTo.File(logfile, rollingInterval: RollingInterval.Day)
                            .CreateLogger();

        //novsi styl, viac linerarny kod, tiez vselico predefinovane
        var builder=Host.CreateApplicationBuilder(args);

        builder.Configuration
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddUserSecrets(Assembly.GetExecutingAssembly())
            .AddJsonFile("pdfsettings.json", optional: false, reloadOnChange: true);

        builder.Services.AddSingleton<MainApp>();
        builder.Services.AddOptions<PdfOptions>().BindConfiguration("PdfOptions");
        builder.Services.AddDbContext<ApplicationDbContext>(options =>
            options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));
        builder.Services.AddScoped<ISdsConfigRepo, SdsConfigRepo>();
        builder.Services.AddScoped<IMsWordTools, MsWordTools>();

        builder.Logging
            .ClearProviders()
            .AddSerilog(seriLogger);
               
        var host=builder.Build();                     
        MainApp mainApp = host.Services.GetRequiredService<MainApp>();
        await mainApp.StartProcess();
    }
}

