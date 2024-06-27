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
using McMaster.Extensions.Hosting.CommandLine;
using Microsoft.Office.Interop.Word;
using McMaster.Extensions.CommandLineUtils;
using Serilog.Sinks.SystemConsole.Themes;
using SeriLogThemesLibrary;

internal class Program
{   
    private static async Task<int> Main(string[] args)
    {               
        var enviroment = Environment.GetEnvironmentVariable("NETCORE_ENVIROMENT");

        string logfile = "loggs\\myapp.txt";
        Logger seriLogger = new LoggerConfiguration()
                            //https://dev.to/karenpayneoregon/serilog-color-themes-58nj
                            .WriteTo.Console(theme: SeriLogCustomThemes.Theme1())
                            //.WriteTo.Console(theme: AnsiConsoleTheme.Code)
                            .WriteTo.File(logfile, rollingInterval: RollingInterval.Day)
                            .CreateLogger();

        //var builder = Host.CreateApplicationBuilder(args);
        var builder = Host.CreateDefaultBuilder()
            .ConfigureAppConfiguration(config =>
            {
                config.SetBasePath(Directory.GetCurrentDirectory());
                config.AddUserSecrets(Assembly.GetExecutingAssembly());
                config.AddJsonFile("pdfsettings.json", optional: false, reloadOnChange: true);
            })
            .ConfigureServices((context, services) =>
            {
                services.AddSingleton<MainApp>();
                services.AddSingleton(context.Configuration); //?
                services.AddOptions<PdfOptions>().BindConfiguration("PdfOptions");
                services.AddDbContext<ApplicationDbContext>(options =>
                    options.UseSqlServer(context.Configuration.GetConnectionString("DefaultConnection")));
                services.AddScoped<ISdsConfigRepo, SdsConfigRepo>();
                services.AddScoped<IMsWordTools, MsWordTools>();
            })
            .ConfigureLogging((context, builder) =>
            {
                builder.ClearProviders();
                builder.AddSerilog(seriLogger);
            });

        string[] testargs = { "--mode", "sql", "--sprac", "202404" };
        return await builder.RunCommandLineApplicationAsync<MainApp>(testargs);
    }
}

