using ListyApp;
using ListyApp.Models;
using ListyApp.Models.Exports;
using ListyApp.Repositories;
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
using static System.Formats.Asn1.AsnWriter;

internal class Program1
{
    private static async Task Main1(string[] args)
    {
        //https://stackoverflow.com/questions/42268265/how-to-get-manage-user-secrets-in-a-net-core-console-application
        //https://andrewlock.net/using-dependency-injection-in-a-net-core-console-application/
        //https://learn.microsoft.com/en-us/aspnet/core/fundamentals/configuration/?view=aspnetcore-8.0
        //https://learn.microsoft.com/en-us/dotnet/core/extensions/workers
        //https://learn.microsoft.com/en-us/dotnet/core/extensions/generic-host?tabs=appbuilder
        var enviroment = Environment.GetEnvironmentVariable("NETCORE_ENVIROMENT");

        string logfile = "loggs\\myapp.txt";
        Logger seriLogger = new LoggerConfiguration()                    
                            .WriteTo.Console()
                            .WriteTo.File(logfile, rollingInterval: RollingInterval.Day)
                            .CreateLogger();

        //novsi styl, viac linerarny kod, tiez vselico predefinvane
        /*var builder2 = Host.CreateApplicationBuilder(args);

        builder2.Configuration
            //.SetBasePath(Directory.GetCurrentDirectory())
            //.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            //.AddJsonFile($"appsettings.{enviroment}.json", optional: true, reloadOnChange: true)
            .AddUserSecrets(Assembly.GetExecutingAssembly())
            .AddJsonFile("pdfsettings.json", optional: false, reloadOnChange: true);

        builder2.Services.AddSingleton<MainApp>();
        builder2.Services.AddOptions<PdfOptions>().BindConfiguration("PdfOptions");
        builder2.Services.AddDbContext<ApplicationDbContext>(options =>
            options.UseSqlServer(builder2.Configuration.GetConnectionString("DefaultConnection")));
        builder2.Services.AddScoped<ISdsConfigRepo, SdsConfigRepo>();

        builder2.Logging
            .ClearProviders()
            .AddSerilog(seriLogger);                  
        
        var host=builder2.Build(); */
        
        //klasicky styl cez callback, ma aj niektore predkonfigurovane veci ako appsettings.json, consoleloger a pod
        var builder = Host.CreateDefaultBuilder(args);       
        builder.ConfigureAppConfiguration(config =>
        {
            //netreba, toto robi DefaultBuilder() automaticky!
            //config.SetBasePath(Directory.GetCurrentDirectory());
            //config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            //config.AddJsonFile($"appsettings.{enviroment}.json", optional: true, reloadOnChange: true);
            config.AddUserSecrets(Assembly.GetExecutingAssembly());
            config.AddJsonFile("pdfsettings.json", optional: false, reloadOnChange: true);
        });

        builder.ConfigureServices((context, services) =>
        {
            services.AddSingleton<MainApp>();
            services.AddSingleton(context.Configuration);
            services.Configure<PdfOptions>(context.Configuration.GetSection("PdfOptions"));

            var connectionString = context.Configuration.GetConnectionString("DefaultConnection");
            services.AddDbContext<ApplicationDbContext>(options => options.UseSqlServer(connectionString));
            services.AddScoped<ISdsConfigRepo, SdsConfigRepo>();
        });

        builder.ConfigureLogging((context, builder) =>
        {
            builder.ClearProviders();
            builder.AddSerilog(seriLogger);
        });
           
        var host=builder.Build();
       
        MainApp mainApp = host.Services.GetRequiredService<MainApp>();
        await mainApp.StartProcess();

        //toto asi nie je teraz takto nutne takto robit, staci vid vyssie
        //using (IServiceScope scope = host.Services.CreateScope())
        //{
        //    MainApp mainApp = scope.ServiceProvider.GetRequiredService<MainApp>();
        //    await mainApp.Start();
        //}
    }
}

/*
var builder = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddJsonFile($"appsettings.{enviroment}.json", optional: true, reloadOnChange: true)
    .AddJsonFile("pdfsettings.json", optional: false, reloadOnChange: true)
    .AddUserSecrets(Assembly.GetExecutingAssembly());

var config=builder.Build();

var services = CreateServices(config);

MainApp mainApp=services.GetRequiredService<MainApp>(); 

mainApp.Start();
*/

/*static IServiceProvider CreateServices(IConfigurationRoot config)
{
    string logfile = "loggs\\myapp.txt";
    
    LoggerFactory loggerFactory = new LoggerFactory();
    Logger slogger = new LoggerConfiguration()
                    .WriteTo.Console()
                    .WriteTo.File(logfile, rollingInterval: RollingInterval.Day)
                    .CreateLogger();

    var serviceProvider = new ServiceCollection()
        .AddLogging(options=>
        {
            options.ClearProviders();
            options.AddSerilog(slogger);
        })
        .AddSingleton<MainApp>()    
        .AddSingleton<IConfiguration>(config)
        .Configure<PdfOptions>(config.GetSection("PdfOptions"))
        .BuildServiceProvider();   
    
    return serviceProvider;
}*/