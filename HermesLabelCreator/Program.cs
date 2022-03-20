using HermesLabelCreator.Configurations;
using HermesLabelCreator.Formatters;
using HermesLabelCreator.Interfaces;
using HermesLabelCreator.Models;
using HermesLabelCreator.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using Serilog.Events;
using Serilog.Formatting.Compact;
using Serilog.Formatting.Raw;
using Serilog.Sinks.File.Header;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HermesLabelCreator
{
    class Program
    {
        public static IConfigurationRoot Configuration { get; set; }
        public static ILogger Logger { get; set; }

        static void Main(string[] args)
        {
            string importFolderFilePath = args != null && args.Length > 0
                ? args[0]
                : string.Empty;

            bool singleFile = false;
            if (args != null && args.Length > 1)
            {
                bool.TryParse(args[1], out singleFile);
            }

            if (string.IsNullOrWhiteSpace(importFolderFilePath))
            {
                Console.WriteLine("Please enter path to _import directory.");
                Console.ReadLine();
                return;
            }

            string[] directories = Directory.GetDirectories(importFolderFilePath, string.Empty, SearchOption.TopDirectoryOnly);
            if (!directories.Any(d => new DirectoryInfo(d).Name.Equals("_import")))
            {
                Console.WriteLine("There is no directory with \"_import\" folder. Please enter valid directory path.");
                Console.ReadLine();
                return;
            }

            string logFileInfoPath = Path.Combine(importFolderFilePath, "IrisReponseWithRemarks.csv");
            string logFileErrorPath = Path.Combine(importFolderFilePath, "SystemErrors.txt");

            if (File.Exists(logFileInfoPath))
            {
                File.Delete(logFileInfoPath);
            }

            if (File.Exists(logFileErrorPath))
            {
                File.Delete(logFileErrorPath);
            }

            var builder = new ConfigurationBuilder();
            BuildConfig(builder);


            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(Configuration)
                .Enrich.FromLogContext()
                .WriteTo.Logger(l => l.Filter.ByIncludingOnly(e => e.Level == LogEventLevel.Information).WriteTo.File(new CsvFormatter(), logFileInfoPath, hooks: new HeaderWriter(CsvFormatter.HeaderFactory)))
                .WriteTo.Logger(l => l.Filter.ByIncludingOnly(e => e.Level == LogEventLevel.Error).WriteTo.File(logFileErrorPath))
                .CreateLogger();

            var host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                    services.AddOptions();
                    services.Configure<ApplicationConfiguration>(Configuration.GetSection("Application"));
                    services.AddSingleton<Iservice, IrisService>();
                    services.AddSingleton<ProcessRunner>();
                })
                .UseSerilog()
                .Build();

            Logger = Log.Logger;
            AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;

            var process = ActivatorUtilities.CreateInstance<ProcessRunner>(host.Services);
            process.Run(importFolderFilePath, singleFile);
        }

        private static void BuildConfig(IConfigurationBuilder builder)
        {
            #if DEBUG
              Environment.SetEnvironmentVariable("ASPNETCORE_ENVIRONMENT", "Development");
            #else
              Environment.SetEnvironmentVariable("ASPNETCORE_ENVIRONMENT", "Production");
            #endif

            builder.SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile($"appsettings.json", false, true)
                .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT")}.json", true, true)
                .AddEnvironmentVariables();

            Configuration = builder.Build();
        }

        static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {
            Logger.Error((Exception)e.ExceptionObject, ((Exception)e.ExceptionObject).Message);
        }
    }
}
