// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;

    /// <summary>
    /// Program class of the company communicator application.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Main function of the company communicator application.
        /// It builds a web host, then launches the company communicator into it.
        /// </summary>
        /// <param name="args">Arguments passed in to the function.</param>
        public static void Main(string[] args)
        {
            // Persistent crash logger: ntdll!ZwTerminateProcess fail-fasts wipe the worker-local
            // CrashDumps folder before we can read it, so write a sentinel to D:\home\LogFiles
            // (the only persisted location accessible from Kudu vfs).
            AppDomain.CurrentDomain.UnhandledException += (sender, e) => WriteFatal("UnhandledException", e.ExceptionObject as Exception, e.IsTerminating);
            TaskScheduler.UnobservedTaskException += (sender, e) => { WriteFatal("UnobservedTaskException", e.Exception, false); e.SetObserved(); };
            AppDomain.CurrentDomain.ProcessExit += (sender, e) => WriteFatal("ProcessExit", null, true);

            CreateHostBuilder(args).Build().Run();
        }

        private static void WriteFatal(string source, Exception ex, bool terminating)
        {
            try
            {
                var dir = Environment.GetEnvironmentVariable("HOME");
                if (string.IsNullOrEmpty(dir))
                {
                    dir = "D:\\home";
                }

                var logDir = Path.Combine(dir, "LogFiles", "fatal");
                Directory.CreateDirectory(logDir);
                var file = Path.Combine(logDir, $"fatal-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Environment.ProcessId}-{source}.log");
                File.WriteAllText(file, $"=== {source} at {DateTime.UtcNow:O} pid={Environment.ProcessId} terminating={terminating} ===\n{ex}\n");
            }
            catch
            {
                // best-effort; never throw from a fatal-path logger.
            }
        }

        /// <summary>
        /// Create the web host builder.
        /// </summary>
        /// <param name="args">Arguments passed into the main function.</param>
        /// <returns>A web host builder instance.</returns>
        public static IHostBuilder CreateHostBuilder(string[] args) =>
           Host.CreateDefaultBuilder(args)
               .ConfigureWebHostDefaults(webBuilder =>
               {
                   webBuilder.UseStartup<Startup>();
               })
               .ConfigureServices(services =>
               {
                services.AddHostedService<SendMessageScheduler>();
               });
    }
}