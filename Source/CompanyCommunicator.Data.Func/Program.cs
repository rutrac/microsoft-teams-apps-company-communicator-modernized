// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

using System;
using System.Globalization;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Adapter;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Secrets;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Data.Func;
using Microsoft.Teams.Apps.CompanyCommunicator.Data.Func.Services.FileCardServices;
using Microsoft.Teams.Apps.CompanyCommunicator.Data.Func.Services.NotificationDataServices;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        services.AddOptions<RepositoryOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.StorageAccountConnectionString = cfg.GetValue<string>("StorageAccountConnectionString");
                opts.StorageAccountName = cfg.GetValue<string>("StorageAccountName");
                opts.EnsureTableExists = !cfg.GetValue<bool>("IsItExpectedThatTableAlreadyExists", true);
            });

        services.AddOptions<BotOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.UserAppId = cfg.GetValue<string>("UserAppId");
                opts.UserAppPassword = cfg.GetValue<string>("UserAppPassword", string.Empty);
                opts.UserAppCertName = cfg.GetValue<string>("UserAppCertName", string.Empty);
                opts.AuthorAppId = cfg.GetValue<string>("AuthorAppId");
                opts.AuthorAppPassword = cfg.GetValue<string>("AuthorAppPassword", string.Empty);
                opts.AuthorAppCertName = cfg.GetValue<string>("AuthorAppCertName", string.Empty);
                opts.GraphAppId = cfg.GetValue<string>("GraphAppId");
                opts.GraphAppCertName = cfg.GetValue<string>("GraphAppCertName", string.Empty);
                opts.UseCertificate = cfg.GetValue<bool>("UseCertificate", false);
            });

        services.AddOptions<CleanUpFileOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.CleanUpFile = cfg.GetValue<string>("CleanUpFile");
            });

        services.AddOptions<DataQueueMessageOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.FirstTenMinutesRequeueMessageDelayInSeconds = cfg.GetValue<double>("FirstTenMinutesRequeueMessageDelayInSeconds", 20);
                opts.RequeueMessageDelayInSeconds = cfg.GetValue<double>("RequeueMessageDelayInSeconds", 120);
            });

        services.AddLocalization();
        services.AddHttpClient();

        var useManagedIdentity = bool.Parse(Environment.GetEnvironmentVariable("UseManagedIdentity") ?? "false");
        services.AddBlobClient(useManagedIdentity);
        services.AddServiceBusClient(useManagedIdentity);

        var culture = Environment.GetEnvironmentVariable("i18n:DefaultCulture");
        if (!string.IsNullOrEmpty(culture))
        {
            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo(culture);
            CultureInfo.DefaultThreadCurrentUICulture = new CultureInfo(culture);
        }

        services.AddSingleton<UserAppCredentials>();
        services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
        services.AddSingleton<ICCBotFrameworkHttpAdapter, CCBotFrameworkHttpAdapter>();
        services.AddSingleton<BotFrameworkHttpAdapter>();

        var keyVaultUrl = Environment.GetEnvironmentVariable("KeyVault:Url");
        services.AddSecretsProvider(keyVaultUrl);

        services.AddSingleton<IFileCardService, FileCardService>();

        services.AddTransient<AggregateSentNotificationDataService>();
        services.AddTransient<UpdateNotificationDataService>();

        services.AddSingleton<INotificationDataRepository, NotificationDataRepository>();
        services.AddSingleton<ISentNotificationDataRepository, SentNotificationDataRepository>();
        services.AddSingleton<IUserDataRepository, UserDataRepository>();
        services.AddSingleton<IExportDataRepository, ExportDataRepository>();
        services.AddTransient<TableRowKeyGenerator>();

        services.AddSingleton<IDataQueue, DataQueue>();
    })
    .Build();

await host.RunAsync();
