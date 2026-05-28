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
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Secrets;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func;
using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        var configuration = context.Configuration;

        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        services.AddOptions<SendFunctionOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.MaxNumberOfAttempts = cfg.GetValue<int>("MaxNumberOfAttempts", 1);
                opts.SendRetryDelayNumberOfSeconds = cfg.GetValue<double>("SendRetryDelayNumberOfSeconds", 660);
            });

        services.AddOptions<BotOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.UserAppId = cfg.GetValue<string>("UserAppId");
                opts.UserAppPassword = cfg.GetValue<string>("UserAppPassword", string.Empty);
                opts.UserAppCertName = cfg.GetValue<string>("UserAppCertName", string.Empty);
                opts.AuthorAppId = cfg.GetValue<string>("AuthorAppId");
                opts.AuthorAppCertName = cfg.GetValue<string>("AuthorAppCertName", string.Empty);
                opts.GraphAppId = cfg.GetValue<string>("GraphAppId");
                opts.GraphAppCertName = cfg.GetValue<string>("GraphAppCertName", string.Empty);
                opts.UseCertificate = cfg.GetValue<bool>("UseCertificate", false);
            });

        services.AddOptions<RepositoryOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.StorageAccountConnectionString = cfg.GetValue<string>("StorageAccountConnectionString");
                opts.StorageAccountName = cfg.GetValue<string>("StorageAccountName");
                opts.EnsureTableExists = !cfg.GetValue<bool>("IsItExpectedThatTableAlreadyExists", true);
            });

        services.AddLocalization();

        var useManagedIdentity = bool.Parse(Environment.GetEnvironmentVariable("UseManagedIdentity") ?? "false");
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

        services.AddTransient<IMessageService, MessageService>();

        services.AddSingleton<ISendingNotificationDataRepository, SendingNotificationDataRepository>();
        services.AddSingleton<IGlobalSendingNotificationDataRepository, GlobalSendingNotificationDataRepository>();
        services.AddSingleton<ISentNotificationDataRepository, SentNotificationDataRepository>();
        services.AddSingleton<INotificationDataRepository, NotificationDataRepository>();
        services.AddTransient<TableRowKeyGenerator>();

        services.AddSingleton<ISendQueue, SendQueue>();

        services.AddTransient<INotificationService, NotificationService>();

        var keyVaultUrl = Environment.GetEnvironmentVariable("KeyVault:Url");
        services.AddSecretsProvider(keyVaultUrl);
    })
    .Build();

await host.RunAsync();
