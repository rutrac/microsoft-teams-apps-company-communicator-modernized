// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

using System;
using System.Globalization;
using System.Text.Json;
using Azure.Core.Serialization;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Adapter;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Clients;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Secrets;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.ExportQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Recipients;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.User;
using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func;
using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Streams;
using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults(workerApp => { }, options =>
    {
        options.Serializer = new JsonObjectSerializer(new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            IncludeFields = true,
        });
    })
    .ConfigureServices((context, services) =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        services.AddOptions<RepositoryOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.StorageAccountConnectionString = cfg.GetValue<string>("StorageAccountConnectionString");
                opts.StorageAccountName = cfg.GetValue<string>("StorageAccountName");
                opts.EnsureTableExists = !cfg.GetValue<bool>("IsItExpectedThatTableAlreadyExists", false);
            });

        services.AddOptions<BotOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.UserAppId = cfg.GetValue<string>("UserAppId");
                opts.UserAppPassword = cfg.GetValue<string>("UserAppPassword", string.Empty);
                opts.AuthorAppId = cfg.GetValue<string>("AuthorAppId");
                opts.AuthorAppPassword = cfg.GetValue<string>("AuthorAppPassword", string.Empty);
                opts.GraphAppId = cfg.GetValue<string>("GraphAppId");
                opts.UseCertificate = cfg.GetValue<bool>("UseCertificate", false);
                opts.AuthorAppCertName = cfg.GetValue<string>("AuthorAppCertName", string.Empty);
                opts.UserAppCertName = cfg.GetValue<string>("UserAppCertName", string.Empty);
                opts.GraphAppCertName = cfg.GetValue<string>("GraphAppCertName", string.Empty);
                opts.TenantId = cfg.GetValue<string>("TenantId", string.Empty);
            });

        services.AddOptions<DataQueueMessageOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.MessageDelayInSeconds = cfg.GetValue<int>("DataQueueMessageDelayInSeconds", 5);
            });

        services.AddOptions<TeamsConversationOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.ProactivelyInstallUserApp = cfg.GetValue<bool>("ProactivelyInstallUserApp", true);
                opts.MaxAttemptsToCreateConversation = cfg.GetValue<int>("MaxAttemptsToCreateConversation", 2);
            });

        services.AddLocalization();

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
        services.AddSingleton<AuthorAppCredentials>();
        services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
        services.AddSingleton<ICCBotFrameworkHttpAdapter, CCBotFrameworkHttpAdapter>();
        services.AddSingleton<BotFrameworkHttpAdapter>();

        services.AddSingleton<INotificationDataRepository, NotificationDataRepository>();
        services.AddSingleton<ISendingNotificationDataRepository, SendingNotificationDataRepository>();
        services.AddSingleton<ISentNotificationDataRepository, SentNotificationDataRepository>();
        services.AddSingleton<IUserDataRepository, UserDataRepository>();
        services.AddSingleton<ITeamDataRepository, TeamDataRepository>();
        services.AddSingleton<IExportDataRepository, ExportDataRepository>();
        services.AddSingleton<IAppConfigRepository, AppConfigRepository>();

        services.AddSingleton<ISendQueue, SendQueue>();
        services.AddSingleton<IDataQueue, DataQueue>();
        services.AddSingleton<IExportQueue, ExportQueue>();

        services.AddTransient<TableRowKeyGenerator>();
        services.AddTransient<AdaptiveCardCreator>();
        services.AddTransient<IAppSettingsService, AppSettingsService>();
        services.AddTransient<IStorageClientFactory, StorageClientFactory>();
        services.AddTransient<IUserTypeService, UserTypeService>();
        services.AddTransient<IRecipientsService, RecipientsService>();

        services.AddTransient<ITeamMembersService, TeamMembersService>();
        services.AddTransient<IConversationService, ConversationService>();

        var keyVaultUrl = Environment.GetEnvironmentVariable("KeyVault:Url");
        services.AddSecretsProvider(keyVaultUrl);

        // Graph services.
        services.AddOptions<ConfidentialClientApplicationOptions>()
            .Configure<IConfiguration>((opts, cfg) =>
            {
                opts.ClientId = cfg.GetValue<string>("GraphAppId");
                opts.ClientSecret = cfg.GetValue<string>("GraphAppPassword", string.Empty);
                opts.TenantId = cfg.GetValue<string>("TenantId");
            });

        var useClientCertificates = bool.Parse(Environment.GetEnvironmentVariable("UseCertificate") ?? "false");
        services.AddConfidentialClient(useClientCertificates);

        services.AddSingleton<MsalAuthenticationProvider>();

        services.AddSingleton<GraphServiceClient>(sp =>
            new GraphServiceClient(
                new Microsoft.Kiota.Abstractions.Authentication.BaseBearerTokenAuthenticationProvider(
                    sp.GetRequiredService<MsalAuthenticationProvider>())));

        services.AddSingleton<IGraphServiceFactory, GraphServiceFactory>();

        services.AddScoped<IUsersService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetUsersService());
        services.AddScoped<IGroupMembersService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetGroupMembersService());
        services.AddScoped<IAppManagerService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetAppManagerService());
        services.AddScoped<IChatsService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetChatsService());

        services.AddTransient<IDataStreamFacade, DataStreamFacade>();
    })
    .Build();

await host.RunAsync();
