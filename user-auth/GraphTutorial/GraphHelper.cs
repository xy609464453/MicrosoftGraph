// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Azure.Core;
using Azure.Identity;

using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;

using System.Text;

partial class GraphHelper
{
    // <UserAuthConfigSnippet>
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        var options = new DeviceCodeCredentialOptions
        {
            ClientId = settings.ClientId,
            TenantId = settings.TenantId,
            DeviceCodeCallback = deviceCodePrompt,
        };

        _deviceCodeCredential = new DeviceCodeCredential(options);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    // </UserAuthConfigSnippet>

    // <GetUserTokenSnippet>
    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }
    // </GetUserTokenSnippet>

    // <GetUserSnippet>
    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me.GetAsync((config) => {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName" };
        });
    }
    // </GetUserSnippet>

    // <GetInboxSnippet>
    public static Task<MessageCollectionResponse?> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            // Only messages from Inbox folder
            .MailFolders["Inbox"]
            .Messages
            .GetAsync((config) => {
                // Only request specific properties
                config.QueryParameters.Select = new[] { "from", "isRead", "receivedDateTime", "subject" };
                // Get at most 25 results
                config.QueryParameters.Top = 25;
                // Sort by received time, newest first
                config.QueryParameters.Orderby = new[] { "receivedDateTime DESC" };
            });
    }
    // </GetInboxSnippet>

    // <SendMailSnippet>
    public static async Task SendMailAsync(string subject, string body, string recipient)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Create a new message
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = BodyType.Text
            },
            ToRecipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient
                    }
                }
            }
        };

        // Send the message
        await _userClient.Me
            .SendMail
            .PostAsync(new SendMailPostRequestBody
            {
                Message = message
            });
    }
    // </SendMailSnippet>
}

partial class GraphHelper
{
#pragma warning disable CS1998
    // <MakeGraphCallSnippet>
    // This function serves as a playground for testing Graph snippets
    // or other code
    public async static Task MakeGraphCallAsync()
    {
        var mapper = new Dictionary<int, (string name, Action action)>();

        mapper.Add(0, ("Exit", () => {
            Console.WriteLine("Goodbye...");
        }
        ));

        mapper.Add(1, ("OneDrive Root Children", async () => {
            _ = _userClient ??
                throw new System.NullReferenceException();
            var driveItem = await _userClient.Me.Drive.GetAsync();
            _ = driveItem ??
              throw new System.NullReferenceException();

            var userDriveId = driveItem.Id;
            // List children in the drive
            var driveRequest = _userClient.Drives[userDriveId];
            var root = await driveRequest.Root.GetAsync();

            var children = await driveRequest.Items[root.Id].Children.GetAsync();
            foreach (var item in children.Value)
            {
                Console.WriteLine(item.Name);
            }
        }
        ));

        mapper.Add(2, ("OneDrive Root Upload", async () => {
            _ = _userClient ??
                throw new System.NullReferenceException();
            var driveItem = await _userClient.Me.Drive.GetAsync();
            _ = driveItem ??
              throw new System.NullReferenceException();

            var userDriveId = driveItem.Id;
            // List children in the drive
            var driveRequest = _userClient.Drives[userDriveId];
            var root = await driveRequest.Root.GetAsync();

            using var stream = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(@"The contents of the file goes here."));
            await driveRequest.Items[root.Id].ItemWithPath("test.txt").Content.PutAsync(stream);
        }
        ));

        int choice = -1;

        while (choice != 0)
        {
            Console.WriteLine("Please choose one of the following options:");
            foreach (var item in mapper)
            {
                Console.WriteLine($"{item.Key}. {item.Value.name}");
            }

            // INSERT YOUR CODE HERE
            try
            {
                choice = int.Parse(Console.ReadLine() ?? string.Empty);
            }
            catch (System.FormatException)
            {
                // Set to invalid value
                choice = -1;
            }

            if (mapper.TryGetValue(choice, out var operation))
            {
                operation.action.Invoke();
            }
        }
    }
    // </MakeGraphCallSnippet>
}
