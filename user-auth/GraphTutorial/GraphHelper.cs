// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Azure.Core;
using Azure.Identity;

using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
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

        mapper.Add(3, ("OneDrive Root Delete", async () => {
            _ = _userClient ??
                throw new System.NullReferenceException();
            var driveItem = await _userClient.Me.Drive.GetAsync();
            _ = driveItem ??
              throw new System.NullReferenceException();

            var userDriveId = driveItem.Id;
            // List children in the drive
            var driveRequest = _userClient.Drives[userDriveId];
            var root = await driveRequest.Root.GetAsync();

            await driveRequest.Items[root.Id].ItemWithPath("test.txt").DeleteAsync();
        }
        ));

        mapper.Add(4, ("OneDrive Root Download", async () => {
            _ = _userClient ??
                throw new NullReferenceException();
            var driveItem = await _userClient.Me.Drive.GetAsync();
            _ = driveItem ??
              throw new NullReferenceException();

            var userDriveId = driveItem.Id;
            // List children in the drive
            var driveRequest = _userClient.Drives[userDriveId];
            var root = await driveRequest.Root.GetAsync();


            var task = driveRequest.Items[root.Id].ItemWithPath("test.txt").Content.GetAsync();

            using var stream = await task;

            _ = stream ??
                throw new NullReferenceException();
            using var streamWriter = new StreamWriter("test.txt");
            await stream.CopyToAsync(streamWriter.BaseStream);
            await Console.Out.WriteLineAsync("Download complete!");
            System.Diagnostics.Process.Start("explorer.exe", Directory.GetCurrentDirectory());
        }
        ));

        mapper.Add(5, ("OneDrive Root Upload Large Files ", async () => {
            _ = _userClient ??
                throw new NullReferenceException();
            var driveItem = await _userClient.Me.Drive.GetAsync();
            _ = driveItem ??
              throw new NullReferenceException();

            var bytes = new byte[1024 * 1024 * 10];
            var fileStream = new System.IO.MemoryStream(bytes);
            var newBytes = Encoding.UTF8.GetBytes(@"The contents of the file goes here.");
            fileStream.Write(newBytes, 0, newBytes.Length);

            var userDriveId = driveItem.Id;
            // List children in the drive
            var driveRequest = _userClient.Drives[userDriveId];
            var root = await driveRequest.Root.GetAsync();

            // Use properties to specify the conflict behavior
            // in this case, replace

            var uploadSessionRequestBody = new CreateUploadSessionPostRequestBody
            {
                Item = new DriveItemUploadableProperties
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" }
                    }
                }
            };
            var uploadSession = await driveRequest.Items[root.Id].ItemWithPath("TestSizeFile.txt").CreateUploadSession.PostAsync(uploadSessionRequestBody);


            // Max slice size must be a multiple of 320 KiB
            int maxSliceSize = 320 * 1024;
            var fileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize, _userClient.RequestAdapter);

            var totalLength = fileStream.Length;
            // Create a callback that is invoked after each slice is uploaded
            IProgress<long> progress = new Progress<long>(prog => {
                Console.WriteLine($"Uploaded {prog} bytes of {totalLength} bytes");
            });

            try
            {
                // Upload the file
                var uploadResult = await fileUploadTask.UploadAsync(progress);

                Console.WriteLine(uploadResult.UploadSucceeded ?
                    $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                    "Upload failed");
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error uploading: {ex.ToString()}");
            }
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
