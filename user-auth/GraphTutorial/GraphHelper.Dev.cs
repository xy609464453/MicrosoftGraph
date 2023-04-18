using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateLink;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;

using System.Text;

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

        mapper.Add(5, ("OneDrive Root Upload Large Files", async () => {
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

        mapper.Add(6, ("OneDrive Create Link", async () => {
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

            var type = "view";

            var password = "ThisIsMyPrivatePassword";

            var scope = "anonymous";

            var link = driveRequest.Items[root.Id].ItemWithPath("TestSizeFile.txt").CreateLink;
            var permission = await link.PostAsync(new CreateLinkPostRequestBody
            {
                Type = type,
                Scope = scope,
                Password = password,
            });

            Console.WriteLine("Link created successfully");
            Console.WriteLine(permission.Link.WebUrl);
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