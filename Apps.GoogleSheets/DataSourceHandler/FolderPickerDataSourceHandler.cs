using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.GoogleSheets.DataSourceHandler
{
    public class FolderPickerDataSourceHandler
        : BaseInvocable, IAsyncFileDataSourceItemHandler
    {
        private const string FolderMime = "application/vnd.google-apps.folder";

        private const string MyDriveVirtualId = "v:mydrive";
        private const string SharedDrivesVirtualId = "v:shared";
        private const string SharedWithMeVirtualId = "v:sharedwithme";

        private const string MyDriveDisplay = "My Drive";
        private const string SharedDrivesDisplay = "Shared drives";
        private const string SharedWithMeDisplay = "Shared with me";

        private const string HomeVirtualId = "v:home";
        private const string HomeDisplay = "Google Drive";

        private IEnumerable<AuthenticationCredentialsProvider> Creds =>
           InvocationContext.AuthenticationCredentialsProviders;

        private GoogleDriveClient? _client;
        private GoogleDriveClient Client => _client ??= new GoogleDriveClient(Creds);

        public FolderPickerDataSourceHandler(InvocationContext invocationContext)
          : base(invocationContext)
        {
        }

        public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(FolderContentDataSourceContext context, CancellationToken cancellationToken)
        {
            var folderId = string.IsNullOrEmpty(context.FolderId) ? HomeVirtualId : context.FolderId;

            if (folderId == HomeVirtualId)
            {
                return new List<FileDataItem>
                {
                    new Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.Folder { Id = MyDriveVirtualId, DisplayName = MyDriveDisplay, IsSelectable = false },
                    new Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.Folder { Id = SharedDrivesVirtualId, DisplayName = SharedDrivesDisplay, IsSelectable = false },
                    new Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.Folder { Id = SharedWithMeVirtualId, DisplayName = SharedWithMeDisplay, IsSelectable = false }
                };
            }

            if (folderId == MyDriveVirtualId)
            {
                var items = await ListItemsInFolderByIdAsync("root", cancellationToken);
                return items
                    .Where(f => string.Equals(f.MimeType, FolderMime, StringComparison.OrdinalIgnoreCase))
                    .Select(ToFolder)
                    .Cast<FileDataItem>()
                    .ToList();
            }

            if (folderId == SharedDrivesVirtualId)
            {
                var drives = await ListSharedDrivesAsync(cancellationToken);
                return drives
                    .Select(d => new Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.Folder
                    {
                        Id = $"d:{d.Id}",
                        DisplayName = d.Name,
                        IsSelectable = false
                    })
                    .Cast<FileDataItem>()
                    .ToList();
            }

            if (folderId.StartsWith("d:", StringComparison.Ordinal))
            {
                var driveId = folderId.Substring(2);
                var items = await ListItemsInSharedDriveRootAsync(driveId, cancellationToken);

                return items
                    .Where(f => string.Equals(f.MimeType, FolderMime, StringComparison.OrdinalIgnoreCase))
                    .Select(ToFolder)
                    .Cast<FileDataItem>()
                    .ToList();
            }

            if (folderId == SharedWithMeVirtualId)
            {
                var sharedFolders = await ListSharedWithMeFoldersAsync(cancellationToken);
                return sharedFolders
                    .Select(ToFolder)
                    .Cast<FileDataItem>()
                    .ToList();
            }


            var children = await ListItemsInFolderByIdAsync(folderId, cancellationToken);
            return children
                .Where(f => string.Equals(f.MimeType, FolderMime, StringComparison.OrdinalIgnoreCase))
                .Select(ToFolder)
                .Cast<FileDataItem>()
                .ToList();
        }

        public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(context?.FileDataItemId))
                return new List<FolderPathItem> { new() { DisplayName = HomeDisplay, Id = HomeVirtualId } };

            var id = context.FileDataItemId;

            if (id == MyDriveVirtualId)
                return new List<FolderPathItem> { new() { DisplayName = HomeDisplay, Id = HomeVirtualId }, new() { DisplayName = MyDriveDisplay, Id = MyDriveVirtualId } };
            if (id == SharedDrivesVirtualId)
                return new List<FolderPathItem> { new() { DisplayName = HomeDisplay, Id = HomeVirtualId }, new() { DisplayName = SharedDrivesDisplay, Id = SharedDrivesVirtualId } };
            if (id == SharedWithMeVirtualId)
                return new List<FolderPathItem> { new() { DisplayName = HomeDisplay, Id = HomeVirtualId }, new() { DisplayName = SharedWithMeDisplay, Id = SharedWithMeVirtualId } };

            try
            {
                var current = await GetFileMetadataByIdAsync(id, cancellationToken);

                if (!string.IsNullOrEmpty(current.DriveId))
                {
                    var drive = await GetDriveAsync(current.DriveId, cancellationToken);
                    var path = new List<FolderPathItem>
                    {
                        new() { DisplayName = HomeDisplay, Id = HomeVirtualId },
                        new() { DisplayName = SharedDrivesDisplay, Id = SharedDrivesVirtualId },
                        new() { DisplayName = drive.Name, Id = $"d:{drive.Id}" }
                    };

                    var parentId = current.Parents?.FirstOrDefault();
                    var stack = new Stack<FolderPathItem>();
                    while (!string.IsNullOrEmpty(parentId) && parentId != drive.Id)
                    {
                        var parent = await GetFileMetadataByIdAsync(parentId!, cancellationToken);
                        stack.Push(new FolderPathItem { DisplayName = parent.Name, Id = parent.Id });
                        parentId = parent.Parents?.FirstOrDefault();
                    }

                    path.AddRange(stack);
                    return path;
                }

                if (await IsUnderMyDriveAsync(current, cancellationToken))
                {
                    var path = new List<FolderPathItem>
                    {
                        new() { DisplayName = HomeDisplay, Id = HomeVirtualId },
                        new() { DisplayName = MyDriveDisplay, Id = MyDriveVirtualId }
                    };
                    var parentId = current.Parents?.FirstOrDefault();
                    var stack = new Stack<FolderPathItem>();

                    while (!string.IsNullOrEmpty(parentId) && parentId != "root")
                    {
                        var parent = await GetFileMetadataByIdAsync(parentId!, cancellationToken);
                        stack.Push(new FolderPathItem { DisplayName = parent.Name, Id = parent.Id });
                        parentId = parent.Parents?.FirstOrDefault();
                    }

                    path.AddRange(stack);
                    return path;
                }
                else
                {
                    var path = new List<FolderPathItem>
                    {
                        new() { DisplayName = HomeDisplay, Id = HomeVirtualId },
                    };

                    var parentId = current.Parents?.FirstOrDefault();
                    var stack = new Stack<FolderPathItem>();

                    while (!string.IsNullOrEmpty(parentId))
                    {
                        var parent = await GetFileMetadataByIdAsync(parentId!, cancellationToken);
                        if (!string.IsNullOrEmpty(parent.DriveId))
                        {
                            var drive = await GetDriveAsync(parent.DriveId, cancellationToken);
                            path.Add(new FolderPathItem { DisplayName = drive.Name, Id = $"d:{drive.Id}" });
                            break;
                        }

                        stack.Push(new FolderPathItem { DisplayName = parent.Name, Id = parent.Id });
                        parentId = parent.Parents?.FirstOrDefault();
                    }

                    path.AddRange(stack);
                    return path;
                }
            }
            catch
            {
                return new List<FolderPathItem> { new() { DisplayName = HomeDisplay, Id = HomeVirtualId } };
            }
        }

        private static Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.Folder ToFolder(Google.Apis.Drive.v3.Data.File f) => new()
        {
            Id = f.Id,
            DisplayName = f.Name,
            Date = f.CreatedTime,
            IsSelectable = true
        };

        private async Task<bool> IsUnderMyDriveAsync(Google.Apis.Drive.v3.Data.File file, CancellationToken ct)
        {
            var parentId = file.Parents?.FirstOrDefault();
            while (!string.IsNullOrEmpty(parentId))
            {
                if (parentId == "root") return true;
                var parent = await GetFileMetadataByIdAsync(parentId!, ct);
                if (!string.IsNullOrEmpty(parent.DriveId)) return false;
                parentId = parent.Parents?.FirstOrDefault();
            }
            return false;
        }

        private async Task<IList<Google.Apis.Drive.v3.Data.File>> ListItemsInFolderByIdAsync(string folderId, CancellationToken ct)
        {
            var files = new List<Google.Apis.Drive.v3.Data.File>();
            string? pageToken = null;

            do
            {
                var req = Client.Files.List();
                req.Q = $"'{folderId}' in parents and trashed = false";
                req.IncludeItemsFromAllDrives = true;
                req.SupportsAllDrives = true;
                req.Spaces = "drive";
                req.Fields = "nextPageToken, files(id, name, mimeType, parents, createdTime, modifiedTime, driveId)";
                req.PageSize = 100;
                req.PageToken = pageToken;

                var resp = await req.ExecuteAsync(ct);
                if (resp.Files is { Count: > 0 }) files.AddRange(resp.Files);
                pageToken = resp.NextPageToken;
            } while (!string.IsNullOrEmpty(pageToken));

            return files;
        }

        private async Task<IList<Google.Apis.Drive.v3.Data.Drive>> ListSharedDrivesAsync(CancellationToken ct)
        {
            var drives = new List<Google.Apis.Drive.v3.Data.Drive>();
            string? pageToken = null;

            do
            {
                var req = Client.Drives.List();
                req.PageSize = 100;
                req.UseDomainAdminAccess = false;
                req.Fields = "nextPageToken, drives(id, name)";
                req.PageToken = pageToken;

                var resp = await req.ExecuteAsync(ct);
                if (resp.Drives is { Count: > 0 }) drives.AddRange(resp.Drives);
                pageToken = resp.NextPageToken;
            } while (!string.IsNullOrEmpty(pageToken));

            return drives;
        }

        private async Task<IList<Google.Apis.Drive.v3.Data.File>> ListSharedWithMeFoldersAsync(CancellationToken ct)
        {
            var files = new List<Google.Apis.Drive.v3.Data.File>();
            string? pageToken = null;

            do
            {
                var req = Client.Files.List();
                req.Q = $"sharedWithMe = true and trashed = false and mimeType = '{FolderMime}'";
                req.IncludeItemsFromAllDrives = true;
                req.SupportsAllDrives = true;
                req.Spaces = "drive";
                req.Fields = "nextPageToken, files(id, name, mimeType, parents, createdTime, modifiedTime, driveId)";
                req.PageSize = 100;
                req.PageToken = pageToken;

                var resp = await req.ExecuteAsync(ct);
                if (resp.Files is { Count: > 0 }) files.AddRange(resp.Files);
                pageToken = resp.NextPageToken;
            } while (!string.IsNullOrEmpty(pageToken));

            return files;
        }

        private async Task<IList<Google.Apis.Drive.v3.Data.File>> ListItemsInSharedDriveRootAsync(string driveId, CancellationToken ct)
        {
            var files = new List<Google.Apis.Drive.v3.Data.File>();
            string? pageToken = null;

            do
            {
                var req = Client.Files.List();
                req.Corpora = "drive";
                req.DriveId = driveId;
                req.Q = $"'{driveId}' in parents and trashed = false";
                req.IncludeItemsFromAllDrives = true;
                req.SupportsAllDrives = true;
                req.Spaces = "drive";
                req.Fields = "nextPageToken, files(id, name, mimeType, parents, createdTime, modifiedTime, driveId)";
                req.PageSize = 100;
                req.PageToken = pageToken;

                var resp = await req.ExecuteAsync(ct);
                if (resp.Files is { Count: > 0 }) files.AddRange(resp.Files);
                pageToken = resp.NextPageToken;
            } while (!string.IsNullOrEmpty(pageToken));

            return files;
        }

        private async Task<Google.Apis.Drive.v3.Data.File> GetFileMetadataByIdAsync(string fileId, CancellationToken ct)
        {
            var req = Client.Files.Get(fileId);
            req.SupportsAllDrives = true;
            req.Fields = "id, name, mimeType, parents, createdTime, modifiedTime, driveId";
            return await req.ExecuteAsync(ct);
        }

        private async Task<Google.Apis.Drive.v3.Data.Drive> GetDriveAsync(string driveId, CancellationToken ct)
        {
            var req = Client.Drives.Get(driveId);
            req.Fields = "id, name";
            return await req.ExecuteAsync(ct);
        }
    }
}
