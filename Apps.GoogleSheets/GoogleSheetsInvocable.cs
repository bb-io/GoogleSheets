using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.GoogleSheets;

public class GoogleSheetsInvocable : BaseInvocable
{
    protected AuthenticationCredentialsProvider[] Creds => InvocationContext.AuthenticationCredentialsProviders.ToArray();
    protected GoogleSheetsClient Client { get; }
    protected GoogleDriveClient DriveClient { get; }

    public GoogleSheetsInvocable(InvocationContext invocationContext) : base(invocationContext)
    {
        Client = new(Creds);
        DriveClient = new(Creds);
    }
}