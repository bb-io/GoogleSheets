using Apps.GoogleSheets.Connections;
using Blackbird.Applications.Sdk.Common.Authentication;
using FluentAssertions;
using Tests.GoogleSheets.Base;

namespace Tests.GoogleSheets;

[TestClass]
public class ValidatorTests : TestBase
{
    [TestMethod]
    public async Task ValidateConnection_ValidEnterpriseConnection_ShouldBeSuccessful()
    {
        var validator = new ConnectionValidator();

        var result = await validator.ValidateConnection(Creds, CancellationToken.None);
        Console.WriteLine(result.Message);
        result.IsValid.Should().Be(true);
    }

    [TestMethod]
    public async Task ValidateConnection_InvalidConnection_ShouldFail()
    {
        var validator = new ConnectionValidator();

        var newCredentials = Creds.Select(x => new AuthenticationCredentialsProvider(x.KeyName, x.Value + "_incorrect"));
        var result = await validator.ValidateConnection(newCredentials, CancellationToken.None);
        result.IsValid.Should().Be(false);
    }
}