namespace Apps.GoogleSheets.Constants;

public static class Urls
{
    private const string GoogleApisOauth = "https://oauth2.googleapis.com";

    public const string Auth = "https://accounts.google.com/o/oauth2/v2/auth";
    public const string Token = GoogleApisOauth + "/token";
    public const string Revoke = GoogleApisOauth + "/revoke";
}