using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Models;
using Outlook.Core.Services;

namespace Outlook.App.ViewModels;

/// <summary>"Add account" form. Validates by logging in before anything is saved.</summary>
public partial class AccountSetupViewModel(IMailService mail, IMailStore store, ICredentialStore credentials,
    Func<Account, Task> onAdded) : ObservableObject
{
    public static TransportSecurity[] SecurityOptions { get; } = Enum.GetValues<TransportSecurity>();

    [ObservableProperty] private string _displayName = "";
    [ObservableProperty] private string _email = "";
    [ObservableProperty] private string _userName = "";
    [ObservableProperty] private string _password = "";
    [ObservableProperty] private string _imapHost = "";
    [ObservableProperty] private int _imapPort = 993;
    [ObservableProperty] private TransportSecurity _imapSecurity = TransportSecurity.SslOnConnect;
    [ObservableProperty] private string _smtpHost = "";
    [ObservableProperty] private int _smtpPort = 587;
    [ObservableProperty] private TransportSecurity _smtpSecurity = TransportSecurity.StartTls;
    [ObservableProperty] private string? _error;

    partial void OnEmailChanged(string? oldValue, string newValue)
    {
        // Most providers use the address as the login and imap./smtp. hosts; prefill those as the user types.
        if (string.IsNullOrEmpty(UserName) || UserName == oldValue)
            UserName = newValue;
        var at = newValue.IndexOf('@');
        if (at < 0) return;
        var domain = newValue[(at + 1)..];
        var oldDomain = oldValue?.Contains('@') == true ? oldValue[(oldValue.IndexOf('@') + 1)..] : null;
        if (string.IsNullOrEmpty(ImapHost) || ImapHost == $"imap.{oldDomain}") ImapHost = $"imap.{domain}";
        if (string.IsNullOrEmpty(SmtpHost) || SmtpHost == $"smtp.{oldDomain}") SmtpHost = $"smtp.{domain}";
    }

    /// <summary>Fills in the GreenMail server from docker-compose.yml.</summary>
    [RelayCommand]
    private void UseLocalTestServer()
    {
        DisplayName = "Dev User";
        Email = "dev@localhost";
        UserName = "dev@localhost";
        Password = "dev";
        ImapHost = "localhost"; ImapPort = 3143; ImapSecurity = TransportSecurity.None;
        SmtpHost = "localhost"; SmtpPort = 3025; SmtpSecurity = TransportSecurity.None;
    }

    [RelayCommand]
    private async Task SignInAsync()
    {
        Error = null;
        if (string.IsNullOrWhiteSpace(Email) || string.IsNullOrWhiteSpace(ImapHost) || string.IsNullOrWhiteSpace(SmtpHost))
        {
            Error = "Email, IMAP server and SMTP server are required.";
            return;
        }

        var account = new Account
        {
            DisplayName = string.IsNullOrWhiteSpace(DisplayName) ? Email : DisplayName,
            Email = Email.Trim(),
            UserName = string.IsNullOrWhiteSpace(UserName) ? Email.Trim() : UserName.Trim(),
            ImapHost = ImapHost.Trim(), ImapPort = ImapPort, ImapSecurity = ImapSecurity,
            SmtpHost = SmtpHost.Trim(), SmtpPort = SmtpPort, SmtpSecurity = SmtpSecurity,
        };

        try
        {
            await mail.VerifyAsync(account, Password);
        }
        catch (Exception ex)
        {
            Error = $"Couldn't sign in: {ex.Message}";
            return;
        }

        account = await store.AddAccountAsync(account);
        credentials.Save(account.Id, Password);
        Password = "";
        await onAdded(account);
    }
}
