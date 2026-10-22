using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Models;
using Outlook.Core.Services;
using Outlook.Mail;

namespace Outlook.App.ViewModels;

/// <summary>
/// Add or edit an account. Typing the address looks the servers up (see <see cref="ServerDiscovery"/>),
/// so most people only enter address and password; the server fields stay available under
/// "Server settings". Nothing is saved until both servers accept the login.
/// </summary>
public partial class AccountSetupViewModel : ObservableObject
{
    private readonly IMailService _mail;
    private readonly IMailStore _store;
    private readonly ICredentialStore _credentials;
    private readonly Func<string, CancellationToken, Task<DiscoveredSettings?>> _discover;
    private readonly Func<Account, bool, Task> _onSaved;
    private readonly Account? _existing;
    private readonly DispatcherTimer _discoverDelay = new() { Interval = TimeSpan.FromMilliseconds(700) };
    private CancellationTokenSource? _discovering;
    private bool _serversEditedByHand;
    private bool _filling;

    /// <param name="onSaved">Called with the saved account and whether it is new.</param>
    public AccountSetupViewModel(IMailService mail, IMailStore store, ICredentialStore credentials,
        Func<string, CancellationToken, Task<DiscoveredSettings?>> discover, Func<Account, bool, Task> onSaved,
        Account? existing = null, bool canCancel = true)
    {
        _mail = mail;
        _store = store;
        _credentials = credentials;
        _discover = discover;
        _onSaved = onSaved;
        _existing = existing;
        CanCancel = canCancel;
        _discoverDelay.Tick += async (_, _) =>
        {
            _discoverDelay.Stop();
            await DiscoverAsync();
        };

        if (existing is not null)
        {
            _filling = true;
            DisplayName = existing.DisplayName;
            Email = existing.Email;
            UserName = existing.UserName;
            ImapHost = existing.ImapHost; ImapPort = existing.ImapPort; ImapSecurity = existing.ImapSecurity;
            SmtpHost = existing.SmtpHost; SmtpPort = existing.SmtpPort; SmtpSecurity = existing.SmtpSecurity;
            _filling = false;
            _serversEditedByHand = true;   // keep what is saved; do not re-detect over it
            ShowServerSettings = true;
        }
    }

    public event Action? CloseRequested;

    public static TransportSecurity[] SecurityOptions { get; } = Enum.GetValues<TransportSecurity>();

    public bool IsNew => _existing is null;
    public bool CanCancel { get; }
    public string Heading => IsNew ? "Add an account" : "Account settings";
    public string SaveLabel => IsNew ? "Sign in" : "Save";
    public string PasswordHint => IsNew ? "Password" : "Password (leave empty to keep the current one)";

#if DEBUG
    public bool ShowDevShortcut => IsNew;
#else
    public bool ShowDevShortcut => false;
#endif

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

    /// <summary>"Found settings for OVH." and the like, under the address.</summary>
    [ObservableProperty] private string _discoveryNote = "";
    [ObservableProperty] private bool _isDiscovering;
    [ObservableProperty] private bool _showServerSettings;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SignInCommand))]
    private bool _isBusy;

    partial void OnEmailChanged(string? oldValue, string newValue)
    {
        if (_filling) return;
        // The login is usually the address itself.
        if (string.IsNullOrEmpty(UserName) || UserName == oldValue)
            UserName = newValue;
        if (_serversEditedByHand) return;
        _discoverDelay.Stop();
        DiscoveryNote = "";
        if (LooksLikeAddress(newValue))
            _discoverDelay.Start();   // look up once typing pauses
    }

    // Typing in a server field means "use what I typed", so detection stops overwriting it.
    partial void OnImapHostChanged(string value) => MarkEdited();
    partial void OnSmtpHostChanged(string value) => MarkEdited();
    partial void OnImapPortChanged(int value) => MarkEdited();
    partial void OnSmtpPortChanged(int value) => MarkEdited();

    private void MarkEdited()
    {
        if (!_filling) _serversEditedByHand = true;
    }

    private static bool LooksLikeAddress(string value)
    {
        var at = value.IndexOf('@');
        return at > 0 && value.IndexOf('.', at) > at + 1 && !value.EndsWith('.');
    }

    /// <summary>Looks the servers up for the typed address and fills them in.</summary>
    [RelayCommand]
    public async Task DiscoverAsync()
    {
        if (!LooksLikeAddress(Email)) return;
        _discovering?.Cancel();
        var cts = _discovering = new CancellationTokenSource();
        var email = Email.Trim();
        IsDiscovering = true;
        DiscoveryNote = "Looking up your mail servers…";
        try
        {
            var found = await _discover(email, cts.Token);
            if (cts.IsCancellationRequested || email != Email.Trim()) return;
            if (found is null)
            {
                DiscoveryNote = "Couldn't find the servers automatically. Enter them under Server settings (your provider's help pages list them).";
                ShowServerSettings = true;
                return;
            }
            _filling = true;
            ImapHost = found.ImapHost; ImapPort = found.ImapPort; ImapSecurity = found.ImapSecurity;
            SmtpHost = found.SmtpHost; SmtpPort = found.SmtpPort; SmtpSecurity = found.SmtpSecurity;
            UserName = found.UserName;
            _filling = false;
            DiscoveryNote = found.NeedsOAuth
                ? $"{found.Provider} needs Microsoft/Google sign-in, which this app does not support yet."
                : $"Found the settings for {found.Provider} ({found.ImapHost}).";
        }
        catch (OperationCanceledException)
        {
            // a newer lookup replaced this one
        }
        finally
        {
            if (_discovering == cts) IsDiscovering = false;
        }
    }

    [RelayCommand]
    private void ToggleServerSettings() => ShowServerSettings = !ShowServerSettings;

    /// <summary>Fills in the GreenMail server from docker-compose.yml (debug builds only).</summary>
    [RelayCommand]
    private void UseLocalTestServer()
    {
        _filling = true;
        DisplayName = "Dev User";
        Email = "dev@localhost";
        UserName = "dev@localhost";
        Password = "dev";
        ImapHost = "localhost"; ImapPort = 3143; ImapSecurity = TransportSecurity.None;
        SmtpHost = "localhost"; SmtpPort = 3025; SmtpSecurity = TransportSecurity.None;
        _filling = false;
        _serversEditedByHand = true;
        ShowServerSettings = true;
    }

    [RelayCommand]
    private void Cancel()
    {
        _discovering?.Cancel();
        CloseRequested?.Invoke();
    }

    private bool CanSignIn() => !IsBusy;

    [RelayCommand(CanExecute = nameof(CanSignIn))]
    private async Task SignInAsync()
    {
        Error = null;
        if (string.IsNullOrWhiteSpace(Email))
        {
            Error = "Enter your e-mail address.";
            return;
        }
        if (string.IsNullOrWhiteSpace(ImapHost) || string.IsNullOrWhiteSpace(SmtpHost))
        {
            // The lookup may still be running, or has not run because the address was pasted at the last moment.
            await DiscoverAsync();
            if (string.IsNullOrWhiteSpace(ImapHost) || string.IsNullOrWhiteSpace(SmtpHost))
            {
                Error = "Enter the incoming and outgoing servers under Server settings.";
                ShowServerSettings = true;
                return;
            }
        }

        var password = Password.Length > 0 ? Password : _existing is null ? "" : _credentials.Load(_existing.Id) ?? "";
        if (password.Length == 0)
        {
            Error = "Enter your password.";
            return;
        }

        var account = new Account
        {
            Id = _existing?.Id ?? 0,
            DisplayName = string.IsNullOrWhiteSpace(DisplayName) ? Email.Trim() : DisplayName.Trim(),
            Email = Email.Trim(),
            UserName = string.IsNullOrWhiteSpace(UserName) ? Email.Trim() : UserName.Trim(),
            ImapHost = ImapHost.Trim(), ImapPort = ImapPort, ImapSecurity = ImapSecurity,
            SmtpHost = SmtpHost.Trim(), SmtpPort = SmtpPort, SmtpSecurity = SmtpSecurity,
        };

        IsBusy = true;
        try
        {
            try
            {
                await _mail.VerifyAsync(account, password);
            }
            catch (Exception ex)
            {
                Error = $"Couldn't sign in to {account.ImapHost}: {ex.Message}";
                ShowServerSettings = true;
                return;
            }
            try
            {
                await _mail.VerifySmtpAsync(account, password);
            }
            catch (Exception ex)
            {
                Error = $"Signed in, but sending through {account.SmtpHost} failed: {ex.Message}";
                ShowServerSettings = true;
                return;
            }

            if (_existing is null)
            {
                if ((await _store.GetAccountsAsync()).Any(a => string.Equals(a.Email, account.Email, StringComparison.OrdinalIgnoreCase)))
                {
                    Error = $"{account.Email} is already set up.";
                    return;
                }
                account = await _store.AddAccountAsync(account);
            }
            else
            {
                await _store.UpdateAccountAsync(account);
                await _mail.ForgetAccountAsync(account.Id);   // reconnect with the new settings
            }
            _credentials.Save(account.Id, password);
            Password = "";
            await _onSaved(account, _existing is null);
            CloseRequested?.Invoke();
        }
        finally
        {
            IsBusy = false;
        }
    }
}
