using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace Outlook.App.ViewModels;

/// <summary>The Settings window: edits a copy of the preferences, applied only on Save.</summary>
public partial class SettingsViewModel : ObservableObject
{
    private readonly AppSettings _settings;
    private readonly string _accountEmail;

    public SettingsViewModel(AppSettings settings, string accountEmail, MainWindowViewModel? accounts = null)
    {
        _settings = settings;
        _accountEmail = accountEmail;
        AccountsHost = accounts;
        _theme = settings.Theme;
        _newMailNotifications = settings.NewMailNotifications;
        _closeToTray = settings.CloseToTray;
        _showAsConversations = settings.ShowAsConversations;
        _markReadWhenOpened = settings.MarkReadWhenOpened;
        _undoSendSeconds = settings.UndoSendSeconds;
        _signature = settings.SignatureFor(accountEmail);
        _signatureOnReplies = settings.SignatureOnReplies;
        _defaultReminder = AppointmentEditorViewModel.ReminderOptions.FirstOrDefault(r => r.Minutes == settings.DefaultReminderMinutes)
                           ?? AppointmentEditorViewModel.ReminderOptions[3];
    }

    public event Action? CloseRequested;

    public static string[] Themes { get; } = ["System", "Light", "Dark"];
    public static ReminderOption[] ReminderOptions => AppointmentEditorViewModel.ReminderOptions;

    /// <summary>Account changes (add, edit, remove, switch) apply at once, like Outlook's Account Settings.</summary>
    public MainWindowViewModel? AccountsHost { get; }
    public bool HasAccountsHost => AccountsHost is not null;

    /// <summary>Which page is open; Accounts is first.</summary>
    [ObservableProperty] private int _selectedTab;

    public string AccountEmail => _accountEmail;
    public bool HasAccount => _accountEmail.Length > 0;

    [ObservableProperty] private string _theme;
    [ObservableProperty] private bool _newMailNotifications;
    [ObservableProperty] private bool _closeToTray;
    [ObservableProperty] private bool _showAsConversations;
    [ObservableProperty] private bool _markReadWhenOpened;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(UndoSendText))]
    private int _undoSendSeconds;

    [ObservableProperty] private string _signature;
    [ObservableProperty] private bool _signatureOnReplies;
    [ObservableProperty] private ReminderOption _defaultReminder;

    public string UndoSendText => UndoSendSeconds == 0 ? "Off: messages are sent immediately" : $"{UndoSendSeconds} seconds to undo";

    [RelayCommand]
    private void Save()
    {
        _settings.Theme = Theme;
        _settings.NewMailNotifications = NewMailNotifications;
        _settings.CloseToTray = CloseToTray;
        _settings.ShowAsConversations = ShowAsConversations;
        _settings.MarkReadWhenOpened = MarkReadWhenOpened;
        _settings.UndoSendSeconds = Math.Clamp(UndoSendSeconds, 0, 10);
        if (HasAccount)
        {
            if (string.IsNullOrWhiteSpace(Signature)) _settings.Signatures.Remove(_accountEmail);
            else _settings.Signatures[_accountEmail] = Signature.TrimEnd();
        }
        _settings.SignatureOnReplies = SignatureOnReplies;
        _settings.DefaultReminderMinutes = DefaultReminder.Minutes;
        _settings.SaveAndNotify();
        CloseRequested?.Invoke();
    }

    [RelayCommand]
    private void Cancel() => CloseRequested?.Invoke();
}
