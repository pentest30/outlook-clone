using Outlook.Core.Models;

namespace Outlook.App.ViewModels;

/// <summary>One account in the Accounts list and the account switcher.</summary>
public sealed record AccountRowViewModel(Account Account, bool IsActive)
{
    public string Email => Account.Email;
    public string Name => Account.DisplayName;
    public string Servers => $"{Account.ImapHost} · {Account.SmtpHost}";
}
