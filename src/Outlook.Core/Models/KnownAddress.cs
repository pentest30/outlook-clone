namespace Outlook.Core.Models;

/// <summary>
/// Someone this account has exchanged mail with: Outlook's "auto-complete list". Built from senders
/// of received mail and recipients of sent mail, and used to suggest addresses while typing.
/// </summary>
public class KnownAddress
{
    public int Id { get; set; }
    public int AccountId { get; set; }
    /// <summary>Lower-case, the identity of the entry.</summary>
    public string Address { get; set; } = "";
    public string Name { get; set; } = "";
    public int SentCount { get; set; }
    public int ReceivedCount { get; set; }
    public DateTimeOffset LastUsedUtc { get; set; }

    public string Display => string.IsNullOrWhiteSpace(Name) ? Address : $"{Name} <{Address}>";
}
