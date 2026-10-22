using Microsoft.EntityFrameworkCore;
using Outlook.Core.Compose;
using Outlook.Core.Models;

namespace Outlook.Data;

/// <summary>
/// Batches updates to the auto-complete list: count each person once per message, keep the best
/// known display name, never list the account itself or no-reply robots.
/// </summary>
public sealed class AddressBook(MailDbContext db, Account account)
{
    private readonly Dictionary<string, KnownAddress> _pending = new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<string, KnownAddress>? _existing;

    /// <summary>Recipients of a message from its formatted To and Cc lines.</summary>
    public static IEnumerable<(string Name, string Address)> People(params string?[] lines) =>
        lines.SelectMany(AddressList.Split).Select(AddressList.Parse);

    public void Record(IEnumerable<(string Name, string Address)> people, bool sent, DateTimeOffset when)
    {
        foreach (var (name, rawAddress) in people)
        {
            var address = rawAddress.Trim().ToLowerInvariant();
            if (!IsUseful(address)) continue;

            var entry = Find(address);
            if (sent) entry.SentCount++;
            else entry.ReceivedCount++;
            if (when > entry.LastUsedUtc) entry.LastUsedUtc = when;
            // Prefer a real name over nothing, and the newest name over an old one.
            if (!string.IsNullOrWhiteSpace(name) && !name.Contains('@') && (entry.Name.Length == 0 || when >= entry.LastUsedUtc))
                entry.Name = name.Trim();
        }
    }

    private bool IsUseful(string address)
    {
        if (address.Length == 0 || !address.Contains('@')) return false;
        if (string.Equals(address, account.Email, StringComparison.OrdinalIgnoreCase)) return false;
        var local = address[..address.IndexOf('@')];
        return !(local.Contains("noreply") || local.Contains("no-reply") || local.Contains("donotreply")
                 || local.Contains("do-not-reply") || local is "mailer-daemon" or "postmaster");
    }

    private KnownAddress Find(string address)
    {
        if (_pending.TryGetValue(address, out var entry)) return entry;
        _existing ??= db.KnownAddresses.Where(a => a.AccountId == account.Id).ToDictionary(a => a.Address, StringComparer.OrdinalIgnoreCase);
        if (!_existing.TryGetValue(address, out entry))
        {
            entry = new KnownAddress { AccountId = account.Id, Address = address };
            db.KnownAddresses.Add(entry);
        }
        _pending[address] = entry;
        return entry;
    }

    public Task SaveAsync(CancellationToken ct) => db.SaveChangesAsync(ct);
}
