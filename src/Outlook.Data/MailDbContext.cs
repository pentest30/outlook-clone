using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using Outlook.Core.Models;

namespace Outlook.Data;

public class MailDbContext(DbContextOptions<MailDbContext> options) : DbContext(options)
{
    public DbSet<Account> Accounts => Set<Account>();
    public DbSet<MailFolder> Folders => Set<MailFolder>();
    public DbSet<MailMessage> Messages => Set<MailMessage>();
    public DbSet<MailAttachment> Attachments => Set<MailAttachment>();
    public DbSet<Category> Categories => Set<Category>();
    public DbSet<Appointment> Appointments => Set<Appointment>();
    public DbSet<KnownAddress> KnownAddresses => Set<KnownAddress>();

    protected override void OnModelCreating(ModelBuilder b)
    {
        // SQLite cannot sort DateTimeOffset, so store UTC ticks: ORDER BY Date works and stays exact.
        var utcTicks = new ValueConverter<DateTimeOffset, long>(
            v => v.UtcTicks,
            v => new DateTimeOffset(v, TimeSpan.Zero));
        var nullableUtcTicks = new ValueConverter<DateTimeOffset?, long?>(
            v => v.HasValue ? v.Value.UtcTicks : null,
            v => v.HasValue ? new DateTimeOffset(v.Value, TimeSpan.Zero) : null);

        b.Entity<Account>(e =>
        {
            e.HasIndex(a => a.Email).IsUnique();
            e.HasMany(a => a.Folders).WithOne(f => f.Account).HasForeignKey(f => f.AccountId).OnDelete(DeleteBehavior.Cascade);
        });

        b.Entity<MailFolder>(e =>
        {
            e.HasIndex(f => new { f.AccountId, f.FullName }).IsUnique();
            e.Ignore(f => f.SortRank);
            e.Property(f => f.LastSyncedAt).HasConversion(nullableUtcTicks);
            e.HasMany(f => f.Messages).WithOne(m => m.Folder).HasForeignKey(m => m.FolderId).OnDelete(DeleteBehavior.Cascade);
        });

        b.Entity<MailMessage>(e =>
        {
            e.HasIndex(m => new { m.FolderId, m.Uid }).IsUnique();
            e.HasIndex(m => new { m.FolderId, m.Date });
            e.HasIndex(m => m.ThreadId);
            e.Property(m => m.Date).HasConversion(utcTicks);
            e.Ignore(m => m.SenderDisplay);
            e.Ignore(m => m.CategoryKeywordList);
            e.HasMany(m => m.Attachments).WithOne(a => a.Message).HasForeignKey(a => a.MessageId).OnDelete(DeleteBehavior.Cascade);
        });

        b.Entity<MailAttachment>(e => e.HasIndex(a => a.MessageId));

        b.Entity<Appointment>(e =>
        {
            e.HasIndex(a => new { a.AccountId, a.Uid }).IsUnique();
            e.HasIndex(a => new { a.AccountId, a.StartUtc });
            e.Property(a => a.StartUtc).HasConversion(utcTicks);
            e.Property(a => a.EndUtc).HasConversion(utcTicks);
            e.Property(a => a.SeriesEndUtc).HasConversion(nullableUtcTicks);
            e.Property(a => a.RemindersDismissedThroughUtc).HasConversion(nullableUtcTicks);
            e.Property(a => a.LastModifiedUtc).HasConversion(utcTicks);
            e.Ignore(a => a.IsMeeting);
            e.HasOne<Account>().WithMany().HasForeignKey(a => a.AccountId).OnDelete(DeleteBehavior.Cascade);
        });

        b.Entity<KnownAddress>(e =>
        {
            e.HasIndex(a => new { a.AccountId, a.Address }).IsUnique();
            e.Property(a => a.LastUsedUtc).HasConversion(utcTicks);
            e.Ignore(a => a.Display);
            e.HasOne<Account>().WithMany().HasForeignKey(a => a.AccountId).OnDelete(DeleteBehavior.Cascade);
        });

        b.Entity<Category>(e =>
        {
            e.HasIndex(c => new { c.AccountId, c.Keyword }).IsUnique();
            e.HasOne<Account>().WithMany().HasForeignKey(c => c.AccountId).OnDelete(DeleteBehavior.Cascade);
        });
    }

    /// <summary>Opens the cache at the given SQLite file. Call <see cref="MailStore.Migrate"/> once at startup.</summary>
    public static MailDbContext Open(string databasePath)
    {
        var options = new DbContextOptionsBuilder<MailDbContext>()
            .UseSqlite($"Data Source={databasePath}")
            .Options;
        return new MailDbContext(options);
    }
}
