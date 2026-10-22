using Microsoft.EntityFrameworkCore.Design;

namespace Outlook.Data;

/// <summary>Lets `dotnet ef migrations add` build the context without starting the app.</summary>
internal sealed class DesignTimeContextFactory : IDesignTimeDbContextFactory<MailDbContext>
{
    public MailDbContext CreateDbContext(string[] args) => MailDbContext.Open("design-time.db");
}
