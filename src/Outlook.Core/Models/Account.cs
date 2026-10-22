namespace Outlook.Core.Models;

public enum TransportSecurity
{
    None,
    StartTls,
    SslOnConnect,
}

public class Account
{
    public int Id { get; set; }
    public string DisplayName { get; set; } = "";
    public string Email { get; set; } = "";
    public string UserName { get; set; } = "";

    public string ImapHost { get; set; } = "";
    public int ImapPort { get; set; } = 993;
    public TransportSecurity ImapSecurity { get; set; } = TransportSecurity.SslOnConnect;

    public string SmtpHost { get; set; } = "";
    public int SmtpPort { get; set; } = 587;
    public TransportSecurity SmtpSecurity { get; set; } = TransportSecurity.StartTls;

    public List<MailFolder> Folders { get; set; } = [];
}
