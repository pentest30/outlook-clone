using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Models;
using Outlook.Core.Services;

namespace Outlook.App.ViewModels;

/// <summary>
/// The meeting card above an invitation e-mail: what, when, where, plus Accept / Tentative / Decline
/// (REQUEST), "Remove from calendar" (CANCEL), or the attendee's answer (REPLY).
/// </summary>
public partial class InvitationViewModel : ObservableObject
{
    private readonly ICalendarService _calendar;
    private readonly Account _account;
    private readonly Action<string?> _changed;

    public InvitationViewModel(ICalendarService calendar, Account account, Invitation invitation, Appointment? existing, Action<string?> changed)
    {
        _calendar = calendar;
        _account = account;
        _changed = changed;
        Invitation = invitation;
        _status = StatusFor(existing);
    }

    public Invitation Invitation { get; }

    public bool IsRequest => Invitation.Method == "REQUEST" && !IsMine;
    public bool IsCancel => Invitation.Method == "CANCEL";
    public bool IsReply => Invitation.Method == "REPLY";
    /// <summary>My own invitation, seen in Sent: nothing to answer.</summary>
    private bool IsMine => string.Equals(Invitation.OrganizerEmail, _account.Email, StringComparison.OrdinalIgnoreCase);

    public string Heading => Invitation.Method switch
    {
        "CANCEL" => "Meeting canceled",
        "REPLY" => $"{Invitation.ReplyFrom} {Invitation.ReplyStatus switch { MeetingResponse.Accepted => "accepted", MeetingResponse.Tentative => "tentatively accepted", MeetingResponse.Declined => "declined", _ => "replied to" }} your meeting",
        _ => IsMine ? "Meeting invitation you sent" : "Meeting invitation",
    };

    public string Title => Invitation.Title;

    public string When
    {
        get
        {
            var culture = CultureInfo.CurrentCulture;
            if (Invitation.IsAllDay)
            {
                var first = Invitation.Start.UtcDateTime;
                var last = Invitation.End.UtcDateTime.AddDays(-1);
                return last.Date <= first.Date ? $"{first.ToString("D", culture)} (all day)" : $"{first.ToString("D", culture)} – {last.ToString("D", culture)}";
            }
            var start = Invitation.Start.ToLocalTime();
            var end = Invitation.End.ToLocalTime();
            return start.Date == end.Date
                ? $"{start.ToString("D", culture)}, {start:t} – {end:t}"
                : $"{start.ToString("f", culture)} – {end.ToString("f", culture)}";
        }
    }

    public string Where => Invitation.Location;
    public bool HasWhere => !string.IsNullOrWhiteSpace(Where);
    public string? Repeats => Invitation.RecurrenceText;
    public bool HasRepeats => Repeats is not null;
    public string Organizer => string.IsNullOrWhiteSpace(Invitation.OrganizerName) ? Invitation.OrganizerEmail : $"{Invitation.OrganizerName} <{Invitation.OrganizerEmail}>";

    [ObservableProperty] private string _status;
    [ObservableProperty] [NotifyCanExecuteChangedFor(nameof(AcceptCommand), nameof(TentativeCommand), nameof(DeclineCommand), nameof(RemoveCommand))] private bool _isBusy;

    /// <summary>Send the answer to the organizer (Outlook's "Send the response now"); off means "Don't send a response".</summary>
    [ObservableProperty] private bool _sendResponse = true;

    private string StatusFor(Appointment? existing) => Invitation.Method switch
    {
        "CANCEL" => existing is null ? "Not in your calendar." : "Still in your calendar.",
        "REPLY" => existing is null ? "" : "Recorded on your meeting.",
        _ when IsMine => "",
        _ => existing?.MyResponse switch
        {
            MeetingResponse.Accepted => "You accepted.",
            MeetingResponse.Tentative => "You tentatively accepted.",
            null or MeetingResponse.NeedsAction or MeetingResponse.None => "You haven't responded yet.",
            _ => "",
        },
    };

    private bool CanAct() => !IsBusy;

    [RelayCommand(CanExecute = nameof(CanAct))]
    private Task AcceptAsync() => RespondAsync(MeetingResponse.Accepted, "You accepted.");

    [RelayCommand(CanExecute = nameof(CanAct))]
    private Task TentativeAsync() => RespondAsync(MeetingResponse.Tentative, "You tentatively accepted.");

    [RelayCommand(CanExecute = nameof(CanAct))]
    private Task DeclineAsync() => RespondAsync(MeetingResponse.Declined, "You declined.");

    private async Task RespondAsync(MeetingResponse response, string done)
    {
        IsBusy = true;
        try
        {
            var problem = await _calendar.RespondAsync(_account, Invitation, response, SendResponse);
            Status = problem ?? done + (SendResponse ? " The organizer was told." : "");
            _changed(problem);
        }
        catch (Exception ex)
        {
            Status = $"Couldn't respond: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand(CanExecute = nameof(CanAct))]
    private async Task RemoveAsync()
    {
        IsBusy = true;
        try
        {
            await _calendar.ApplyAsync(_account, Invitation);
            Status = "Removed from your calendar.";
            _changed(null);
        }
        catch (Exception ex)
        {
            Status = $"Couldn't update the calendar: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Replies are recorded as soon as they are opened, as Outlook does.</summary>
    public async Task AutoApplyAsync()
    {
        if (!IsReply) return;
        try
        {
            await _calendar.ApplyAsync(_account, Invitation);
            _changed(null);
        }
        catch (Exception)
        {
            // Purely informational; the reply mail is still there.
        }
    }
}
