using Avalonia.Controls;
using Outlook.Core.Compose;

namespace Outlook.App.Controls;

/// <summary>
/// Glue for recipient fields that hold several addresses: suggestions are for the entry being
/// typed (after the last ; or ,), and picking one completes only that entry.
/// </summary>
public static class RecipientCompletion
{
    public static AutoCompleteSelector<string> Selector { get; } = (typed, picked) => AddressList.CompleteLast(typed, picked);
}
