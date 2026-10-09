extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Template;

/// <summary>
/// Another presentation opened read-only and windowless in the session's PowerPoint for reading.
/// When the file is already open in that PowerPoint it is reused and left open; only a copy this
/// class opened is closed. The session's own presentation is never opened twice. STA thread only.
/// </summary>
internal sealed class SourcePresentation : IDisposable
{
    private PowerPoint.Presentation? _presentation;
    private readonly bool _owned;

    private SourcePresentation(PowerPoint.Presentation presentation, bool owned)
    {
        _presentation = presentation;
        _owned = owned;
    }

    public PowerPoint.Presentation Presentation => _presentation ?? throw new ObjectDisposedException(nameof(SourcePresentation));

    /// <summary>Opens <paramref name="path"/>; throws <see cref="ArgumentException"/> for a missing file or the session's own file.</summary>
    public static SourcePresentation Open(PowerPoint.Presentation current, string path)
    {
        string full;
        try
        {
            full = System.IO.Path.GetFullPath(path);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new ArgumentException($"Invalid path '{path}': {ex.Message}");
        }
        if (!File.Exists(full))
            throw new ArgumentException($"The file '{full}' does not exist.");
        if (string.Equals(full, current.FullName, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("That file is the presentation open in this session; pass a different file.");

        PowerPoint.Application? application = null;
        PowerPoint.Presentations? presentations = null;
        try
        {
            application = current.Application;
            presentations = application.Presentations;
            for (int index = 1; index <= presentations.Count; index++)
            {
                var open = presentations[index];
                if (string.Equals(open.FullName, full, StringComparison.OrdinalIgnoreCase))
                    return new SourcePresentation(open, owned: false);
                ComUtilities.Release(ref open!);
            }
            var opened = presentations.Open(full, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
            return new SourcePresentation(opened, owned: true);
        }
        finally
        {
            if (presentations is not null) ComUtilities.Release(ref presentations);
            if (application is not null) ComUtilities.Release(ref application);
        }
    }

    public void Dispose()
    {
        if (_presentation is null)
            return;
        if (_owned)
            _presentation.Close();
        ComUtilities.Release(ref _presentation!);
        _presentation = null;
    }
}
