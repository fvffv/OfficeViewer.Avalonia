using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeViewer.Avalonia;

/// <summary>Creates an owned, seekable document snapshot from an input stream.</summary>
internal static class DocumentData
{
    public static async Task<byte[]> ReadAsync(Stream source, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(source);

        if (source.CanSeek)
        {
            var remaining = source.Length - source.Position;
            if (remaining is >= 0 and <= int.MaxValue)
            {
                var document = GC.AllocateUninitializedArray<byte>((int)remaining);
                await source.ReadExactlyAsync(document.AsMemory(), cancellationToken).ConfigureAwait(false);
                return document;
            }
        }

        using var buffer = new MemoryStream();
        await source.CopyToAsync(buffer, cancellationToken).ConfigureAwait(false);
        return buffer.ToArray();
    }
}
