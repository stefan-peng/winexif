using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using WinExif.Models;

namespace WinExif.Services;

public sealed class ExifToolService
{
    private static readonly string[] DateFormats =
    [
        "yyyy:MM:dd HH:mm:sszzz",
        "yyyy:MM:dd HH:mm:ssK",
        "yyyy:MM:dd HH:mm:ss",
        "yyyy-MM-dd HH:mm:sszzz",
        "yyyy-MM-dd HH:mm:ss"
    ];

    public async Task<IReadOnlyList<PhotoMetadata>> ReadMetadataAsync(
        IReadOnlyList<string> photoPaths,
        string? configuredPath,
        CancellationToken cancellationToken)
    {
        var executable = await ResolveExecutableAsync(configuredPath, cancellationToken);
        if (string.IsNullOrWhiteSpace(executable))
        {
            return [];
        }

        var startInfo = CreateStartInfo(executable);
        startInfo.ArgumentList.Add("-json");
        startInfo.ArgumentList.Add("-n");
        startInfo.ArgumentList.Add("-DateTimeOriginal");
        startInfo.ArgumentList.Add("-CreateDate");
        startInfo.ArgumentList.Add("-GPSLatitude");
        startInfo.ArgumentList.Add("-GPSLongitude");

        foreach (var photoPath in photoPaths)
        {
            startInfo.ArgumentList.Add(photoPath);
        }

        var result = await RunAsync(startInfo, cancellationToken);
        if (result.ExitCode != 0 || string.IsNullOrWhiteSpace(result.StandardOutput))
        {
            return [];
        }

        using var document = JsonDocument.Parse(result.StandardOutput);
        var items = new List<PhotoMetadata>();
        foreach (var element in document.RootElement.EnumerateArray())
        {
            var sourcePath = element.TryGetProperty("SourceFile", out var pathProperty)
                ? pathProperty.GetString()
                : null;

            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                continue;
            }

            var captureTime = ParseDate(element, "DateTimeOriginal") ?? ParseDate(element, "CreateDate");
            var latitude = TryGetDouble(element, "GPSLatitude");
            var longitude = TryGetDouble(element, "GPSLongitude");
            items.Add(new PhotoMetadata(sourcePath, captureTime, latitude, longitude));
        }

        return items;
    }

    public async Task<string?> ResolveExecutableAsync(string? configuredPath, CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(configuredPath) && File.Exists(configuredPath))
        {
            return configuredPath;
        }

        var startInfo = new ProcessStartInfo("where.exe")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add("exiftool");

        var result = await RunAsync(startInfo, cancellationToken);
        if (result.ExitCode != 0)
        {
            return null;
        }

        return result.StandardOutput
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault(File.Exists);
    }

    public async Task<SaveResult> SaveAsync(
        IReadOnlyList<PhotoItem> photos,
        string? configuredPath,
        bool rewriteCaptureTime,
        CancellationToken cancellationToken)
    {
        var executable = await ResolveExecutableAsync(configuredPath, cancellationToken);
        if (string.IsNullOrWhiteSpace(executable))
        {
            return new SaveResult(false, "ExifTool was not found. Install it or specify its full path.");
        }

        var failures = new List<string>();
        foreach (var photo in photos)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var startInfo = CreateStartInfo(executable);
            startInfo.ArgumentList.Add("-overwrite_original");
            startInfo.ArgumentList.Add("-P");
            startInfo.ArgumentList.Add("-m");

            if (photo.HasCoordinates && photo.Latitude.HasValue && photo.Longitude.HasValue)
            {
                startInfo.ArgumentList.Add("-GPSLatitude=" + photo.Latitude.Value.ToString("F8", CultureInfo.InvariantCulture));
                startInfo.ArgumentList.Add("-GPSLongitude=" + photo.Longitude.Value.ToString("F8", CultureInfo.InvariantCulture));
                startInfo.ArgumentList.Add("-GPSLatitudeRef=" + (photo.Latitude.Value >= 0 ? "N" : "S"));
                startInfo.ArgumentList.Add("-GPSLongitudeRef=" + (photo.Longitude.Value >= 0 ? "E" : "W"));
            }

            if (rewriteCaptureTime && photo.AdjustedCaptureTime.HasValue)
            {
                var formatted = photo.AdjustedCaptureTime.Value.ToLocalTime().ToString("yyyy:MM:dd HH:mm:ss", CultureInfo.InvariantCulture);
                startInfo.ArgumentList.Add("-DateTimeOriginal=" + formatted);
                startInfo.ArgumentList.Add("-CreateDate=" + formatted);
            }

            startInfo.ArgumentList.Add(photo.Path);

            var result = await RunAsync(startInfo, cancellationToken);
            if (result.ExitCode != 0)
            {
                failures.Add($"{photo.FileName}: {result.StandardError.Trim()}");
            }
        }

        if (failures.Count > 0)
        {
            return new SaveResult(false, string.Join(Environment.NewLine, failures));
        }

        return new SaveResult(true, $"Saved {photos.Count} file(s) with ExifTool.");
    }

    private static ProcessStartInfo CreateStartInfo(string executable)
    {
        return new ProcessStartInfo(executable)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
    }

    private static DateTimeOffset? ParseDate(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property) || string.IsNullOrWhiteSpace(property.GetString()))
        {
            return null;
        }

        var value = property.GetString()!;
        foreach (var format in DateFormats)
        {
            if (DateTimeOffset.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var timestamp))
            {
                return timestamp;
            }
        }

        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed))
        {
            return new DateTimeOffset(parsed);
        }

        return null;
    }

    private static double? TryGetDouble(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property))
        {
            return null;
        }

        if (property.ValueKind == JsonValueKind.Number && property.TryGetDouble(out var value))
        {
            return value;
        }

        if (property.ValueKind == JsonValueKind.String &&
            double.TryParse(property.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
            return value;
        }

        return null;
    }

    private static async Task<ProcessResult> RunAsync(ProcessStartInfo startInfo, CancellationToken cancellationToken)
    {
        using var process = new Process { StartInfo = startInfo, EnableRaisingEvents = true };
        process.Start();

        var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        await process.WaitForExitAsync(cancellationToken);

        return new ProcessResult(
            process.ExitCode,
            await standardOutputTask,
            await standardErrorTask);
    }
}

public sealed record PhotoMetadata(string SourceFile, DateTimeOffset? CaptureTime, double? Latitude, double? Longitude);

public sealed record SaveResult(bool Success, string Message);

internal sealed record ProcessResult(int ExitCode, string StandardOutput, string StandardError);
