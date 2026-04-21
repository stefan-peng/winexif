using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using WinExif.Models;

namespace WinExif.Services;

public enum PhotoSortField
{
    FileName,
    CaptureTime,
    AdjustedCaptureTime,
    Latitude,
    Longitude,
    CoordinateSource,
    Status,
}

public sealed class PhotoSortService
{
    public IReadOnlyList<PhotoItem> Sort(
        IEnumerable<PhotoItem> photos,
        PhotoSortField field,
        ListSortDirection direction)
    {
        return field switch
        {
            PhotoSortField.CaptureTime => SortByDate(photos, photo => photo.CaptureTime, direction),
            PhotoSortField.AdjustedCaptureTime => SortByDate(photos, photo => photo.AdjustedCaptureTime, direction),
            PhotoSortField.Latitude => SortByNumber(photos, photo => photo.Latitude, direction),
            PhotoSortField.Longitude => SortByNumber(photos, photo => photo.Longitude, direction),
            PhotoSortField.CoordinateSource => SortByText(photos, photo => photo.CoordinateSource, direction),
            PhotoSortField.Status => SortByText(photos, photo => photo.Status, direction),
            _ => SortByText(photos, photo => photo.FileName, direction),
        };
    }

    private static IReadOnlyList<PhotoItem> SortByDate(
        IEnumerable<PhotoItem> photos,
        Func<PhotoItem, DateTimeOffset?> selector,
        ListSortDirection direction)
    {
        var ordered = photos.OrderBy(photo => selector(photo).HasValue ? 0 : 1);
        ordered = direction == ListSortDirection.Descending
            ? ordered.ThenByDescending(selector)
            : ordered.ThenBy(selector);

        return ordered
            .ThenBy(photo => photo.FileName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IReadOnlyList<PhotoItem> SortByNumber(
        IEnumerable<PhotoItem> photos,
        Func<PhotoItem, double?> selector,
        ListSortDirection direction)
    {
        var ordered = photos.OrderBy(photo => selector(photo).HasValue ? 0 : 1);
        ordered = direction == ListSortDirection.Descending
            ? ordered.ThenByDescending(selector)
            : ordered.ThenBy(selector);

        return ordered
            .ThenBy(photo => photo.FileName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IReadOnlyList<PhotoItem> SortByText(
        IEnumerable<PhotoItem> photos,
        Func<PhotoItem, string> selector,
        ListSortDirection direction)
    {
        var ordered = direction == ListSortDirection.Descending
            ? photos.OrderByDescending(selector, StringComparer.OrdinalIgnoreCase)
            : photos.OrderBy(selector, StringComparer.OrdinalIgnoreCase);

        return ordered
            .ThenBy(photo => photo.FileName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
