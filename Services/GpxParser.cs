using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace WinExif.Services;

public sealed class GpxParser
{
    public GpxTrack Parse(string path)
    {
        var document = XDocument.Load(path, LoadOptions.None);
        var points = document
            .Descendants()
            .Where(x => x.Name.LocalName is "trkpt" or "rtept")
            .Select(ParsePoint)
            .Where(x => x is not null)
            .Cast<GpxTrackPoint>()
            .OrderBy(x => x.Timestamp)
            .ToList();

        if (points.Count == 0)
        {
            throw new InvalidOperationException("The GPX file does not contain timed track points.");
        }

        return new GpxTrack(points);
    }

    private static GpxTrackPoint? ParsePoint(XElement element)
    {
        var latitudeValue = element.Attribute("lat")?.Value;
        var longitudeValue = element.Attribute("lon")?.Value;
        var timeValue = element.Elements().FirstOrDefault(x => x.Name.LocalName == "time")?.Value;

        if (!double.TryParse(latitudeValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var latitude) ||
            !double.TryParse(longitudeValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var longitude) ||
            !DateTimeOffset.TryParse(timeValue, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var timestamp))
        {
            return null;
        }

        return new GpxTrackPoint(timestamp.ToUniversalTime(), latitude, longitude);
    }
}

public sealed class GpxTrack
{
    public GpxTrack(IReadOnlyList<GpxTrackPoint> points)
    {
        Points = points;
        StartTime = points[0].Timestamp;
        EndTime = points[^1].Timestamp;
    }

    public IReadOnlyList<GpxTrackPoint> Points { get; }

    public DateTimeOffset StartTime { get; }

    public DateTimeOffset EndTime { get; }

    public GeotagMatch? Match(DateTimeOffset? timestamp)
    {
        if (timestamp is null)
        {
            return null;
        }

        var utc = timestamp.Value.ToUniversalTime();
        if (utc < StartTime || utc > EndTime)
        {
            return null;
        }

        var points = Points;
        for (var index = 0; index < points.Count - 1; index++)
        {
            var current = points[index];
            var next = points[index + 1];

            if (utc == current.Timestamp)
            {
                return new GeotagMatch(current.Latitude, current.Longitude, "GPX");
            }

            if (utc <= next.Timestamp)
            {
                var total = (next.Timestamp - current.Timestamp).TotalSeconds;
                if (total <= 0)
                {
                    return new GeotagMatch(current.Latitude, current.Longitude, "GPX");
                }

                var elapsed = (utc - current.Timestamp).TotalSeconds;
                var fraction = elapsed / total;
                var latitude = current.Latitude + ((next.Latitude - current.Latitude) * fraction);
                var longitude = current.Longitude + ((next.Longitude - current.Longitude) * fraction);
                return new GeotagMatch(latitude, longitude, "Interpolated GPX");
            }
        }

        var finalPoint = points[^1];
        return utc == finalPoint.Timestamp
            ? new GeotagMatch(finalPoint.Latitude, finalPoint.Longitude, "GPX")
            : null;
    }
}

public sealed record GpxTrackPoint(DateTimeOffset Timestamp, double Latitude, double Longitude);

public sealed record GeotagMatch(double Latitude, double Longitude, string Source);
