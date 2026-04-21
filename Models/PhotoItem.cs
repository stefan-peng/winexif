using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace WinExif.Models;

public sealed class PhotoItem : INotifyPropertyChanged
{
    private DateTimeOffset? _captureTime;
    private DateTimeOffset? _adjustedCaptureTime;
    private double? _latitude;
    private double? _longitude;
    private string _coordinateSource = "Unassigned";
    private string _status = "Ready";

    public PhotoItem(string path)
    {
        Path = path;
        FileName = System.IO.Path.GetFileName(path);
    }

    public string Id { get; } = Guid.NewGuid().ToString("N");

    public string Path { get; }

    public string FileName { get; }

    public DateTimeOffset? CaptureTime
    {
        get => _captureTime;
        set
        {
            if (SetProperty(ref _captureTime, value))
            {
                OnPropertyChanged(nameof(CaptureTimeDisplay));
            }
        }
    }

    public DateTimeOffset? AdjustedCaptureTime
    {
        get => _adjustedCaptureTime;
        set
        {
            if (SetProperty(ref _adjustedCaptureTime, value))
            {
                OnPropertyChanged(nameof(AdjustedCaptureTimeDisplay));
            }
        }
    }

    public double? Latitude
    {
        get => _latitude;
        set
        {
            if (SetProperty(ref _latitude, value))
            {
                OnPropertyChanged(nameof(LatitudeDisplay));
                OnPropertyChanged(nameof(HasCoordinates));
            }
        }
    }

    public double? Longitude
    {
        get => _longitude;
        set
        {
            if (SetProperty(ref _longitude, value))
            {
                OnPropertyChanged(nameof(LongitudeDisplay));
                OnPropertyChanged(nameof(HasCoordinates));
            }
        }
    }

    public string CoordinateSource
    {
        get => _coordinateSource;
        set => SetProperty(ref _coordinateSource, value);
    }

    public string Status
    {
        get => _status;
        set => SetProperty(ref _status, value);
    }

    public bool HasCoordinates => Latitude.HasValue && Longitude.HasValue;

    public string CaptureTimeDisplay => FormatDate(CaptureTime);

    public string AdjustedCaptureTimeDisplay => FormatDate(AdjustedCaptureTime);

    public string LatitudeDisplay => Latitude?.ToString("F6", CultureInfo.InvariantCulture) ?? "—";

    public string LongitudeDisplay => Longitude?.ToString("F6", CultureInfo.InvariantCulture) ?? "—";

    public event PropertyChangedEventHandler? PropertyChanged;

    public void ApplyTimeOffset(TimeSpan offset)
    {
        AdjustedCaptureTime = CaptureTime?.Add(offset);
    }

    public void SetCoordinates(double latitude, double longitude, string source)
    {
        Latitude = latitude;
        Longitude = longitude;
        CoordinateSource = source;
    }

    public void ClearCoordinates()
    {
        Latitude = null;
        Longitude = null;
        CoordinateSource = "Unassigned";
    }

    private static string FormatDate(DateTimeOffset? value)
    {
        return value?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture) ?? "—";
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}
