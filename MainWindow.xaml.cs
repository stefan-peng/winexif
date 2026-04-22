using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using Microsoft.Windows.ApplicationModel.Resources;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.Web.WebView2.Core;
using Windows.Storage;
using Windows.Storage.FileProperties;
using Windows.Storage.Pickers;
using WinExif.Models;
using WinExif.Services;
using WinRT.Interop;

namespace WinExif;

public sealed partial class MainWindow : Window, INotifyPropertyChanged
{
    private const string MapVirtualHostName = "appassets.winexif.example";
    private static ResourceLoader? _resourceLoader;
    private static readonly string[] SupportedPhotoExtensions =
    [
        ".jpg", ".jpeg", ".dng", ".arw", ".cr2", ".cr3", ".nef", ".nrw", ".orf", ".raf", ".rw2", ".pef", ".srw", ".x3f", ".erf", ".kdc"
    ];

    private readonly ExifToolService _exifToolService = new();
    private readonly GpxParser _gpxParser = new();
    private readonly PhotoSortService _photoSortService = new();
    private readonly PhotoSortStateService _photoSortStateService = new();
    private GpxTrack? _track;
    private bool _mapReady;
    private bool _mapInitializationStarted;
    private bool _isHorizontalResizeActive;
    private bool _isHorizontalSplitterHovered;
    private bool _isPreviewMapResizeActive;
    private bool _isPreviewMapSplitterHovered;
    private bool _isVerticalResizeActive;
    private bool _isVerticalSplitterHovered;
    private string _statusMessage = string.Empty;
    private InfoBarSeverity _statusSeverity = InfoBarSeverity.Informational;
    private string _gpxSummary = GetString("Summary.GpxNone");
    private string _exifToolPath = string.Empty;
    private string _mapClickSummary = GetString("Summary.MapClickDefault");
    private double _offsetMinutes;
    private int _previewRequestVersion;
    private bool _rewriteCaptureTime = true;
    private double? _lastClickedLatitude;
    private double? _lastClickedLongitude;
    private ImageSource? _selectedPhotoPreviewSource;
    private string _selectedPhotoPreviewStatus = GetString("Preview.NoSelectionStatus");
    private string _selectedPhotoPreviewTitle = GetString("Preview.NoSelectionTitle");
    private double _horizontalResizeStartY;
    private double _initialPhotoPaneHeight;
    private double _initialMapPaneHeight;
    private double _initialPreviewPaneHeight;
    private double _initialSidebarWidth;
    private PhotoSortField _photoSortField = PhotoSortField.FileName;
    private ListSortDirection _photoSortDirection = ListSortDirection.Ascending;
    private double _previewMapResizeStartY;
    private double _verticalResizeStartX;

    private const double MinimumMapWidth = 420;
    private const double MinimumMapPreviewHeight = 180;
    private const double MinimumPhotoPaneHeight = 180;
    private const double MinimumPreviewPaneHeight = 120;
    private const double MinimumSidebarWidth = 280;
    private const double MinimumTopPaneHeight = 220;

    public MainWindow()
    {
        InitializeComponent();

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);

        Photos.CollectionChanged += (_, _) => RefreshUiState();
        Activated += MainWindow_Activated;

        UpdateSortHeaderVisuals();
        UpdateVerticalSplitterVisual(isHovered: false);
        UpdateHorizontalSplitterVisual(isHovered: false);
        UpdatePreviewMapSplitterVisual(isHovered: false);
    }

    public ObservableCollection<PhotoItem> Photos { get; } = [];

    public string StatusMessage
    {
        get => _statusMessage;
        private set
        {
            if (SetProperty(ref _statusMessage, value))
            {
                OnPropertyChanged(nameof(IsStatusOpen));
            }
        }
    }

    public bool IsStatusOpen => !string.IsNullOrWhiteSpace(StatusMessage);

    public InfoBarSeverity StatusSeverity
    {
        get => _statusSeverity;
        private set => SetProperty(ref _statusSeverity, value);
    }

    public string GpxSummary
    {
        get => _gpxSummary;
        private set => SetProperty(ref _gpxSummary, value);
    }

    public string ExifToolPath
    {
        get => _exifToolPath;
        set => SetProperty(ref _exifToolPath, value);
    }

    public double OffsetMinutes
    {
        get => _offsetMinutes;
        set
        {
            if (SetProperty(ref _offsetMinutes, value))
            {
                foreach (var photo in Photos)
                {
                    photo.ApplyTimeOffset(TimeSpan.FromMinutes(_offsetMinutes));
                }

                RefreshDerivedMatches();
                OnPropertyChanged(nameof(OffsetSummary));
            }
        }
    }

    public bool RewriteCaptureTime
    {
        get => _rewriteCaptureTime;
        set => SetProperty(ref _rewriteCaptureTime, value);
    }

    public string OffsetSummary => FormatString("Summary.OffsetFormat", OffsetMinutes);

    public string MapClickSummary
    {
        get => _mapClickSummary;
        private set => SetProperty(ref _mapClickSummary, value);
    }

    public ImageSource? SelectedPhotoPreviewSource
    {
        get => _selectedPhotoPreviewSource;
        private set => SetProperty(ref _selectedPhotoPreviewSource, value);
    }

    public string SelectedPhotoPreviewStatus
    {
        get => _selectedPhotoPreviewStatus;
        private set => SetProperty(ref _selectedPhotoPreviewStatus, value);
    }

    public string SelectedPhotoPreviewTitle
    {
        get => _selectedPhotoPreviewTitle;
        private set => SetProperty(ref _selectedPhotoPreviewTitle, value);
    }

    public string PhotoCountSummary => FormatString("Summary.PhotoCountFormat", Photos.Count);

    public string SelectionSummary
    {
        get
        {
            var selected = GetSelectedPhotos();
            if (selected.Count == 0)
            {
                return GetString("Summary.SelectionNone");
            }

            var matched = selected.Count(photo => photo.HasCoordinates);
            return FormatString("Summary.SelectionFormat", selected.Count, matched);
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private async void MainWindow_Activated(object sender, WindowActivatedEventArgs args)
    {
        if (_mapInitializationStarted)
        {
            return;
        }

        _mapInitializationStarted = true;
        await MapView.EnsureCoreWebView2Async();
        MapView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
        var mapFolder = Path.Combine(AppContext.BaseDirectory, "Assets", "Map");
        MapView.CoreWebView2.SetVirtualHostNameToFolderMapping(
            MapVirtualHostName,
            mapFolder,
            CoreWebView2HostResourceAccessKind.Deny);
        MapView.Source = new Uri($"https://{MapVirtualHostName}/index.html");
    }

    private async void AddPhotos_Click(object sender, RoutedEventArgs e)
    {
        var picker = CreateOpenPicker();
        picker.FileTypeFilter.Add(".jpg");
        picker.FileTypeFilter.Add(".jpeg");
        picker.FileTypeFilter.Add(".dng");
        picker.FileTypeFilter.Add(".arw");
        picker.FileTypeFilter.Add(".cr2");
        picker.FileTypeFilter.Add(".cr3");
        picker.FileTypeFilter.Add(".nef");
        picker.FileTypeFilter.Add(".nrw");
        picker.FileTypeFilter.Add(".orf");
        picker.FileTypeFilter.Add(".raf");
        picker.FileTypeFilter.Add(".rw2");
        picker.FileTypeFilter.Add(".pef");
        picker.FileTypeFilter.Add(".srw");
        picker.FileTypeFilter.Add(".x3f");

        var files = await picker.PickMultipleFilesAsync();
        if (files is null || files.Count == 0)
        {
            return;
        }

        await AddPhotosAsync(files.Select(file => file.Path));
    }

    private async void LoadGpx_Click(object sender, RoutedEventArgs e)
    {
        var picker = CreateOpenPicker();
        picker.FileTypeFilter.Add(".gpx");

        var file = await picker.PickSingleFileAsync();
        if (file is null)
        {
            return;
        }

        try
        {
            _track = _gpxParser.Parse(file.Path);
            GpxSummary = FormatString(
                "Summary.GpxLoadedFormat",
                Path.GetFileName(file.Path),
                _track.Points.Count,
                _track.StartTime.LocalDateTime,
                _track.EndTime.LocalDateTime);
            SetStatus(FormatString("Status.GpxLoaded", _track.Points.Count), InfoBarSeverity.Success);
            PushMapState(fit: true);
        }
        catch (Exception ex)
        {
            SetStatus(ex.Message, InfoBarSeverity.Error);
        }
    }

    private void ApplyGpx_Click(object sender, RoutedEventArgs e)
    {
        if (_track is null)
        {
            SetStatus(GetString("Status.GpxRequired"), InfoBarSeverity.Warning);
            return;
        }

        var matchedCount = RefreshDerivedMatches();

        SetStatus(FormatString("Status.ApplyGpxComplete", matchedCount), InfoBarSeverity.Success);
        PushMapState(fit: true);
    }

    private void RemoveSelected_Click(object sender, RoutedEventArgs e)
    {
        var selected = GetSelectedPhotos();
        if (selected.Count == 0)
        {
            SetStatus(GetString("Status.RemoveSelectionRequired"), InfoBarSeverity.Warning);
            return;
        }

        foreach (var photo in selected)
        {
            Photos.Remove(photo);
        }

        ApplyPhotoSort();
        SetStatus(FormatString("Status.RemoveComplete", selected.Count), InfoBarSeverity.Informational);
        PushMapState(fit: true);
    }

    private async void SaveSelected_Click(object sender, RoutedEventArgs e)
    {
        var selected = GetSelectedPhotos();
        if (selected.Count == 0)
        {
            SetStatus(GetString("Status.SaveSelectionRequired"), InfoBarSeverity.Warning);
            return;
        }

        var result = await _exifToolService.SaveAsync(selected, ExifToolPath, RewriteCaptureTime, default);
        foreach (var photo in selected)
        {
            photo.Status = result.Success ? "Saved" : "Save failed";
        }

        ApplyPhotoSort();
        SetStatus(result.Message, result.Success ? InfoBarSeverity.Success : InfoBarSeverity.Error);
    }

    private void FitMap_Click(object sender, RoutedEventArgs e)
    {
        PushMapState(fit: true);
    }

    private async void BrowseExifTool_Click(object sender, RoutedEventArgs e)
    {
        var picker = CreateOpenPicker();
        picker.FileTypeFilter.Add(".exe");

        var file = await picker.PickSingleFileAsync();
        if (file is null)
        {
            return;
        }

        ExifToolPath = file.Path;
    }

    private void PhotoListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        RefreshUiState();
        RefreshSelectedPhotoPreview();

        if (PhotoListView.SelectedItems.Count == 1 && PhotoListView.SelectedItem is PhotoItem photo && photo.HasCoordinates)
        {
            FocusPhotoOnMap(photo);
        }
        else
        {
            PushMapState(fit: false);
        }
    }

    private void PhotoListView_DragItemsStarting(object sender, DragItemsStartingEventArgs e)
    {
        e.Data.SetText("WinExif photo selection");
        e.Data.Properties.Add("WinExifSelection", "photos");
    }

    private void MapDropTarget_DragOver(object sender, DragEventArgs e)
    {
        e.AcceptedOperation = Windows.ApplicationModel.DataTransfer.DataPackageOperation.Copy;
    }

    private void MapDropTarget_Drop(object sender, DragEventArgs e)
    {
        TagSelectedAtLastClick();
    }

    private void VerticalSplitter_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        _isVerticalResizeActive = true;
        _verticalResizeStartX = e.GetCurrentPoint(WorkspaceGrid).Position.X;
        _initialSidebarWidth = SidebarColumn.ActualWidth;
        ((UIElement)sender).CapturePointer(e.Pointer);
        UpdateVerticalSplitterVisual(isHovered: true);
        e.Handled = true;
    }

    private void VerticalSplitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isVerticalResizeActive)
        {
            return;
        }

        var splitterWidth = Math.Max(VerticalSplitter.ActualWidth, 8);
        var maximumSidebarWidth = Math.Max(
            MinimumSidebarWidth,
            WorkspaceGrid.ActualWidth - splitterWidth - MinimumMapWidth);
        var delta = e.GetCurrentPoint(WorkspaceGrid).Position.X - _verticalResizeStartX;
        var requestedWidth = Math.Clamp(_initialSidebarWidth + delta, MinimumSidebarWidth, maximumSidebarWidth);
        SidebarColumn.Width = new GridLength(requestedWidth);
        e.Handled = true;
    }

    private void VerticalSplitter_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        EndVerticalResize((UIElement)sender);
        e.Handled = true;
    }

    private void VerticalSplitter_PointerCanceled(object sender, PointerRoutedEventArgs e)
    {
        EndVerticalResize((UIElement)sender);
        e.Handled = true;
    }

    private void VerticalSplitter_PointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        EndVerticalResize((UIElement)sender);
    }

    private void HorizontalSplitter_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        _isHorizontalResizeActive = true;
        _horizontalResizeStartY = e.GetCurrentPoint(PaneGrid).Position.Y;
        _initialPhotoPaneHeight = PhotoPaneRow.ActualHeight;
        ((UIElement)sender).CapturePointer(e.Pointer);
        UpdateHorizontalSplitterVisual(isHovered: true);
        e.Handled = true;
    }

    private void HorizontalSplitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isHorizontalResizeActive)
        {
            return;
        }

        var delta = e.GetCurrentPoint(PaneGrid).Position.Y - _horizontalResizeStartY;
        var splitterHeight = Math.Max(HorizontalSplitter.ActualHeight, 8);
        var resizeAreaHeight = WorkspaceRow.ActualHeight + splitterHeight + PhotoPaneRow.ActualHeight;
        var maximumPhotoPaneHeight = Math.Max(
            MinimumPhotoPaneHeight,
            resizeAreaHeight - splitterHeight - MinimumTopPaneHeight);
        var requestedHeight = Math.Clamp(
            _initialPhotoPaneHeight - delta,
            MinimumPhotoPaneHeight,
            maximumPhotoPaneHeight);
        PhotoPaneRow.Height = new GridLength(requestedHeight);
        e.Handled = true;
    }

    private void HorizontalSplitter_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        EndHorizontalResize((UIElement)sender);
        e.Handled = true;
    }

    private void HorizontalSplitter_PointerCanceled(object sender, PointerRoutedEventArgs e)
    {
        EndHorizontalResize((UIElement)sender);
        e.Handled = true;
    }

    private void HorizontalSplitter_PointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        EndHorizontalResize((UIElement)sender);
    }

    private void PreviewMapSplitter_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        _isPreviewMapResizeActive = true;
        _previewMapResizeStartY = e.GetCurrentPoint(PreviewMapGrid).Position.Y;
        _initialPreviewPaneHeight = PreviewPaneRow.ActualHeight;
        _initialMapPaneHeight = MapPaneRow.ActualHeight;
        ((UIElement)sender).CapturePointer(e.Pointer);
        UpdatePreviewMapSplitterVisual(isHovered: true);
        e.Handled = true;
    }

    private void PreviewMapSplitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isPreviewMapResizeActive)
        {
            return;
        }

        var delta = e.GetCurrentPoint(PreviewMapGrid).Position.Y - _previewMapResizeStartY;
        var requestedPreviewHeight = Math.Clamp(
            _initialPreviewPaneHeight + delta,
            MinimumPreviewPaneHeight,
            _initialPreviewPaneHeight + _initialMapPaneHeight - MinimumMapPreviewHeight);
        var requestedMapHeight = Math.Max(
            MinimumMapPreviewHeight,
            _initialPreviewPaneHeight + _initialMapPaneHeight - requestedPreviewHeight);

        PreviewPaneRow.Height = new GridLength(requestedPreviewHeight);
        MapPaneRow.Height = new GridLength(requestedMapHeight);
        e.Handled = true;
    }

    private void PreviewMapSplitter_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        EndPreviewMapResize((UIElement)sender);
        e.Handled = true;
    }

    private void PreviewMapSplitter_PointerCanceled(object sender, PointerRoutedEventArgs e)
    {
        EndPreviewMapResize((UIElement)sender);
        e.Handled = true;
    }

    private void PreviewMapSplitter_PointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        EndPreviewMapResize((UIElement)sender);
    }

    private void TagSelectedAtClick_Click(object sender, RoutedEventArgs e)
    {
        TagSelectedAtLastClick();
    }

    private void PhotoSortHeader_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.Tag is not string tag)
        {
            return;
        }

        var requestedField = Enum.Parse<PhotoSortField>(tag, ignoreCase: false);
        var nextState = _photoSortStateService.Advance(_photoSortField, _photoSortDirection, requestedField);
        _photoSortField = nextState.Field;
        _photoSortDirection = nextState.Direction;

        ApplyPhotoSort();
    }

    private void VerticalSplitter_PointerEnteredVisual(object sender, PointerRoutedEventArgs e)
    {
        _isVerticalSplitterHovered = true;
        UpdateVerticalSplitterVisual(isHovered: true);
    }

    private void VerticalSplitter_PointerExitedVisual(object sender, PointerRoutedEventArgs e)
    {
        _isVerticalSplitterHovered = false;
        if (!_isVerticalResizeActive)
        {
            UpdateVerticalSplitterVisual(isHovered: false);
        }
    }

    private void HorizontalSplitter_PointerEnteredVisual(object sender, PointerRoutedEventArgs e)
    {
        _isHorizontalSplitterHovered = true;
        UpdateHorizontalSplitterVisual(isHovered: true);
    }

    private void HorizontalSplitter_PointerExitedVisual(object sender, PointerRoutedEventArgs e)
    {
        _isHorizontalSplitterHovered = false;
        if (!_isHorizontalResizeActive)
        {
            UpdateHorizontalSplitterVisual(isHovered: false);
        }
    }

    private void PreviewMapSplitter_PointerEnteredVisual(object sender, PointerRoutedEventArgs e)
    {
        _isPreviewMapSplitterHovered = true;
        UpdatePreviewMapSplitterVisual(isHovered: true);
    }

    private void PreviewMapSplitter_PointerExitedVisual(object sender, PointerRoutedEventArgs e)
    {
        _isPreviewMapSplitterHovered = false;
        if (!_isPreviewMapResizeActive)
        {
            UpdatePreviewMapSplitterVisual(isHovered: false);
        }
    }

    private async void RefreshSelectedPhotoPreview()
    {
        if (PhotoListView.SelectedItems.Count != 1 || PhotoListView.SelectedItem is not PhotoItem photo)
        {
            _previewRequestVersion++;
            SetSelectedPhotoPreview(
                null,
                PhotoListView.SelectedItems.Count > 1
                    ? GetString("Preview.MultipleSelectionTitle")
                    : GetString("Preview.NoSelectionTitle"),
                PhotoListView.SelectedItems.Count > 1
                    ? GetString("Preview.MultipleSelectionStatus")
                    : GetString("Preview.NoSelectionStatus"));
            return;
        }

        var requestVersion = ++_previewRequestVersion;
        SetSelectedPhotoPreview(
            null,
            photo.FileName,
            GetString("Preview.LoadingStatus"));

        try
        {
            var file = await StorageFile.GetFileFromPathAsync(photo.Path);
            using var thumbnail = await file.GetThumbnailAsync(
                ThumbnailMode.PicturesView,
                1200,
                ThumbnailOptions.UseCurrentScale);

            if (requestVersion != _previewRequestVersion)
            {
                return;
            }

            if (thumbnail is null)
            {
                SetSelectedPhotoPreview(
                    null,
                    photo.FileName,
                    GetString("Preview.UnavailableStatus"));
                return;
            }

            var bitmap = new BitmapImage();
            await bitmap.SetSourceAsync(thumbnail);

            if (requestVersion != _previewRequestVersion)
            {
                return;
            }

            SetSelectedPhotoPreview(
                bitmap,
                photo.FileName,
                FormatString("Preview.LineFormat", photo.CaptureTimeDisplay, photo.Status));
        }
        catch (Exception)
        {
            if (requestVersion != _previewRequestVersion)
            {
                return;
            }

            SetSelectedPhotoPreview(
                null,
                photo.FileName,
                GetString("Preview.FailedStatus"));
        }
    }

    private void MapView_NavigationCompleted(WebView2 sender, CoreWebView2NavigationCompletedEventArgs args)
    {
        _mapReady = args.IsSuccess;
        if (_mapReady)
        {
            PushMapState(fit: true);
        }
        else
        {
            SetStatus(GetString("Status.MapInitFailed"), InfoBarSeverity.Warning);
        }
    }

    private void CoreWebView2_WebMessageReceived(CoreWebView2 sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        using var document = JsonDocument.Parse(args.TryGetWebMessageAsString());
        var root = document.RootElement;
        var type = root.GetProperty("type").GetString();

        if (type == "mapClick")
        {
            _lastClickedLatitude = root.GetProperty("latitude").GetDouble();
            _lastClickedLongitude = root.GetProperty("longitude").GetDouble();
            MapClickSummary = FormatString(
                "Summary.MapClickFormat",
                _lastClickedLatitude.Value,
                _lastClickedLongitude.Value);
        }
        else if (type == "markerClick")
        {
            var id = root.GetProperty("id").GetString();
            var photo = Photos.FirstOrDefault(item => item.Id == id);
            if (photo is not null)
            {
                PhotoListView.SelectedItem = photo;
                PhotoListView.ScrollIntoView(photo);
                FocusPhotoOnMap(photo);
                RefreshUiState();
            }
        }
    }

    private async System.Threading.Tasks.Task AddPhotosAsync(IEnumerable<string> paths)
    {
        var candidatePaths = paths
            .Where(path => SupportedPhotoExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(path => Photos.All(existing => !string.Equals(existing.Path, path, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (candidatePaths.Count == 0)
        {
            SetStatus(GetString("Status.NoSupportedPhotos"), InfoBarSeverity.Warning);
            return;
        }

        foreach (var path in candidatePaths)
        {
            Photos.Add(new PhotoItem(path));
        }

        string? resolvedExifTool = null;
        try
        {
            resolvedExifTool = await _exifToolService.ResolveExecutableAsync(ExifToolPath, default);
            if (!string.IsNullOrWhiteSpace(resolvedExifTool) && string.IsNullOrWhiteSpace(ExifToolPath))
            {
                ExifToolPath = resolvedExifTool;
            }

            var metadata = string.IsNullOrWhiteSpace(resolvedExifTool)
                ? []
                : await _exifToolService.ReadMetadataAsync(candidatePaths, resolvedExifTool, default);
            var metadataByPath = metadata.ToDictionary(item => item.SourceFile, StringComparer.OrdinalIgnoreCase);

            foreach (var photo in Photos.Where(photo => candidatePaths.Contains(photo.Path, StringComparer.OrdinalIgnoreCase)))
            {
                if (metadataByPath.TryGetValue(photo.Path, out var item))
                {
                    photo.CaptureTime = item.CaptureTime ?? new DateTimeOffset(File.GetLastWriteTime(photo.Path));
                    photo.ApplyTimeOffset(TimeSpan.FromMinutes(OffsetMinutes));

                    if (item.Latitude.HasValue && item.Longitude.HasValue)
                    {
                        photo.SetCoordinates(item.Latitude.Value, item.Longitude.Value, "Existing metadata");
                        photo.Status = "Loaded from EXIF";
                    }
                    else
                    {
                        photo.Status = "Ready";
                    }
                }
                else
                {
                    photo.CaptureTime = new DateTimeOffset(File.GetLastWriteTime(photo.Path));
                    photo.ApplyTimeOffset(TimeSpan.FromMinutes(OffsetMinutes));
                    photo.Status = "Loaded without EXIF";
                }
            }

            var message = string.IsNullOrWhiteSpace(resolvedExifTool)
                ? FormatString("Status.AddPhotosWithoutExifTool", candidatePaths.Count)
                : FormatString("Status.AddPhotosComplete", candidatePaths.Count);
            ApplyPhotoSort();
            SetStatus(message, string.IsNullOrWhiteSpace(resolvedExifTool) ? InfoBarSeverity.Warning : InfoBarSeverity.Success);
            PushMapState(fit: true);
        }
        catch (Exception ex)
        {
            SetStatus(FormatString("Status.MetadataLoadFailed", ex.Message), InfoBarSeverity.Error);
        }
    }

    private void TagSelectedAtLastClick()
    {
        var selected = GetSelectedPhotos();
        if (selected.Count == 0)
        {
            SetStatus(GetString("Status.SelectPhotosFirst"), InfoBarSeverity.Warning);
            return;
        }

        if (!_lastClickedLatitude.HasValue || !_lastClickedLongitude.HasValue)
        {
            SetStatus(GetString("Status.SelectMapPointFirst"), InfoBarSeverity.Warning);
            return;
        }

        foreach (var photo in selected)
        {
            photo.SetCoordinates(_lastClickedLatitude.Value, _lastClickedLongitude.Value, "Manual map pin");
            photo.Status = "Manual pin";
        }

        ApplyPhotoSort();
        SetStatus(FormatString("Status.PinComplete", selected.Count), InfoBarSeverity.Success);
        PushMapState(fit: false);
    }

    private int RefreshDerivedMatches()
    {
        var matchedCount = 0;
        foreach (var photo in Photos)
        {
            if (photo.CoordinateSource.StartsWith("Manual", StringComparison.OrdinalIgnoreCase))
            {
                photo.Status = "Manual pin retained";
                continue;
            }

            if (photo.CoordinateSource == "Existing metadata")
            {
                photo.Status = "Loaded from EXIF";
                continue;
            }

            if (_track is null)
            {
                if (photo.CoordinateSource is "GPX" or "Interpolated GPX")
                {
                    photo.ClearCoordinates();
                }

                photo.Status = "Ready";
                continue;
            }

            var match = _track.Match(photo.AdjustedCaptureTime);
            if (match is null)
            {
                if (photo.CoordinateSource is "GPX" or "Interpolated GPX")
                {
                    photo.ClearCoordinates();
                }

                photo.Status = "Outside GPX range";
                continue;
            }

            photo.SetCoordinates(match.Latitude, match.Longitude, match.Source);
            photo.Status = "Matched to GPX";
            matchedCount++;
        }

        ApplyPhotoSort();
        PushMapState(fit: false);
        return matchedCount;
    }

    private void FocusPhotoOnMap(PhotoItem photo)
    {
        if (!_mapReady || !photo.Latitude.HasValue || !photo.Longitude.HasValue || MapView.CoreWebView2 is null)
        {
            return;
        }

        var message = JsonSerializer.Serialize(new
        {
            type = "focus",
            latitude = photo.Latitude,
            longitude = photo.Longitude
        });
        MapView.CoreWebView2.PostWebMessageAsJson(message);
        PushMapState(fit: false);
    }

    private void PushMapState(bool fit)
    {
        if (!_mapReady || MapView.CoreWebView2 is null)
        {
            return;
        }

        var payload = JsonSerializer.Serialize(new
        {
            type = "setState",
            fit,
            track = _track?.Points.Select(point => new[] { point.Latitude, point.Longitude }) ?? [],
            photos = Photos.Select(photo => new
            {
                id = photo.Id,
                fileName = photo.FileName,
                latitude = photo.Latitude,
                longitude = photo.Longitude,
                coordinateSource = photo.CoordinateSource
            }),
            selectedIds = GetSelectedPhotos().Select(photo => photo.Id)
        });

        MapView.CoreWebView2.PostWebMessageAsJson(payload);
    }

    private List<PhotoItem> GetSelectedPhotos()
    {
        return PhotoListView.SelectedItems.Cast<PhotoItem>().ToList();
    }

    private void RefreshUiState()
    {
        OnPropertyChanged(nameof(PhotoCountSummary));
        OnPropertyChanged(nameof(SelectionSummary));
    }

    private void ApplyPhotoSort()
    {
        var sorted = _photoSortService.Sort(
            Photos,
            _photoSortField,
            _photoSortDirection);
        var currentOrder = Photos.Select(photo => photo.Id).ToArray();
        var nextOrder = sorted.Select(photo => photo.Id).ToArray();

        var selectedIds = GetSelectedPhotos()
            .Select(photo => photo.Id)
            .ToHashSet(StringComparer.Ordinal);

        if (!currentOrder.SequenceEqual(nextOrder))
        {
            Photos.Clear();
            foreach (var photo in sorted)
            {
                Photos.Add(photo);
            }

            PhotoListView.SelectedItems.Clear();
            foreach (var photo in Photos.Where(photo => selectedIds.Contains(photo.Id)))
            {
                PhotoListView.SelectedItems.Add(photo);
            }
        }

        UpdateSortHeaderVisuals();
        RefreshUiState();
    }

    private void SetStatus(string message, InfoBarSeverity severity)
    {
        StatusSeverity = severity;
        StatusMessage = message;
    }

    private FileOpenPicker CreateOpenPicker()
    {
        var picker = new FileOpenPicker();
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
        picker.SuggestedStartLocation = PickerLocationId.PicturesLibrary;
        return picker;
    }

    private void SetSelectedPhotoPreview(ImageSource? source, string title, string status)
    {
        SelectedPhotoPreviewSource = source;
        SelectedPhotoPreviewTitle = title;
        SelectedPhotoPreviewStatus = status;
    }

    private void EndVerticalResize(UIElement splitter)
    {
        if (!_isVerticalResizeActive)
        {
            return;
        }

        _isVerticalResizeActive = false;
        splitter.ReleasePointerCaptures();
        UpdateVerticalSplitterVisual(_isVerticalSplitterHovered);
    }

    private void EndHorizontalResize(UIElement splitter)
    {
        if (!_isHorizontalResizeActive)
        {
            return;
        }

        _isHorizontalResizeActive = false;
        splitter.ReleasePointerCaptures();
        UpdateHorizontalSplitterVisual(_isHorizontalSplitterHovered);
    }

    private void EndPreviewMapResize(UIElement splitter)
    {
        if (!_isPreviewMapResizeActive)
        {
            return;
        }

        _isPreviewMapResizeActive = false;
        splitter.ReleasePointerCaptures();
        UpdatePreviewMapSplitterVisual(_isPreviewMapSplitterHovered);
    }

    private void UpdateSortHeaderVisuals()
    {
        foreach (var header in GetSortHeaders())
        {
            var isActive = header.Field == _photoSortField;
            header.Button.Background = isActive
                ? GetThemeBrush("SubtleFillColorSecondaryBrush")
                : null;
            header.Button.Foreground = isActive
                ? GetThemeBrush("TextFillColorPrimaryBrush")
                : GetThemeBrush("TextFillColorSecondaryBrush");
            header.Indicator.Text = _photoSortDirection == ListSortDirection.Descending ? "↓" : "↑";
            header.Indicator.Visibility = isActive ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private IEnumerable<(PhotoSortField Field, Button Button, TextBlock Indicator)> GetSortHeaders()
    {
        yield return (PhotoSortField.FileName, FileSortHeaderButton, FileSortHeaderIndicator);
        yield return (PhotoSortField.CaptureTime, CaptureTimeSortHeaderButton, CaptureTimeSortHeaderIndicator);
        yield return (PhotoSortField.AdjustedCaptureTime, UpdatedTimeSortHeaderButton, UpdatedTimeSortHeaderIndicator);
        yield return (PhotoSortField.Latitude, LatitudeSortHeaderButton, LatitudeSortHeaderIndicator);
        yield return (PhotoSortField.Longitude, LongitudeSortHeaderButton, LongitudeSortHeaderIndicator);
        yield return (PhotoSortField.CoordinateSource, SourceSortHeaderButton, SourceSortHeaderIndicator);
        yield return (PhotoSortField.Status, StatusSortHeaderButton, StatusSortHeaderIndicator);
    }

    private void UpdateVerticalSplitterVisual(bool isHovered)
    {
        VerticalSplitterGlow.Opacity = _isVerticalResizeActive ? 0.48 : isHovered ? 0.24 : 0;
        VerticalSplitterHandle.Opacity = _isVerticalResizeActive ? 0.95 : isHovered ? 0.72 : 0.4;
    }

    private void UpdateHorizontalSplitterVisual(bool isHovered)
    {
        HorizontalSplitterGlow.Opacity = _isHorizontalResizeActive ? 0.48 : isHovered ? 0.24 : 0;
        HorizontalSplitterHandle.Opacity = _isHorizontalResizeActive ? 0.95 : isHovered ? 0.72 : 0.4;
    }

    private void UpdatePreviewMapSplitterVisual(bool isHovered)
    {
        PreviewMapSplitterGlow.Opacity = _isPreviewMapResizeActive ? 0.48 : isHovered ? 0.24 : 0;
        PreviewMapSplitterHandle.Opacity = _isPreviewMapResizeActive ? 0.95 : isHovered ? 0.72 : 0.4;
    }

    private static Brush GetThemeBrush(string key)
    {
        return (Brush)Application.Current.Resources[key];
    }

    private static string GetString(string key)
    {
        try
        {
            _resourceLoader ??= new ResourceLoader();
            var value = _resourceLoader.GetString(key);
            return string.IsNullOrWhiteSpace(value) ? key : value;
        }
        catch
        {
            return key;
        }
    }

    private static string FormatString(string key, params object[] arguments)
    {
        return string.Format(CultureInfo.CurrentCulture, GetString(key), arguments);
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}
