using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WinExif.Models;
using WinExif.Services;

namespace WinExif.Tests.Services;

[TestClass]
public class PhotoSortServiceTests
{
    [TestMethod]
    public void Sort_ByFileNameAscending_ReturnsAlphabeticalOrder()
    {
        // Arrange
        var service = new PhotoSortService();
        var photos = new List<PhotoItem>
        {
            new("c.jpg"),
            new("A.jpg"),
            new("b.jpg"),
        };

        // Act
        var sorted = service.Sort(photos, PhotoSortField.FileName, ListSortDirection.Ascending);

        // Assert
        CollectionAssert.AreEqual(
            new[] { "A.jpg", "b.jpg", "c.jpg" },
            GetFileNames(sorted));
    }

    [TestMethod]
    public void Sort_ByAdjustedCaptureTimeDescending_PlacesNewestBeforeOlderAndNullLast()
    {
        // Arrange
        var service = new PhotoSortService();
        var newest = new PhotoItem("newest.jpg")
        {
            AdjustedCaptureTime = new DateTimeOffset(2026, 4, 21, 12, 0, 0, TimeSpan.Zero),
        };
        var older = new PhotoItem("older.jpg")
        {
            AdjustedCaptureTime = new DateTimeOffset(2026, 4, 20, 12, 0, 0, TimeSpan.Zero),
        };
        var missing = new PhotoItem("missing.jpg");

        // Act
        var sorted = service.Sort(
            new[] { older, missing, newest },
            PhotoSortField.AdjustedCaptureTime,
            ListSortDirection.Descending);

        // Assert
        CollectionAssert.AreEqual(
            new[] { "newest.jpg", "older.jpg", "missing.jpg" },
            GetFileNames(sorted));
    }

    [TestMethod]
    public void Sort_ByLatitudeAscending_PlacesMissingCoordinatesLast()
    {
        // Arrange
        var service = new PhotoSortService();
        var west = new PhotoItem("west.jpg");
        west.SetCoordinates(-15.2, 0, "GPX");

        var east = new PhotoItem("east.jpg");
        east.SetCoordinates(34.1, 0, "GPX");

        var missing = new PhotoItem("missing.jpg");

        // Act
        var sorted = service.Sort(
            new[] { east, missing, west },
            PhotoSortField.Latitude,
            ListSortDirection.Ascending);

        // Assert
        CollectionAssert.AreEqual(
            new[] { "west.jpg", "east.jpg", "missing.jpg" },
            GetFileNames(sorted));
    }

    [TestMethod]
    public void Sort_ByStatusDescending_UsesCaseInsensitiveOrdering()
    {
        // Arrange
        var service = new PhotoSortService();
        var alpha = new PhotoItem("alpha.jpg")
        {
            Status = "outside GPX range",
        };
        var beta = new PhotoItem("beta.jpg")
        {
            Status = "Saved",
        };
        var gamma = new PhotoItem("gamma.jpg")
        {
            Status = "manual pin",
        };

        // Act
        var sorted = service.Sort(
            new[] { gamma, alpha, beta },
            PhotoSortField.Status,
            ListSortDirection.Descending);

        // Assert
        CollectionAssert.AreEqual(
            new[] { "beta.jpg", "alpha.jpg", "gamma.jpg" },
            GetFileNames(sorted));
    }

    private static string[] GetFileNames(IReadOnlyList<PhotoItem> photos)
    {
        var results = new string[photos.Count];
        for (var index = 0; index < photos.Count; index++)
        {
            results[index] = photos[index].FileName;
        }

        return results;
    }
}
