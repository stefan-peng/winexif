using System.ComponentModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WinExif.Services;

namespace WinExif.Tests.Services;

[TestClass]
public class PhotoSortStateServiceTests
{
    [TestMethod]
    public void Advance_WhenSwitchingFields_UsesAscendingOrder()
    {
        // Arrange
        var service = new PhotoSortStateService();

        // Act
        var next = service.Advance(
            PhotoSortField.FileName,
            ListSortDirection.Descending,
            PhotoSortField.Status);

        // Assert
        Assert.AreEqual(PhotoSortField.Status, next.Field);
        Assert.AreEqual(ListSortDirection.Ascending, next.Direction);
    }

    [TestMethod]
    public void Advance_WhenClickingSameField_TogglesToDescending()
    {
        // Arrange
        var service = new PhotoSortStateService();

        // Act
        var next = service.Advance(
            PhotoSortField.FileName,
            ListSortDirection.Ascending,
            PhotoSortField.FileName);

        // Assert
        Assert.AreEqual(PhotoSortField.FileName, next.Field);
        Assert.AreEqual(ListSortDirection.Descending, next.Direction);
    }

    [TestMethod]
    public void Advance_WhenClickingSameFieldAgain_TogglesBackToAscending()
    {
        // Arrange
        var service = new PhotoSortStateService();

        // Act
        var next = service.Advance(
            PhotoSortField.FileName,
            ListSortDirection.Descending,
            PhotoSortField.FileName);

        // Assert
        Assert.AreEqual(PhotoSortField.FileName, next.Field);
        Assert.AreEqual(ListSortDirection.Ascending, next.Direction);
    }
}
