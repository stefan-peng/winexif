using System.ComponentModel;

namespace WinExif.Services;

internal readonly record struct PhotoSortState(PhotoSortField Field, ListSortDirection Direction);

internal sealed class PhotoSortStateService
{
    public PhotoSortState Advance(
        PhotoSortField currentField,
        ListSortDirection currentDirection,
        PhotoSortField requestedField)
    {
        if (currentField == requestedField)
        {
            return new PhotoSortState(requestedField, Flip(currentDirection));
        }

        return new PhotoSortState(requestedField, ListSortDirection.Ascending);
    }

    private static ListSortDirection Flip(ListSortDirection direction)
    {
        return direction == ListSortDirection.Descending
            ? ListSortDirection.Ascending
            : ListSortDirection.Descending;
    }
}
