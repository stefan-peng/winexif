using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;

namespace WinExif;

public sealed class ResizeCursorGrid : Grid
{
    public static readonly DependencyProperty CursorShapeProperty =
        DependencyProperty.Register(
            nameof(CursorShape),
            typeof(InputSystemCursorShape),
            typeof(ResizeCursorGrid),
            new PropertyMetadata(InputSystemCursorShape.Arrow));

    public InputSystemCursorShape CursorShape
    {
        get => (InputSystemCursorShape)GetValue(CursorShapeProperty);
        set => SetValue(CursorShapeProperty, value);
    }

    public ResizeCursorGrid()
    {
        PointerEntered += OnPointerEntered;
        PointerExited += OnPointerExited;
        PointerCaptureLost += OnPointerCaptureLost;
    }

    private void OnPointerEntered(object sender, PointerRoutedEventArgs e)
    {
        ProtectedCursor = InputSystemCursor.Create(CursorShape);
    }

    private void OnPointerExited(object sender, PointerRoutedEventArgs e)
    {
        ProtectedCursor = null;
    }

    private void OnPointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        ProtectedCursor = null;
    }
}
