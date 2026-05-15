using System.Windows;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>표의 특정 열 속성을 편집하는 다이얼로그 (너비).</summary>
public partial class ColumnPropertiesWindow : Window
{
    private readonly Table _table;
    private readonly int _colIndex;

    public ColumnPropertiesWindow(Table table, int colIndex)
    {
        if (colIndex < 0 || colIndex >= table.Columns.Count)
            throw new ArgumentOutOfRangeException(nameof(colIndex));

        InitializeComponent();
        _table = table;
        _colIndex = colIndex;
        LoadValues();
        WireUpEvents();
    }

    private void LoadValues()
    {
        var col = _table.Columns[_colIndex];
        WidthBox.Text = col.WidthMm > 0 ? col.WidthMm.ToString("F1") : "0";
    }

    private void WireUpEvents()
    {
        OkButton.Click += OkButton_Click;
        CancelButton.Click += CancelButton_Click;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(WidthBox.Text, out double width))
        {
            MessageBox.Show("너비는 숫자여야 합니다.", "입력 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var col = _table.Columns[_colIndex];
        col.WidthMm = width;

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
