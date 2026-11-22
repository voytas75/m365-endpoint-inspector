using System;
using System.Collections;
using System.ComponentModel;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Wpf;

public partial class EndpointGrid : DataGrid
{
    private GridLayout layout = GridLayout.Default;
    public bool UpdatingLayout { get; private set; }
    public event EventHandler? LayoutChanged;

    public EndpointGrid()
    {
        InitializeComponent();
        Sorting += SortRows;
        ColumnReordered += (_, _) => NotifyLayout();
        foreach (var column in Columns)
            DependencyPropertyDescriptor.FromProperty(DataGridColumn.WidthProperty, typeof(DataGridColumn))
                .AddValueChanged(column, (_, _) => NotifyLayout());
    }

    private string Id(DataGridColumn column) => GridLayout.Ids[Columns.IndexOf(column)];
    private DataGridColumn Column(string id) => Columns[GridLayout.Ids.IndexOf(id)];

    public void RestoreLayout(GridLayout? restored)
    {
        UpdatingLayout = true;
        try
        {
            layout = restored?.IsValid() == true ? restored : GridLayout.Default;
            foreach (var saved in layout.Columns.OrderBy(c => c.Order))
            {
                var column = Column(saved.Id);
                column.DisplayIndex = saved.Order;
                column.Width = new DataGridLength(saved.Width);
            }
            ApplySort();
        }
        finally { UpdatingLayout = false; }
    }

    public void SetRows(System.Collections.Generic.IReadOnlyList<BrowserRow> rows)
    {
        UpdatingLayout = true;
        try
        {
            // Each grid owns its view; never mutate or sort the immutable filter result.
            ItemsSource = new ListCollectionView((IList)rows.ToArray());
            ApplySort();
        }
        finally { UpdatingLayout = false; }
    }

    private void ApplySort()
    {
        var selected = SelectedItem;
        if (ItemsSource is ICollectionView rows)
        {
            using (rows.DeferRefresh())
            {
                rows.SortDescriptions.Clear();
                foreach (var sort in layout.Sorts)
                    rows.SortDescriptions.Add(new SortDescription(Column(sort.Id).SortMemberPath,
                        sort.Descending ? ListSortDirection.Descending : ListSortDirection.Ascending));
            }
        }
        foreach (var column in Columns) column.SortDirection = null;
        foreach (var sort in layout.Sorts)
            Column(sort.Id).SortDirection = sort.Descending ? ListSortDirection.Descending : ListSortDirection.Ascending;
        if (selected != null) SelectedItem = selected;
    }

    private void SortRows(object sender, DataGridSortingEventArgs e)
    {
        e.Handled = true;
        string id = Id(e.Column);
        var previous = layout.Sorts.FirstOrDefault(s => s.Id == id);
        bool descending = previous != null && !previous.Descending;
        var sorts = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)
            ? layout.Sorts.Where(s => s.Id != id).Append(new GridSort(id, descending)).ToArray()
            : new[] { new GridSort(id, descending) };
        layout = layout with { Sorts = sorts };
        UpdatingLayout = true;
        try { ApplySort(); }
        finally { UpdatingLayout = false; }
        NotifyLayout();
    }

    private void NotifyLayout()
    {
        if (UpdatingLayout) return;
        CaptureLayout();
        LayoutChanged?.Invoke(this, EventArgs.Empty);
    }

    public GridLayout CaptureLayout()
    {
        if (UpdatingLayout) return layout;
        var captured = new GridLayout(Columns.Select(c => new GridColumn(Id(c), c.DisplayIndex,
            Math.Clamp(c.Width.IsAbsolute ? c.Width.Value : c.ActualWidth, 20, 2000))).ToArray(), layout.Sorts);
        // Record equality prevents no-op width notifications from queuing another save.
        if (!layout.Columns.SequenceEqual(captured.Columns)) layout = captured;
        return layout;
    }
}

internal static class GridColumnIds
{
    internal static int IndexOf(this System.Collections.Generic.IReadOnlyList<string> ids, string id)
    {
        for (int i = 0; i < ids.Count; i++) if (ids[i] == id) return i;
        throw new ArgumentException("Unknown column identifier.", nameof(id));
    }
}
