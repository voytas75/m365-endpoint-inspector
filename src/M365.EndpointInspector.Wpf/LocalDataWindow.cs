using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using M365.EndpointInspector.Feed;

namespace M365.EndpointInspector.Wpf;

internal sealed class LocalDataWindow : Window
{
    internal LocalDataWindow(string build, bool canSave, string workspace, string cacheDirectory, CacheView cache)
    {
        Title = "Local data · M365 Endpoint Inspector";
        Width = 700;
        Height = 520;
        MinWidth = 540;
        MinHeight = 380;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var root = new DockPanel { Margin = new Thickness(20) };
        Content = root;

        var footer = new DockPanel();
        var close = Button("_Close", Close);
        close.IsCancel = true;
        DockPanel.SetDock(close, Dock.Right);
        footer.Children.Add(close);
        DockPanel.SetDock(footer, Dock.Bottom);
        root.Children.Add(footer);
        var host = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        root.Children.Add(host);

        var status = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 10, 0, 0) };
        string workspaceState = canSave ? "Saving enabled" : "Saving suspended";
        string cacheState = cache.Generation switch
        {
            CacheGeneration.Active => "Active copy",
            CacheGeneration.Previous => "Previous copy",
            _ => cache.GenerationsMissing ? "No local copy" : "No readable local copy"
        };
        string cacheDetails = cache.Snapshot?.Observation is { } o
            ? $"{cacheState}\nVersion: {o.ObservedVersion}\nRetrieved: {o.RetrievedAt.UtcDateTime:u}\nAge at inspection: {cache.Age:d\\.hh\\:mm} (days.hours:minutes)"
            : cacheState;

        footer.Children.Add(status);
        var sections = new (string Title, string Details, Func<Button> Action)[]
        {
            ("Workspace", $"Current Windows user · {workspaceState}", () => Folder(workspace)),
            ("Worldwide cache", $"Shared by Windows users on this computer\n{cacheDetails}", () => Folder(cacheDirectory)),
            ("Diagnostics", build, () => Button("_Copy diagnostics", CopyDiagnostics))
        };
        var stack = new StackPanel();
        var cards = new System.Windows.Controls.Primitives.UniformGrid { Columns = 2 };
        cards.Children.Add(Frame(Body(0)));
        cards.Children.Add(Frame(Body(1)));
        stack.Children.Add(cards);
        stack.Children.Add(Frame(Body(2)));
        host.Content = stack;

        void CopyDiagnostics()
        {
            try
            {
                Clipboard.SetText($"M365 Endpoint Inspector · {build}\nWorkspace: {workspaceState}\nWorldwide cache:\n{cacheDetails}");
                status.Text = "Diagnostics copied.";
            }
            catch (ExternalException) { status.Text = "Clipboard unavailable. Try again."; }
        }

        StackPanel Body(int index)
        {
            var section = sections[index];
            var panel = new StackPanel();
            panel.Children.Add(Heading(section.Title));
            panel.Children.Add(new TextBlock { Text = section.Details, TextWrapping = TextWrapping.Wrap });
            panel.Children.Add(section.Action());
            return panel;
        }


        Button Folder(string path)
        {
            var button = Button("Open folder", () =>
            {
                try
                {
                    if (!Directory.Exists(path)) { status.Text = "Folder does not exist or is unavailable."; return; }
                    using var process = Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                }
                catch (Exception ex) when (ex is Win32Exception or InvalidOperationException)
                { status.Text = "Could not open folder."; }
            });
            button.ToolTip = path;
            return button;
        }

        static TextBlock Heading(string title) => new() { Text = title, FontWeight = FontWeights.SemiBold,
            FontSize = 16, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 6, 4) };

        static Border Frame(UIElement child) => new() { Child = child, Padding = new Thickness(12),
            BorderThickness = new Thickness(1), BorderBrush = SystemColors.ActiveBorderBrush,
            CornerRadius = new CornerRadius(5), Margin = new Thickness(0, 0, 6, 8) };

        static Button Button(string label, Action action)
        {
            var button = new Button { Content = label, HorizontalAlignment = HorizontalAlignment.Left,
                Padding = new Thickness(10, 4, 10, 4), Margin = new Thickness(0, 6, 0, 0) };
            button.Click += (_, _) => action();

            return button;
        }
    }
}
