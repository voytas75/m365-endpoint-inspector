using System.Xml.Linq;
using M365.EndpointInspector.Domain;
using M365.EndpointInspector.Feed.Tests;
using M365.EndpointInspector.Presentation;

namespace M365.EndpointInspector.Presentation.Tests;

internal static class FilterValidationCases
{
    public static void Run(string repo)
    {
        // Explicitly synthetic, offline data; no production workspace is opened.
        var document = BrowserDocument.OpenSynthetic(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "synthetic-2.json"), TimeProvider.System, CancellationToken.None);
        var applied = new BrowserFilter { Area = ServiceArea.Exchange, Protocol = Transport.Tcp, Ports = "5002-5004" };
        var workspace = BrowserWorkspace.Empty.Open(document) with { Applied = applied, Draft = applied };
        var previous = workspace.Present();
        void Invalid(BrowserFilter draft, params FilterField[] fields)
        {
            var rejected = (workspace with { Draft = draft }).Present(applyDraft: true);
            Check.Equal(string.Join(",", fields.Order()), string.Join(",", rejected.FieldErrors.Select(e => e.Field).Order()));
            Check.True(rejected.FieldErrors.All(e => e.Code.Length > 0 && e.Message.Length > 20));
            Check.Equal(applied, rejected.Workspace.Applied);
            Check.Equal(string.Join(",", previous.Results.Hits.Select(r => r.Entry.Id)),
                string.Join(",", rejected.Results.Hits.Select(r => r.Entry.Id)));
            var restored = rejected.Workspace.Present();
            Check.Equal(string.Join("|", rejected.FieldErrors), string.Join("|", restored.FieldErrors));
            foreach (var valid in new[] { applied, new BrowserFilter() })
            {
                var cleared = (rejected.Workspace with { Draft = valid }).Present(applyDraft: true);
                Check.True(cleared.DraftError == null && cleared.FieldErrors.Count == 0 && cleared.Results.FieldErrors.Count == 0);
                Check.Equal(valid, cleared.Workspace.Applied);
            }
        }
        Invalid(new() { Ports = "70000" }, FilterField.Protocol, FilterField.Ports);
        Invalid(new() { Ports = "443" }, FilterField.Protocol);
        foreach (string ports in new[] { "0", "5004-5000", "443,", "01" })
            Invalid(new() { Protocol = Transport.Tcp, Ports = ports }, FilterField.Ports);
        foreach (string domain in new[] { "https://mail.example.test", "bad..example.test", "*.unpublished.example.test", "*" })
            Invalid(new() { Domain = domain }, FilterField.Domain);
        foreach (string address in new[] { "192.0.2.1/24", "192.0.2.0/33", "2001:db8::/129", "127.1" })
            Invalid(new() { Address = address }, FilterField.Address);
        Invalid(new() { Pointer = "extra", JsonValue = "{" }, FilterField.Pointer, FilterField.JsonValue);
        Invalid(new() { Pointer = "/extra/~2" }, FilterField.Pointer);
        Invalid(new() { JsonValue = "true" }, FilterField.Pointer);
        Invalid(new() { JsonValue = "{" }, FilterField.Pointer, FilterField.JsonValue);
        Invalid(new() { Ports = "70000", Domain = "bad..test", Address = "127.1", Pointer = "bad", JsonValue = "{" },
            FilterField.Protocol, FilterField.Ports, FilterField.Domain, FilterField.Address, FilterField.Pointer, FilterField.JsonValue);
        foreach (var valid in new BrowserFilter[] {
            new() { Protocol = Transport.Tcp, Ports = "1,65535" },
            new() { Address = "192.0.2.0/24" }, new() { Address = "2001:db8::/32" },
            new() { Pointer = "/extra/~0~1", JsonValue = "null" }, new() { Pointer = "/extra" },
            new() { Domain = document.Snapshot!.Entries.Values.SelectMany(e => e.DnsPatterns).First() } })
            Check.True(BrowserResults.Select(document, valid).FilterError == null);

        XNamespace x = "http://schemas.microsoft.com/winfx/2006/xaml";
        var window = XDocument.Load(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/MainWindow.xaml"));
        XElement Named(string name) => window.Descendants().Single(e => (string?)e.Attribute(x + "Name") == name);
        string code = File.ReadAllText(Path.Combine(repo, "src/M365.EndpointInspector.Wpf/MainWindow.xaml.cs"));
        foreach (var (field, input) in new[] { ("Protocol", "Protocol"), ("Ports", "Ports"), ("Domain", "DomainBox"),
            ("Address", "Address"), ("Pointer", "Pointer"), ("JsonValue", "JsonValue"), ("Change", "Change") })
        {
            Check.Equal("{StaticResource InvalidFilter" + Named(input).Name.LocalName + "}", (string?)Named(input).Attribute("Style"));
            var message = Named(input + "Error");
            Check.Equal("{StaticResource FieldErrorText}", (string?)message.Attribute("Style"));
            if (Named(input).Name.LocalName == "ComboBox")
            {
                var wrapper = Named(input + "ValidationBorder");
                Check.Equal("Border", wrapper.Name.LocalName);
                Check.True(Named(input).Parent == wrapper && wrapper.Elements().Single() == Named(input));
                Check.Equal("{StaticResource InvalidFilterComboBoxBorder}", (string?)wrapper.Attribute("Style"));
                Check.Equal("{Binding Tag, ElementName=" + input + "}", (string?)wrapper.Attribute("Tag"));
                Check.True(wrapper.ElementsAfterSelf().First() == message);
                Check.True(Named(input).Attribute("Template") == null);
            }
            else Check.True(Named(input).ElementsAfterSelf().First() == message);
            Check.True(code.Contains($"Show(FilterField.{field}, {input}, {input}Error);"));
        }
        XElement Style(string key) => window.Descendants().Single(e => (string?)e.Attribute(x + "Key") == key);
        bool Sets(XElement element, string property, string value) => element.Descendants().Any(e =>
            e.Name.LocalName == "Setter" && (string?)e.Attribute("Property") == property && (string?)e.Attribute("Value") == value);
        foreach (string key in new[] { "InvalidFilterTextBox", "InvalidFilterComboBoxBorder" })
        {
            var invalid = Style(key).Descendants().Single(e => e.Name.LocalName == "Trigger" &&
                (string?)e.Attribute("Property") == "Tag" && (string?)e.Attribute("Value") == "Invalid");
            Check.True(Sets(invalid, "BorderThickness", "3") && Sets(invalid, "BorderBrush", "#9B251C"));
        }
        var comboStyle = Style("InvalidFilterComboBox");
        Check.Equal("{StaticResource {x:Type ComboBox}}", (string?)comboStyle.Attribute("BasedOn"));
        Check.True(!comboStyle.Descendants().Any(e => (string?)e.Attribute("Property") == "Template" ||
            (string?)e.Attribute("Property") == "BorderBrush" || (string?)e.Attribute("Property") == "Background"));
        Check.True(Sets(comboStyle, "FontWeight", "Bold"));
        Check.True(Sets(Style("FieldErrorText"), "Foreground", "#9B251C"));
        Check.True(Sets(Style("FieldErrorText"), "FontWeight", "SemiBold"));
        foreach (string key in new[] { "InvalidFilterTextBox", "InvalidFilterComboBoxBorder", "FieldErrorText", "FilterWarningBanner" })
        {
            var contrast = Style(key).Descendants().Single(e =>
                (e.Name.LocalName == "DataTrigger" && (string?)e.Attribute("Binding") == "{Binding Path=(SystemParameters.HighContrast)}" &&
                    (string?)e.Attribute("Value") == "True") ||
                (e.Name.LocalName == "MultiDataTrigger" && e.Descendants().Any(c => c.Name.LocalName == "Condition" &&
                    (string?)c.Attribute("Binding") == "{Binding Path=(SystemParameters.HighContrast)}" && (string?)c.Attribute("Value") == "True")));
            string foreground = key is "InvalidFilterComboBoxBorder" or "FilterWarningBanner" ? "TextElement.Foreground" : "Foreground";
            Check.True(Sets(contrast, foreground, "{DynamicResource {x:Static SystemColors.WindowTextBrushKey}}"));
            Check.True(Sets(contrast, "Background", "{DynamicResource {x:Static SystemColors.WindowBrushKey}}"));
            if (key != "FieldErrorText")
                Check.True(Sets(contrast, "BorderBrush", "{DynamicResource {x:Static SystemColors.WindowTextBrushKey}}"));
            Check.True(!contrast.Descendants().Any(e => (string?)e.Attribute("Property") == "BorderThickness" ||
                (string?)e.Attribute("Property") == "FontWeight"));
        }
        var banner = Named("FilterProblemBanner");
        Check.Equal("Border", banner.Name.LocalName);
        Check.Equal("{StaticResource FilterWarningBanner}", (string?)banner.Attribute("Style"));
        Check.True(Named("FilterProblem").Ancestors().Contains(banner) && Named("FilterProblemHeading").Ancestors().Contains(banner));
        Check.Equal("⚠ Filter error", (string?)Named("FilterProblemHeading").Attribute("Text"));
        Check.Equal("Bold", (string?)Named("FilterProblemHeading").Attribute("FontWeight"));
        Check.Equal("", (string?)Named("FilterProblem").Attribute("Text"));
        var bannerStyle = Style("FilterWarningBanner");
        Check.True(Sets(bannerStyle, "BorderThickness", "2") && Sets(bannerStyle, "BorderBrush", "#9B251C") &&
            Sets(bannerStyle, "Background", "#FFF0EE") && Sets(bannerStyle, "TextElement.Foreground", "#9B251C"));
        Check.True(Sets(bannerStyle, "Margin", "0,6,0,6"));
        var collapse = bannerStyle.Descendants().Single(e => e.Name.LocalName == "DataTrigger" &&
            (string?)e.Attribute("Binding") == "{Binding Text, ElementName=FilterProblem}" && (string?)e.Attribute("Value") == "");
        Check.True(Sets(collapse, "Visibility", "Collapsed"));
        Check.True(code.Contains("The table retains the last valid result") && code.Contains("ShowFilterErrors(display.FieldErrors)"));
        Check.True(code.Contains("message.Text = text.Length == 0 ? \"\" : $\"⚠ {text}\"") && code.Contains("SourceFields.IsExpanded = true"));
        Check.Equal("Bold", (string?)Named("FilterProblem").Attribute("FontWeight"));
        Check.Equal("Wrap", (string?)Named("FilterProblem").Attribute("TextWrapping"));
    }
}
