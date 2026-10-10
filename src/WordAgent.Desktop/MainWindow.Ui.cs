using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace WordAgent.Desktop;

public partial class MainWindow
{
    private readonly CancellationTokenSource _windowLifetime = new();
    private readonly UiLogBuffer _uiLog = new();
    private DispatcherTimer? _logFlushTimer;
    private Task? _runtimeInitializationTask;
    private volatile bool _uiClosed;
    private bool _closeRequested;
    private bool _layoutQueued;
    private const int MaxDisplayedLogCharacters = 98304;

    private void InitializeUiLifecycle()
    {
        var work = SystemParameters.WorkArea;
        MinWidth = Math.Min(MinWidth, Math.Max(320, work.Width));
        MinHeight = Math.Min(MinHeight, Math.Max(240, work.Height));
        Width = Math.Max(MinWidth, Math.Min(Width, work.Width - 24));
        Height = Math.Max(MinHeight, Math.Min(Height, work.Height - 24));
        CancelButton.IsEnabled = false;
        _logFlushTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(100)
        };
        _logFlushTimer.Tick += (_, _) => FlushUiLog();
        _logFlushTimer.Start();
        Loaded += (_, _) =>
        {
            if (_uiClosed) return;
            ApplyResponsiveLayout();
            _runtimeInitializationTask ??= InitializeRuntimeAsync();
        };
        SizeChanged += (_, _) =>
        {
            if (_layoutQueued || _uiClosed) return;
            _layoutQueued = true;
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
            {
                _layoutQueued = false;
                if (!_uiClosed) ApplyResponsiveLayout();
            }));
        };
    }

    private void ApplyResponsiveLayout()
    {
        double width = Math.Max(320, ActualWidth > 0 ? ActualWidth : Width);
        bool sideBySide = width >= 1500;
        ResponsiveWorkspaceGrid.ColumnDefinitions[1].Width = new GridLength(sideBySide ? 310 : 0);
        Grid.SetColumn(MainWorkspacePanel, 0);
        Grid.SetRow(MainWorkspacePanel, 0);
        Grid.SetColumnSpan(MainWorkspacePanel, sideBySide ? 1 : 2);
        MainWorkspacePanel.Margin = new Thickness(0, 0, sideBySide ? 16 : 0, 0);
        Grid.SetColumn(SidebarPanel, sideBySide ? 1 : 0);
        Grid.SetRow(SidebarPanel, sideBySide ? 0 : 1);
        Grid.SetColumnSpan(SidebarPanel, sideBySide ? 1 : 2);
        SidebarPanel.Margin = new Thickness(0, sideBySide ? 0 : 14, 0, 0);

        bool compactHeader = width < 1180;
        HeaderTagline.Visibility = width >= 1800 ? Visibility.Visible : Visibility.Collapsed;
        Grid.SetRow(HeaderUtilities, compactHeader ? 1 : 0);
        Grid.SetColumn(HeaderUtilities, 0);
        Grid.SetColumnSpan(HeaderUtilities, compactHeader ? 3 : 1);
        HeaderUtilities.Margin = new Thickness(0, compactHeader ? 10 : 0, 0, 0);
        Grid.SetColumn(HeaderBrand, compactHeader ? 0 : 2);
        Grid.SetColumnSpan(HeaderBrand, compactHeader ? 3 : 1);
        HeaderBrand.HorizontalAlignment = HorizontalAlignment.Right;

        double mainWidth = width - 48 - (sideBySide ? 326 : 0);
        bool compactPickers = mainWidth < 850;
        Grid.SetColumn(InputPickerRtlGrid, compactPickers ? 0 : 1);
        Grid.SetRow(InputPickerRtlGrid, 0);
        Grid.SetColumnSpan(InputPickerRtlGrid, compactPickers ? 2 : 1);
        Grid.SetColumn(OutputPickerRtlGrid, 0);
        Grid.SetRow(OutputPickerRtlGrid, compactPickers ? 1 : 0);
        Grid.SetColumnSpan(OutputPickerRtlGrid, compactPickers ? 2 : 1);
        OutputPickerRtlGrid.Margin = compactPickers ? new Thickness(0, 12, 0, 0) : new Thickness(0, 0, 12, 0);

        int columns = mainWidth >= 1130 ? 3 : mainWidth >= 820 ? 2 : 1;
        if (OperationsGrid.ColumnDefinitions.Count != columns)
        {
            OperationsGrid.ColumnDefinitions.Clear();
            for (int i = 0; i < columns; i++) OperationsGrid.ColumnDefinitions.Add(new ColumnDefinition());
        }
        if (columns == 3)
        {
            Place(SmartCard, 0, 2, 1, new Thickness(5, 0, 0, 0));
            Place(GenericCard, 0, 1, 1, new Thickness(5, 0, 5, 0));
            Place(TemplateCard, 0, 0, 1, new Thickness(0, 0, 5, 0));
        }
        else if (columns == 2)
        {
            Place(SmartCard, 0, 1, 1, new Thickness(5, 0, 0, 10));
            Place(GenericCard, 0, 0, 1, new Thickness(0, 0, 5, 10));
            Place(TemplateCard, 1, 0, 2, new Thickness(0));
        }
        else
        {
            Place(SmartCard, 0, 0, 1, new Thickness(0, 0, 0, 10));
            Place(GenericCard, 1, 0, 1, new Thickness(0, 0, 0, 10));
            Place(TemplateCard, 2, 0, 1, new Thickness(0));
        }
    }

    private static void Place(FrameworkElement element, int row, int column, int span, Thickness margin)
    {
        Grid.SetRow(element, row);
        Grid.SetColumn(element, column);
        Grid.SetColumnSpan(element, span);
        element.Margin = margin;
    }

    private void FlushUiLog()
    {
        if (_uiClosed) return;
        var batch = _uiLog.Drain();
        if (batch.Lines.Count == 0 && batch.Dropped == 0) return;
        bool followTail = LogBox.VerticalOffset >= LogBox.ExtentHeight - LogBox.ViewportHeight - 24;
        var text = new StringBuilder();
        if (batch.Dropped != 0) text.AppendLine("\u200F[نمایش گزارش] بخشی از پیام‌های قدیمی برای حفظ پاسخ‌گویی نمایش داده نشد.");
        foreach (var line in batch.Lines)
        {
            string localized = LocalizeTechnicalLine(line);
            if (!string.IsNullOrWhiteSpace(localized) && !localized.StartsWith('\u200F')) text.Append('\u200F');
            text.AppendLine(localized);
        }
        LogBox.AppendText(text.ToString());
        if (LogBox.Text.Length > MaxDisplayedLogCharacters)
        {
            string existing = LogBox.Text;
            int start = existing.Length - 65536;
            int newline = existing.IndexOf('\n', start);
            if (newline >= start && newline < existing.Length - 1) start = newline + 1;
            if (start < existing.Length && char.IsLowSurrogate(existing[start])) start++;
            LogBox.Text = "\u200F[نمایش گزارش] فقط جدیدترین پیام‌ها نمایش داده می‌شوند.\n" + existing[start..];
        }
        if (followTail) LogBox.ScrollToEnd();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (_isBusy && _operationCts is not null)
        {
            e.Cancel = true;
            _closeRequested = true;
            _operationCts.Cancel();
            _runner.CancelActive();
            StatusText.Text = "در حال لغو عملیات جاری؛ پنجره پس از پایان پاک‌سازی بسته می‌شود.";
        }
        base.OnClosing(e);
    }

    protected override void OnClosed(EventArgs e)
    {
        _uiClosed = true;
        _timer.Stop();
        _logFlushTimer?.Stop();
        _windowLifetime.Cancel();
        var initialization = _runtimeInitializationTask;
        if (initialization is null || initialization.IsCompleted) _windowLifetime.Dispose();
        else _ = initialization.ContinueWith(_ => _windowLifetime.Dispose(), CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
        base.OnClosed(e);
    }

    private sealed class BufferedUiProgress(Action<string> report) : IProgress<string>
    {
        public void Report(string value) => report(value);
    }

    private sealed class UiLogBuffer
    {
        private readonly object _sync = new();
        private readonly Queue<string> _lines = new();
        private int _characters;
        private long _dropped;
        public void Enqueue(string? value)
        {
            value ??= string.Empty;
            if (value.Length > 4096)
            {
                int take = char.IsHighSurrogate(value[4095]) ? 4095 : 4096;
                value = value[..take] + " [سطر بلند کوتاه شد]";
            }
            lock (_sync)
            {
                while (_lines.Count > 0 && (_characters + value.Length > 65536 || _lines.Count >= 1024))
                {
                    _characters -= _lines.Dequeue().Length;
                    _dropped++;
                }
                _lines.Enqueue(value);
                _characters += value.Length;
            }
        }
        public (IReadOnlyList<string> Lines, long Dropped) Drain()
        {
            lock (_sync)
            {
                var lines = _lines.ToArray(); long dropped = _dropped;
                _lines.Clear(); _characters = 0; _dropped = 0;
                return (lines, dropped);
            }
        }
    }
}
