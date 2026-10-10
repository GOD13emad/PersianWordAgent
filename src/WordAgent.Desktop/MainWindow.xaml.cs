using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace WordAgent.Desktop;

public partial class MainWindow : Window
{
    private sealed record ProfileOption(string Id, string DisplayName)
    {
        public override string ToString() => DisplayName;
    }

    private const string AutoOutputProfileId = "__AUTO_OUTPUT_PROFILE__";

    public string ProductVersionText => "نسخه " + ToPersianDigits(ReleaseContract.DesktopVersion);

    private static readonly IReadOnlyDictionary<string, string> OutputProfileDisplayNames =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["formal-admin-technical-proposal-fa"] = "گزارش / سند رسمی، اداری و فنی",
            ["questionnaire-assessment-fa"] = "پرسشنامه / فرم ارزیابی",
            ["thesis-dissertation-academic-fa"] = "پایان‌نامه / رساله دانشگاهی",
            ["proposal.industrial.fa"] = "پروپوزال صنعتی",
            ["proposal.governance.fa"] = "پروپوزال حاکمیتی / راهبری",
            ["proposal.academic.fa"] = "پروپوزال دانشگاهی / دانشجویی",
            ["thesis.dissertation.phd.fa"] = "رساله دکتری",
            ["book.university.fa"] = "کتاب دانشگاهی"
        };

    private static readonly IReadOnlyDictionary<string, string> ProfileResourceKeys =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["formal-admin-technical-proposal-fa"] = "ProfileFormalAdminTechnicalProposalFa",
            ["questionnaire-assessment-fa"] = "ProfileQuestionnaireAssessmentFa",
            ["thesis-dissertation-academic-fa"] = "ProfileThesisDissertationAcademicFa",
            ["proposal.industrial.fa"] = "ProfileProposalIndustrialFa",
            ["proposal.governance.fa"] = "ProfileProposalGovernanceFa",
            ["proposal.academic.fa"] = "ProfileProposalAcademicFa",
            ["thesis.dissertation.phd.fa"] = "ProfileThesisDissertationPhdFa",
            ["book.university.fa"] = "ProfileBookUniversityFa"
        };

    private readonly CliProcessRunner _runner = new();
    private readonly DispatcherTimer _timer;
    private WordAgentRuntime? _runtime;
    private CancellationTokenSource? _operationCts;
    private DateTimeOffset? _operationStarted;
    private string? _suggestedProfile;
    private string? _lastResultPath;
    private string? _lastReportPath;
    private bool _auditReady;
    private bool _safePlanReady;
    private bool _resultPromoted;
    private bool _layoutPlanReady;
    private bool _isBusy;

    public MainWindow()
    {
        InitializeComponent();

        _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _timer.Tick += (_, _) => UpdateElapsed();

        InitializeUiLifecycle();
        UpdateButtons();
    }

    private string T(string key) => TryFindResource(key) as string ?? $"[{key}]";

    private async Task InitializeRuntimeAsync()
    {
        try
        {
            var resolved = await Task.Run(() =>
            {
                bool ok = RuntimeLocator.TryResolve(out var runtime, out var error, _windowLifetime.Token);
                var files = ok && Directory.Exists(runtime.ProfilesDirectory)
                    ? Directory.EnumerateFiles(runtime.ProfilesDirectory, "*.json", SearchOption.TopDirectoryOnly).Take(1025).ToArray()
                    : Array.Empty<string>();
                if (files.Length > 1024) throw new InvalidDataException("PROFILE_CATALOG_LIMIT");
                return (Ok: ok, Runtime: runtime, Error: error, Files: files);
            }, _windowLifetime.Token);
            if (_uiClosed || _windowLifetime.IsCancellationRequested) return;
            if (!resolved.Ok)
            {
                RuntimeText.Text = T("UiRuntimeError");
                RuntimeText.Foreground = Brushes.DarkRed;
                StatusText.Text = T("UiRuntimeError");
                StatusText.Foreground = Brushes.DarkRed;
                AppendLog("[INIT][ERROR] " + resolved.Error);
                return;
            }
            _runtime = resolved.Runtime;
            LoadProfiles(resolved.Files);
            int count = ProfileCombo.Items.OfType<ProfileOption>().Count(p => p.Id != AutoOutputProfileId);
            RuntimeText.Text = string.Format(CultureInfo.CurrentCulture, T("UiRuntimeReadyFormat"), ToPersianDigits(count.ToString(CultureInfo.InvariantCulture)));
            StatusText.Text = "هسته آماده است — فایل ورودی را انتخاب کنید.";
            StatusText.Foreground = Brushes.DarkSlateGray;
            AppendLog("[INIT] Engine: " + _runtime.EngineExe);
            AppendLog("[INIT] Profiles: " + _runtime.ProfilesDirectory);
        }
        catch (OperationCanceledException) when (_windowLifetime.IsCancellationRequested) { }
        catch (Exception ex)
        {
            if (!_uiClosed)
            {
                _runtime = null;
                RuntimeText.Text = T("UiRuntimeError");
                StatusText.Text = T("UiRuntimeError");
                StatusText.Foreground = Brushes.DarkRed;
                AppendLog("[INIT][ERROR] " + ex.Message);
            }
        }
        finally { if (!_uiClosed) UpdateButtons(); }
    }

    private void LoadProfiles(IReadOnlyList<string> files)
    {
        ProfileCombo.Items.Clear();
        ProfileCombo.Items.Add(new ProfileOption(
            AutoOutputProfileId,
            "هوشمند — تشخیص از متن و ساختار"));

        var customIndex = 0;
        foreach (var file in files.OrderBy(Path.GetFileNameWithoutExtension, StringComparer.OrdinalIgnoreCase))
        {
            var id = Path.GetFileNameWithoutExtension(file);
            if (string.IsNullOrWhiteSpace(id))
                continue;

            string displayName;
            if (OutputProfileDisplayNames.TryGetValue(id, out var friendly))
            {
                displayName = friendly;
            }
            else if (ProfileResourceKeys.TryGetValue(id, out var resourceKey))
            {
                displayName = T(resourceKey);
            }
            else
            {
                customIndex++;
                displayName = string.Format(
                    CultureInfo.CurrentCulture,
                    T("UiCustomProfileFormat"),
                    ToPersianDigits(customIndex.ToString(CultureInfo.InvariantCulture)));
            }

            ProfileCombo.Items.Add(new ProfileOption(id, displayName));
        }

        ProfileCombo.SelectedIndex = 0;
    }

    private void BrowseInput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = T("UiOpenFileTitle"),
            Filter = T("UiOpenFileFilter"),
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
            SetInput(dialog.FileName);
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = T("UiOpenFolderTitle"),
            Multiselect = false
        };

        if (Directory.Exists(OutputDirBox.Text))
            dialog.InitialDirectory = OutputDirBox.Text;

        if (dialog.ShowDialog(this) == true)
            OutputDirBox.Text = dialog.FolderName;
    }

    private void BrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = T("UiOpenTemplateTitle"),
            Filter = T("UiOpenFileFilter"),
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) == true)
        {
            TemplatePathBox.Text = Path.GetFullPath(dialog.FileName);
            AppendLog("[TEMPLATE] " + TemplatePathBox.Text);
            UpdateButtons();
        }
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = !_isBusy && HasDocx(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (_isBusy || !HasDocx(e.Data)) return;

        var files = (string[])e.Data.GetData(DataFormats.FileDrop);
        var docx = files.First(x => x.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));
        SetInput(docx);
    }

    private static bool HasDocx(IDataObject data)
    {
        if (!data.GetDataPresent(DataFormats.FileDrop)) return false;
        var files = data.GetData(DataFormats.FileDrop) as string[];
        return files?.Any(x => x.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) == true;
    }

    private void SetInput(string path)
    {
        if (_isBusy || _uiClosed) return;
        InputPathBox.Text = Path.GetFullPath(path);
        OutputDirBox.Text = Path.Combine(Path.GetDirectoryName(InputPathBox.Text)!, "WordAgent_Desktop_Output");

        _suggestedProfile = null;
        SuggestionText.Text = "—";
        UseSuggestionButton.IsEnabled = false;

        _auditReady = false;
        _safePlanReady = false;
        _resultPromoted = false;
        _layoutPlanReady = false;
        _lastResultPath = null;
        _lastReportPath = null;

        AppendLog("[INPUT] " + InputPathBox.Text);
        UpdateButtons();
    }

    private string? SelectedProfile
    {
        get
        {
            var id = (ProfileCombo.SelectedItem as ProfileOption)?.Id;
            return string.IsNullOrWhiteSpace(id)
                || id.Equals(AutoOutputProfileId, StringComparison.OrdinalIgnoreCase)
                ? null
                : id;
        }
    }

    private bool IsAutoOutputProfile =>
        (ProfileCombo.SelectedItem as ProfileOption)?.Id.Equals(
            AutoOutputProfileId,
            StringComparison.OrdinalIgnoreCase) == true;

    private void ProfileChip_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy || _uiClosed) return;
        if (sender is not System.Windows.Controls.Button button ||
            button.Tag is not string profileId ||
            string.IsNullOrWhiteSpace(profileId))
        {
            return;
        }

        var option = ProfileCombo.Items
            .OfType<ProfileOption>()
            .FirstOrDefault(item => item.Id.Equals(profileId, StringComparison.OrdinalIgnoreCase));

        if (option is null)
        {
            StatusText.Text = "پروفایل انتخاب‌شده در نصب فعلی موجود نیست.";
            StatusText.Foreground = Brushes.DarkRed;
            AppendLog("[PROFILE_CHIP][MISSING] " + profileId);
            return;
        }

        ProfileCombo.SelectedItem = option;
        StatusText.Text = profileId.Equals(AutoOutputProfileId, StringComparison.OrdinalIgnoreCase)
            ? "حالت هوشمند انتخاب شد — پس از انتخاب فایل، نوع سند تشخیص داده می‌شود."
            : $"پروفایل «{option.DisplayName}» انتخاب شد.";
        StatusText.Foreground = Brushes.DarkSlateGray;
        AppendLog("[PROFILE_CHIP][SELECT] " + profileId);
    }

    private void UpdateProfileSelectionUi()
    {
        var selected = ProfileCombo.SelectedItem as ProfileOption;
        var selectedId = selected?.Id ?? AutoOutputProfileId;
        var selectedName = selected?.DisplayName ?? "هوشمند — تشخیص از متن و ساختار";

        SelectedProfileSummaryText.Text = selectedName;

        foreach (var button in ProfileChipPanel.Children.OfType<System.Windows.Controls.Button>())
        {
            var id = button.Tag as string ?? string.Empty;
            var isSelected = id.Equals(selectedId, StringComparison.OrdinalIgnoreCase);

            button.Background = isSelected
                ? (Brush)FindResource("PrimarySoftBrush")
                : Brushes.White;
            button.BorderBrush = isSelected
                ? new SolidColorBrush(Color.FromRgb(188, 217, 248))
                : (Brush)FindResource("BorderBrushModern");
            button.Foreground = isSelected
                ? new SolidColorBrush(Color.FromRgb(20, 94, 168))
                : (Brush)FindResource("TextBrush");
            button.FontWeight = isSelected ? FontWeights.SemiBold : FontWeights.Normal;
        }
    }

    private void ProfileCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        _auditReady = false;
        _safePlanReady = false;
        _resultPromoted = false;
        _layoutPlanReady = false;
        UpdateProfileSelectionUi();
        UpdateButtons();
    }

    private void UseSuggestion_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_suggestedProfile)) return;

        var option = ProfileCombo.Items
            .OfType<ProfileOption>()
            .FirstOrDefault(item => item.Id.Equals(_suggestedProfile, StringComparison.OrdinalIgnoreCase));

        if (option is not null)
        {
            ProfileCombo.SelectedItem = option;
            AppendLog("[PROFILE] Suggested profile explicitly accepted by user: " + _suggestedProfile);
        }
    }


    private async void GenericPersianize_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateRuntimeAndInput()) return;

        var args = new List<string>
        {
            "persianize-generic",
            "--input", InputPathBox.Text,
            "--confirm-generic",
            "--fast-verify"
        };

        if (!string.IsNullOrWhiteSpace(OutputDirBox.Text))
        {
            args.Add("--outdir");
            args.Add(OutputDirBox.Text);
        }

        await RunCliAsync(
            T("UiGenericOperation"),
            args,
            acceptableExitCodes: new[] { 0 });
    }

    private async void TemplateAnalyze_Click(object sender, RoutedEventArgs e)
    {
        if (_runtime is null)
        {
            MessageBox.Show(this, T("UiRuntimeMissingMessage"), T("UiDialogTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (string.IsNullOrWhiteSpace(TemplatePathBox.Text) || !File.Exists(TemplatePathBox.Text))
        {
            MessageBox.Show(this, T("UiChooseTemplateMessage"), T("UiDialogTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var args = new List<string>
        {
            "template-analyze",
            "--template", TemplatePathBox.Text
        };

        if (!string.IsNullOrWhiteSpace(OutputDirBox.Text))
        {
            args.Add("--outdir");
            args.Add(OutputDirBox.Text);
        }

        await RunCliAsync(
            T("UiTemplateOperation"),
            args,
            acceptableExitCodes: new[] { 0 });
    }

    private async void TemplateApply_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateRuntimeAndInput())
            return;

        if (string.IsNullOrWhiteSpace(TemplatePathBox.Text) || !File.Exists(TemplatePathBox.Text))
        {
            MessageBox.Show(this, T("UiChooseTemplateMessage"), T("UiDialogTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var args = new List<string>
        {
            "template-apply",
            "--input", InputPathBox.Text,
            "--template", TemplatePathBox.Text,
            "--confirm-template-apply"
        };

        if (!string.IsNullOrWhiteSpace(OutputDirBox.Text))
        {
            args.Add("--outdir");
            args.Add(OutputDirBox.Text);
        }

        await RunCliAsync(
            T("UiTemplateApplyOperation"),
            args,
            acceptableExitCodes: new[] { 0 });
    }

    private async void Detect_Click(object sender, RoutedEventArgs e)
    {
        _ = await DetectAndApplyProfileAsync();
    }

    private async Task<bool> DetectAndApplyProfileAsync()
    {
        if (!ValidateRuntimeAndInput()) return false;

        var autoSelection = IsAutoOutputProfile;
        var result = await RunCliAsync(
            T("UiDetect"),
            new[]
            {
                "detect-profile",
                "--input", InputPathBox.Text,
                "--profiles", _runtime!.ProfilesDirectory
            },
            acceptableExitCodes: new[] { 0, 3, 4 });

        if (result is null) return false;

        var status = ExtractValue(result.StdOut, "DETECTION_STATUS=");
        var profile = ExtractValue(result.StdOut, "PROFILE_ID=");
        var rawConfidence = ExtractValue(result.StdOut, "CONFIDENCE=");
        var confidencePercent = ExtractValue(result.StdOut, "CONFIDENCE_PERCENT=");
        var basis = ExtractValue(result.StdOut, "DETECTION_BASIS=");
        var textEvidence = ExtractValue(result.StdOut, "TEXT_EVIDENCE_COUNT=");

        if (string.IsNullOrWhiteSpace(confidencePercent)
            && double.TryParse(rawConfidence, NumberStyles.Float, CultureInfo.InvariantCulture, out var confidenceValue))
        {
            confidencePercent = (confidenceValue * 100.0).ToString("0.0", CultureInfo.InvariantCulture);
        }

        if (string.IsNullOrWhiteSpace(confidencePercent))
            confidencePercent = "0.0";

        var confidenceText = string.Format(
            CultureInfo.CurrentCulture,
            T("UiConfidenceFormat"),
            ToPersianDigits(confidencePercent));

        var contentFirst = basis.Equals("TEXT_CONTENT", StringComparison.OrdinalIgnoreCase);
        var basisText = contentFirst ? T("UiContentDetectionPrefix") : T("UiDetectionDetected");

        if (status.Equals("DETECTED", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(profile))
        {
            var option = ProfileCombo.Items
                .OfType<ProfileOption>()
                .FirstOrDefault(item => item.Id.Equals(profile, StringComparison.OrdinalIgnoreCase));

            if (option is not null)
            {
                SuggestionText.Text = $"{basisText}: {ProfileDisplayName(profile)} | {confidenceText}";

                if (autoSelection)
                {
                    ProfileCombo.SelectedItem = option;
                    _suggestedProfile = null;
                    UseSuggestionButton.IsEnabled = false;
                    StatusText.Text = $"پروفایل خروجی هوشمند: {ProfileDisplayName(profile)}";
                    AppendLog(
                        $"[OUTPUT_PROFILE][AUTO] {basisText}: «{ProfileDisplayName(profile)}» | {confidenceText}"
                        + (string.IsNullOrWhiteSpace(textEvidence) ? string.Empty : $" | شواهد متنی={ToPersianDigits(textEvidence)}"));
                    UpdateButtons();
                    return true;
                }

                _suggestedProfile = profile;
                UseSuggestionButton.IsEnabled = true;
                StatusText.Text = $"پیشنهاد هوشمند: {ProfileDisplayName(profile)} — انتخاب صریح فعلی حفظ شد.";
                AppendLog(
                    $"[OUTPUT_PROFILE][SUGGESTION_ONLY] «{ProfileDisplayName(profile)}» | explicit={ProfileDisplayName(SelectedProfile!)}");
                UpdateButtons();
                return true;
            }
        }

        _suggestedProfile = null;
        SuggestionText.Text = $"{DetectionStatusFa(status)} | {confidenceText} | {T("UiManualSelectionShort")}";
        UseSuggestionButton.IsEnabled = false;
        StatusText.Text = T("UiDetectionCompletedManual");
        UpdateButtons();
        return false;
    }

    private async void SmartRepair_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateRuntimeAndInput()) return;
        if (!EnsureInputAvailableForWorkflow()) return;

        if (string.IsNullOrWhiteSpace(SelectedProfile))
        {
            var detected = await DetectAndApplyProfileAsync();
            if (!detected || string.IsNullOrWhiteSpace(SelectedProfile))
            {
                MessageBox.Show(
                    this,
                    "حالت هوشمند نتوانست یک پروفایل خروجی یکتا و مطمئن انتخاب کند. یک پروفایل خروجی را صریحاً انتخاب کنید.",
                    T("UiDialogTitle"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }
        }

        var profileId = SelectedProfile!;
        var profileName = ProfileDisplayName(profileId);

        if (profileId.Equals("thesis.dissertation.phd.fa", StringComparison.OrdinalIgnoreCase) ||
            profileId.Equals("thesis-dissertation-academic-fa", StringComparison.OrdinalIgnoreCase))
        {
            var thesisConfirm = MessageBox.Show(
                this,
                $"پروفایل خروجی «{profileName}» انتخاب شده است.\n\nساختار دانشگاهی، فهرست مطالب و Captionهای جدول/شکل روی یک کپی جدید اعمال شود؟",
                T("UiDialogTitle"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No);

            if (thesisConfirm != MessageBoxResult.Yes)
                return;

            var thesisArgs = new List<string>
            {
                "academicize-thesis",
                "--input", InputPathBox.Text,
                "--confirm-academicize",
                "--profile", profileId
            };
            if (!string.IsNullOrWhiteSpace(OutputDirBox.Text))
            {
                thesisArgs.Add("--outdir");
                thesisArgs.Add(OutputDirBox.Text);
            }

            var thesis = await RunCliAsync(
                $"تبدیل به {profileName}",
                thesisArgs,
                acceptableExitCodes: new[] { 0 });

            if (thesis is not null && thesis.ExitCode == 0)
            {
                StatusText.Text = $"پروفایل خروجی «{profileName}» روی یک کپی جدید اعمال شد.";
                StatusText.Foreground = Brushes.DarkGreen;
            }
            return;
        }

        var confirm = MessageBox.Show(
            this,
            $"پروفایل خروجی «{profileName}» انتخاب شده است.\n\nاصلاح کنترل‌شده روی یک کپی جدید انجام شود؟",
            T("UiDialogTitle"),
            MessageBoxButton.YesNo,
            MessageBoxImage.Question,
            MessageBoxResult.No);

        if (confirm != MessageBoxResult.Yes)
            return;

        var args = BuildProfileCommand("fix-controlled").ToList();
        args.Add("--confirm-smart-repair");

        var fix = await RunCliAsync(
            $"اعمال پروفایل خروجی: {profileName}",
            args,
            acceptableExitCodes: new[] { 0 });

        if (fix is null || fix.ExitCode != 0)
            return;

        _auditReady = false;
        _safePlanReady = false;
        _resultPromoted = false;
        _layoutPlanReady = false;
        StatusText.Text = $"پروفایل خروجی «{profileName}» با موفقیت روی کپی جدید اعمال شد.";
        StatusText.Foreground = Brushes.DarkGreen;
        UpdateButtons();

        MessageBox.Show(
            this,
            StatusText.Text,
            T("UiDialogTitle"),
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private async void Audit_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateProfile()) return;
        if (!EnsureInputAvailableForWorkflow()) return;

        var result = await RunCliAsync(
            T("UiAudit"),
            BuildProfileCommand("audit"),
            acceptableExitCodes: new[] { 0, 3 });

        var auditBlockReason = string.Empty;
        _auditReady = result is not null && IsAuditCompleteForWorkflow(result, out auditBlockReason);
        _safePlanReady = false;
        _resultPromoted = false;
        _layoutPlanReady = false;
        UpdateButtons();

        if (result is not null && !_auditReady)
            ShowWorkflowBlocked(auditBlockReason);
    }

    private async void Plan_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateProfile() || !_auditReady) return;
        if (!EnsureInputAvailableForWorkflow()) return;

        var result = await RunCliAsync(
            T("UiSafePlan"),
            BuildProfileCommand("plan"),
            acceptableExitCodes: new[] { 0 });

        _safePlanReady = result is not null
            && result.ExitCode == 0
            && ExtractValue(result.StdOut, "SAFE_FIX_GATE=").Equals("READY", StringComparison.OrdinalIgnoreCase);
        UpdateButtons();
    }

    private async void Fix_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateProfile() || !_safePlanReady) return;
        if (!EnsureInputAvailableForWorkflow()) return;

        var args = BuildProfileCommand("fix-controlled").ToList();
        args.Add("--confirm-plan-ready");

        var result = await RunCliAsync(
            T("UiControlledFix"),
            args,
            acceptableExitCodes: new[] { 0 });

        if (result is not null && result.ExitCode == 0)
        {
            _safePlanReady = false;
            _resultPromoted = false;
            _layoutPlanReady = false;
        }

        UpdateButtons();
    }

    private async void LayoutPlan_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateProfile() || !_resultPromoted) return;

        var result = await RunCliAsync(
            T("UiLayoutPlan"),
            BuildProfileCommand("layout-plan"),
            acceptableExitCodes: new[] { 0 });

        _layoutPlanReady = result is not null && result.ExitCode == 0;
        UpdateButtons();
    }

    private async void LayoutFix_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateProfile() || !_layoutPlanReady) return;

        var args = BuildProfileCommand("layout-fix-controlled").ToList();
        args.Add("--confirm-layout-plan");

        var result = await RunCliAsync(
            T("UiLayoutFix"),
            args,
            acceptableExitCodes: new[] { 0 });

        if (result is not null && result.ExitCode == 0)
            _layoutPlanReady = false;

        UpdateButtons();
    }

    private IReadOnlyList<string> BuildProfileCommand(string command)
    {
        var args = new List<string>
        {
            command,
            "--input", InputPathBox.Text,
            "--profile", SelectedProfile!,
            "--profiles", _runtime!.ProfilesDirectory
        };

        if (!string.IsNullOrWhiteSpace(OutputDirBox.Text))
        {
            args.Add("--outdir");
            args.Add(OutputDirBox.Text);
        }

        return args;
    }

    private async Task<CliRunResult?> RunCliAsync(
        string operation,
        IReadOnlyList<string> arguments,
        IReadOnlyCollection<int> acceptableExitCodes)
    {
        if (_runtime is null || _isBusy || _uiClosed)
            return null;

        _operationCts = new CancellationTokenSource();
        _operationStarted = DateTimeOffset.Now;
        SetBusy(true, operation);

        AppendLog("");
        AppendLog("============================================================");
        AppendLog($"[START] {operation} @ {_operationStarted:yyyy-MM-dd HH:mm:ss}");
        AppendLog("[CMD] " + string.Join(" ", arguments.Select(QuoteForDisplay)));

        var progress = new BufferedUiProgress(AppendLog);

        try
        {
            var result = await _runner.RunAsync(
                _runtime.EngineExe,
                arguments,
                progress,
                _operationCts.Token);

            CapturePaths(result.StdOut);
            CapturePaths(result.StdErr);

            AppendLog($"[END] {operation} | exit={result.ExitCode} | elapsed={result.Elapsed.TotalSeconds:0.0}s");

            if (!acceptableExitCodes.Contains(result.ExitCode))
            {
                StatusText.Text = string.Format(
                    CultureInfo.CurrentCulture,
                    T("UiStatusFailedFormat"),
                    operation,
                    ToPersianDigits(result.ExitCode.ToString(CultureInfo.InvariantCulture)));
                StatusText.Foreground = Brushes.DarkRed;
                return null;
            }

            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("UiStatusSuccessFormat"), operation);
            StatusText.Foreground = Brushes.DarkGreen;
            return result;
        }
        catch (OperationCanceledException)
        {
            AppendLog($"[CANCELLED] {operation}");
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("UiStatusCanceledFormat"), operation);
            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(133, 77, 14));
            return null;
        }
        catch (Exception ex)
        {
            AppendLog("[ERROR] " + ex);
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("UiStatusErrorFormat"), operation);
            StatusText.Foreground = Brushes.DarkRed;
            return null;
        }
        finally
        {
            _operationCts?.Dispose();
            _operationCts = null;
            _operationStarted = null;
            SetBusy(false, operation);
            UpdateButtons();
            if (_closeRequested && !_uiClosed) _ = Dispatcher.BeginInvoke(new Action(Close));
        }
    }

    private void CapturePaths(string text)
    {
        foreach (var raw in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var line = raw.Trim();

            foreach (var prefix in new[] { "Controlled output: ", "Layout output: ", "Generic output: ", "Template output: " })
            {
                if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    var path = line[prefix.Length..].Trim();
                    if (File.Exists(path))
                        _lastResultPath = path;
                }
            }

            foreach (var prefix in new[]
            {
                "Audit report: ",
                "Audit JSON: ",
                "Safe-Fix plan: ",
                "Fix report: ",
                "Layout plan: ",
                "Layout report: ",
                "Post-layout plan: ",
                "Generic report: ",
                "Template contract report: ",
                "Template apply report: "
            })
            {
                if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    var path = line[prefix.Length..].Trim();
                    if (File.Exists(path))
                        _lastReportPath = path;
                }
            }
        }

        OpenResultButton.IsEnabled = !_isBusy && File.Exists(_lastResultPath);
        UseResultButton.IsEnabled = !_isBusy && File.Exists(_lastResultPath);
        OpenReportButton.IsEnabled = !_isBusy && File.Exists(_lastReportPath);
    }

    private static string ExtractValue(string text, string prefix)
    {
        foreach (var raw in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var line = raw.Trim();
            if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return line[prefix.Length..].Trim();
        }
        return string.Empty;
    }

    private string ProfileDisplayName(string id)
    {
        if (OutputProfileDisplayNames.TryGetValue(id, out var friendly))
            return friendly;
        if (ProfileResourceKeys.TryGetValue(id, out var resourceKey))
            return T(resourceKey);
        return T("UiCustomProfilePlain");
    }

    private string DetectionStatusFa(string status) => status.ToUpperInvariant() switch
    {
        "DETECTED" => T("UiDetectionDetected"),
        "AMBIGUOUS" => T("UiDetectionAmbiguous"),
        "NO_MATCH" => T("UiDetectionNoMatch"),
        "LOW_CONFIDENCE" => T("UiDetectionLowConfidence"),
        _ => T("UiDetectionUnknown")
    };

    private bool EnsureInputAvailableForWorkflow()
    {
        var path = InputPathBox.Text;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return false;

        try
        {
            using var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.None);
            return true;
        }
        catch (IOException ex)
        {
            AppendLog("[INPUT][LOCKED] " + ex.Message);
            StatusText.Text = "فایل ورودی در برنامه دیگری باز است؛ ابتدا آن را ببندید.";
            StatusText.Foreground = Brushes.DarkRed;
            MessageBox.Show(
                this,
                "فایل ورودی در برنامه دیگری (برای مثال Microsoft Word) باز است و WordAgent نمی‌تواند با اطمینان عملیات ممیزی/اصلاح را ادامه دهد.\n\nفایل را در Word ببندید و سپس دوباره عملیات را اجرا کنید. WordAgent هیچ برنامه‌ای را خودکار نمی‌بندد.",
                T("UiDialogTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }
        catch (UnauthorizedAccessException ex)
        {
            AppendLog("[INPUT][ACCESS] " + ex.Message);
            StatusText.Text = "فایل ورودی برای پردازش قابل دسترسی نیست.";
            StatusText.Foreground = Brushes.DarkRed;
            MessageBox.Show(
                this,
                "WordAgent به فایل ورودی دسترسی لازم ندارد. دسترسی فایل را بررسی کنید و دوباره تلاش کنید.",
                T("UiDialogTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }
    }

    private bool IsAuditCompleteForWorkflow(CliRunResult audit, out string reason)
    {
        reason = string.Empty;
        var auditJsonPath = ExtractValue(audit.StdOut, "Audit JSON: ");
        if (string.IsNullOrWhiteSpace(auditJsonPath) || !File.Exists(auditJsonPath))
        {
            reason = "داده ممیزی کامل تولید نشد؛ ادامه عملیات اصلاح برای حفظ ایمنی متوقف شد.";
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(auditJsonPath));
            var root = document.RootElement;

            foreach (var propertyName in new[] { "Word", "OpenXml", "Compliance" })
            {
                if (!root.TryGetProperty(propertyName, out var property) ||
                    property.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
                {
                    reason = "ممیزی فایل کامل نشده است. اگر سند در Microsoft Word باز است، آن را ببندید و دوباره تلاش کنید.";
                    return false;
                }
            }

            if (root.TryGetProperty("Warnings", out var warnings) &&
                warnings.ValueKind == JsonValueKind.Array)
            {
                foreach (var warning in warnings.EnumerateArray())
                {
                    var text = warning.GetString() ?? string.Empty;
                    if (text.Contains("being used by another process", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("process cannot access the file", StringComparison.OrdinalIgnoreCase))
                    {
                        reason = "فایل ورودی توسط برنامه دیگری در حال استفاده است. فایل را در Microsoft Word یا برنامه دیگر ببندید و دوباره تلاش کنید.";
                        return false;
                    }
                }
            }

            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            AppendLog("[AUDIT][INCOMPLETE] " + ex.Message);
            reason = "نتیجه ممیزی قابل اعتبارسنجی نیست؛ ادامه عملیات اصلاح متوقف شد.";
            return false;
        }
    }

    private void ShowWorkflowBlocked(string reason)
    {
        StatusText.Text = reason;
        StatusText.Foreground = Brushes.DarkRed;
        AppendLog("[WORKFLOW][BLOCKED] " + reason);
        MessageBox.Show(
            this,
            reason,
            T("UiDialogTitle"),
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
    }

    private bool ValidateRuntimeAndInput()
    {
        if (_isBusy || _uiClosed) return false;
        if (_runtime is null)
        {
            MessageBox.Show(this, T("UiRuntimeMissingMessage"), T("UiDialogTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }

        if (string.IsNullOrWhiteSpace(InputPathBox.Text) || !File.Exists(InputPathBox.Text))
        {
            MessageBox.Show(this, T("UiChooseDocxMessage"), T("UiDialogTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private bool ValidateProfile()
    {
        if (!ValidateRuntimeAndInput()) return false;

        if (string.IsNullOrWhiteSpace(SelectedProfile))
        {
            MessageBox.Show(
                this,
                T("UiChooseProfileMessage"),
                T("UiDialogTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        if (_operationCts is null) return;

        AppendLog("[CANCEL] User requested cancellation of the owned WordAgent child process.");
        _operationCts.Cancel();
        _runner.CancelActive();
    }

    private void UseResult_Click(object sender, RoutedEventArgs e)
    {
        if (File.Exists(_lastResultPath))
        {
            var profile = SelectedProfile;
            SetInput(_lastResultPath!);

            if (!string.IsNullOrWhiteSpace(profile))
            {
                var option = ProfileCombo.Items
                    .OfType<ProfileOption>()
                    .FirstOrDefault(item => item.Id.Equals(profile, StringComparison.OrdinalIgnoreCase));
                if (option is not null)
                    ProfileCombo.SelectedItem = option;
            }

            _resultPromoted = true;
                _layoutPlanReady = false;
            AppendLog("[INPUT] Result explicitly promoted to current input by user.");
            UpdateButtons();
        }
    }

    private void OpenResult_Click(object sender, RoutedEventArgs e) =>
        OpenPath(_lastResultPath);

    private void OpenReport_Click(object sender, RoutedEventArgs e) =>
        OpenPath(_lastReportPath);

    private void OpenFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = OutputDirBox.Text;
        if (!Directory.Exists(folder))
        {
            MessageBox.Show(this, T("UiOutputFolderMissingMessage"), T("UiDialogTitle"), MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        OpenPath(folder);
    }

    private static void OpenPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path) || (!File.Exists(path) && !Directory.Exists(path)))
            return;

        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    private void SetBusy(bool busy, string operation)
    {
        _isBusy = busy;
        BusyBar.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        CancelButton.IsEnabled = busy;

        if (busy)
        {
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("UiStatusRunningFormat"), operation);
            StatusText.Foreground = Brushes.DarkBlue;
            _timer.Start();
        }
        else
        {
            _timer.Stop();
            ElapsedText.Text = string.Empty;
        }

        UpdateButtons();
    }

    private void UpdateButtons()
    {
        var hasRuntime = _runtime is not null;
        var hasInput = File.Exists(InputPathBox.Text);
        var hasProfile = !string.IsNullOrWhiteSpace(SelectedProfile);
        var hasProfiles = _runtime is not null &&
                          Directory.Exists(_runtime.ProfilesDirectory) &&
                          ProfileCombo.Items
                              .OfType<ProfileOption>()
                              .Any(item => !item.Id.Equals(AutoOutputProfileId, StringComparison.OrdinalIgnoreCase));
        var idle = !_isBusy && !_uiClosed;
        AllowDrop = idle;
        InputPickerRtlGrid.IsEnabled = idle;
        OutputPickerRtlGrid.IsEnabled = idle;
        TemplatePickerGrid.IsEnabled = idle;
        ProfileChipPanel.IsEnabled = idle && hasProfiles;
        CancelButton.IsEnabled = _isBusy && !_uiClosed;

        SmartRepairButton.IsEnabled = idle && hasRuntime && hasInput && hasProfiles;
        GenericPersianizeButton.IsEnabled = idle && hasRuntime && hasInput;
        TemplateAnalyzeButton.IsEnabled = idle && hasRuntime && File.Exists(TemplatePathBox.Text);
        TemplateApplyButton.IsEnabled = idle && hasRuntime && hasInput && File.Exists(TemplatePathBox.Text);

        DetectButton.IsEnabled = idle && hasRuntime && hasInput && hasProfiles;
        AuditButton.IsEnabled = idle && hasRuntime && hasInput && hasProfiles && hasProfile;
        PlanButton.IsEnabled = idle && hasRuntime && hasInput && hasProfiles && hasProfile && _auditReady;
        LayoutPlanButton.IsEnabled = idle && hasRuntime && hasInput && hasProfiles && hasProfile && _resultPromoted;
        FixButton.IsEnabled = idle && hasRuntime && hasInput && hasProfiles && hasProfile && _safePlanReady;
        LayoutFixButton.IsEnabled = idle && hasRuntime && hasInput && hasProfiles && hasProfile && _layoutPlanReady;

        UseSuggestionButton.IsEnabled = idle && !string.IsNullOrWhiteSpace(_suggestedProfile);
        OpenResultButton.IsEnabled = idle && File.Exists(_lastResultPath);
        UseResultButton.IsEnabled = idle && File.Exists(_lastResultPath);
        OpenReportButton.IsEnabled = idle && File.Exists(_lastReportPath);
    }

    private void UpdateElapsed()
    {
        if (_operationStarted is null)
        {
            ElapsedText.Text = string.Empty;
            return;
        }

        var elapsed = DateTimeOffset.Now - _operationStarted.Value;
        var elapsedText = ToPersianDigits(elapsed.ToString(@"hh\:mm\:ss", CultureInfo.InvariantCulture));
        ElapsedText.Text = string.Format(CultureInfo.CurrentCulture, T("UiElapsedFormat"), elapsedText);
    }

    private void AppendLog(string text)
    {
        if (!_uiClosed) _uiLog.Enqueue(text);
    }

    private static string LocalizeTechnicalLine(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var localized = text
            .Replace("[START]", "[شروع]", StringComparison.OrdinalIgnoreCase)
            .Replace("[END]", "[پایان]", StringComparison.OrdinalIgnoreCase)
            .Replace("[CMD]", "[فرمان]", StringComparison.OrdinalIgnoreCase)
            .Replace("[ERROR]", "[خطا]", StringComparison.OrdinalIgnoreCase)
            .Replace("[CANCELLED]", "[لغوشده]", StringComparison.OrdinalIgnoreCase)
            .Replace("[CANCEL]", "[لغو]", StringComparison.OrdinalIgnoreCase)
            .Replace("[INIT]", "[راه‌اندازی]", StringComparison.OrdinalIgnoreCase)
            .Replace("[PROFILE]", "[پروفایل]", StringComparison.OrdinalIgnoreCase)
            .Replace("[INPUT]", "[ورودی]", StringComparison.OrdinalIgnoreCase)
            .Replace("DETECTION_STATUS=DETECTED", "وضعیت تشخیص=شناسایی قطعی", StringComparison.OrdinalIgnoreCase)
            .Replace("DETECTION_STATUS=AMBIGUOUS", "وضعیت تشخیص=مبهم", StringComparison.OrdinalIgnoreCase)
            .Replace("DETECTION_STATUS=NOMATCH", "وضعیت تشخیص=بدون تطابق", StringComparison.OrdinalIgnoreCase)
            .Replace("PROFILE_ID=", "شناسه پروفایل=", StringComparison.OrdinalIgnoreCase)
            .Replace("CONFIDENCE=", "اطمینان=", StringComparison.OrdinalIgnoreCase)
            .Replace("CANDIDATE |", "نامزد |", StringComparison.OrdinalIgnoreCase)
            .Replace("EVIDENCE |", "شاهد |", StringComparison.OrdinalIgnoreCase)
            .Replace("SAFE_FIX_GATE=READY", "گیت اصلاح امن=آماده", StringComparison.OrdinalIgnoreCase)
            .Replace("SAFE_FIX_GATE=BLOCKED", "گیت اصلاح امن=مسدود", StringComparison.OrdinalIgnoreCase)
            .Replace("SAFE_FIX_UNCLASSIFIED=", "پاراگراف طبقه‌بندی‌نشده=", StringComparison.OrdinalIgnoreCase)
            .Replace("SAFE_FIX_STYLE_FAILURES=", "Style الزامی نامنطبق/گمشده=", StringComparison.OrdinalIgnoreCase)
            .Replace("CONTROLLED_STAGE=", "مرحله اصلاح کنترل‌شده=", StringComparison.OrdinalIgnoreCase)
            .Replace("Audit report:", "گزارش ممیزی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Audit JSON:", "داده ممیزی JSON:", StringComparison.OrdinalIgnoreCase)
            .Replace("Safe-Fix plan:", "طرح اصلاح امن:", StringComparison.OrdinalIgnoreCase)
            .Replace("Fix report:", "گزارش اصلاح:", StringComparison.OrdinalIgnoreCase)
            .Replace("Controlled output:", "خروجی اصلاح کنترل‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Generic output:", "خروجی اصلاح عمومی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Generic report:", "گزارش اصلاح عمومی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Layout output:", "خروجی صفحه‌آرایی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Layout report:", "گزارش صفحه‌آرایی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Layout plan:", "طرح صفحه‌آرایی:", StringComparison.OrdinalIgnoreCase)
            .Replace("OUTPUT STATUS:", "وضعیت خروجی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Integrity gate:", "گیت یکپارچگی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Mutation engine:", "موتور اعمال تغییر:", StringComparison.OrdinalIgnoreCase)
            .Replace("Profile detection:", "تشخیص پروفایل:", StringComparison.OrdinalIgnoreCase)
            .Replace("Persian paragraphs:", "پاراگراف‌های فارسی:", StringComparison.OrdinalIgnoreCase)
            .Replace("Paragraphs changed:", "پاراگراف‌های اصلاح‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Runs changed:", "Runهای اصلاح‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Tables examined:", "جدول‌های بررسی‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("Tables centered:", "جدول‌های وسط‌چین‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("TOC styles changed:", "Styleهای فهرست اصلاح‌شده:", StringComparison.OrdinalIgnoreCase)
            .Replace("AUTO_FIELD_UPDATE_DISABLED=", "به‌روزرسانی خودکار Field هنگام بازشدن غیرفعال=", StringComparison.OrdinalIgnoreCase)
            .Replace("PASS", "موفق", StringComparison.OrdinalIgnoreCase)
            .Replace("FAIL", "ناموفق", StringComparison.OrdinalIgnoreCase)
            .Replace("BLOCK", "مسدود", StringComparison.OrdinalIgnoreCase)
            .Replace("exit=", "کد خروج=", StringComparison.OrdinalIgnoreCase)
            .Replace("elapsed=", "زمان=", StringComparison.OrdinalIgnoreCase);

        return localized;
    }

    private static string ToPersianDigits(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "—";

        return value
            .Replace('0', '۰')
            .Replace('1', '۱')
            .Replace('2', '۲')
            .Replace('3', '۳')
            .Replace('4', '۴')
            .Replace('5', '۵')
            .Replace('6', '۶')
            .Replace('7', '۷')
            .Replace('8', '۸')
            .Replace('9', '۹')
            .Replace('.', '٫');
    }

    private static string QuoteForDisplay(string value) =>
        value.Any(char.IsWhiteSpace) ? "\"" + value.Replace("\"", "\\\"") + "\"" : value;
}

