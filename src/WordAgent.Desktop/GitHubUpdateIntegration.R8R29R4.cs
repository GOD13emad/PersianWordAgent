namespace WordAgent.Desktop;

// Private release policy: the app must never spawn the unqualified updater.
// Keep UI consistent with the updater script's deny-all execution policy.
public partial class MainWindow
{
    private void IntegratedUpdateButton_Click(object sender, System.Windows.RoutedEventArgs e)
        => LaunchIntegratedGitHubUpdater();

    private void LaunchIntegratedGitHubUpdater()
    {
        System.Windows.MessageBox.Show(
            this,
            "به‌روزرسانی خودکار در نسخه خصوصی تا تأیید اصالت ناشر و امکان بازگشت ایمن غیرفعال است. برای نصب نسخه جدید از نصب‌کننده آفلاین تأییدشده استفاده کنید.",
            "WordAgent Update",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Information);
    }
}