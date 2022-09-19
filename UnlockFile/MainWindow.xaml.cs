using System.Diagnostics;
using System.IO;
using System.Management;
using System.Windows;
using Meziantou.Framework.Win32;
using Microsoft.Win32;

namespace UnlockFile;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public sealed partial class MainWindow : Window
{
    private const string DefaultTitle = "Unlock File";

    public MainWindow()
    {
        Title = DefaultTitle;
        InitializeComponent();
        MenuItemRestartAsAdmin.IsEnabled = CanRunAsAdmin();
        UpdateShellIntegrationMenuItems();
        Loaded += MainWindow_Loaded;
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        var arg = Environment.GetCommandLineArgs().LastOrDefault();
        if (!string.IsNullOrEmpty(arg))
        {
            OpenPath(arg);
        }
    }

    private void MenuItemExit_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void OpenPath(string path)
    {
        if (!File.Exists(path))
        {
            MessageBox.Show(path + " is not a valid file path");
            return;
        }

        Title = DefaultTitle + " - " + path;
        ListViewProcess.Items.Clear();

        try
        {
            using var session = RestartManager.CreateSession();
            session.RegisterFile(path);
            var processes = session.GetProcessesLockingResources();
            if (processes.Count == 0)
            {
                MessageBox.Show("The file is not locked");
            }
            else
            {
                foreach (var process in processes)
                {
                    ListViewProcess.Items.Add(process);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("An error occured: " + ex);
        }
    }

    private static bool CanRunAsAdmin()
    {
        using (var token = AccessToken.OpenCurrentProcessToken(TokenAccessLevels.Query))
        {
            if (token.GetElevationType() == TokenElevationType.Limited)
            {
                return true;
            }
        }

        return false;
    }

    private void MenuItemRestartAsAdmin_Click(object sender, RoutedEventArgs e)
    {
        string commandLine;
        using (var searcher = new ManagementObjectSearcher($"SELECT Name, CommandLine FROM Win32_Process WHERE ProcessId = {Environment.ProcessId}"))
        using (var objects = searcher.Get())
        {
            commandLine = objects.Cast<ManagementBaseObject>().SingleOrDefault()?["CommandLine"]?.ToString() ?? string.Empty;
        }

        var (fileName, arguments) = SplitCommandLine(commandLine);

        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            Verb = "runas",
            UseShellExecute = true,
        };
        Process.Start(psi);
        Close();

        static (string fileName, string arguments) SplitCommandLine(string command)
        {
            var isQuoted = false;
            var i = 0;
            for (; i < command.Length; i++)
            {
                switch (command[i])
                {
                    case ' ':
                        if (!isQuoted)
                        {
                            return (command[..i].Trim('"'), command[i..].TrimStart());
                        }
                        break;

                    case '"':
                        isQuoted = !isQuoted;
                        break;
                }
            }

            return (command[..i].Trim('"'), command[i..]);
        }
    }

    private void UpdateShellIntegrationMenuItems()
    {
        var isIntegrated = IsShellIntegrationPresent();
        MenuItemAddShellIntegration.Visibility = isIntegrated ? Visibility.Collapsed : Visibility.Visible;
        MenuItemRemoveShellIntegration.Visibility = !isIntegrated ? Visibility.Collapsed : Visibility.Visible;
    }

    private static bool IsShellIntegrationPresent()
    {
        using var key = Registry.CurrentUser.OpenSubKey("Software\\Classes\\*\\shell\\UnlockFile", writable: false);
        return key is not null;
    }

    private void MenuItemAddShellIntegration_Click(object sender, RoutedEventArgs e)
    {
        using var reg = Registry.CurrentUser.CreateSubKey("Software\\Classes\\*\\shell\\UnlockFile", writable: true);
        reg.SetValue("", "Unlock File");

        using var command = reg.CreateSubKey("command");
        command.SetValue("", $@"""{Environment.ProcessPath}"" ""%1""");

        UpdateShellIntegrationMenuItems();
    }

    private void MenuItemRemoveShellIntegration_Click(object sender, RoutedEventArgs e)
    {
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Classes\\*\\shell\\UnlockFile", throwOnMissingSubKey: false);

        UpdateShellIntegrationMenuItems();
    }

    private void ExecuteOpenCommand(object sender, System.Windows.Input.ExecutedRoutedEventArgs e)
    {
        var dialog = new OpenFileDialog();
        if (dialog.ShowDialog() == true)
        {
            OpenPath(dialog.FileName);
        }
    }

    private void ExecuteDeleteCommand(object sender, System.Windows.Input.ExecutedRoutedEventArgs e)
    {
        if (ListViewProcess.SelectedItem is not Process process)
            return;

        var result = MessageBox.Show($"Are you sure you want to kill the process {process.ProcessName}?", "Kill process?", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (result == MessageBoxResult.Yes)
        {
            process.Kill();
            ListViewProcess.Items.Remove(process);
        }
    }
}
