using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Forms = System.Windows.Forms;
using WpfBrush = System.Windows.Media.Brush;
using WpfColor = System.Windows.Media.Color;
using WpfDataObject = System.Windows.IDataObject;
using WpfDragEventArgs = System.Windows.DragEventArgs;
using WpfSolidColorBrush = System.Windows.Media.SolidColorBrush;

namespace PDFExtractor;

public partial class MainWindow : Window
{
    private static readonly WpfBrush DropBorderBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0xD8, 0xD1, 0xC4));
    private static readonly WpfBrush DropActiveBorderBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0x2E, 0x6F, 0x72));
    private static readonly WpfBrush ValidationNeutralBackgroundBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0xF7, 0xFB, 0xFA));
    private static readonly WpfBrush ValidationNeutralBorderBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0xD7, 0xE7, 0xE2));
    private static readonly WpfBrush ValidationNeutralTextBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0x1B, 0x1F, 0x23));
    private static readonly WpfBrush ValidationValidBackgroundBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0xEC, 0xF7, 0xF1));
    private static readonly WpfBrush ValidationValidBorderBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0x97, 0xCA, 0xB7));
    private static readonly WpfBrush ValidationValidTextBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0x1D, 0x4D, 0x4F));
    private static readonly WpfBrush ValidationInvalidBackgroundBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0xFD, 0xF1, 0xF0));
    private static readonly WpfBrush ValidationInvalidBorderBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0xE8, 0xB5, 0xB0));
    private static readonly WpfBrush ValidationInvalidTextBrush = new WpfSolidColorBrush(WpfColor.FromRgb(0x93, 0x2F, 0x2A));

    private readonly ObservableCollection<CreatedFileItem> _createdFiles = [];
    private readonly AppSettings _settings;

    private int _sourceInfoRequestId;
    private int? _sourcePageCount;
    private bool _isBusy;
    private bool _isRangeValid;

    public MainWindow()
    {
        InitializeComponent();

        _settings = AppSettingsService.Load();
        OpenFolderCheckBox.IsChecked = _settings.OpenFolderAfterExtract;

        CreatedFilesListBox.ItemsSource = _createdFiles;
        System.Windows.DataObject.AddPastingHandler(RangeTextBox, RangeTextBox_OnPaste);

        SetStatus("준비됨", "원본 PDF와 페이지 범위를 입력한 뒤 추출을 실행하세요.");
        SetRangeValidationVisual(
            PdfSplitService.ValidateRangeText(null, null),
            ValidationMode.Neutral);
        UpdateOutputPreview();
        UpdateCreatedFilesUi();
        UpdateExtractButtonState();
    }

    private async void SourcePdfTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateOutputPreview();
        await RefreshSourceInfoAsync();
    }

    private void RangeTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateRangeValidation();
    }

    private void RangeTextBox_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (string.IsNullOrEmpty(e.Text))
        {
            return;
        }

        if (e.Text.Any(character => !IsAllowedRangeCharacter(character)))
        {
            e.Handled = true;
            SetStatus("입력 제한", "페이지 범위에는 숫자, 쉼표, 하이픈, 공백만 입력할 수 있습니다.");
        }
    }

    private void RangeTextBox_OnPaste(object sender, DataObjectPastingEventArgs e)
    {
        if (!e.SourceDataObject.GetDataPresent(System.Windows.DataFormats.UnicodeText))
        {
            e.CancelCommand();
            return;
        }

        var pastedText = e.SourceDataObject.GetData(System.Windows.DataFormats.UnicodeText) as string ?? string.Empty;
        if (pastedText.Any(character => !IsAllowedRangeCharacter(character)))
        {
            e.CancelCommand();
            SetStatus("붙여넣기 제한", "페이지 범위에는 숫자, 쉼표, 하이픈, 공백만 사용할 수 있습니다.");
        }
    }

    private void OutputFolderTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateOutputPreview();
    }

    private void OpenFolderCheckBox_OnChanged(object sender, RoutedEventArgs e)
    {
        _settings.OpenFolderAfterExtract = OpenFolderCheckBox.IsChecked == true;
        SaveSettingsSafely();
    }

    private void BrowseSourceButton_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "원본 PDF 선택",
            Filter = "PDF 파일 (*.pdf)|*.pdf",
            CheckFileExists = true,
            Multiselect = false
        };

        var initialDirectory = TryGetExistingDirectory(SourcePdfTextBox.Text)
            ?? TryGetExistingDirectory(_settings.RecentSourceFolder)
            ?? TryGetExistingDirectory(_settings.RecentOutputFolder);

        if (!string.IsNullOrWhiteSpace(initialDirectory))
        {
            dialog.InitialDirectory = initialDirectory;
        }

        if (dialog.ShowDialog() == true)
        {
            SourcePdfTextBox.Text = dialog.FileName;
            RememberSourceDirectory(dialog.FileName);
            SetStatus("원본 PDF 선택됨", Path.GetFileName(dialog.FileName));
        }
    }

    private void BrowseOutputFolderButton_OnClick(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "PDF를 저장할 폴더를 선택하세요.",
            UseDescriptionForTitle = true,
            SelectedPath = TryGetExistingDirectory(OutputFolderTextBox.Text)
                ?? TryGetExistingDirectory(_settings.RecentOutputFolder)
                ?? TryGetExistingDirectory(SourcePdfTextBox.Text)
                ?? TryGetExistingDirectory(_settings.RecentSourceFolder)
                ?? string.Empty
        };

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
        {
            OutputFolderTextBox.Text = dialog.SelectedPath;
            RememberOutputDirectory(dialog.SelectedPath);
        }
    }

    private void SourceDropBorder_OnPreviewDragOver(object sender, WpfDragEventArgs e)
    {
        if (TryGetDroppedPdfPath(e.Data, out _))
        {
            e.Effects = System.Windows.DragDropEffects.Copy;
            SourceDropBorder.BorderBrush = DropActiveBorderBrush;
        }
        else
        {
            e.Effects = System.Windows.DragDropEffects.None;
            SourceDropBorder.BorderBrush = DropBorderBrush;
        }

        e.Handled = true;
    }

    private void SourceDropBorder_OnDragLeave(object sender, WpfDragEventArgs e)
    {
        SourceDropBorder.BorderBrush = DropBorderBrush;
    }

    private void SourceDropBorder_OnDrop(object sender, WpfDragEventArgs e)
    {
        SourceDropBorder.BorderBrush = DropBorderBrush;

        if (!TryGetDroppedPdfPath(e.Data, out var pdfPath))
        {
            SetStatus("드래그 앤 드롭 실패", "PDF 파일만 놓을 수 있습니다.");
            return;
        }

        SourcePdfTextBox.Text = pdfPath;
        RememberSourceDirectory(pdfPath);
        SetStatus("원본 PDF 선택됨", $"{Path.GetFileName(pdfPath)} 파일을 불러왔습니다.");
    }

    private async void ExtractButton_OnClick(object sender, RoutedEventArgs e)
    {
        UpdateRangeValidation();

        if (string.IsNullOrWhiteSpace(SourcePdfTextBox.Text))
        {
            SetStatus("입력 확인 필요", "원본 PDF를 먼저 선택하세요.");
            System.Windows.MessageBox.Show(
                this,
                "원본 PDF를 먼저 선택하세요.",
                "입력 확인",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!_isRangeValid)
        {
            SetStatus("입력 확인 필요", "페이지 범위 형식을 먼저 확인하세요.");
            System.Windows.MessageBox.Show(
                this,
                "페이지 범위 형식을 먼저 확인하세요.",
                "입력 확인",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var sourcePdfPath = SourcePdfTextBox.Text.Trim();
        var rangeText = RangeTextBox.Text.Trim();
        var outputFolderText = OutputFolderTextBox.Text.Trim();

        SetBusyState(true);
        SetStatus("처리 중", "페이지 범위를 해석하고 PDF를 분리하고 있습니다.");

        try
        {
            var result = await Task.Run(() => PdfSplitService.Split(sourcePdfPath, rangeText, outputFolderText));

            UpdateOutputPreview(result.OutputDirectory);
            PopulateCreatedFiles(result.CreatedFiles);
            RememberSourceDirectory(sourcePdfPath);
            RememberOutputDirectory(result.OutputDirectory);

            var statusDetail = $"{result.CreatedFiles.Count}개의 PDF를 '{result.OutputDirectory}' 폴더에 저장했습니다.";

            if (OpenFolderCheckBox.IsChecked == true)
            {
                TryOpenFolder(result.OutputDirectory, ref statusDetail);
            }

            SetStatus("완료", statusDetail);
        }
        catch (Exception ex)
        {
            SetStatus("오류", ex.Message);
            System.Windows.MessageBox.Show(
                this,
                ex.Message,
                "처리 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetBusyState(false);
            UpdateExtractButtonState();
        }
    }

    private void CreatedFilesListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateCreatedFilesUi();
    }

    private void CreatedFilesListBox_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        OpenSelectedCreatedFile();
    }

    private void OpenSelectedFileButton_OnClick(object sender, RoutedEventArgs e)
    {
        OpenSelectedCreatedFile();
    }

    private void OpenSelectedFolderButton_OnClick(object sender, RoutedEventArgs e)
    {
        var selectedFile = CreatedFilesListBox.SelectedItem as CreatedFileItem;
        if (selectedFile is null)
        {
            return;
        }

        OpenFileInExplorer(selectedFile.FullPath);
    }

    private void SetBusyState(bool isBusy)
    {
        _isBusy = isBusy;
        UpdateExtractButtonState();
        Mouse.OverrideCursor = isBusy ? System.Windows.Input.Cursors.Wait : null;
    }

    private void SetStatus(string title, string detail)
    {
        StatusTextBlock.Text = title;
        StatusDetailTextBlock.Text = detail;
    }

    private async Task RefreshSourceInfoAsync()
    {
        var requestId = ++_sourceInfoRequestId;
        var sourcePath = SourcePdfTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            _sourcePageCount = null;
            SourceInfoTextBlock.Text = "PDF를 끌어놓거나 직접 선택하세요.";
            UpdateRangeValidation();
            return;
        }

        SourceInfoTextBlock.Text = "PDF 정보를 확인하는 중입니다...";

        try
        {
            var pageCount = await Task.Run(() => PdfSplitService.GetPageCount(sourcePath));
            if (requestId != _sourceInfoRequestId)
            {
                return;
            }

            _sourcePageCount = pageCount;
            SourceInfoTextBlock.Text = $"PDF를 끌어놓거나 직접 선택하세요. 총 {pageCount}페이지입니다.";
        }
        catch (Exception ex)
        {
            if (requestId != _sourceInfoRequestId)
            {
                return;
            }

            _sourcePageCount = null;
            SourceInfoTextBlock.Text = ex.Message;
        }

        UpdateRangeValidation();
    }

    private void UpdateRangeValidation()
    {
        var validation = PdfSplitService.ValidateRangeText(RangeTextBox.Text, _sourcePageCount);

        _isRangeValid = validation.IsValid;

        var mode = !validation.HasInput
            ? ValidationMode.Neutral
            : validation.IsValid
                ? ValidationMode.Valid
                : ValidationMode.Invalid;

        SetRangeValidationVisual(validation, mode);
        UpdateExtractButtonState();
    }

    private void SetRangeValidationVisual(
        PdfSplitService.RangeValidationResult validation,
        ValidationMode mode)
    {
        RangeValidationTextBlock.Text = validation.Message;

        switch (mode)
        {
            case ValidationMode.Valid:
                RangeValidationBorder.Background = ValidationValidBackgroundBrush;
                RangeValidationBorder.BorderBrush = ValidationValidBorderBrush;
                RangeValidationTextBlock.Foreground = ValidationValidTextBrush;
                break;
            case ValidationMode.Invalid:
                RangeValidationBorder.Background = ValidationInvalidBackgroundBrush;
                RangeValidationBorder.BorderBrush = ValidationInvalidBorderBrush;
                RangeValidationTextBlock.Foreground = ValidationInvalidTextBrush;
                break;
            default:
                RangeValidationBorder.Background = ValidationNeutralBackgroundBrush;
                RangeValidationBorder.BorderBrush = ValidationNeutralBorderBrush;
                RangeValidationTextBlock.Foreground = ValidationNeutralTextBrush;
                break;
        }
    }

    private void UpdateExtractButtonState()
    {
        ExtractButton.IsEnabled = !_isBusy && _isRangeValid && !string.IsNullOrWhiteSpace(SourcePdfTextBox.Text);
    }

    private void UpdateCreatedFilesUi()
    {
        CreatedFilesEmptyTextBlock.Visibility = _createdFiles.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        CreatedFilesListBox.Visibility = _createdFiles.Count == 0 ? Visibility.Collapsed : Visibility.Visible;

        var hasSelection = CreatedFilesListBox.SelectedItem is CreatedFileItem;
        OpenSelectedFileButton.IsEnabled = hasSelection;
        OpenSelectedFolderButton.IsEnabled = hasSelection;
    }

    private void PopulateCreatedFiles(IReadOnlyList<string> filePaths)
    {
        _createdFiles.Clear();

        foreach (var filePath in filePaths)
        {
            _createdFiles.Add(new CreatedFileItem(Path.GetFileName(filePath), filePath));
        }

        if (_createdFiles.Count > 0)
        {
            CreatedFilesListBox.SelectedIndex = 0;
        }

        UpdateCreatedFilesUi();
    }

    private void UpdateOutputPreview(string? explicitPath = null)
    {
        if (!string.IsNullOrWhiteSpace(explicitPath))
        {
            OutputPreviewTextBlock.Text = explicitPath;
            return;
        }

        try
        {
            var previewPath = PdfSplitService.ResolvePreviewOutputDirectory(
                SourcePdfTextBox.Text,
                OutputFolderTextBox.Text);

            OutputPreviewTextBlock.Text = string.IsNullOrWhiteSpace(previewPath)
                ? "원본 PDF를 선택하면 여기에 출력 위치가 표시됩니다."
                : previewPath;
        }
        catch
        {
            OutputPreviewTextBlock.Text = "경로를 해석할 수 없습니다. 입력값을 확인하세요.";
        }
    }

    private void OpenSelectedCreatedFile()
    {
        var selectedFile = CreatedFilesListBox.SelectedItem as CreatedFileItem;
        if (selectedFile is null)
        {
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = selectedFile.FullPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            SetStatus("파일 열기 실패", ex.Message);
        }
    }

    private void OpenFileInExplorer(string filePath)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{filePath}\"",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            SetStatus("폴더 열기 실패", ex.Message);
        }
    }

    private void RememberSourceDirectory(string sourceFilePath)
    {
        var directoryPath = TryGetExistingDirectory(sourceFilePath);
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            return;
        }

        _settings.RecentSourceFolder = directoryPath;
        SaveSettingsSafely();
    }

    private void RememberOutputDirectory(string outputDirectoryPath)
    {
        var directoryPath = TryGetExistingDirectory(outputDirectoryPath);
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            return;
        }

        _settings.RecentOutputFolder = directoryPath;
        SaveSettingsSafely();
    }

    private void SaveSettingsSafely()
    {
        try
        {
            AppSettingsService.Save(_settings);
        }
        catch
        {
            // Ignore settings persistence failures so extraction still works.
        }
    }

    private static bool TryGetDroppedPdfPath(WpfDataObject dataObject, out string pdfPath)
    {
        pdfPath = string.Empty;

        if (!dataObject.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            return false;
        }

        if (dataObject.GetData(System.Windows.DataFormats.FileDrop) is not string[] droppedPaths || droppedPaths.Length == 0)
        {
            return false;
        }

        var firstPdfPath = droppedPaths.FirstOrDefault(
            path => string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase));

        if (string.IsNullOrWhiteSpace(firstPdfPath))
        {
            return false;
        }

        pdfPath = firstPdfPath;
        return true;
    }

    private static bool IsAllowedRangeCharacter(char character)
    {
        return char.IsDigit(character) || character == ',' || character == '-' || char.IsWhiteSpace(character);
    }

    private static string? TryGetExistingDirectory(string? pathText)
    {
        if (string.IsNullOrWhiteSpace(pathText))
        {
            return null;
        }

        try
        {
            var fullPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(pathText.Trim()));

            if (Directory.Exists(fullPath))
            {
                return fullPath;
            }

            var directoryPath = Path.GetDirectoryName(fullPath);
            return !string.IsNullOrWhiteSpace(directoryPath) && Directory.Exists(directoryPath)
                ? directoryPath
                : null;
        }
        catch
        {
            return null;
        }
    }

    private static void TryOpenFolder(string outputDirectory, ref string statusDetail)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = outputDirectory,
                UseShellExecute = true
            });

            statusDetail += " 폴더를 자동으로 열었습니다.";
        }
        catch (Exception ex)
        {
            statusDetail += $" 폴더를 자동으로 열지 못했습니다: {ex.Message}";
        }
    }

    private enum ValidationMode
    {
        Neutral,
        Valid,
        Invalid
    }

    private sealed record CreatedFileItem(string FileName, string FullPath);
}
