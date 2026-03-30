using System.IO;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PDFExtractor;

internal static class PdfSplitService
{
    private const int MaxDefaultFolderTitleLength = 36;
    private const int MaxOutputFileStemLength = 80;

    public static SplitResult Split(string sourcePdfPath, string rangeText, string? outputFolderInput)
    {
        var normalizedSourcePath = NormalizeSourcePath(sourcePdfPath);
        var segments = ParseSegments(rangeText);
        var outputDirectory = ResolveOutputDirectory(normalizedSourcePath, outputFolderInput);

        Directory.CreateDirectory(outputDirectory);

        using var sourceDocument = PdfReader.Open(normalizedSourcePath, PdfDocumentOpenMode.Import);
        EnsurePagesExist(segments, sourceDocument.PageCount);

        var createdFiles = new List<string>(segments.Count);
        var sourceTitle = Path.GetFileNameWithoutExtension(normalizedSourcePath);
        var safeSourceTitle = ShortenPathSegment(SanitizePathSegment(sourceTitle), MaxOutputFileStemLength);

        foreach (var segment in segments)
        {
            using var outputDocument = new PdfDocument();

            foreach (var pageNumber in segment.Pages)
            {
                outputDocument.AddPage(sourceDocument.Pages[pageNumber - 1]);
            }

            var outputStem = ShortenPathSegment(
                SanitizePathSegment($"{safeSourceTitle}_{segment.DisplayLabel}"),
                MaxOutputFileStemLength);

            var outputPath = GetUniqueFilePath(outputDirectory, outputStem, ".pdf");
            outputDocument.Save(outputPath);
            createdFiles.Add(outputPath);
        }

        return new SplitResult(outputDirectory, createdFiles);
    }

    public static string ResolvePreviewOutputDirectory(string? sourcePdfPath, string? outputFolderInput)
    {
        if (!string.IsNullOrWhiteSpace(outputFolderInput))
        {
            return ResolveOutputDirectory(sourcePdfPath ?? string.Empty, outputFolderInput);
        }

        if (string.IsNullOrWhiteSpace(sourcePdfPath))
        {
            return string.Empty;
        }

        var normalizedSourcePath = NormalizePath(sourcePdfPath);
        if (string.IsNullOrWhiteSpace(normalizedSourcePath))
        {
            return string.Empty;
        }

        var sourceDirectory = Path.GetDirectoryName(normalizedSourcePath);
        if (string.IsNullOrWhiteSpace(sourceDirectory))
        {
            return string.Empty;
        }

        return Path.Combine(
            sourceDirectory,
            BuildDefaultDirectoryName(Path.GetFileNameWithoutExtension(normalizedSourcePath)));
    }

    public static int GetPageCount(string sourcePdfPath)
    {
        var normalizedSourcePath = NormalizeSourcePath(sourcePdfPath);
        using var sourceDocument = PdfReader.Open(normalizedSourcePath, PdfDocumentOpenMode.Import);
        return sourceDocument.PageCount;
    }

    public static RangeValidationResult ValidateRangeText(string? rangeText, int? sourcePageCount = null)
    {
        if (string.IsNullOrWhiteSpace(rangeText))
        {
            return new RangeValidationResult(
                IsValid: false,
                HasInput: false,
                Message: "페이지 범위를 입력하면 바로 형식을 확인합니다.");
        }

        try
        {
            var segments = ParseSegments(rangeText);

            if (sourcePageCount.HasValue)
            {
                EnsurePagesExist(segments, sourcePageCount.Value);
            }

            var selectedPageCount = segments.Sum(segment => segment.Pages.Count);
            var message = sourcePageCount.HasValue
                ? $"{segments.Count}개 PDF 생성 예정, 총 {selectedPageCount}페이지 선택, 문서 전체 {sourcePageCount.Value}페이지 기준"
                : $"{segments.Count}개 PDF 생성 예정, 총 {selectedPageCount}페이지 선택";

            return new RangeValidationResult(
                IsValid: true,
                HasInput: true,
                Message: message);
        }
        catch (Exception ex)
        {
            return new RangeValidationResult(
                IsValid: false,
                HasInput: true,
                Message: ex.Message);
        }
    }

    private static string NormalizeSourcePath(string sourcePdfPath)
    {
        if (string.IsNullOrWhiteSpace(sourcePdfPath))
        {
            throw new ArgumentException("원본 PDF 경로를 입력하세요.");
        }

        var normalizedPath = NormalizePath(sourcePdfPath);
        if (string.IsNullOrWhiteSpace(normalizedPath))
        {
            throw new ArgumentException("원본 PDF 경로를 확인할 수 없습니다.");
        }

        if (!File.Exists(normalizedPath))
        {
            throw new FileNotFoundException("원본 PDF 파일을 찾을 수 없습니다.", normalizedPath);
        }

        if (!string.Equals(Path.GetExtension(normalizedPath), ".pdf", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("원본 파일은 PDF여야 합니다.");
        }

        return normalizedPath;
    }

    private static string ResolveOutputDirectory(string sourcePdfPath, string? outputFolderInput)
    {
        if (!string.IsNullOrWhiteSpace(outputFolderInput))
        {
            var normalizedOutput = NormalizePath(outputFolderInput);
            if (string.IsNullOrWhiteSpace(normalizedOutput))
            {
                throw new ArgumentException("저장 폴더 경로를 확인할 수 없습니다.");
            }

            if (string.Equals(Path.GetExtension(normalizedOutput), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                var containingDirectory = Path.GetDirectoryName(normalizedOutput);
                if (string.IsNullOrWhiteSpace(containingDirectory))
                {
                    throw new ArgumentException("저장 폴더 경로를 확인할 수 없습니다.");
                }

                return containingDirectory;
            }

            return normalizedOutput;
        }

        var normalizedSourcePath = NormalizeSourcePath(sourcePdfPath);
        var sourceDirectory = Path.GetDirectoryName(normalizedSourcePath)
            ?? throw new ArgumentException("원본 PDF의 폴더를 확인할 수 없습니다.");

        return Path.Combine(
            sourceDirectory,
            BuildDefaultDirectoryName(Path.GetFileNameWithoutExtension(normalizedSourcePath)));
    }

    private static string BuildDefaultDirectoryName(string sourceTitle)
    {
        var safeTitle = ShortenPathSegment(SanitizePathSegment(sourceTitle), MaxDefaultFolderTitleLength);
        return $"분리-{safeTitle}";
    }

    private static List<PageSegment> ParseSegments(string rangeText)
    {
        if (string.IsNullOrWhiteSpace(rangeText))
        {
            throw new ArgumentException("페이지 범위를 입력하세요.");
        }

        var parts = rangeText.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
        {
            throw new ArgumentException("쉼표로 구분된 페이지 범위를 입력하세요.");
        }

        return parts.Select(ParseSegment).ToList();
    }

    private static PageSegment ParseSegment(string text)
    {
        var token = text.Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            throw new ArgumentException("빈 페이지 범위가 포함되어 있습니다.");
        }

        if (!token.Contains('-'))
        {
            if (!int.TryParse(token, out var singlePage) || singlePage < 1)
            {
                throw new ArgumentException($"페이지 범위 '{token}' 형식이 올바르지 않습니다.");
            }

            return new PageSegment(token.Replace(" ", string.Empty), new[] { singlePage });
        }

        var rangeParts = token.Split('-', StringSplitOptions.TrimEntries);
        if (rangeParts.Length != 2
            || !int.TryParse(rangeParts[0], out var startPage)
            || !int.TryParse(rangeParts[1], out var endPage))
        {
            throw new ArgumentException($"페이지 범위 '{token}' 형식이 올바르지 않습니다.");
        }

        if (startPage < 1 || endPage < 1)
        {
            throw new ArgumentException("페이지 번호는 1 이상이어야 합니다.");
        }

        if (startPage > endPage)
        {
            throw new ArgumentException($"페이지 범위 '{token}'는 시작 페이지가 끝 페이지보다 클 수 없습니다.");
        }

        return new PageSegment(
            $"{startPage}-{endPage}",
            Enumerable.Range(startPage, endPage - startPage + 1).ToArray());
    }

    private static void EnsurePagesExist(IEnumerable<PageSegment> segments, int sourcePageCount)
    {
        foreach (var segment in segments)
        {
            if (segment.Pages.Any(page => page > sourcePageCount))
            {
                throw new ArgumentException(
                    $"페이지 범위 '{segment.DisplayLabel}'가 PDF 전체 페이지 수({sourcePageCount})를 벗어났습니다.");
            }
        }
    }

    private static string GetUniqueFilePath(string directoryPath, string fileStem, string extension)
    {
        var candidatePath = Path.Combine(directoryPath, fileStem + extension);
        var suffix = 2;

        while (File.Exists(candidatePath))
        {
            candidatePath = Path.Combine(directoryPath, $"{fileStem} ({suffix}){extension}");
            suffix++;
        }

        return candidatePath;
    }

    private static string SanitizePathSegment(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var sanitized = new string(
            value.Select(character => invalidCharacters.Contains(character) ? '_' : character).ToArray());

        sanitized = sanitized.Trim().TrimEnd('.', ' ');
        return string.IsNullOrWhiteSpace(sanitized) ? "PDF" : sanitized;
    }

    private static string ShortenPathSegment(string value, int maxLength)
    {
        if (value.Length <= maxLength)
        {
            return value;
        }

        return value[..maxLength].TrimEnd();
    }

    private static string NormalizePath(string pathText)
    {
        return Path.GetFullPath(Environment.ExpandEnvironmentVariables(pathText.Trim()));
    }

    internal sealed record SplitResult(string OutputDirectory, IReadOnlyList<string> CreatedFiles);

    internal sealed record RangeValidationResult(bool IsValid, bool HasInput, string Message);

    private sealed record PageSegment(string DisplayLabel, IReadOnlyList<int> Pages);
}
