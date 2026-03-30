# PDFExtractor

WPF 기반의 PDF 페이지 분리 도구입니다.  
원본 PDF를 선택하고 페이지 범위를 입력하면, 쉼표로 구분된 각 범위마다 별도의 PDF 파일을 생성합니다.

## 주요 기능

- 원본 PDF 선택
- 페이지 범위 입력 예시: `1-3, 5, 7-9`
- 쉼표 기준으로 여러 개의 PDF 파일 생성
- 저장 폴더를 비워두면 원본 PDF 폴더 아래에 `분리-원본제목` 폴더 자동 생성
- 원본 PDF 드래그 앤 드롭 지원
- 페이지 범위 실시간 검증
- 완료 후 출력 폴더 자동 열기
- 최근 사용 폴더 기억
- 생성된 파일 목록에서 바로 파일 열기 / 폴더 열기

## 사용 환경

- Windows
- .NET 8
- WPF

## 실행 방법

솔루션 루트에서 아래 명령으로 실행할 수 있습니다.

```powershell
dotnet run --project .\PDFExtractor\PDFExtractor.csproj
```

빌드는 아래 명령으로 확인할 수 있습니다.

```powershell
dotnet build .\PDFExtractor.sln
```

## 사용 예시

원본 PDF가 `report.pdf`이고 페이지 범위를 아래처럼 입력하면:

```text
1-3, 5, 7-9
```

다음과 같이 3개의 PDF가 생성됩니다.

- `report_1-3.pdf`
- `report_5.pdf`
- `report_7-9.pdf`

저장 폴더를 직접 지정하지 않으면 원본 PDF가 있는 위치에 아래 형태의 폴더가 자동 생성됩니다.

```text
분리-report
```

## 프로젝트 구조

```text
PDFExtractor.sln
PDFExtractor/
  App.xaml
  MainWindow.xaml
  MainWindow.xaml.cs
  PdfSplitService.cs
  AppSettingsService.cs
```

## 참고

- 출력 파일 이름이 이미 존재하면 자동으로 번호를 붙여 중복을 피합니다.
- `bin`, `obj`, `.vs` 폴더는 Git 추적 대상에서 제외되어 있습니다.
