# SharePoint File Downloader

A PowerShell script to download files from SharePoint links. Supports single downloads, command-line automation, and batch downloads via a text file.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Network access to the SharePoint instance

## Usage

### Interactive mode

Run without any parameters. The script will ask for the URL and save location.

```powershell
.\SharePoint-Downloader.ps1
```

### Single file via command line

```powershell
.\SharePoint-Downloader.ps1 -Url "https://yourcompany.sharepoint.com/:u:/s/..." -Destination "C:\Downloads"
```

### Batch download from a text file

```powershell
.\SharePoint-Downloader.ps1 -SourceFile "sources.txt" -Destination "C:\Downloads"
```

Downloads all URLs listed in `sources.txt` sequentially and shows a summary at the end.

## Parameters

| Parameter | Description |
|---|---|
| `-Url` | A single SharePoint download URL |
| `-Destination` | Target folder for downloaded files. Created automatically if it does not exist. Defaults to `%USERPROFILE%\Downloads` in interactive mode. |
| `-SourceFile` | Path to a text file containing one URL per line |

## sources.txt format

One URL per line. Empty lines and lines starting with `#` are skipped.

```text
# Production builds
https://yourcompany.sharepoint.com/:u:/s/public/file1.zip?download=1
https://yourcompany.sharepoint.com/:u:/s/public/file2.zip?download=1

# Test builds
https://yourcompany.sharepoint.com/:u:/s/public/test.zip?download=1
```

## One-liner alternative

If you just need to download a single file without any extras, PowerShell can do it in one line:

```powershell
Invoke-WebRequest -Uri "https://yourcompany.sharepoint.com/:u:/s/public/yourfile.zip?download=1" -OutFile "C:\temp\yourfile.zip"
```

Note that you need to know the filename in advance — there is no automatic detection.

## How it works

The script detects the original filename from the server's `Content-Disposition` response header. It supports all three standard formats (`filename*=UTF-8''...`, `filename="..."`, and `filename=...`). If no filename can be detected, the file is saved as `download.bin`.

## License

MIT
